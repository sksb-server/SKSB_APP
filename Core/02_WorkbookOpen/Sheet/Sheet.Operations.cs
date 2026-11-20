namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;

public static class SheetOperations
{
    public static bool IsTemplateEditMode { get; set; } = false;
    public const string LookupsSheetName = "Lookups";
    public const string DataSheetName = "Data";

    public static string ResolveSheetName(string[] candidates, string defaultFallback)
    {
        if (candidates == null || candidates.Length == 0)
            return defaultFallback;

        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                foreach (object item in sheetArray)
                {
                    string rawName = item?.ToString() ?? string.Empty;
                    int bracketIdx = rawName.LastIndexOf(']');
                    string cleanName = bracketIdx >= 0 ? rawName.Substring(bracketIdx + 1) : rawName;
                    foreach (var candidate in candidates)
                    {
                        if (string.Equals(cleanName, candidate, StringComparison.OrdinalIgnoreCase))
                        {
                            return cleanName;
                        }
                    }
                }
            }
        }
        catch { }

        return defaultFallback;
    }

    public static bool IsMasterTemplateFile() => IsMasterTemplateFile(null);

    public static bool IsMasterTemplateFile(dynamic? wb)
    {
        try
        {
            if (wb == null)
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                wb = xlApp.ActiveWorkbook;
            }
            if (wb == null) return false;

            string fileName = wb.Name?.ToString() ?? string.Empty;
            return fileName.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase) ||
                   fileName.EndsWith(".xltm", StringComparison.OrdinalIgnoreCase);
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Universal workbook guard: true when <paramref name="wb"/> (or the active workbook) contains one of
    /// the alias sheets carrying the module's signature ListObject table.
    /// </summary>
    public static bool IsModuleWorkbook(dynamic? wb, string[] sheetAliases, string signatureTableName)
    {
        if (sheetAliases == null || sheetAliases.Length == 0 || string.IsNullOrWhiteSpace(signatureTableName))
            return false;

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? targetWb = wb ?? xlApp.ActiveWorkbook;
            if (targetWb == null) return false;

            foreach (string alias in sheetAliases)
            {
                dynamic? ws = null;
                try { ws = targetWb.Sheets[alias]; } catch { }
                if (ws == null) continue;

                dynamic? table = null;
                try { table = ws.ListObjects[signatureTableName]; } catch { }
                if (table != null) return true;
            }
        }
        catch { }

        return false;
    }

    public static bool IsModuleWorkbook(dynamic? wb, ISaveContract contract) =>
        contract != null && IsModuleWorkbook(wb, contract.MasterSheetAliases, contract.MasterTableName);

    public static void ShowStatusConfirmation(string message)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp != null)
            {
                bool offlineTag = SKSB_App.Core.Integration.Database.ServerEnvironment.IsOffline
                    && !message.Contains("[OFFLINE MODE]", StringComparison.Ordinal);
                xlApp.StatusBar = $"[SKSB System] {(offlineTag ? "[OFFLINE MODE] " : string.Empty)}{message}";
            }
        }
        catch { }
    }

    public static void SetAuxiliarySheetsVisibility(bool visible, bool synchronous = false)
    {
        void Execute()
        {
            using var _ = AppLogger.Scope($"SetAuxiliarySheetsVisibility(visible={visible})");
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            bool origUpdating = true;
            bool origEvents = true;
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

            string activeSheetBefore = xlApp.ActiveSheet?.Name?.ToString() ?? string.Empty;

            try
            {
                SetSingleSheetVisibility(LookupsSheetName, visible);
                SetSingleSheetVisibility(DataSheetName, visible);
                MoveAuxiliarySheetsToEnd();

                // Restore original active sheet if Excel shifted focus
                if (!string.IsNullOrEmpty(activeSheetBefore))
                {
                    try { wb.Sheets[activeSheetBefore].Activate(); } catch { }
                }
            }
            finally
            {
                try { xlApp.EnableEvents = origEvents; } catch { }
                try { xlApp.ScreenUpdating = origUpdating; } catch { }
            }
        }

        if (synchronous || SyncHelper.IsSynchronousScope) Execute();
        else ExcelAsyncUtil.QueueAsMacro(Execute);
    }

    private static void SetSingleSheetVisibility(string sheetName, bool visible)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic? ws = null;
            try { ws = wb.Sheets[sheetName]; } catch { }
            if (ws == null) return;

            int targetVis = visible ? -1 /* xlSheetVisible */ : 0 /* xlSheetHidden */;
            if ((int)ws.Visible == targetVis)
            {
                return; // Already in desired state, do not trigger UI repaint
            }

            string currentActive = wb.ActiveSheet?.Name?.ToString() ?? string.Empty;

            if (!visible && string.Equals(currentActive, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                foreach (dynamic candidate in wb.Sheets)
                {
                    string cName = candidate.Name?.ToString() ?? string.Empty;
                    if (!cName.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase) &&
                        !cName.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        candidate.Activate();
                        break;
                    }
                }
            }

            ws.Visible = targetVis;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"SetSingleSheetVisibility('{sheetName}', visible={visible}) failed", ex);
        }
    }

    public static void MoveAuxiliarySheetsToEnd()
    {
        MoveSingleSheetToEnd(LookupsSheetName);
        MoveSingleSheetToEnd(DataSheetName);
    }

    private static void MoveSingleSheetToEnd(string targetSheetName)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic sheets = wb.Sheets;
            dynamic? targetSheet = null;
            try { targetSheet = sheets[targetSheetName]; } catch { }
            if (targetSheet == null) return;

            int totalSheets = (int)sheets.Count;
            int currentIdx = (int)targetSheet.Index;

            // Lookups and Data together occupy the final two slots (totalSheets - 1 and totalSheets).
            // If the sheet is already in the final 2 positions, do NOT move it!
            if (currentIdx >= totalSheets - 1)
            {
                return;
            }

            dynamic? activeSheetBeforeMove = wb.ActiveSheet;

            targetSheet.Move(After: sheets[totalSheets]);
            AppLogger.Info($"[MoveSingleSheetToEnd] Moved '{targetSheetName}' from {currentIdx} to end ({totalSheets}).");

            if (activeSheetBeforeMove != null)
            {
                try { activeSheetBeforeMove.Activate(); } catch { }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"MoveSingleSheetToEnd('{targetSheetName}') failed", ex);
        }
    }

    /// <summary>
    /// Safely purges all operational data rows across all sheets, preserving:
    /// 1. Headers on any row (Row 1, Row 2, Row 3, etc.)
    /// 2. Exactly 1 clean archetype row per table (keeping calculated formulas and styles intact)
    /// 3. Non-table sheet parameter anchor cells (e.g., Dates, Range Filters in rows 1-2)
    /// 4. System sheets ('Lookups' and 'Data')
    /// </summary>
    public static void PurgeAllSheetsData()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return;

        bool origEvents = true;
        bool origUpdating = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

        try
        {
            foreach (dynamic ws in wb.Sheets)
            {
                string sheetName = ws.Name?.ToString() ?? string.Empty;

                // Never purge Lookups or Data sheets
                if (sheetName.Equals("Lookups", StringComparison.OrdinalIgnoreCase) ||
                    sheetName.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                bool isBoard = sheetName.Equals("Daily Schedule", StringComparison.OrdinalIgnoreCase) ||
                               sheetName.Equals("Daily_Schedule", StringComparison.OrdinalIgnoreCase) ||
                               sheetName.Equals("Schedule", StringComparison.OrdinalIgnoreCase);
                if (isBoard)
                {
                    try { PurgeScheduleBoard(ws); }
                    catch (Exception ex) { AppLogger.Error($"PurgeScheduleBoard failed for '{sheetName}'", ex); }
                }

                int tableCount = 0;
                try { tableCount = (int)ws.ListObjects.Count; } catch { }

                if (!isBoard && tableCount > 0)
                {
                    int minHeaderRow = int.MaxValue;
                    for (int t = 1; t <= tableCount; t++)
                    {
                        try { minHeaderRow = Math.Min(minHeaderRow, (int)ws.ListObjects[t].HeaderRowRange.Row); } catch { }
                    }
                    if (minHeaderRow != int.MaxValue && minHeaderRow >= 3)
                    {
                        try { PurgeTopBarRow(ws, minHeaderRow - 1); }
                        catch (Exception ex) { AppLogger.Error($"PurgeTopBarRow failed for '{sheetName}'", ex); }
                    }
                }

                for (int t = 1; t <= tableCount; t++)
                {
                    dynamic table = ws.ListObjects[t];
                    PurgeSingleTable(ws, table);
                }            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("PurgeAllSheetsData failed", ex);
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    private static bool IsFormulaCell(dynamic cell)
    {
        string formula = string.Empty;
        try { formula = cell.Formula2?.ToString() ?? string.Empty; } catch { }
        if (string.IsNullOrWhiteSpace(formula))
        {
            try { formula = cell.Formula?.ToString() ?? string.Empty; } catch { }
        }
        return formula.StartsWith("=");
    }

    private static int GetLastUsedColumn(dynamic ws)
    {
        try
        {
            dynamic used = ws.UsedRange;
            return (int)used.Column + (int)used.Columns.Count - 1;
        }
        catch { return 1; }
    }

    private static void PurgeTopBarRow(dynamic ws, int topBarMaxRow)
    {
        int lastCol = Math.Min(GetLastUsedColumn(ws), 50);
        for (int r = 2; r <= topBarMaxRow; r++)
        {
            for (int c = 1; c <= lastCol; c++)
            {
                try
                {
                    dynamic cell = ws.Cells[r, c];
                    object? v = cell.Value2;
                    if (v == null || string.IsNullOrWhiteSpace(v.ToString())) continue;
                    if (IsFormulaCell(cell)) continue;

                    bool hasLabel = false;
                    try
                    {
                        object? label = ws.Cells[1, c].Value2;
                        hasLabel = label != null && !string.IsNullOrWhiteSpace(label.ToString());
                    }
                    catch { }

                    bool hasValidation = false;
                    try { var vt = cell.Validation.Type; hasValidation = true; } catch { }

                    bool isDropDown = false;
                    try { isDropDown = string.Equals(cell.Style.Name?.ToString(), "DropDown", StringComparison.OrdinalIgnoreCase); } catch { }

                    if (hasLabel || hasValidation || isDropDown)
                        cell.Value2 = null;
                }
                catch { }
            }
        }
    }

    private static void PurgeScheduleBoard(dynamic ws)
    {
        PurgeTopBarRow(ws, 2);

        // Slot 0 (B:D) for Day Work (4..18) and Night Work (20..34): keep formatting, clear values/comments
        int[][] blocks = { new[] { 4, 18 }, new[] { 20, 34 } };
        foreach (var blk in blocks)
        {
            dynamic slotRange = ws.Range[ws.Cells[blk[0], 2], ws.Cells[blk[1], 4]];
            try { slotRange.ClearComments(); } catch { }
            for (int r = blk[0]; r <= blk[1]; r++)
            {
                for (int c = 2; c <= 4; c++)
                {
                    try
                    {
                        dynamic cell = ws.Cells[r, c];
                        if (!IsFormulaCell(cell)) cell.Value2 = null;
                    }
                    catch { }
                }
            }
        }

        // Extra dynamic slots (Column E onwards)
        int lastCol = GetLastUsedColumn(ws);
        if (lastCol >= 5)
        {
            dynamic extra = ws.Range[ws.Cells[4, 5], ws.Cells[34, lastCol]];
            try { extra.ClearComments(); } catch { }
            try { extra.Clear(); } catch { }
        }
    }
    private static void PurgeSingleTable(dynamic ws, dynamic table)
    {
        if (table == null) return;

        bool hadTotals = false;
        try { hadTotals = (bool)table.ShowTotals; } catch { }

        try { if (table.AutoFilter != null && (bool)table.AutoFilter.FilterMode) table.AutoFilter.ShowAllData(); } catch { }

        try
        {
            // 1. Dynamically read the table's actual coordinates from Excel COM
            dynamic headerRange = table.HeaderRowRange;
            int headerRow = (int)headerRange.Row;       // e.g. Row 1, 2, or 3!
            int startCol = (int)headerRange.Column;
            int colCount = (int)table.ListColumns.Count;
            int archetypeRow = headerRow + 1;           // First data row directly beneath header
            int endCol = startCol + colCount - 1;

            // 2. Temporarily disable total row (restored in finally)
            if (hadTotals) { try { table.ShowTotals = false; } catch { } }

            // 3. Clear data from all rows BELOW the archetype row (if any exist)
            dynamic? body = table.DataBodyRange;
            int currentBodyRows = body != null ? (int)body.Rows.Count : 0;

            if (currentBodyRows > 1)
            {
                int clearStartRow = archetypeRow + 1;
                int clearEndRow = headerRow + currentBodyRows;

                // Resize table back down to 1 row first
                dynamic singleRowRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[archetypeRow, endCol]];
                table.Resize(singleRowRange);

                // Clear any lingering cells outside the table
                dynamic staleRange = ws.Range[ws.Cells[clearStartRow, startCol], ws.Cells[clearEndRow, endCol]];
                try { staleRange.Clear(); } catch { }
            }
            else if (currentBodyRows == 0)
            {
                // If the table was shrunk to 0 body rows, re-expand to exactly 1 archetype row
                dynamic singleRowRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[archetypeRow, endCol]];
                table.Resize(singleRowRange);
            }

            // 4. Clean the remaining single archetype row (archetypeRow)
            // CRITICAL: Preserve cell formulas, clear only user data values
            for (int c = 1; c <= colCount; c++)
            {
                int absCol = startCol + c - 1;
                dynamic cell = ws.Cells[archetypeRow, absCol];

                string formula = cell.Formula2?.ToString() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(formula))
                    formula = cell.Formula?.ToString() ?? string.Empty;

                // If the cell contains a formula (e.g. =XLOOKUP, =IF, =AND), DO NOT clear it!
                if (!formula.StartsWith("="))
                {
                    cell.Value2 = null;
                }
            }
        }
        catch (Exception ex)
        {
            string tName = table?.Name?.ToString() ?? "Table";
            AppLogger.Error($"PurgeSingleTable failed for '{tName}'", ex);
        }
        finally
        {
            if (hadTotals) { try { table.ShowTotals = true; } catch { } }
        }
    }
}