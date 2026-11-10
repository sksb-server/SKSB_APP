namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;

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

    public static bool IsMasterTemplateFile()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp?.ActiveWorkbook;
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

    public static void ShowStatusConfirmation(string message)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp != null)
            {
                xlApp.StatusBar = $"[SKSB System] {message}";
            }
        }
        catch { }
    }

    public static void SetAuxiliarySheetsVisibility(bool visible)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            using var _ = AppLogger.Scope($"SetAuxiliarySheetsVisibility(visible={visible})");
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp?.ActiveWorkbook;
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
        });
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
        dynamic wb = xlApp?.ActiveWorkbook;
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

                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
                    int tableCount = 0;
                    try { tableCount = (int)ws.ListObjects.Count; } catch { }

                    for (int t = 1; t <= tableCount; t++)
                    {
                        dynamic table = ws.ListObjects[t];
                        PurgeSingleTable(ws, table);
                    }
                }));
            }
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

    private static void PurgeSingleTable(dynamic ws, dynamic table)
    {
        if (table == null) return;

        bool hadTotals = false;
        try { hadTotals = (bool)table.ShowTotals; } catch { }

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