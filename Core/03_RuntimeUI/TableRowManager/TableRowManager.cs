namespace SKSB_App.Core.RuntimeUI.TableRowManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;

public enum TableInsertMode
{
    InsertAtBottom,
    InsertAtTop
}

public enum TableNature
{
    Absolute,
    Additive
}

public static partial class TableRowManager
{
    private static bool _isManagingRows = false;
    private const string InsertAtTopDefinedName = "__SKSB_INSERT_AT_TOP";
    private static bool _inMemoryFallback = false;

    private static readonly ConcurrentDictionary<string, TableNature> TableNatures =
        new(StringComparer.OrdinalIgnoreCase);

    static TableRowManager()
    {
        // =========================================================================
        // PRE-CONFIGURED TABLE NATURE CLASSIFICATIONS
        // =========================================================================
        // Absolute Tables (Master Entities / Registries - Full Set Reconciliation)
        TableNatures["AllVehiclesTable"] = TableNature.Absolute;
        TableNatures["WorkshopsTable"] = TableNature.Absolute;
        TableNatures["FuelCardsTable"] = TableNature.Absolute;
        TableNatures["FuelCompanyNamesTable"] = TableNature.Absolute;
        TableNatures["InsuranceCompaniesTable"] = TableNature.Absolute;
        TableNatures["StaffMasterTable"] = TableNature.Absolute;
        TableNatures["AllStaffsTable"] = TableNature.Absolute;
        TableNatures["AllWorkersTable"] = TableNature.Absolute;
        TableNatures["SkillLicenseTypesTable"] = TableNature.Absolute;
        TableNatures["ProjectsTable"] = TableNature.Absolute;

        // Additive Tables (Append-Only Event Logs / Transaction Ledgers)
        TableNatures["OdometersTable"] = TableNature.Additive;
        TableNatures["MaintenancesTable"] = TableNature.Additive;
        TableNatures["RoadTaxesTable"] = TableNature.Additive;
        TableNatures["InspectionsTable"] = TableNature.Additive;
        TableNatures["PermitsTable"] = TableNature.Additive;
        TableNatures["InsurancesTable"] = TableNature.Additive;
        TableNatures["InsuranceClaimsTable"] = TableNature.Additive;
        TableNatures["StaffLicensesTable"] = TableNature.Additive;
        TableNatures["StaffLeavesTable"] = TableNature.Additive;
    }

    public static TableNature GetTableNature(dynamic table)
    {
        try
        {
            string tableName = table?.Name?.ToString() ?? string.Empty;
            if (!string.IsNullOrEmpty(tableName))
            {
                if (TableNatures.TryGetValue(tableName, out var nature))
                    return nature;

                return ResolveDefaultTableNature(tableName);
            }
        }
        catch { }

        return TableNature.Absolute;
    }

    private static TableNature ResolveDefaultTableNature(string tableName)
    {
        string norm = tableName.Replace("_", "").Replace(" ", "").ToLowerInvariant();

        // Additive log heuristics
        if (norm.Contains("odometer") ||
            norm.Contains("maintenance") ||
            norm.Contains("inspection") ||
            norm.Contains("roadtax") ||
            norm.Contains("permit") ||
            (norm.Contains("insurance") && !norm.Contains("company") && !norm.Contains("companies")) ||
            norm.Contains("claim") ||
            norm.Contains("log") ||
            norm.Contains("transaction") ||
            norm.Contains("attendance") ||
            norm.Contains("history") ||
            norm.Contains("schedule") ||
            norm.Contains("timesheet"))
        {
            return TableNature.Additive;
        }

        return TableNature.Absolute;
    }

    public static bool IsInsertAtTopEnabled
    {
        get => GetWorkbookInsertAtTop();
        set => SetWorkbookInsertAtTop(value);
    }

    public static TableInsertMode GetTableInsertMode()
    {
        return IsInsertAtTopEnabled ? TableInsertMode.InsertAtTop : TableInsertMode.InsertAtBottom;
    }

    /// <summary>
    /// Inserts a new formatted row at the top or bottom of the active table
    /// (per Insert at Top toggle) and selects it. Accessible via Ctrl+Shift+I.
    /// </summary>
    [ExcelCommand(Name = "InsertTableRow", ShortCut = "^+I")]
    public static void InsertRowCommand()
    {
        try
        {
            if (_isManagingRows) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic selection = xlApp.Selection;

            dynamic? table = null;
            try { table = selection?.ListObject; } catch { }
            if (table == null)
            {
                try { table = selection?.Cells[1, 1].ListObject; } catch { }
            }
            if (table == null)
            {
                try
                {
                    dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                    if (ws != null && (int)ws.ListObjects.Count == 1) table = ws.ListObjects[1];
                }
                catch { }
            }

            if (table == null)
            {
                MessageBox.Show("Please select a cell inside a table to insert a row.",
                    "No Table Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            _isManagingRows = true;

            int oldCalc = -4105;
            bool oldEvents = true;
            bool oldScreen = true;
            try { oldCalc = (int)xlApp.Calculation; } catch { }
            try { oldEvents = (bool)xlApp.EnableEvents; } catch { }
            try { oldScreen = (bool)xlApp.ScreenUpdating; } catch { }

            bool atTop = IsInsertAtTopEnabled;
            string tableName = "Table";
            dynamic? newRowFirstCell = null;

            try
            {
                try { xlApp.Calculation = -4135; } catch { }
                try { xlApp.EnableEvents = false; } catch { }
                try { xlApp.ScreenUpdating = false; } catch { }

                tableName = table.Name?.ToString() ?? "Table";
                dynamic worksheet = table.Parent;
                int startCol = (int)table.HeaderRowRange.Column;
                int colCount = (int)table.ListColumns.Count;
                int headerRow = (int)table.HeaderRowRange.Row;

                SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, atTop);

                int rowCount = (int)table.ListRows.Count;
                if (rowCount > 0)
                {
                    newRowFirstCell = (atTop ? table.ListRows[1].Range : table.ListRows[rowCount].Range).Cells[1, 1];
                }
            }
            finally
            {
                try { xlApp.Calculation = oldCalc; } catch { }
                try { xlApp.EnableEvents = oldEvents; } catch { }
                try { xlApp.ScreenUpdating = oldScreen; } catch { }
                _isManagingRows = false;
            }

            if (newRowFirstCell != null)
            {
                try { newRowFirstCell.Select(); } catch { }
            }

            SheetOperations.ShowStatusConfirmation(atTop
                ? $"Inserted new row at top of '{tableName}'."
                : $"Inserted new row at bottom of '{tableName}'.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("InsertRowCommand failed", ex);
        }
    }

    /// <summary>
    /// Explicit row deletion command. Only executes on Absolute tables; blocks on Additive tables.
    /// Accessible via Manager Ribbon and Ctrl+Shift+D.
    /// </summary>
    [ExcelCommand(Name = "DeleteSelectedTableRow", ShortCut = "^+D")]
    public static void DeleteSelectedRowCommand()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic selection = xlApp.Selection;
            if (selection == null) return;

            dynamic? table = null;
            try { table = selection.ListObject; } catch { }
            if (table == null)
            {
                try { table = selection.Cells[1, 1].ListObject; } catch { }
            }

            if (table == null)
            {
                MessageBox.Show("Please select a cell inside a table to delete a row.",
                    "No Table Selected", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string tableName = table.Name?.ToString() ?? "Table";
            var nature = GetTableNature(table);

            // GUARD: Block row deletion if table is Additive
            if (nature == TableNature.Additive)
            {
                MessageBox.Show(
                    $"Row deletion is disabled for '{tableName}'.\n\n" +
                    $"This table is categorized as an Additive / Transaction Log table.\n" +
                    $"Individual rows cannot be deleted directly to safeguard historical operational records.",
                    "Deletion Blocked (Additive Table)",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                SheetOperations.ShowStatusConfirmation($"Deletion blocked: '{tableName}' is an Additive log table.");
                return;
            }

            dynamic worksheet = table.Parent;
            dynamic dataBody = table.DataBodyRange;
            if (dataBody == null) return;

            int bodyStartRow = (int)dataBody.Row;
            int totalBodyRows = (int)dataBody.Rows.Count;

            var rowIndicesToDelete = new HashSet<int>();
            foreach (dynamic cell in selection.Cells)
            {
                int r = (int)cell.Row;
                int listRowIdx = r - bodyStartRow + 1;
                if (listRowIdx >= 1 && listRowIdx <= totalBodyRows)
                {
                    rowIndicesToDelete.Add(listRowIdx);
                }
            }

            if (rowIndicesToDelete.Count == 0) return;

            var confirmResult = MessageBox.Show(
                $"Are you sure you want to delete {rowIndicesToDelete.Count} selected row(s) from '{tableName}'?\n\n" +
                "This change will be permanently committed to the database upon save.",
                "Confirm Delete Row",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );
            if (confirmResult != DialogResult.Yes) return;

            foreach (int idx in rowIndicesToDelete.OrderByDescending(x => x))
            {
                if ((int)table.ListRows.Count > 1)
                {
                    table.ListRows[idx].Delete();
                }
                else
                {
                    // Keep 1 clean archetype row if deleting the final row
                    dynamic lastRowRange = table.ListRows[1].Range;
                    int colCount = (int)table.ListColumns.Count;
                    int headerRow = (int)table.HeaderRowRange.Row;
                    int startCol = (int)table.HeaderRowRange.Column;

                    for (int c = 1; c <= colCount; c++)
                    {
                        dynamic archCell = worksheet.Cells[headerRow + 1, startCol + c - 1];
                        string archFormula = archCell.Formula?.ToString() ?? string.Empty;
                        if (!archFormula.StartsWith("="))
                        {
                            lastRowRange.Cells[1, c].Value2 = null;
                        }
                    }
                }
            }

            SheetOperations.ShowStatusConfirmation($"Deleted {rowIndicesToDelete.Count} row(s) from '{tableName}'.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("DeleteSelectedRowCommand failed", ex);
        }
    }

    public static void HandleSheetChange(dynamic target)
    {
        if (target == null || SheetOperations.IsTemplateEditMode || _isManagingRows) return;

        dynamic? table = null;
        try { table = target!.ListObject; } catch { }
        if (table == null)
        {
            try { table = target!.Cells[1, 1].ListObject; } catch { }
        }
        if (table == null) return;

        dynamic dataBody = table.DataBodyRange;

        // Ignore edits touching the Total Row or falling outside the DataBodyRange
        try
        {
            dynamic xlAppForBounds = ExcelDnaUtil.Application;
            if ((bool)table.ShowTotals && table.TotalsRowRange != null &&
                xlAppForBounds.Intersect(target, table.TotalsRowRange) != null)
            {
                return;
            }
            if (dataBody != null && xlAppForBounds.Intersect(target, dataBody) == null)
            {
                return;
            }
        }
        catch { }

        int bodyStartRow = dataBody != null ? (int)dataBody.Row : (int)table.HeaderRowRange.Row + 1;
        int rowCount = dataBody != null ? (int)dataBody.Rows.Count : 0;
        int lastDataRow = bodyStartRow + Math.Max(1, rowCount) - 1;

        TableInsertMode insertMode = GetTableInsertMode();

        int targetTopRow = (int)target!.Row;
        int targetRowCount = (int)target.Rows.Count;
        int targetBottomRow = targetTopRow + targetRowCount - 1;

        bool touchesBottom = (targetBottomRow >= lastDataRow);
        bool touchesTop = (targetTopRow <= bodyStartRow);
        bool hasEmptiedCells = false;

        try
        {
            foreach (dynamic cell in target.Cells)
            {
                if (string.IsNullOrEmpty(cell.Value2?.ToString()?.Trim()))
                {
                    hasEmptiedCells = true;
                    break;
                }
            }
        }
        catch { }

        if (!touchesBottom && !touchesTop && !hasEmptiedCells) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (_isManagingRows) return;
            _isManagingRows = true;

            dynamic xlApp = ExcelDnaUtil.Application;
            bool origEnableEvents = true;
            try { origEnableEvents = (bool)xlApp.EnableEvents; } catch { }

            var auditTrail = new List<string>();

            try
            {
                xlApp.EnableEvents = false;
                dynamic worksheet = table.Parent;

                dynamic currentBody = table.DataBodyRange;
                int startCol = (int)table.HeaderRowRange.Column;
                int colCount = (int)table.ListColumns.Count;
                int headerRow = (int)table.HeaderRowRange.Row;

                if (currentBody == null)
                {
                    SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, false, auditTrail);
                    return;
                }

                int curRowCount = (int)currentBody.Rows.Count;
                int curBodyStartRow = (int)currentBody.Row;

                if (insertMode == TableInsertMode.InsertAtBottom)
                {
                    int lastRowIdx = curBodyStartRow + curRowCount - 1;
                    bool lastRowHasInput = CheckRowHasUserInput(worksheet, table, lastRowIdx, startCol, colCount, headerRow);
                    if (lastRowHasInput && touchesBottom)
                    {
                        SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, false, auditTrail);
                    }
                }
                else if (insertMode == TableInsertMode.InsertAtTop)
                {
                    int topRowIdx = curBodyStartRow;
                    bool topRowHasInput = CheckRowHasUserInput(worksheet, table, topRowIdx, startCol, colCount, headerRow);
                    if (topRowHasInput && touchesTop)
                    {
                        SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, true, auditTrail);
                    }
                }

                // Auto-delete / compact empty buffer rows
                CleanTableEmptyRows(table, insertMode, auditTrail);
            }
            catch (Exception ex)
            {
                AppLogger.Error("HandleSheetChange error", ex);
            }
            finally
            {
                try { xlApp.EnableEvents = origEnableEvents; } catch { }
                _isManagingRows = false;
            }
        });
    }

    public static int CleanActiveSheetTables(dynamic worksheet)
    {
        if (worksheet == null) return 0;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEnableEvents = true;
        try { origEnableEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        int tablesCleaned = 0;
        try
        {
            foreach (dynamic table in worksheet.ListObjects)
            {
                TableInsertMode mode = GetTableInsertMode();
                CleanTableEmptyRows(table, mode, null);
                tablesCleaned++;
            }
        }
        finally
        {
            try { xlApp.EnableEvents = origEnableEvents; } catch { }
        }

        return tablesCleaned;
    }

    public static void CleanTableEmptyRows(dynamic table, TableInsertMode insertMode, List<string>? auditLog = null)
    {
        if (table == null) return;

        dynamic worksheet = table.Parent;
        dynamic currentBody = table.DataBodyRange;
        if (currentBody == null) return;

        int curRowCount = (int)currentBody.Rows.Count;
        int startCol = (int)table.HeaderRowRange.Column;
        int colCount = (int)table.ListColumns.Count;
        int headerRow = (int)table.HeaderRowRange.Row;
        int curBodyStartRow = (int)currentBody.Row;

        var emptyRowIndices = new List<int>();

        for (int r = 1; r <= curRowCount; r++)
        {
            int absRow = curBodyStartRow + r - 1;
            bool hasInput = CheckRowHasUserInput(worksheet, table, absRow, startCol, colCount, headerRow);
            if (!hasInput)
            {
                emptyRowIndices.Add(r);
            }
        }

        if (emptyRowIndices.Count > 1)
        {
            int toDeleteCount = emptyRowIndices.Count - 1;
            List<int> rowsToDelete = insertMode == TableInsertMode.InsertAtBottom
                ? emptyRowIndices.Take(toDeleteCount).OrderByDescending(x => x).ToList()
                : emptyRowIndices.Skip(1).OrderByDescending(x => x).ToList();

            foreach (int idx in rowsToDelete)
            {
                DeleteTableRowSafely(table, idx);
                auditLog?.Add($"[DELETE] Trimmed redundant empty row at index {idx}.");
            }
        }
        else if (emptyRowIndices.Count == 1 && curRowCount > 1)
        {
            int emptyIdx = emptyRowIndices[0];
            if (insertMode == TableInsertMode.InsertAtBottom && emptyIdx != curRowCount)
            {
                DeleteTableRowSafely(table, emptyIdx);
                SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, false, auditLog);
            }
            else if (insertMode == TableInsertMode.InsertAtTop && emptyIdx != 1)
            {
                DeleteTableRowSafely(table, emptyIdx);
                SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, true, auditLog);
            }
        }
        else if (emptyRowIndices.Count == 0)
        {
            SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, insertMode == TableInsertMode.InsertAtTop, auditLog);
        }
    }

    public static bool CheckRowHasUserInput(dynamic worksheet, dynamic table, int rowIndex, int startCol, int colCount, int headerRow)
    {
        dynamic driverCell = worksheet.Cells[rowIndex, startCol];
        string driverVal = driverCell.Value2?.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(driverVal))
        {
            int nonDefaultCount = 0;
            int archetypeRow = headerRow + 1;

            for (int c = 2; c <= colCount; c++)
            {
                int absCol = startCol + c - 1;
                dynamic cell = worksheet.Cells[rowIndex, absCol];
                dynamic archCell = worksheet.Cells[archetypeRow, absCol];

                string archFormula = archCell.Formula?.ToString() ?? string.Empty;
                if (archFormula.StartsWith("=")) continue;

                string cellVal = cell.Value2?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(cellVal)) continue;

                string headerText = worksheet.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                var rule = RuleRegistry.GetRule(table.Name?.ToString() ?? string.Empty, headerText);
                if (rule?.IsStrictlyCalculated == true) continue;
                if (RuleRegistry.IsZeroOrEmptyValue(cell.Value2, rule)) continue;

                if (rule?.DefaultRule?.Kind == DefaultKind.Literal &&
                    string.Equals(cellVal, rule!.DefaultRule!.LiteralValue, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                nonDefaultCount++;
            }

            if (nonDefaultCount == 0) return false;
        }

        return true;
    }

    public static void SpawnRowSafely(
        dynamic worksheet,
        dynamic table,
        int startCol,
        int colCount,
        int headerRow,
        bool atTop,
        List<string>? auditLog = null)
    {
        string tableName = table.Name?.ToString() ?? string.Empty;
        int existingRows = 0;
        try { existingRows = (int)table.ListRows.Count; } catch { }
        bool insertAtTop = atTop && existingRows > 0;

        dynamic newRow = insertAtTop ? table.ListRows.Add(1) : table.ListRows.Add();
        dynamic newRowRange = newRow.Range;

        int sourceArchetypeRow = insertAtTop ? headerRow + 2 : headerRow + 1;

        // Inside TableRowManager.SpawnRowSafely:

        for (int c = 1; c <= colCount; c++)
        {
            int absCol = startCol + c - 1;
            dynamic archetypeCell = worksheet.Cells[sourceArchetypeRow, absCol];
            dynamic newCell = newRowRange.Cells[1, c];

            string headerText = worksheet.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
            var rule = RuleRegistry.GetRule(tableName, headerText);

            if (rule != null && rule.HasDropdown)
            {
                try { newCell.Validation.Delete(); } catch { }
            }

            // Explicitly mirror NumberFormat
            try
            {
                string archFormat = archetypeCell.NumberFormat?.ToString() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(archFormat))
                {
                    newCell.NumberFormat = archFormat;
                }
                newCell.NumberFormatLocal = archetypeCell.NumberFormatLocal;
            }
            catch { }

            try
            {
                newCell.HorizontalAlignment = archetypeCell.HorizontalAlignment;
                newCell.VerticalAlignment = archetypeCell.VerticalAlignment;
            }
            catch { }

            var cfg = Metadata.ReadConfig(archetypeCell);
            if (!cfg.IsEmpty)
            {
                Metadata.WriteConfig(newCell, cfg);
            }

            string archFormula = archetypeCell.Formula?.ToString() ?? string.Empty;
            if (!archFormula.StartsWith("=") || rule?.IsStrictlyCalculated == true)
            {
                newCell.Value2 = null;
                RuleRegistry.ApplyCalculatedColumnLock(newCell, rule);
            }

            // Declarative [ColumnFormat]: units + dash for zero on the new row (text template cells stay text)
            if (RuleRegistry.ShouldApplyFormat(rule) && rule!.EffectiveFormatKind != ColumnFormatKind.OdometerDynamic &&
                !string.Equals(archetypeCell.NumberFormat?.ToString(), "@", StringComparison.Ordinal))
            {
                try { newCell.NumberFormat = rule!.ExcelNumberFormat; } catch { }
            }
        }

        auditLog?.Add(atTop
            ? "[INSERT] Spawned new entry row at TOP."
            : "[INSERT] Spawned new entry row at BOTTOM.");
    }

    public static void DeleteTableRowSafely(dynamic table, int listRowIndex)
    {
        table.ListRows[listRowIndex].Delete();
    }

    private static bool GetWorkbookInsertAtTop()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb != null)
            {
                dynamic? nameObj = null;
                try { nameObj = wb.Names.Item(InsertAtTopDefinedName); } catch { }
                if (nameObj != null)
                {
                    string refersTo = nameObj.RefersTo?.ToString() ?? string.Empty;
                    return refersTo.IndexOf("TRUE", StringComparison.OrdinalIgnoreCase) >= 0;
                }
            }
        }
        catch { }
        return _inMemoryFallback;
    }

    private static void SetWorkbookInsertAtTop(bool enabled)
    {
        _inMemoryFallback = enabled;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb != null)
            {
                try { wb.Names.Item(InsertAtTopDefinedName).Delete(); } catch { }
                wb.Names.Add(
                    Name: InsertAtTopDefinedName,
                    RefersTo: enabled ? "=\"TRUE\"" : "=\"FALSE\"",
                    Visible: false
                );
            }
        }
        catch { }
    }

    /// <summary>
    /// High-speed bulk deduplication: reads the table body in 1 COM call into a 2D memory array,
    /// detects duplicates in RAM, and deletes redundant rows with calculation and screen updating suspended.
    /// </summary>
    private static bool TryGetSeconds(object? raw, out double seconds)
    {
        seconds = 0;
        if (raw is double d) { seconds = Math.Round((d - Math.Floor(d)) * 86400.0) % 86400; return true; }
        if (raw != null && TimeSpan.TryParse(raw.ToString(), CultureInfo.InvariantCulture, out var ts)) { seconds = ts.TotalSeconds; return true; }
        return false;
    }

    public static int CleanTableDuplicateRows(dynamic table, List<string> warnings, int[]? keyColumnIndices = null)
    {
        if (table == null) return 0;

        dynamic? currentBody = null;
        try { currentBody = table.DataBodyRange; } catch { }
        if (currentBody == null) return 0;

        int curRowCount = 0;
        try { curRowCount = (int)currentBody.Rows.Count; } catch { }
        int colCount = (int)table.ListColumns.Count;
        if (curRowCount <= 1 || colCount <= 0) return 0;

        int startCol = (int)table.HeaderRowRange.Column;
        int headerRow = (int)table.HeaderRowRange.Row;
        int curBodyStartRow = (int)currentBody.Row;
        string tableName = table.Name?.ToString() ?? "Table";
        dynamic ws = table.Parent;
        dynamic xlApp = ws.Application;

        // 1. FREEZE COM ENVIRONMENT (Eliminates repaint & formula cascade lag)
        bool origUpdating = true;
        bool origEvents = true;
        int origCalc = -4105; // xlCalculationAutomatic
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origCalc = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }

        try
        {
            // 2. Identify formula columns in 1 COM call across archetype row
            int archetypeRow = headerRow + 1;
            var archFormulas = SheetMapperHelper.ReadArchetypeFormulas((object)ws, archetypeRow, startCol, colCount);
            var isFormulaCol = new bool[colCount + 1];
            for (int c = 1; c <= colCount; c++)
                isFormulaCol[c] = archFormulas[c].StartsWith("=");

            // 3. Bulk-read entire table body in 1 COM call
            object[,] dataGrid = TableMapper.GetRange2DValues(currentBody, curRowCount, colCount);

            // Header row in 1 COM call: calculated columns are ignored, time columns are normalized to HH:mm.
            object[,] headerGrid = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, colCount);
            var headers = new string[colCount + 1];
            var ignoreCol = new bool[colCount + 1];
            var isTimeCol = new bool[colCount + 1];
            for (int c = 1; c <= colCount; c++)
            {
                headers[c] = headerGrid[1, c]?.ToString()?.Trim() ?? string.Empty;
                var colRule = RuleRegistry.GetRule(tableName, headers[c]);
                ignoreCol[c] = colRule?.IsStrictlyCalculated == true;
                isTimeCol[c] = colRule?.EffectiveFormatKind == ColumnFormatKind.Time
                    || headers[c].Equals("Time", StringComparison.OrdinalIgnoreCase)
                    || headers[c].Equals("Clock-In Time", StringComparison.OrdinalIgnoreCase)
                    || headers[c].Equals("Clock-Out Time", StringComparison.OrdinalIgnoreCase);
            }

            string CellText(int r, int c)
            {
                object? raw = dataGrid[r, c];
                if (isTimeCol[c] && raw is double d)
                    return TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(Math.Round((d - Math.Floor(d)) * 86400.0) % 86400))
                        .ToString("HH:mm", CultureInfo.InvariantCulture);
                return SheetMapperHelper.CleanCellText(raw);
            }

            int ColOf(string header)
            {
                for (int c = 1; c <= colCount; c++)
                    if (headers[c].Equals(header, StringComparison.OrdinalIgnoreCase)) return c;
                return -1;
            }

            // AttendanceEntriesTable: natural key + Thumbprint double-scan detection (even if Entry Type was flipped).
            bool isEntries = tableName.Equals("AttendanceEntriesTable", StringComparison.OrdinalIgnoreCase);
            int[]? effectiveKeys = keyColumnIndices;
            int sourceCol = -1, staffCol = -1, dateCol = -1, timeCol = -1;
            if (isEntries)
            {
                dateCol = ColOf("Date"); timeCol = ColOf("Time"); staffCol = ColOf("Staff ID");
                int typeCol = ColOf("Entry Type"); sourceCol = ColOf("Source");
                if (keyColumnIndices == null)
                {
                    var natural = new[] { dateCol, timeCol, staffCol, typeCol, sourceCol }.Where(i => i > 0).ToArray();
                    if (natural.Length > 0) effectiveKeys = natural;
                }
            }
            bool detectDoubleScan = isEntries && keyColumnIndices == null && sourceCol > 0 && staffCol > 0 && dateCol > 0 && timeCol > 0;
            var keptThumbprint = new Dictionary<string, List<double>>(StringComparer.OrdinalIgnoreCase);

            var seenSignatures = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            var duplicateIndices = new List<int>();

            // 4. In-Memory Analysis (0 COM calls inside the loop)
            for (int r = 1; r <= curRowCount; r++)
            {
                // Fast check: verify row contains actual user data
                bool rowHasInput = false;
                for (int c = 1; c <= colCount; c++)
                {
                    if (isFormulaCol[c]) continue;
                    string cleanVal = SheetMapperHelper.CleanCellText(dataGrid[r, c]);
                    if (string.IsNullOrWhiteSpace(cleanVal) || cleanVal == "0" || cleanVal == "0.00" ||
                        cleanVal == "-" || cleanVal == "\u2014" || cleanVal.Equals("None", StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }
                    rowHasInput = true;
                    break;
                }

                if (!rowHasInput) continue;

                // Build unique signature using data values in memory
                var sigParts = new List<string>(colCount);
                for (int c = 1; c <= colCount; c++)
                {
                    if (effectiveKeys != null && !effectiveKeys.Contains(c)) continue;
                    if (isFormulaCol[c] || ignoreCol[c]) continue; // Skip calculated columns

                    sigParts.Add(CellText(r, c));
                }

                string signature = string.Join("|", sigParts);
                if (string.IsNullOrWhiteSpace(signature) || signature.All(ch => ch == '|'))
                {
                    continue;
                }

                int absRow = curBodyStartRow + r - 1;
                bool isDoubleScan = false;
                if (detectDoubleScan &&
                    CellText(r, sourceCol).Equals("Thumbprint", StringComparison.OrdinalIgnoreCase) &&
                    TryGetSeconds(dataGrid[r, timeCol], out double totalSeconds))
                {
                    string staffDay = CellText(r, staffCol) + "|" + CellText(r, dateCol);
                    if (!keptThumbprint.TryGetValue(staffDay, out var kept)) keptThumbprint[staffDay] = kept = new List<double>();
                    if (kept.Any(prevSeconds => Math.Abs(totalSeconds - prevSeconds) < 60)) isDoubleScan = true;
                    else kept.Add(totalSeconds);
                }

                if (isDoubleScan || !seenSignatures.Add(signature))
                {
                    duplicateIndices.Add(r);
                    warnings.Add($"Table '{tableName}' Row {absRow}: Duplicate entry removed ({signature.Trim('|')}).");
                }
            }

            // 5. Delete duplicate rows from bottom up under protection bypass
            if (duplicateIndices.Count > 0)
            {
                foreach (int idx in duplicateIndices.OrderByDescending(x => x))
                {
                    table.ListRows[idx].Delete();
                }
            }

            return duplicateIndices.Count;
        }
        finally
        {
            // 6. Restore Excel environment in correct sequence
            try { xlApp.Calculation = origCalc; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

}
