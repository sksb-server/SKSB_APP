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
using System.Text;
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
    private const string DebounceSigDefinedName = "__SKSB_ROW_SIG";
    private static bool _inMemoryFallback = false;

    private static readonly ConcurrentDictionary<string, TableInsertMode> TableModes =
        new(StringComparer.OrdinalIgnoreCase);

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
        TableNatures["StaffLicensesTable"] = TableNature.Additive;
        TableNatures["StaffLeavesTable"] = TableNature.Additive;
    }

    public static void SetTableNature(string tableName, TableNature nature)
    {
        if (!string.IsNullOrWhiteSpace(tableName))
        {
            TableNatures[tableName.Trim()] = nature;
        }
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

    public static void SetTableInsertMode(string tableName, TableInsertMode mode)
    {
        if (!string.IsNullOrWhiteSpace(tableName))
        {
            TableModes[tableName.Trim()] = mode;
        }
    }

    public static TableInsertMode GetTableInsertMode(dynamic table)
    {
        try
        {
            string tableName = table?.Name?.ToString() ?? string.Empty;
            if (!string.IsNullOrEmpty(tableName) && TableModes.TryGetValue(tableName, out var mode))
            {
                return mode;
            }
        }
        catch { }

        return IsInsertAtTopEnabled ? TableInsertMode.InsertAtTop : TableInsertMode.InsertAtBottom;
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

            SheetProtections.ExecuteWithBypass((object)worksheet, new Action(() =>
            {
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
            }));

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
        try { table = target.ListObject; } catch { }
        if (table == null)
        {
            try { table = target.Cells[1, 1].ListObject; } catch { }
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

        TableInsertMode insertMode = GetTableInsertMode(table);

        int targetTopRow = (int)target.Row;
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

                SheetProtections.ExecuteWithBypass((object)worksheet, new Action(() =>
                {
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
                }));
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
            SheetProtections.ExecuteWithBypass((object)worksheet, new Action(() =>
            {
                foreach (dynamic table in worksheet.ListObjects)
                {
                    TableInsertMode mode = GetTableInsertMode(table);
                    CleanTableEmptyRows(table, mode, null);
                    tablesCleaned++;
                }
            }));
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

                if (rule?.DefaultRule?.Kind == DefaultKind.Literal &&
                    string.Equals(cellVal, rule.DefaultRule.LiteralValue, StringComparison.OrdinalIgnoreCase))
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
        dynamic newRow = atTop ? table.ListRows.Add(1) : table.ListRows.Add();
        dynamic newRowRange = newRow.Range;

        int sourceArchetypeRow = atTop ? headerRow + 2 : headerRow + 1;

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
                newCell.Style = Commands.DropDownStyleName;
                try { newCell.Validation.Delete(); } catch { }
            }
            else
            {
                string styleName = archetypeCell.Style?.Name?.ToString() ?? string.Empty;
                if (!string.IsNullOrEmpty(styleName))
                {
                    try { newCell.Style = styleName; } catch { }
                }
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

            // Ensure cell is unlocked if the style is whitelisted
            string appliedStyle = newCell.Style?.Name?.ToString() ?? string.Empty;
            newCell.Locked = !SheetProtections.IsUnlockedStyle(appliedStyle);

            var cfg = Metadata.ReadConfig(archetypeCell);
            if (!cfg.IsEmpty)
            {
                Metadata.WriteConfig(newCell, cfg);
            }

            string archFormula = archetypeCell.Formula?.ToString() ?? string.Empty;
            if (!archFormula.StartsWith("="))
            {
                newCell.Value2 = null;
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
            dynamic wb = xlApp?.ActiveWorkbook;
            if (wb != null)
            {
                dynamic nameObj = null;
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
            dynamic wb = xlApp?.ActiveWorkbook;
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
            var isFormulaCol = new bool[colCount + 1];
            try
            {
                dynamic archRange = ws.Range[ws.Cells[archetypeRow, startCol], ws.Cells[archetypeRow, startCol + colCount - 1]];
                object? rawFormulas = archRange.Formula;
                if (rawFormulas is object[,] fMat)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        string f = fMat[1, c]?.ToString() ?? string.Empty;
                        isFormulaCol[c] = f.StartsWith("=");
                    }
                }
                else if (rawFormulas is string singleF)
                {
                    isFormulaCol[1] = singleF.StartsWith("=");
                }
            }
            catch { }

            // 3. Bulk-read entire table body in 1 COM call
            object[,] dataGrid = TableMapper.GetRange2DValues(currentBody, curRowCount, colCount);

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
                    string cleanVal = CleanCellText(dataGrid[r, c]);
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
                    if (keyColumnIndices != null && !keyColumnIndices.Contains(c)) continue;
                    if (isFormulaCol[c]) continue; // Skip calculated columns

                    string cleanVal = CleanCellText(dataGrid[r, c]);
                    sigParts.Add(cleanVal);
                }

                string signature = string.Join("|", sigParts);
                if (string.IsNullOrWhiteSpace(signature) || signature.All(ch => ch == '|'))
                {
                    continue;
                }

                int absRow = curBodyStartRow + r - 1;
                if (!seenSignatures.Add(signature))
                {
                    duplicateIndices.Add(r);
                    warnings.Add($"Table '{tableName}' Row {absRow}: Duplicate entry removed ({signature.Trim('|')}).");
                }
            }

            // 5. Delete duplicate rows from bottom up under protection bypass
            if (duplicateIndices.Count > 0)
            {
                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
                    foreach (int idx in duplicateIndices.OrderByDescending(x => x))
                    {
                        table.ListRows[idx].Delete();
                    }
                }));
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

    private static string CleanCellText(object? val)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        if (val is double d && Math.Abs(d - Math.Round(d)) < 0.00001 && d < 1e15)
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (s.StartsWith("#") || s == " ") return string.Empty;
        return s;
    }
}