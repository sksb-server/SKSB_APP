namespace SKSB_App.Core.RuntimeUI.TableRowManager;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookOpen.Sheet;

public enum TableInsertMode
{
    InsertAtBottom,
    InsertAtTop
}

public static class TableRowManager
{
    private static bool _isManagingRows = false;
    private const string InsertAtTopDefinedName = "__SKSB_INSERT_AT_TOP";
    private const string DebounceSigDefinedName = "__SKSB_ROW_SIG";
    private static bool _inMemoryFallback = false;

    private static readonly ConcurrentDictionary<string, TableInsertMode> TableModes =
        new(StringComparer.OrdinalIgnoreCase);

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

    public static void HandleSheetChange(dynamic singleCell)
    {
        if (singleCell == null || SheetOperations.IsTemplateEditMode || _isManagingRows) return;

        dynamic? table = null;
        try { table = singleCell.ListObject; } catch { }
        if (table == null) return;

        dynamic dataBody = table.DataBodyRange;
        int editedRow = (int)singleCell.Row;
        int bodyStartRow = dataBody != null ? (int)dataBody.Row : (int)table.HeaderRowRange.Row + 1;
        int rowCount = dataBody != null ? (int)dataBody.Rows.Count : 0;
        int lastDataRow = bodyStartRow + Math.Max(1, rowCount) - 1;

        TableInsertMode insertMode = GetTableInsertMode(table);
        bool isAtBottom = (editedRow == lastDataRow);
        bool isAtTop = (editedRow == bodyStartRow);
        bool cellIsEmpty = string.IsNullOrEmpty(singleCell.Value2?.ToString()?.Trim());

        if (!isAtBottom && !isAtTop && !cellIsEmpty) return;

        string cellAddress = string.Empty;
        string cellVal = string.Empty;
        try
        {
            cellAddress = singleCell.Address?.ToString() ?? "Unknown";
            cellVal = singleCell.Value2?.ToString()?.Trim() ?? string.Empty;
        }
        catch { }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (_isManagingRows) return;
            _isManagingRows = true;

            dynamic xlApp = ExcelDnaUtil.Application;
            bool origEnableEvents = true;
            try { origEnableEvents = (bool)xlApp.EnableEvents; } catch { }

            var auditTrail = new List<string>();
            string tableName = string.Empty;

            try
            {
                xlApp.EnableEvents = false;
                dynamic worksheet = table.Parent;
                dynamic wb = worksheet.Parent;
                tableName = table.Name?.ToString() ?? "Table";

                string currentSig = $"{tableName}_{cellAddress}_{cellVal}";
                if (IsDebounced(wb, currentSig))
                {
                    return;
                }
                RecordDebounceSignature(wb, currentSig);

                insertMode = GetTableInsertMode(table);

                SheetProtections.ExecuteWithBypass((object)worksheet, new Action(() =>
                {
                    // Ensure ShowTotals is permanently disabled
                    try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }

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
                        if (lastRowHasInput && editedRow == lastRowIdx)
                        {
                            SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, false, auditTrail);
                        }
                    }
                    else if (insertMode == TableInsertMode.InsertAtTop)
                    {
                        int topRowIdx = curBodyStartRow;
                        bool topRowHasInput = CheckRowHasUserInput(worksheet, table, topRowIdx, startCol, colCount, headerRow);
                        if (topRowHasInput && editedRow == topRowIdx)
                        {
                            SpawnRowSafely(worksheet, table, startCol, colCount, headerRow, true, auditTrail);
                        }
                    }

                    CleanTableEmptyRows(table, insertMode, auditTrail);
                }));
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TableRowManager.HandleSheetChange Error] {ex.Message}");
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
                    try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }
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

        try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }

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
        int archetypeRow = headerRow + 1;

        for (int c = 1; c <= colCount; c++)
        {
            int absCol = startCol + c - 1;
            dynamic cell = worksheet.Cells[rowIndex, absCol];
            dynamic archCell = worksheet.Cells[archetypeRow, absCol];

            string archFormula = archCell.Formula?.ToString() ?? string.Empty;
            if (archFormula.StartsWith("=")) continue;

            string formula = cell.Formula?.ToString() ?? string.Empty;
            if (formula.StartsWith("=")) continue;

            bool isLocked = true;
            try { isLocked = (bool)cell.Locked; } catch { }
            if (isLocked) continue;

            object? valObj = cell.Value2;
            if (valObj == null) continue;

            if (valObj is int iErr && iErr < -2000000000) continue;
            if (valObj is long lErr && lErr < -2000000000) continue;

            string valStr = valObj.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(valStr))
            {
                return true;
            }
        }

        return false;
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
        try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }

        string tableName = table.Name?.ToString() ?? string.Empty;
        dynamic newRow = atTop ? table.ListRows.Add(1) : table.ListRows.Add();
        dynamic newRowRange = newRow.Range;

        int sourceArchetypeRow = atTop ? headerRow + 2 : headerRow + 1;

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
                    newCell.Style = styleName;
                }
            }

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

            newCell.Locked = false;

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
        try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }
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

    private static bool IsDebounced(dynamic wb, string currentSignature)
    {
        try
        {
            dynamic nameObj = null;
            try { nameObj = wb.Names.Item(DebounceSigDefinedName); } catch { }
            if (nameObj != null)
            {
                string stored = nameObj.RefersTo?.ToString() ?? string.Empty;
                var tokens = stored.TrimStart('=').Trim('"').Split('|');
                if (tokens.Length == 2)
                {
                    string sig = tokens[0];
                    if (long.TryParse(tokens[1], out long timestampTicks))
                    {
                        var elapsed = DateTime.UtcNow - new DateTime(timestampTicks, DateTimeKind.Utc);
                        if (string.Equals(sig, currentSignature, StringComparison.OrdinalIgnoreCase) && elapsed.TotalMilliseconds < 1200)
                        {
                            return true;
                        }
                    }
                }
            }
        }
        catch { }
        return false;
    }

    private static void RecordDebounceSignature(dynamic wb, string currentSignature)
    {
        try
        {
            try { wb.Names.Item(DebounceSigDefinedName).Delete(); } catch { }
            string val = $"\"{currentSignature}|{DateTime.UtcNow.Ticks}\"";
            wb.Names.Add(Name: DebounceSigDefinedName, RefersTo: $"={val}", Visible: false);
        }
        catch { }
    }
}