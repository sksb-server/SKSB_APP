namespace SKSB_App.Core.RuntimeUI.Undo;

using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

public static class UndoManager
{
    // =========================================================================
    // 1. STATE & DATA STRUCTURES
    // =========================================================================

    private const int MaxHistoryDepth = 50;
    private const int MaxSnapshotCells = 5000;
    private const int MaxTableExpansionRows = 200;

    public static bool IsApplyingUndoRedo { get; private set; }

    private static readonly Dictionary<string, WorkbookUndoState> _workbookStates =
        new(StringComparer.OrdinalIgnoreCase);

    private static PendingSnapshot? _pendingSnapshot;

    private sealed class WorkbookUndoState
    {
        public List<UndoEntry> UndoStack { get; } = new();
        public List<UndoEntry> RedoStack { get; } = new();
    }

    private sealed class UndoEntry
    {
        public required string WorkbookKey { get; init; }
        public required string SheetName { get; init; }
        public required int TopRow { get; init; }
        public required int LeftCol { get; init; }
        public required int RowCount { get; init; }
        public required int ColCount { get; init; }
        public required object?[,] BeforeValues { get; init; }
        public required object?[,] AfterValues { get; init; }
        public required bool[,] SkipMask { get; init; }
        public required string Description { get; init; }
    }

    private sealed class PendingSnapshot
    {
        public required string WorkbookKey { get; init; }
        public required string SheetName { get; init; }
        public required int TopRow { get; init; }
        public required int LeftCol { get; init; }
        public required int RowCount { get; init; }
        public required int ColCount { get; init; }
        public required object?[,] Values { get; set; }
        public required bool[,] SkipMask { get; set; }
        public string? TableName { get; init; }
    }

    // =========================================================================
    // 2. SNAPSHOT CAPTURE & COMMIT
    // =========================================================================

    public static void CaptureSelectionSnapshot(dynamic sh, dynamic target)
    {
        try
        {
            if (IsApplyingUndoRedo || sh == null || target == null || SheetOperations.IsTemplateEditMode)
                return;

            int areasCount = 0;
            try { areasCount = (int)target!.Areas.Count; } catch { return; }
            if (areasCount != 1)
            {
                _pendingSnapshot = null;
                return;
            }

            int topRow = (int)target!.Row;
            int leftCol = (int)target.Column;
            int rowCount = (int)target.Rows.Count;
            int colCount = (int)target.Columns.Count;

            if (rowCount <= 0 || colCount <= 0 || (long)rowCount * colCount > MaxSnapshotCells)
            {
                _pendingSnapshot = null;
                return;
            }

            dynamic? wb = null;
            try { wb = sh!.Parent; } catch { }
            string wbKey = GetWorkbookKey(wb);
            string sheetName = sh!.Name?.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(wbKey) || string.IsNullOrEmpty(sheetName))
            {
                _pendingSnapshot = null;
                return;
            }

            string? tableName = null;
            if (rowCount <= MaxTableExpansionRows)
            {
                try
                {
                    dynamic? table = target.Cells[1, 1].ListObject;
                    if (table != null)
                    {
                        dynamic? body = table.DataBodyRange;
                        if (body != null)
                        {
                            dynamic xlApp = ExcelDnaUtil.Application;
                            dynamic? hit = xlApp.Intersect(target, body);
                            if (hit != null &&
                                (int)hit!.Row == topRow &&
                                (int)hit.Column == leftCol &&
                                (int)hit.Rows.Count == rowCount &&
                                (int)hit.Columns.Count == colCount)
                            {
                                leftCol = (int)table.HeaderRowRange.Column;
                                colCount = (int)table.ListColumns.Count;
                                tableName = table.Name?.ToString();
                            }
                        }
                    }
                }
                catch { }
            }

            if ((long)rowCount * colCount > MaxSnapshotCells)
            {
                _pendingSnapshot = null;
                return;
            }

            dynamic snapRange = sh.Range[sh.Cells[topRow, leftCol], sh.Cells[topRow + rowCount - 1, leftCol + colCount - 1]];
            object?[,] values = TableMapper.GetRange2DValues(snapRange, rowCount, colCount);
            bool[,] skipMask = ReadFormulaSkipMask(snapRange, rowCount, colCount);

            _pendingSnapshot = new PendingSnapshot
            {
                WorkbookKey = wbKey,
                SheetName = sheetName,
                TopRow = topRow,
                LeftCol = leftCol,
                RowCount = rowCount,
                ColCount = colCount,
                Values = values,
                SkipMask = skipMask,
                TableName = tableName
            };
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[UndoManager.CaptureSelectionSnapshot] {ex.Message}");
            _pendingSnapshot = null;
        }
    }

    public static void CaptureCellSnapshot(dynamic sh, dynamic singleCell)
    {
        CaptureSelectionSnapshot(sh, singleCell);
    }

    public static void CommitEdit(dynamic sh, dynamic target, string? description = null)
    {
        try
        {
            if (IsApplyingUndoRedo || sh == null || target == null || SheetOperations.IsTemplateEditMode)
                return;

            var snap = _pendingSnapshot;
            if (snap == null)
            {
                CaptureSelectionSnapshot(sh, target);
                return;
            }

            dynamic? wb = null;
            try { wb = sh!.Parent; } catch { }
            string wbKey = GetWorkbookKey(wb);
            string sheetName = sh!.Name?.ToString() ?? string.Empty;

            if (!string.Equals(snap.WorkbookKey, wbKey, StringComparison.OrdinalIgnoreCase) ||
                !string.Equals(snap.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                CaptureSelectionSnapshot(sh, target);
                return;
            }

            int areasCount = 0;
            try { areasCount = (int)target!.Areas.Count; } catch { return; }
            if (areasCount != 1)
            {
                _pendingSnapshot = null;
                return;
            }

            int targetTop = (int)target!.Row;
            int targetLeft = (int)target.Column;
            int targetRows = (int)target.Rows.Count;
            int targetCols = (int)target.Columns.Count;
            int targetBottom = targetTop + targetRows - 1;
            int targetRight = targetLeft + targetCols - 1;

            int snapBottom = snap.TopRow + snap.RowCount - 1;
            int snapRight = snap.LeftCol + snap.ColCount - 1;

            if (targetTop < snap.TopRow || targetBottom > snapBottom ||
                targetLeft < snap.LeftCol || targetRight > snapRight)
            {
                CaptureSelectionSnapshot(sh, target);
                return;
            }

            dynamic snapRange = sh.Range[sh.Cells[snap.TopRow, snap.LeftCol], sh.Cells[snapBottom, snapRight]];
            object?[,] afterValues = TableMapper.GetRange2DValues(snapRange, snap.RowCount, snap.ColCount);
            bool[,] afterFormulas = ReadFormulaSkipMask(snapRange, snap.RowCount, snap.ColCount);

            var combinedSkip = new bool[snap.RowCount + 1, snap.ColCount + 1];
            bool anyChange = false;

            for (int r = 1; r <= snap.RowCount; r++)
            {
                for (int c = 1; c <= snap.ColCount; c++)
                {
                    bool skip = snap.SkipMask[r, c] || afterFormulas[r, c];
                    combinedSkip[r, c] = skip;
                    if (!skip && !ValuesEqual(snap.Values[r, c], afterValues[r, c]))
                    {
                        anyChange = true;
                    }
                }
            }

            if (!anyChange) return;

            string desc = description ?? BuildEditDescription(target, targetRows, targetCols);

            var entry = new UndoEntry
            {
                WorkbookKey = wbKey,
                SheetName = sheetName,
                TopRow = snap.TopRow,
                LeftCol = snap.LeftCol,
                RowCount = snap.RowCount,
                ColCount = snap.ColCount,
                BeforeValues = CloneGrid(snap.Values, snap.RowCount, snap.ColCount),
                AfterValues = CloneGrid(afterValues, snap.RowCount, snap.ColCount),
                SkipMask = combinedSkip,
                Description = desc
            };

            var state = GetOrCreateState(wbKey);
            state.UndoStack.Add(entry);
            while (state.UndoStack.Count > MaxHistoryDepth)
            {
                state.UndoStack.RemoveAt(0);
            }
            state.RedoStack.Clear();

            snap.Values = afterValues;
            snap.SkipMask = combinedSkip;

            ReArmOnUndoIfAvailable();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[UndoManager.CommitEdit] {ex.Message}");
        }
    }

    // =========================================================================
    // 3. UNDO / REDO EXECUTION & EXCEL TOOLBAR INTEGRATION
    // =========================================================================

    [ExcelCommand(Name = "SKSB_UndoMacro")]
    public static void ExecuteUndoCommand()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = null;
            try { wb = xlApp.ActiveWorkbook; } catch { }
            if (wb == null) return;

            if (SheetOperations.IsTemplateEditMode || SheetOperations.IsMasterTemplateFile(wb))
            {
                try { xlApp.Undo(); } catch { }
                return;
            }

            string wbKey = GetWorkbookKey(wb);
            if (!_workbookStates.TryGetValue(wbKey, out var state) || state.UndoStack.Count == 0)
            {
                try { xlApp.Undo(); } catch { }
                return;
            }

            int lastIdx = state.UndoStack.Count - 1;
            UndoEntry entry = state.UndoStack[lastIdx];
            state.UndoStack.RemoveAt(lastIdx);

            ApplyEntry(xlApp, wb, state, entry, isUndo: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error("ExecuteUndoCommand failed", ex);
        }
    }

    [ExcelCommand(Name = "SKSB_RedoMacro")]
    public static void ExecuteRedoCommand()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = null;
            try { wb = xlApp.ActiveWorkbook; } catch { }
            if (wb == null) return;

            if (SheetOperations.IsTemplateEditMode || SheetOperations.IsMasterTemplateFile(wb))
            {
                try { xlApp.Repeat(); } catch { }
                return;
            }

            string wbKey = GetWorkbookKey(wb);
            if (!_workbookStates.TryGetValue(wbKey, out var state) || state.RedoStack.Count == 0)
            {
                try { xlApp.Repeat(); } catch { }
                return;
            }

            int lastIdx = state.RedoStack.Count - 1;
            UndoEntry entry = state.RedoStack[lastIdx];
            state.RedoStack.RemoveAt(lastIdx);

            ApplyEntry(xlApp, wb, state, entry, isUndo: false);
        }
        catch (Exception ex)
        {
            AppLogger.Error("ExecuteRedoCommand failed", ex);
        }
    }

    public static void ReArmOnUndoIfAvailable()
    {
        try
        {
            if (SheetOperations.IsTemplateEditMode) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = null;
            try { wb = xlApp.ActiveWorkbook; } catch { }
            if (wb == null) return;

            string wbKey = GetWorkbookKey(wb);
            if (_workbookStates.TryGetValue(wbKey, out var state) && state.UndoStack.Count > 0)
            {
                UndoEntry top = state.UndoStack[state.UndoStack.Count - 1];
                try { xlApp.OnUndo($"Undo {top.Description}", "SKSB_UndoMacro"); } catch { }
            }
        }
        catch { }
    }

    public static void ClearWorkbook(dynamic? wb)
    {
        try
        {
            string wbKey = GetWorkbookKey(wb);
            if (!string.IsNullOrEmpty(wbKey))
            {
                _workbookStates.Remove(wbKey);
                if (_pendingSnapshot != null &&
                    string.Equals(_pendingSnapshot.WorkbookKey, wbKey, StringComparison.OrdinalIgnoreCase))
                {
                    _pendingSnapshot = null;
                }
            }
        }
        catch { }
    }

    public static void ClearSheet(dynamic? worksheet)
    {
        try
        {
            if (worksheet == null) return;
            dynamic? wb = null;
            try { wb = worksheet.Parent; } catch { }
            string wbKey = GetWorkbookKey(wb);
            string sheetName = worksheet.Name?.ToString() ?? string.Empty;
            if (string.IsNullOrEmpty(wbKey) || string.IsNullOrEmpty(sheetName)) return;

            if (_workbookStates.TryGetValue(wbKey, out var state))
            {
                state.UndoStack.RemoveAll(e => string.Equals(e.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));
                state.RedoStack.RemoveAll(e => string.Equals(e.SheetName, sheetName, StringComparison.OrdinalIgnoreCase));
            }

            if (_pendingSnapshot != null &&
                string.Equals(_pendingSnapshot.WorkbookKey, wbKey, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_pendingSnapshot.SheetName, sheetName, StringComparison.OrdinalIgnoreCase))
            {
                _pendingSnapshot = null;
            }
        }
        catch { }
    }

    public static void ClearAll()
    {
        _workbookStates.Clear();
        _pendingSnapshot = null;
    }

    // =========================================================================
    // 4. INTERNAL HELPERS
    // =========================================================================

    private static void ApplyEntry(dynamic xlApp, dynamic wb, WorkbookUndoState state, UndoEntry entry, bool isUndo)
    {
        dynamic? ws = null;
        try { ws = wb.Worksheets[entry.SheetName]; } catch { }
        if (ws == null) return;

        bool origUpdating = true;
        bool origEvents = true;
        int origCalc = -4105;
        dynamic? targetRange = null;

        IsApplyingUndoRedo = true;
        try
        {
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
            try { origCalc = (int)xlApp.Calculation; xlApp.Calculation = -4135; } catch { }

            int top = entry.TopRow;
            int left = entry.LeftCol;
            int rows = entry.RowCount;
            int cols = entry.ColCount;

            targetRange = ws.Range[ws.Cells[top, left], ws.Cells[top + rows - 1, left + cols - 1]];
            bool[,] liveFormulaMask = ReadFormulaSkipMask(targetRange, rows, cols);
            object?[,] sourceGrid = isUndo ? entry.BeforeValues : entry.AfterValues;

            bool anySkipped = false;
            for (int r = 1; r <= rows && !anySkipped; r++)
            {
                for (int c = 1; c <= cols; c++)
                {
                    if (entry.SkipMask[r, c] || liveFormulaMask[r, c])
                    {
                        anySkipped = true;
                        break;
                    }
                }
            }

            if (!anySkipped)
            {
                if (rows == 1 && cols == 1)
                {
                    targetRange.Value2 = sourceGrid[1, 1];
                }
                else
                {
                    var bulk = new object?[rows, cols];
                    for (int r = 1; r <= rows; r++)
                        for (int c = 1; c <= cols; c++)
                            bulk[r - 1, c - 1] = sourceGrid[r, c];
                    targetRange.Value2 = bulk;
                }
            }
            else
            {
                for (int r = 1; r <= rows; r++)
                {
                    int c = 1;
                    while (c <= cols)
                    {
                        if (entry.SkipMask[r, c] || liveFormulaMask[r, c])
                        {
                            c++;
                            continue;
                        }

                        int runStart = c;
                        while (c <= cols && !entry.SkipMask[r, c] && !liveFormulaMask[r, c])
                        {
                            c++;
                        }
                        int runLen = c - runStart;

                        int absRow = top + r - 1;
                        int absColStart = left + runStart - 1;
                        if (runLen == 1)
                        {
                            ws.Cells[absRow, absColStart].Value2 = sourceGrid[r, runStart];
                        }
                        else
                        {
                            var runArr = new object?[1, runLen];
                            for (int k = 0; k < runLen; k++)
                                runArr[0, k] = sourceGrid[r, runStart + k];
                            dynamic runRange = ws.Range[ws.Cells[absRow, absColStart], ws.Cells[absRow, absColStart + runLen - 1]];
                            runRange.Value2 = runArr;
                        }
                    }
                }
            }

            if (isUndo)
            {
                state.RedoStack.Add(entry);
                while (state.RedoStack.Count > MaxHistoryDepth)
                    state.RedoStack.RemoveAt(0);
            }
            else
            {
                state.UndoStack.Add(entry);
                while (state.UndoStack.Count > MaxHistoryDepth)
                    state.UndoStack.RemoveAt(0);
            }

            if (_pendingSnapshot != null &&
                string.Equals(_pendingSnapshot.WorkbookKey, entry.WorkbookKey, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(_pendingSnapshot.SheetName, entry.SheetName, StringComparison.OrdinalIgnoreCase) &&
                _pendingSnapshot.TopRow == entry.TopRow &&
                _pendingSnapshot.LeftCol == entry.LeftCol &&
                _pendingSnapshot.RowCount == entry.RowCount &&
                _pendingSnapshot.ColCount == entry.ColCount)
            {
                _pendingSnapshot.Values = CloneGrid(sourceGrid, rows, cols);
            }

            if (targetRange != null)
            {
                ParameterCells.NotifyChanged(ws, targetRange);
            }
            Arrow.HideAll();
        }
        finally
        {
            try { xlApp.Calculation = origCalc; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            IsApplyingUndoRedo = false;
        }

        try
        {
            dynamic? sel = xlApp.Selection;
            if (sel != null)
            {
                CaptureSelectionSnapshot(ws, sel);
            }
        }
        catch { }

        SheetOperations.ShowStatusConfirmation(isUndo ? $"Undone: {entry.Description}" : $"Redone: {entry.Description}");
        ReArmOnUndoIfAvailable();
    }

    private static WorkbookUndoState GetOrCreateState(string wbKey)
    {
        if (!_workbookStates.TryGetValue(wbKey, out var state))
        {
            state = new WorkbookUndoState();
            _workbookStates[wbKey] = state;
        }
        return state;
    }

    private static string GetWorkbookKey(dynamic? wb)
    {
        if (wb == null) return string.Empty;
        try
        {
            string? full = wb.FullName?.ToString();
            if (!string.IsNullOrWhiteSpace(full)) return full;
        }
        catch { }
        try
        {
            return wb.Name?.ToString() ?? string.Empty;
        }
        catch { return string.Empty; }
    }

    private static bool[,] ReadFormulaSkipMask(dynamic range, int rows, int cols)
    {
        var mask = new bool[rows + 1, cols + 1];
        try
        {
            object? raw = range.Formula;
            if (raw is object[,] arr)
            {
                int rLow = arr.GetLowerBound(0);
                int cLow = arr.GetLowerBound(1);
                for (int r = 1; r <= rows; r++)
                {
                    for (int c = 1; c <= cols; c++)
                    {
                        string? f = arr[rLow + r - 1, cLow + c - 1]?.ToString();
                        mask[r, c] = !string.IsNullOrEmpty(f) && f.StartsWith("=", StringComparison.Ordinal);
                    }
                }
            }
            else if (rows == 1 && cols == 1)
            {
                string? f = raw?.ToString();
                mask[1, 1] = !string.IsNullOrEmpty(f) && f.StartsWith("=", StringComparison.Ordinal);
            }
        }
        catch { }
        return mask;
    }

    private static object?[,] CloneGrid(object?[,] src, int rows, int cols)
    {
        var copy = new object?[rows + 1, cols + 1];
        for (int r = 1; r <= rows; r++)
            for (int c = 1; c <= cols; c++)
                copy[r, c] = src[r, c];
        return copy;
    }

    private static bool ValuesEqual(object? a, object? b)
    {
        bool aEmpty = a == null || (a is string sa && sa.Length == 0);
        bool bEmpty = b == null || (b is string sb && sb.Length == 0);
        if (aEmpty && bEmpty) return true;
        if (aEmpty || bEmpty) return false;

        if (TryToDouble(a, out double da) && TryToDouble(b, out double db))
        {
            return Math.Abs(da - db) < 1e-9;
        }

        if (a is bool ba && b is bool bb)
        {
            return ba == bb;
        }

        return string.Equals(
            Convert.ToString(a, CultureInfo.InvariantCulture) ?? string.Empty,
            Convert.ToString(b, CultureInfo.InvariantCulture) ?? string.Empty,
            StringComparison.Ordinal);
    }

    private static bool TryToDouble(object? val, out double d)
    {
        switch (val)
        {
            case double dbl:
                d = dbl;
                return true;
            case float flt:
                d = flt;
                return true;
            case decimal dec:
                d = (double)dec;
                return true;
            case int i:
                d = i;
                return true;
            case long l:
                d = l;
                return true;
            default:
                d = 0;
                return false;
        }
    }

    private static string BuildEditDescription(dynamic target, int rows, int cols)
    {
        try
        {
            string addr = target.Address[false, false]?.ToString() ?? "Cells";
            return rows == 1 && cols == 1 ? $"Typing in {addr}" : $"Edit {addr}";
        }
        catch
        {
            return "Cell Edit";
        }
    }
}
