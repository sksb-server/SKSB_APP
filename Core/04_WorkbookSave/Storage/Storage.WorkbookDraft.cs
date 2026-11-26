namespace SKSB_App.Core.WorkbookSave.Storage;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.Json.Serialization;
using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.Undo;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;

// =========================================================================
// 1. DRAFT DTOs
// =========================================================================

/// <summary>Lossless, type-tagged snapshot of one Excel <c>Value2</c> (string / number / bool).</summary>
public sealed class DraftCellValue
{
    /// <summary><c>"s"</c> string, <c>"n"</c> number (dates as OADate), <c>"b"</c> boolean.</summary>
    public string Kind { get; set; } = "s";
    public string? S { get; set; }
    public double? N { get; set; }
    public bool? B { get; set; }

    /// <summary>Null for empty cells and COM errors (#N/A, #REF!, ...).</summary>
    public static DraftCellValue? FromExcel(object? raw)
    {
        if (raw == null || SheetMapperHelper.IsComError(raw)) return null;
        switch (raw)
        {
            case string s: return new DraftCellValue { Kind = "s", S = s };
            case bool b: return new DraftCellValue { Kind = "b", B = b };
            case double d: return new DraftCellValue { Kind = "n", N = d };
            case int i: return new DraftCellValue { Kind = "n", N = i };
            case long l: return new DraftCellValue { Kind = "n", N = l };
            case decimal m: return new DraftCellValue { Kind = "n", N = (double)m };
            case DateTime dt: return new DraftCellValue { Kind = "n", N = dt.ToOADate() };
            default: return new DraftCellValue { Kind = "s", S = Convert.ToString(raw, CultureInfo.InvariantCulture) };
        }
    }

    /// <summary>COM-safe <c>Value2</c>: <c>string</c>, <c>double</c>, <c>bool</c> or <c>null</c>.</summary>
    public object? ToExcel() => Kind switch
    {
        "n" => N,
        "b" => B,
        _ => S
    };
}

/// <summary>Every row (complete or not) of one ListObject, keyed by header text.</summary>
public sealed class SheetTableDraft
{
    public string SheetName { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;
    public List<string> Headers { get; set; } = new();
    public List<List<DraftCellValue?>> Rows { get; set; } = new();
}

/// <summary>Daily Schedule board: date/day top bar plus the slot grid (rows 4..34, cols 2..1+SlotCount*3).</summary>
public sealed class ScheduleBoardDraft
{
    public string SheetName { get; set; } = string.Empty;
    public DraftCellValue? DateCell { get; set; }
    public DraftCellValue? DayCell { get; set; }
    public int SlotCount { get; set; } = 1;
    public List<List<DraftCellValue?>> Grid { get; set; } = new();
}

public sealed class WorkbookDraftPayload
{
    public string ModuleName { get; set; } = string.Empty;
    public DateTime SavedUtc { get; set; }
    public int ErrorCount { get; set; }
    public int WarningCount { get; set; }
    public List<SheetTableDraft> Tables { get; set; } = new();
    public ScheduleBoardDraft? ScheduleBoard { get; set; }
}

/// <summary>
/// Local "Save Draft &amp; Close" store (<c>%LOCALAPPDATA%\SKSB_App\Data\&lt;type&gt;_Draft.json</c>). A draft is a raw
/// cell-level copy of the operational sheets taken while validation still fails; it is restored onto the fresh
/// template on next open and is NEVER uploaded to PostgreSQL. It is cleared once a valid save commits.
/// </summary>
public static class WorkbookDraft
{
    private const int BoardFirstRow = 4;
    private const int BoardLastRow = 34;
    private const int BoardFirstCol = 2;
    private const int ColsPerSlot = 3;
    private const int MaxSlots = 100;

    private static readonly string[] ScheduleSheetNames = { "Daily Schedule", "Daily_Schedule", "Schedule" };
    private static readonly string[] SkippedSheetNames = { "Lookups", "Data" };

    private static readonly JsonSerializerOptions JsonOptions = new()
    {
        WriteIndented = false,
        PropertyNameCaseInsensitive = true,
        DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull
    };

    // =========================================================================
    // 2. FILE MANAGEMENT
    // =========================================================================
    private static string ResolveTypeName(Type type)
    {
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>))
            return type.GetGenericArguments()[0].Name;
        if (type.IsArray)
            return type.GetElementType()?.Name ?? "Object";
        return type.Name;
    }

    public static string GetDraftPath<T>() =>
        FolderRegistry.Data.GetFilePath($"{ResolveTypeName(typeof(T)).ToLowerInvariant()}_Draft.json");

    public static bool HasDraft<T>()
    {
        try { return File.Exists(GetDraftPath<T>()); }
        catch { return false; }
    }

    public static WorkbookDraftPayload? Load<T>()
    {
        try
        {
            string path = GetDraftPath<T>();
            if (!File.Exists(path)) return null;
            return JsonSerializer.Deserialize<WorkbookDraftPayload>(File.ReadAllText(path), JsonOptions);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Could not read draft for {typeof(T).Name}: {ex.Message}");
            return null;
        }
    }

    public static void Clear<T>()
    {
        try
        {
            string path = GetDraftPath<T>();
            if (File.Exists(path))
            {
                File.Delete(path);
                AppLogger.Info($"[WorkbookDraft] Cleared local draft '{Path.GetFileName(path)}'.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Could not clear draft for {typeof(T).Name}: {ex.Message}");
        }
    }

    // =========================================================================
    // 3. SAVE (bulk reads only; macro thread)
    // =========================================================================
    public static void Save<T>(object? wbObj, string moduleName, int errorCount, int warningCount)
    {
        if (wbObj == null) throw new ArgumentNullException(nameof(wbObj), "No active workbook to save as a draft.");

        var payload = new WorkbookDraftPayload
        {
            ModuleName = moduleName,
            SavedUtc = DateTime.UtcNow,
            ErrorCount = errorCount,
            WarningCount = warningCount
        };

        dynamic wb = wbObj;
        foreach (dynamic ws in wb.Sheets)
        {
            string sheetName = string.Empty;
            try { sheetName = Convert.ToString((object)ws.Name) ?? string.Empty; } catch { }
            if (SkippedSheetNames.Any(n => string.Equals(n, sheetName.Trim(), StringComparison.OrdinalIgnoreCase))) continue;

            int visible = -1;
            try { visible = Convert.ToInt32((object)ws.Visible); } catch { }
            if (visible != -1) continue; // xlSheetVisible only

            try
            {
                foreach (dynamic lo in ws.ListObjects)
                {
                    var tableDraft = CaptureTable(lo, sheetName);
                    if (tableDraft != null) payload.Tables.Add(tableDraft);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[WorkbookDraft] Could not enumerate tables on '{sheetName}': {ex.Message}");
            }

            if (payload.ScheduleBoard == null &&
                ScheduleSheetNames.Any(n => string.Equals(n, sheetName.Trim(), StringComparison.OrdinalIgnoreCase)))
            {
                payload.ScheduleBoard = CaptureScheduleBoard(ws, sheetName);
            }
        }

        FolderRegistry.Data.EnsureExists();
        string path = GetDraftPath<T>();
        File.WriteAllText(path, JsonSerializer.Serialize(payload, JsonOptions));
        AppLogger.Info($"[WorkbookDraft] Saved local draft '{Path.GetFileName(path)}': {payload.Tables.Count} table(s), " +
                       $"{payload.Tables.Sum(t => t.Rows.Count)} row(s), board={(payload.ScheduleBoard != null ? payload.ScheduleBoard.SlotCount + " slot(s)" : "none")}, " +
                       $"{errorCount} error(s), {warningCount} warning(s).");
    }

    private static SheetTableDraft? CaptureTable(dynamic lo, string sheetName)
    {
        string tableName = string.Empty;
        try { tableName = Convert.ToString((object)lo.Name) ?? string.Empty; } catch { }
        try
        {
            dynamic header = lo.HeaderRowRange;
            int colCount = (int)header.Columns.Count;
            object[,] headerValues = TableMapper.GetRange2DValues(header, 1, colCount);

            var draft = new SheetTableDraft { SheetName = sheetName, TableName = tableName };
            for (int c = 1; c <= colCount; c++)
                draft.Headers.Add((Convert.ToString(headerValues[1, c], CultureInfo.InvariantCulture) ?? string.Empty).Trim());

            dynamic? body = null;
            try { body = lo.DataBodyRange; } catch { }
            if (body == null) return draft;

            int rowCount = (int)body.Rows.Count;
            object[,] values = TableMapper.GetRange2DValues(body, rowCount, colCount);
            for (int r = 1; r <= rowCount; r++)
            {
                var row = new List<DraftCellValue?>(colCount);
                for (int c = 1; c <= colCount; c++) row.Add(DraftCellValue.FromExcel(values[r, c]));
                draft.Rows.Add(row);
            }
            return draft;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Could not capture table '{tableName}' on '{sheetName}': {ex.Message}");
            return null;
        }
    }

    private static ScheduleBoardDraft? CaptureScheduleBoard(dynamic ws, string sheetName)
    {
        try
        {
            var board = new ScheduleBoardDraft { SheetName = sheetName };
            try { board.DateCell = DraftCellValue.FromExcel((object?)ws.Range["A2"].Value2); } catch { }
            try { board.DayCell = DraftCellValue.FromExcel((object?)ws.Range["B2"].Value2); } catch { }

            int lastCol = BoardFirstCol + ColsPerSlot - 1;
            try
            {
                dynamic used = ws.UsedRange;
                lastCol = (int)used.Column + (int)used.Columns.Count - 1;
            }
            catch { }
            board.SlotCount = Math.Min(MaxSlots, Math.Max(1, (int)Math.Ceiling((lastCol - 1) / 3.0)));

            int rows = BoardLastRow - BoardFirstRow + 1;
            int cols = board.SlotCount * ColsPerSlot;
            dynamic grid = ws.Range[ws.Cells[BoardFirstRow, BoardFirstCol], ws.Cells[BoardLastRow, 1 + cols]];
            object[,] values = TableMapper.GetRange2DValues(grid, rows, cols);
            for (int r = 1; r <= rows; r++)
            {
                var row = new List<DraftCellValue?>(cols);
                for (int c = 1; c <= cols; c++) row.Add(DraftCellValue.FromExcel(values[r, c]));
                board.Grid.Add(row);
            }
            return board;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Could not capture schedule board on '{sheetName}': {ex.Message}");
            return null;
        }
    }

    // =========================================================================
    // 4. RESTORE (bulk column writes, formulas preserved; macro thread)
    // =========================================================================
    public static void Restore<T>(object? wbObj, WorkbookDraftPayload draft)
    {
        if (wbObj == null || draft == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        object? origCalc = null;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }
        try { origCalc = (object)xlApp.Calculation; } catch { }

        try
        {
            try { xlApp.ScreenUpdating = false; } catch { }
            try { xlApp.EnableEvents = false; } catch { }
            try { xlApp.Calculation = -4135; } catch { } // xlCalculationManual

            dynamic wb = wbObj;

            if (draft.ScheduleBoard != null)
            {
                dynamic? boardWs = FindSheet(wb, draft.ScheduleBoard.SheetName);
                if (boardWs != null) RestoreScheduleBoard(boardWs, draft.ScheduleBoard);
                else AppLogger.Warn($"[WorkbookDraft] Board sheet '{draft.ScheduleBoard.SheetName}' not found; board skipped.");
            }

            int restored = 0;
            foreach (var tableDraft in draft.Tables)
            {
                dynamic? ws = FindSheet(wb, tableDraft.SheetName);
                if (ws == null)
                {
                    AppLogger.Warn($"[WorkbookDraft] Sheet '{tableDraft.SheetName}' not found; table '{tableDraft.TableName}' skipped.");
                    continue;
                }
                dynamic? table = TableMapper.ResolveTable(ws, tableDraft.TableName);
                if (table == null)
                {
                    AppLogger.Warn($"[WorkbookDraft] Table '{tableDraft.TableName}' not found on '{tableDraft.SheetName}'; skipped.");
                    continue;
                }
                if (RestoreTable(ws, table, tableDraft)) restored++;
            }

            UndoManager.ClearWorkbook(wbObj);
            AppLogger.Info($"[WorkbookDraft] Restored {restored}/{draft.Tables.Count} table(s) from local draft saved {draft.SavedUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}.");
        }
        finally
        {
            try { if (origCalc != null) xlApp.Calculation = origCalc; else xlApp.Calculation = -4105; } catch { } // xlCalculationAutomatic
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    private static dynamic? FindSheet(dynamic wb, string sheetName)
    {
        if (string.IsNullOrWhiteSpace(sheetName)) return null;
        try
        {
            foreach (dynamic ws in wb.Worksheets)
            {
                string name = string.Empty;
                try { name = Convert.ToString((object)ws.Name) ?? string.Empty; } catch { }
                if (string.Equals(name.Trim(), sheetName.Trim(), StringComparison.OrdinalIgnoreCase)) return ws;
            }
        }
        catch { }
        return null;
    }

    /// <summary>Bulk-reads <c>Formula</c> of a range as a 1-based grid (a single cell yields a 1x1 grid).</summary>
    private static string[,] ReadFormulaGrid(dynamic range, int rows, int cols)
    {
        var result = new string[rows + 1, cols + 1];
        object? raw = null;
        try { raw = range.Formula; } catch { }
        if (raw is object[,] matrix)
        {
            int lr = matrix.GetLowerBound(0), lc = matrix.GetLowerBound(1);
            for (int r = 1; r <= rows; r++)
                for (int c = 1; c <= cols; c++)
                    result[r, c] = Convert.ToString(matrix[r - 1 + lr, c - 1 + lc], CultureInfo.InvariantCulture) ?? string.Empty;
        }
        else if (raw != null && rows >= 1 && cols >= 1)
        {
            result[1, 1] = Convert.ToString(raw, CultureInfo.InvariantCulture) ?? string.Empty;
        }
        return result;
    }

    private static void RestoreScheduleBoard(dynamic ws, ScheduleBoardDraft board)
    {
        try
        {
            // 1. Top bar (never overwrite formulas)
            WriteCellUnlessFormula(ws.Range["A2"], board.DateCell);
            WriteCellUnlessFormula(ws.Range["B2"], board.DayCell);

            // 2. Extra slots: clone slot 0 formatting / borders / dropdowns so every slot matches the template
            int slotCount = Math.Min(MaxSlots, Math.Max(1, board.SlotCount));
            if (slotCount > 1)
            {
                dynamic slot0 = ws.Range[ws.Cells[BoardFirstRow, BoardFirstCol], ws.Cells[BoardLastRow, BoardFirstCol + ColsPerSlot - 1]];
                for (int s = 1; s < slotCount; s++)
                {
                    int destCol = BoardFirstCol + s * ColsPerSlot;
                    try { slot0.Copy(ws.Cells[BoardFirstRow, destCol]); }
                    catch (Exception ex) { AppLogger.Warn($"[WorkbookDraft] Could not clone board slot {s}: {ex.Message}"); }
                }
            }

            // 3. Grid values: one Formula read, then one write per row (formulas carried through untouched)
            int rows = BoardLastRow - BoardFirstRow + 1;
            int cols = slotCount * ColsPerSlot;
            dynamic gridRange = ws.Range[ws.Cells[BoardFirstRow, BoardFirstCol], ws.Cells[BoardLastRow, 1 + cols]];
            string[,] formulas = ReadFormulaGrid(gridRange, rows, cols);

            var full = new object?[rows, cols];
            for (int r = 0; r < rows; r++)
            {
                var draftRow = r < board.Grid.Count ? board.Grid[r] : null;
                for (int c = 0; c < cols; c++)
                {
                    string f = formulas[r + 1, c + 1] ?? string.Empty;
                    full[r, c] = f.StartsWith("=", StringComparison.Ordinal)
                        ? f
                        : (draftRow != null && c < draftRow.Count ? draftRow[c]?.ToExcel() : null);
                }
            }

            try
            {
                gridRange.Value2 = full;
            }
            catch
            {
                // Merged cells can reject a whole-block write: fall back to one bulk write per row.
                for (int r = 0; r < rows; r++)
                {
                    var rowArr = new object?[1, cols];
                    for (int c = 0; c < cols; c++) rowArr[0, c] = full[r, c];
                    try
                    {
                        dynamic rowRange = ws.Range[ws.Cells[BoardFirstRow + r, BoardFirstCol], ws.Cells[BoardFirstRow + r, 1 + cols]];
                        rowRange.Value2 = rowArr;
                    }
                    catch (Exception ex) { AppLogger.Warn($"[WorkbookDraft] Board row {BoardFirstRow + r} restore failed: {ex.Message}"); }
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Schedule board restore failed: {ex.Message}");
        }
    }

    private static void WriteCellUnlessFormula(dynamic cell, DraftCellValue? value)
    {
        try
        {
            string f = Convert.ToString((object?)cell.Formula, CultureInfo.InvariantCulture) ?? string.Empty;
            if (f.StartsWith("=", StringComparison.Ordinal)) return;
            cell.Value2 = value?.ToExcel();
        }
        catch { }
    }

    private static bool IsTimeRule(PropertyRuleDescriptor? rule) =>
        rule?.Property != null && rule.PropertyType == typeof(string) &&
        rule.Property.Name is "Time" or "TimeStart" or "TimeEnd";

    private static bool RestoreTable(dynamic ws, dynamic table, SheetTableDraft tableDraft)
    {
        string tableName = tableDraft.TableName;
        try { tableName = Convert.ToString((object)table.Name) ?? tableName; } catch { }

        bool hadTotals = false;
        try { hadTotals = (bool)table.ShowTotals; } catch { }

        try
        {
            // 1. Clear filters so every row is written / visible
            try
            {
                dynamic? af = table.AutoFilter;
                if (af != null && (bool)af.FilterMode) af.ShowAllData();
            }
            catch { }

            if (hadTotals) { try { table.ShowTotals = false; } catch { } }

            int rowCount = Math.Max(1, tableDraft.Rows.Count);
            SheetMapperHelper.ResizeTable(table, rowCount);

            dynamic header = table.HeaderRowRange;
            int headerRow = (int)header.Row;
            int startCol = (int)header.Column;
            int colCount = (int)table.ListColumns.Count;

            // 2. Live header -> column offset
            object[,] liveHeaders = TableMapper.GetRange2DValues(header, 1, colCount);
            var liveMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= colCount; c++)
            {
                string h = (Convert.ToString(liveHeaders[1, c], CultureInfo.InvariantCulture) ?? string.Empty).Trim();
                if (h.Length > 0 && !liveMap.ContainsKey(h)) liveMap[h] = c;
            }

            // 3. Archetype row formulas (one bulk read)
            dynamic archRow = ws.Range[ws.Cells[headerRow + 1, startCol], ws.Cells[headerRow + 1, startCol + colCount - 1]];
            string[,] archFormulas = ReadFormulaGrid(archRow, 1, colCount);

            // 4. One bulk write per mapped, non-formula column
            for (int d = 0; d < tableDraft.Headers.Count; d++)
            {
                string draftHeader = (tableDraft.Headers[d] ?? string.Empty).Trim();
                if (draftHeader.Length == 0 || !liveMap.TryGetValue(draftHeader, out int liveCol)) continue;

                var rule = RuleRegistry.GetRule(tableName, draftHeader);
                string archFormula = archFormulas[1, liveCol] ?? string.Empty;
                if (archFormula.StartsWith("=", StringComparison.Ordinal) && rule?.HasCSharpFormula != true) continue;

                var colArr = new object?[rowCount, 1];
                for (int r = 0; r < tableDraft.Rows.Count; r++)
                {
                    var row = tableDraft.Rows[r];
                    colArr[r, 0] = d < row.Count ? row[d]?.ToExcel() : null;
                }

                int absCol = startCol + liveCol - 1;
                dynamic colRange = ws.Range[ws.Cells[headerRow + 1, absCol], ws.Cells[headerRow + rowCount, absCol]];

                string archFmt = string.Empty;
                try { archFmt = Convert.ToString((object?)ws.Cells[headerRow + 1, absCol].NumberFormat, CultureInfo.InvariantCulture) ?? string.Empty; } catch { }

                bool isNric = string.Equals(rule?.Property?.Name, "NRIC", StringComparison.OrdinalIgnoreCase) ||
                              draftHeader.Equals("IC Number", StringComparison.OrdinalIgnoreCase);
                bool isTextColumn = !IsTimeRule(rule) &&
                                    (isNric || archFmt == "@" ||
                                     (rule != null && rule.PropertyType == typeof(string) && !RuleRegistry.ShouldApplyFormat(rule)) ||
                                     rule?.EffectiveFormatKind == ColumnFormatKind.Text);

                if (isTextColumn)
                {
                    // Keep leading zeros ("0001") and IC numbers as text.
                    try { colRange.NumberFormat = "@"; } catch { }
                }
                else if (RuleRegistry.ShouldApplyFormat(rule) &&
                         rule!.EffectiveFormatKind is not (ColumnFormatKind.OdometerDynamic or ColumnFormatKind.WageDynamic))
                {
                    try { colRange.NumberFormat = rule.ExcelNumberFormat; } catch { }
                }

                try { colRange.Value2 = colArr; }
                catch (Exception ex) { AppLogger.Warn($"[WorkbookDraft] Could not restore column '{draftHeader}' in '{tableName}': {ex.Message}"); }
            }

            StackedEntryGuard.CaptureTableSnapshot(ws, table);
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookDraft] Table '{tableName}' restore failed: {ex.Message}");
            return false;
        }
        finally
        {
            if (hadTotals) { try { table.ShowTotals = true; } catch { } }
        }
    }
}
