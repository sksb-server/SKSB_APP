namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookOpen.Sheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public static class SheetMapperHelper
{
    /// <summary>
    /// Resizes a ListObject table cleanly, logging its physical range and row count before and after.
    /// </summary>
    public static void ResizeTable(dynamic table, int dataRowCount)
    {
        if (table == null) return;
        try
        {
            dynamic ws = table.Parent;
            if (ws == null) return;

            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int finalColSpan = (int)table.ListColumns.Count;
            if (finalColSpan < 1) finalColSpan = 1;

            int endCol = startCol + finalColSpan - 1;
            dynamic? oldBody = table.DataBodyRange;
            int oldRowCount = oldBody != null ? (int)oldBody.Rows.Count : 0;

            int effectiveDataRows = Math.Max(1, dataRowCount);
            int endRow = headerRow + effectiveDataRows;

            // Skip COM Resize if dimensions are already identical
            if (oldRowCount == effectiveDataRows && (int)table.ListColumns.Count == finalColSpan)
            {
                return;
            }

            // Cache Total Row state; disable during resize to avoid COM collisions, then restore
            bool hadTotals = false;
            try { hadTotals = (bool)table.ShowTotals; } catch { }
            try
            {
                if (hadTotals) { try { table.ShowTotals = false; } catch { } }

                if (oldRowCount > effectiveDataRows)
                {
                    int clearStartRow = headerRow + effectiveDataRows + 1;
                    int clearEndRow = headerRow + oldRowCount;
                    if (clearStartRow <= clearEndRow)
                    {
                        dynamic staleRange = ws.Range[ws.Cells[clearStartRow, startCol], ws.Cells[clearEndRow, endCol]];
                        staleRange.Clear();
                    }
                }

                dynamic newRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[endRow, endCol]];
                try
                {
                    table.Resize(newRange);
                }
                catch (Exception exResize)
                {
                    System.Diagnostics.Debug.WriteLine($"[{table.Name}] table.Resize threw COM error: {exResize.Message}");
                }
            }
            finally
            {
                if (hadTotals)
                {
                    try { table.ShowTotals = true; } catch { }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"ResizeTable outer exception: {ex.Message}");
        }
    }

    public static void ClampAllTablesToActualData(dynamic wb)
    {
        if (wb == null) return;

        try
        {
            foreach (dynamic ws in wb.Sheets)
            {
                string sName = ws.Name?.ToString() ?? string.Empty;
                if (sName.Equals(SheetOperations.LookupsSheetName, StringComparison.OrdinalIgnoreCase) ||
                    sName.Equals(SheetOperations.DataSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int tableCount = 0;
                try { tableCount = (int)ws.ListObjects.Count; } catch { }
                if (tableCount == 0) continue;

                for (int t = 1; t <= tableCount; t++)
                {
                    dynamic table = ws.ListObjects[t];
                    dynamic? body = table.DataBodyRange;
                    if (body == null) continue;

                    int currentRows = (int)body.Rows.Count;
                    if (currentRows <= 5) continue;

                    dynamic col1Range = table.ListColumns[1].DataBodyRange;
                    object? rawCol1 = col1Range.Value2;
                    int lastNonEmptyRow = 0;

                    if (rawCol1 is object[,] mat)
                    {
                        int rMax = mat.GetUpperBound(0);
                        int cMin = mat.GetLowerBound(1);
                        for (int r = rMax; r >= 1; r--)
                        {
                            string val = CleanString(mat[r, cMin]);
                            if (!string.IsNullOrWhiteSpace(val))
                            {
                                lastNonEmptyRow = r;
                                break;
                            }
                        }
                    }

                    int targetRows = Math.Max(1, lastNonEmptyRow);
                    if (currentRows > targetRows + 1)
                    {
                        AppLogger.Warn($"[UNINTENDED EXPANSION DETECTED] Table '{table.Name}' was inflated to {currentRows} rows! Clamping down to {targetRows} data rows.");
                        ResizeTable(table, targetRows);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("ClampAllTablesToActualData error", ex);
        }
    }

    /// <summary>
    /// Synchronizes a distinct list of strings into an Excel ListObject (Table) and registers a Named Range.
    /// Operates strictly on the target table's parent worksheet via COM to prevent active-sheet bleed-through.
    /// </summary>
    public static bool SyncReferenceTable(
        dynamic workbook,
        string tableName,
        string namedRange,
        IEnumerable<string> items)
    {
        if (workbook == null || string.IsNullOrWhiteSpace(tableName) || items == null) return false;

        try
        {
            var distinctList = items!
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            dynamic? targetTable = RuleDataCache.FindTableInWorkbook(workbook, tableName);

            if (targetTable == null)
            {
                AppLogger.Warn($"[SyncReferenceTable] Table '{tableName}' not found in workbook.");
                return false;
            }

            dynamic sheet = targetTable.Parent;
            int headerRow = (int)targetTable.HeaderRowRange.Row;
            int col = (int)targetTable.HeaderRowRange.Column;

            int rowCount = distinctList.Count;
            int effectiveRows = Math.Max(1, rowCount);

            // 1. Resize table boundary directly
            ResizeTable(targetTable, effectiveRows);

            dynamic dataRange = sheet.Range[
                sheet.Cells[headerRow + 1, col],
                sheet.Cells[headerRow + effectiveRows, col]
            ];

            // 2. Write data without touching any other worksheet
            if (rowCount > 0)
            {
                object[,] data = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++)
                {
                    data[i, 0] = distinctList[i];
                }

                try { dataRange.NumberFormat = "@"; } catch { }
                dataRange.Value2 = data;

                // 3. Register Excel Defined Name if requested
                if (!string.IsNullOrWhiteSpace(namedRange))
                {
                    try { XlCall.Excel(XlCall.xlcDeleteName, namedRange); } catch { }
                    string sName = (string)sheet.Name;
                    int rStart = headerRow + 1;
                    int rEnd = headerRow + rowCount;
                    string refersTo = $"='{sName}'!R{rStart}C{col}:R{rEnd}C{col}";
                    XlCall.Excel(XlCall.xlcDefineName, namedRange, refersTo, 3);
                }
            }
            else
            {
                dataRange.Value2 = null;
            }

            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[SyncReferenceTable] Failed on '{tableName}'", ex);
            return false;
        }
    }

    // =========================================================================
    // SHARED MANAGER MAPPING UTILITIES
    // =========================================================================

    /// <summary>
    /// Resolves a worksheet inside <paramref name="wbObj"/> by its alias list (case-insensitive),
    /// returning null when none of the aliases exist. Never activates or selects the sheet.
    /// </summary>
    public static dynamic? ResolveSheetSafely(object wbObj, string[] aliases, string defaultName)
    {
        if (wbObj == null) return null;
        dynamic wb = wbObj;
        aliases ??= Array.Empty<string>();

        string resolvedName = SheetOperations.ResolveSheetName(aliases, defaultName);
        try
        {
            dynamic s = wb.Sheets[resolvedName];
            if (s != null) return s;
        }
        catch { }

        foreach (dynamic ws in wb.Sheets)
        {
            string name = ws.Name?.ToString() ?? string.Empty;
            if (aliases.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase)))
            {
                return ws;
            }
        }
        return null;
    }

    /// <summary>
    /// DEBUG-only audit listing the location and dimensions of each inspected table after a populate pass.
    /// Written to the runtime log via <see cref="AppLogger"/> (no popup); compiles to a no-op in Release builds.
    /// </summary>
    public static void ShowPopulateDiagnostics(object wbObj, string auditTitle, string masterLabel, int masterCount, params string[] inspectTables)
    {
#if DEBUG
        try
        {
            if (wbObj == null) return;

            dynamic wb = wbObj;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"=== {auditTitle.ToUpperInvariant()} ===");
            sb.AppendLine($"Total {masterLabel} Loaded: {masterCount}");
            sb.AppendLine();

            foreach (var tName in inspectTables ?? Array.Empty<string>())
            {
                dynamic? tbl = null;
                string sName = "Unknown";
                foreach (dynamic ws in wb.Sheets)
                {
                    try
                    {
                        tbl = ws.ListObjects[tName];
                        if (tbl != null) { sName = ws.Name?.ToString() ?? "Unknown"; break; }
                    }
                    catch { }
                }

                if (tbl != null)
                {
                    int cols = (int)tbl.ListColumns.Count;
                    int rows = tbl.DataBodyRange != null ? (int)tbl.DataBodyRange.Rows.Count : 0;
                    int startCol = (int)tbl.HeaderRowRange.Column;
                    int startRow = (int)tbl.HeaderRowRange.Row;
                    sb.AppendLine($"  [{tName}] on Sheet '{sName}':");
                    sb.AppendLine($"      Dimensions : {rows} data row(s) × {cols} column(s)");
                    sb.AppendLine($"      Coordinates: Row {startRow}, Columns {startCol} to {startCol + cols - 1}");
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine($"  [{tName}]: NOT FOUND on any worksheet");
                    sb.AppendLine();
                }
            }

            AppLogger.Info($"[{auditTitle}]\n{sb}");
        }
        catch { }
#endif
    }

    internal static bool IsComError(object? val) =>
        val == null ||
        (val is int i && i < -2000000000) ||
        (val is long l && l < -2000000000) ||
        val.ToString()?.StartsWith("#") == true;

    /// <summary>
    /// Reads the archetype (first data) row's formulas in ONE COM call. Returns a 1-based array (index 0 unused) of
    /// length <paramref name="colSpan"/> + 1; unreadable cells are empty strings.
    /// </summary>
    public static string[] ReadArchetypeFormulas(dynamic worksheet, int archRow, int startCol, int colSpan)
    {
        var formulas = new string[colSpan + 1];
        Array.Fill(formulas, string.Empty);
        try
        {
            dynamic archRange = worksheet.Range[worksheet.Cells[archRow, startCol], worksheet.Cells[archRow, startCol + colSpan - 1]];
            object? raw = archRange.Formula;
            if (raw is object[,] mat)
            {
                for (int c = 1; c <= colSpan; c++)
                    formulas[c] = mat[1, c]?.ToString() ?? string.Empty;
            }
            else if (raw is string single)
            {
                formulas[1] = single;
            }
        }
        catch { }
        return formulas;
    }

    /// <summary>Row-input form: the em-dash null placeholder counts as empty (no user input).</summary>
    public static string CleanCellText(object? val) => CleanString(val, treatDashAsEmpty: true);

    /// <summary>
    /// Single cell-text normaliser: COM errors / "#..." / "System.__ComObject" become empty, whole doubles lose the
    /// ".0", the rest is trimmed. <paramref name="treatDashAsEmpty"/> also maps the em-dash null placeholder to empty.
    /// </summary>
    public static string CleanString(object? val, bool treatDashAsEmpty = false)
    {
        if (val == null || IsComError(val)) return string.Empty;
        if (val is double d && Math.Abs(d - Math.Round(d)) < 0.00001 && d < 1e15)
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (s.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        if (treatDashAsEmpty && s == "—") return string.Empty;
        if (long.TryParse(s, out long num) && num < -2000000000) return string.Empty;
        return s;
    }

}