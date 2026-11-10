namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookOpen.Sheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public static class SheetMapperHelper
{
    private static readonly string[] UniversalDateFormats =
    {
        "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyy-MM-dd"
    };

    /// <summary>
    /// Resizes a ListObject table cleanly, logging its physical range and row count before and after.
    /// </summary>
    public static void ResizeTable(dynamic table, int dataRowCount, int? explicitColSpan = null)
    {
        if (table == null) return;
        try
        {
            dynamic ws = table.Parent;
            if (ws == null) return;

            // Lift any legacy sheet protection; it is never re-applied.
            try
            {
                if ((bool)ws.ProtectContents) ws.Unprotect("SKSB_Internal_Guard");
            }
            catch { }

            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int finalColSpan = explicitColSpan ?? (int)table.ListColumns.Count;
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

    public static void ResizeTable(
        string sheetName,
        int dataRowCount,
        int columnSpan,
        [CallerMemberName] string callerMember = "",
        [CallerFilePath] string callerFile = "",
        [CallerLineNumber] int callerLine = 0)
    {
        string sourceFile = System.IO.Path.GetFileName(callerFile);
        AppLogger.Warn($"[LEGACY OVERLOAD] ResizeTable('{sheetName}', targetRows={dataRowCount}, span={columnSpan}) called by [{sourceFile}:{callerLine} {callerMember}]");

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = xlApp.ActiveWorkbook.Sheets[sheetName];
            if (ws == null || ws.ListObjects.Count == 0) return;
            dynamic table = ws.ListObjects[1];
            ResizeTable(table, dataRowCount, columnSpan, callerMember, callerFile, callerLine);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"Legacy ResizeTable('{sheetName}') error", ex);
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
            var distinctList = items
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x)
                .ToList();

            dynamic? targetTable = null;
            foreach (dynamic ws in workbook.Sheets)
            {
                try
                {
                    targetTable = ws.ListObjects[tableName];
                    if (targetTable != null) break;
                }
                catch { }
            }

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

    public static List<string> ExtractDistinctColumnValues(dynamic workbook, string tableName, int scanRowLimit = 100)
    {
        var distinctValues = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        dynamic? targetTable = null;

        foreach (dynamic ws in workbook.Sheets)
        {
            try
            {
                targetTable = ws.ListObjects[tableName];
                if (targetTable != null) break;
            }
            catch { }
        }

        if (targetTable == null) return distinctValues.ToList();

        dynamic sheet = targetTable.Parent;
        dynamic headerRange = targetTable.HeaderRowRange;
        int headerRow = (int)headerRange.Row;
        int col = (int)headerRange.Column;

        try
        {
            dynamic scanRange = sheet.Range[sheet.Cells[headerRow + 1, col], sheet.Cells[headerRow + scanRowLimit, col]];
            object[,] colGrid = scanRange.Value2;

            if (colGrid != null)
            {
                int rCount = colGrid.GetLength(0);
                for (int r = 1; r <= rCount; r++)
                {
                    string val = CleanString(colGrid[r, 1]);
                    if (!string.IsNullOrWhiteSpace(val) &&
                        !val.StartsWith("#") &&
                        !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                    {
                        distinctValues.Add(val);
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SheetMapperHelper.ExtractDistinctColumnValues Error] {ex.Message}");
        }

        return distinctValues.OrderBy(x => x).ToList();
    }

    internal static bool IsComError(object? val) =>
        val == null ||
        (val is int i && i < -2000000000) ||
        (val is long l && l < -2000000000) ||
        val.ToString()?.StartsWith("#") == true;

    public static string CleanString(object? val)
    {
        if (val == null || IsComError(val)) return string.Empty;
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (long.TryParse(s, out long num) && num < -2000000000) return string.Empty;
        return s;
    }

    public static DateOnly ParseDate(object? cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        string clean = CleanString(cell);
        if (string.IsNullOrWhiteSpace(clean)) return default;
        if (double.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedD) && parsedD > 30000 && parsedD < 80000)
            return DateOnly.FromDateTime(DateTime.FromOADate(parsedD));
        if (DateOnly.TryParseExact(clean, UniversalDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly dOnly))
            return dOnly;
        if (DateTime.TryParse(clean, new CultureInfo("en-GB"), DateTimeStyles.None, out DateTime dt))
            return DateOnly.FromDateTime(dt);
        return default;
    }

    public static string FormatDate(DateOnly date) =>
        date == default ? string.Empty : date.ToString("dd/MM/yyyy");
}