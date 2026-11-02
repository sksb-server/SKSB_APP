namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using ExcelDna.Integration;
using global::SKSB_App.Core.Diagnostics;
using global::SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookOpen.Sheet;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;

public static class SheetMapperHelper
{
    private static readonly string[] UniversalDateFormats =
    {
        "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyy-MM-dd"
    };

    public static void ResizeTable(
        dynamic table,
        int dataRowCount,
        int? explicitColSpan = null,
        [CallerMemberName] string callerMember = "",
        [CallerFilePath] string callerFile = "",
        [CallerLineNumber] int callerLine = 0)
    {
        if (table == null) return;

        string tableName = "Table";
        try { tableName = table.Name?.ToString() ?? "Table"; } catch { }

        string sourceFile = System.IO.Path.GetFileName(callerFile);
        AppLogger.Info($"ResizeTable({tableName}, targetRows={dataRowCount}) requested by [{sourceFile}:{callerLine} {callerMember}]");

        if (dataRowCount >= 500)
        {
            AppLogger.Warn($"[SUSPICIOUS RESIZE] Large row count ({dataRowCount}) requested for '{tableName}' by [{sourceFile}:{callerLine} {callerMember}]!\nStack: {new StackTrace(1, true)}");
        }

        using var _ = AppLogger.Scope($"ResizeTable({tableName}, targetRows={dataRowCount})");

        try
        {
            dynamic ws = table.Parent;
            if (ws == null) return;

            bool wasProtected = false;
            try
            {
                wasProtected = (bool)ws.ProtectContents;
                if (wasProtected) ws.Unprotect("SKSB_Internal_Guard");
            }
            catch { }

            try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }

            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;

            int scannedColSpan = 0;
            while (true)
            {
                int c = startCol + scannedColSpan;
                dynamic cell = ws.Cells[headerRow, c];
                string val = cell.Value2?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrEmpty(val)) break;

                dynamic? otherTable = null;
                try { otherTable = cell.ListObject; } catch { }
                if (otherTable != null)
                {
                    string otherName = otherTable.Name?.ToString() ?? string.Empty;
                    if (!string.Equals(otherName, tableName, StringComparison.OrdinalIgnoreCase))
                    {
                        break;
                    }
                }

                scannedColSpan++;
            }

            int finalColSpan = explicitColSpan ?? Math.Max((int)table.ListColumns.Count, scannedColSpan);
            if (finalColSpan < 1) finalColSpan = 1;

            int endCol = startCol + finalColSpan - 1;
            dynamic? oldBody = table.DataBodyRange;
            int oldRowCount = oldBody != null ? (int)oldBody.Rows.Count : 0;

            int effectiveDataRows = Math.Max(1, dataRowCount);
            int endRow = headerRow + effectiveDataRows;

            AppLogger.Info($"[{tableName}] HeaderRow={headerRow}, StartCol={startCol}, ColSpan={finalColSpan}, OldRows={oldRowCount} -> NewRows={effectiveDataRows}");

            if (oldRowCount > effectiveDataRows)
            {
                int deleteStartRow = endRow + 1;
                int deleteEndRow = headerRow + oldRowCount;
                try
                {
                    dynamic excessTableCells = ws.Range[ws.Cells[deleteStartRow, startCol], ws.Cells[deleteEndRow, endCol]];
                    excessTableCells.Delete(-4162); // xlShiftUp
                    AppLogger.Info($"[{tableName}] Deleted excess rows {deleteStartRow}..{deleteEndRow} via xlShiftUp.");
                }
                catch (Exception exDel)
                {
                    AppLogger.Warn($"[{tableName}] excessTableCells.Delete warning: {exDel.Message}");
                }
            }

            dynamic newRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[endRow, endCol]];
            try
            {
                table.Resize(newRange);
            }
            catch (Exception exResize)
            {
                AppLogger.Error($"[{tableName}] table.Resize threw COM error", exResize);
            }

            dynamic? postBody = table.DataBodyRange;
            int postRows = postBody != null ? (int)postBody.Rows.Count : 0;
            AppLogger.Info($"[{tableName}] Final DataBodyRange rows = {postRows}");

            if (wasProtected)
            {
                try { SheetProtections.ApplyStyleDrivenProtection(ws); } catch { }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"ResizeTable({tableName}) outer exception", ex);
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
                            if (!string.IsNullOrWhiteSpace(val) && !val.Equals("Total", StringComparison.OrdinalIgnoreCase))
                            {
                                lastNonEmptyRow = r;
                                break;
                            }
                        }
                    }

                    int targetRows = Math.Max(1, lastNonEmptyRow);
                    if (currentRows > targetRows + 1)
                    {
                        AppLogger.Info($"[ClampAllTables] Clamping '{table.Name}' from {currentRows} rows down to {targetRows} rows.");
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

    public static bool SyncReferenceTable(
        dynamic workbook,
        string tableName,
        string namedRange,
        IEnumerable<string> items)
    {
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

            if (targetTable == null) return false;

            dynamic sheet = targetTable.Parent;
            dynamic headerRange = targetTable.HeaderRowRange;
            int headerRow = (int)headerRange.Row;
            int col = (int)headerRange.Column;

            int clearStartRow = headerRow + 1;
            var clearRef = new ExcelReference(clearStartRow - 1, clearStartRow + 999, col - 1, col - 1);
            object[,] emptyGrid = new object[1001, 1];
            for (int r = 0; r <= 1000; r++) emptyGrid[r, 0] = string.Empty;
            clearRef.SetValue(emptyGrid);

            int rowCount = distinctList.Count;
            int effectiveRows = Math.Max(1, rowCount);
            int endRow = headerRow + effectiveRows;

            dynamic newRange = sheet.Range[sheet.Cells[headerRow, col], sheet.Cells[endRow, col]];
            targetTable.Resize(newRange);

            if (rowCount > 0)
            {
                object[,] data = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++)
                {
                    data[i, 0] = distinctList[i];
                }
                var writeRef = new ExcelReference(clearStartRow - 1, clearStartRow + rowCount - 2, col - 1, col - 1);
                writeRef.SetValue(data);

                if (!string.IsNullOrWhiteSpace(namedRange))
                {
                    try { XlCall.Excel(XlCall.xlcDeleteName, namedRange); } catch { }
                    string sName = (string)sheet.Name;
                    string refersTo = $"='{sName}'!R{clearStartRow}C{col}:R{clearStartRow + rowCount - 1}C{col}";
                    XlCall.Excel(XlCall.xlcDefineName, namedRange, refersTo, 3);
                }
            }

            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SheetMapperHelper.SyncReferenceTable Error] {ex.Message}");
            return false;
        }
    }
}