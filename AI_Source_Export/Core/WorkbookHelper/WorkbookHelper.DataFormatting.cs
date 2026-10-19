namespace SKSB_App.Core.WorkbookHelper;

using System;
using ExcelDna.Integration;

public static class DataFormatting
{
    /// <summary>
    /// Resizes any Excel ListObject (Table) to match data row count, guaranteeing at least 1 empty row beneath headers.
    /// </summary>
    public static void ResizeTable(string sheetName, int dataRowCount, int columnSpan)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = xlApp.ActiveWorkbook.Sheets[sheetName];
            if (ws == null || ws.ListObjects.Count == 0) return;

            dynamic table = ws.ListObjects[1];
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int endCol = startCol + columnSpan - 1;

            int effectiveDataRows = Math.Max(1, dataRowCount);
            int endRow = headerRow + effectiveDataRows;

            dynamic newRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[endRow, endCol]];
            table.Resize(newRange);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ExcelDataFormatting.ResizeTable] Error on '{sheetName}': {ex.Message}");
        }
    }

    /// <summary>
    /// Converts Excel COM cell objects safely into trimmed strings without throwing COM exceptions.
    /// </summary>
    public static string CleanString(object? cell) =>
        (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            ? string.Empty
            : cell.ToString()?.Trim() ?? string.Empty;

    /// <summary>
    /// Safely parses an Excel cell value into DateOnly, supporting both OADate doubles and formatted date strings.
    /// </summary>
    public static DateOnly ParseDate(object? cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return default;
    }
}