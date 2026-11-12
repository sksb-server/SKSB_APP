namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;

/// <summary>
/// Table self-healing WITHOUT native sheet protection (worksheet.Protect is forbidden: it breaks
/// ListObject sorting/filtering/row insertion). Read-only cells are authored as =FALSE Custom Data
/// Validation in the templates; HealFormulas/HealHeaders restore anything that slips through
/// (paste, delete, undo).
/// </summary>
public static class SheetProtections
{
    /// <summary>Legacy pass-through kept for call-site compatibility. No protection is lifted or re-applied.</summary>
    public static void ExecuteWithBypass(dynamic worksheet, Action action) => action?.Invoke();

    // =========================================================================
    // HEALING
    // =========================================================================

    /// <summary>
    /// For every column whose archetype cell (first body row) holds a formula, ensures all body cells
    /// still hold that column formula, restoring cells that were cleared or overwritten.
    /// </summary>
    public static void HealFormulas(dynamic worksheet, dynamic table)
    {
        if (table == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic bodyRange = table.DataBodyRange;
            if (bodyRange == null) return;

            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            int rows = (int)bodyRange.Rows.Count;

            for (int c = 1; c <= colCount; c++)
            {
                try
                {
                    dynamic archetype = worksheet.Cells[headerRow + 1, startCol + c - 1];
                    string formula = archetype.Formula?.ToString() ?? string.Empty;
                    if (!formula.StartsWith("=")) continue;

                    dynamic colRange = table.ListColumns[c].DataBodyRange;
                    if (colRange == null) continue;

                    bool dirty = false;
                    if (rows > 1)
                    {
                        object? raw = colRange.Formula;
                        if (raw is object[,] arr)
                        {
                            int lo = arr.GetLowerBound(0), hi = arr.GetUpperBound(0);
                            for (int r = lo; r <= hi; r++)
                            {
                                string f = arr[r, arr.GetLowerBound(1)]?.ToString() ?? string.Empty;
                                if (!f.StartsWith("=")) { dirty = true; break; }
                            }
                        }
                    }

                    if (dirty) colRange.Formula = formula;
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[HealFormulas Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    /// <summary>Ensures table header cells match the expected column names (by position).</summary>
    public static void HealHeaders(dynamic table, IEnumerable<string> expectedHeaders)
    {
        if (table == null || expectedHeaders == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            int c = 0;
            int colCount = (int)table.ListColumns.Count;
            foreach (var expected in expectedHeaders)
            {
                c++;
                if (c > colCount) break;
                try
                {
                    dynamic col = table.ListColumns[c];
                    string actual = col.Name?.ToString() ?? string.Empty;
                    if (!string.Equals(actual, expected, StringComparison.Ordinal))
                        table.HeaderRowRange.Cells[1, c].Value2 = expected;
                }
                catch { }
            }
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }
}
