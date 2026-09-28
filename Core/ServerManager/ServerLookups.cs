namespace SKSB_App.Core.ServerManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookManager;

public static class ServerLookups
{
    public const bool KeepLookupsSheetVisible = false;
    private const string TargetSheetName = "Lookups";

    /// <summary>
    /// Projects an in-memory collection loaded from the server/disk down into the Lookups
    /// worksheet and registers an Excel Defined Name ([namePrefix]List).
    /// </summary>
    public static void SyncValues<TSource, TProp>(string namePrefix, IEnumerable<TSource>? items, Func<TSource, TProp> selector)
    {
        if (items == null) return;

        string[] values = items
            .Select(selector)
            .Where(val => val != null)
            .Select(val => val!.ToString()!.Trim())
            .Where(val => !string.IsNullOrEmpty(val) && val != "0")
            .Distinct()
            .ToArray();

        SyncValues($"{namePrefix}List", values);
    }

    /// <summary>
    /// Writes raw values to the Lookups sheet using WorkbookSession's SheetId without UI flickering.
    /// </summary>
    public static void SyncValues(string listName, string[] values)
    {
        if (values == null || values.Length == 0) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                XlCall.Excel(XlCall.xlcEcho, false);

                EnsureLookupsSheetExists();

                IntPtr sheetId = WorkbookSession.GetSheetId(TargetSheetName);
                if (sheetId == IntPtr.Zero) return;

                int targetCol = 0;
                for (int c = 0; c < 100; c += 2)
                {
                    var cellRef = new ExcelReference(0, 0, c, c, sheetId);
                    object val = cellRef.GetValue();
                    string header = val?.ToString()?.Trim() ?? string.Empty;

                    if (string.IsNullOrEmpty(header) || string.Equals(header, listName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetCol = c;
                        break;
                    }
                }

                var wipeRef = new ExcelReference(0, 100, targetCol, targetCol, sheetId);
                object[,] clearGrid = new object[101, 1];
                for (int r = 0; r <= 100; r++) clearGrid[r, 0] = string.Empty;
                wipeRef.SetValue(clearGrid);

                var headerRef = new ExcelReference(0, 0, targetCol, targetCol, sheetId);
                headerRef.SetValue(new object[,] { { listName } });

                int rowCount = values.Length;
                var valRef = new ExcelReference(1, rowCount, targetCol, targetCol, sheetId);
                object[,] cellData = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++) cellData[i, 0] = values[i];
                valRef.SetValue(cellData);

                try { XlCall.Excel(XlCall.xlcDeleteName, listName); } catch { }
                int r1c1StartRow = 2;
                int r1c1EndRow = rowCount + 1;
                int r1c1Col = targetCol + 1;
                string refersTo = $"='{TargetSheetName}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";
                XlCall.Excel(XlCall.xlcDefineName, listName, refersTo, 3);

                if (!KeepLookupsSheetVisible)
                {
                    try { XlCall.Excel(XlCall.xlcWorkbookHide, TargetSheetName); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[ServerLookups Error] {ex.Message}");
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, true);
            }
        });
    }

    private static void EnsureLookupsSheetExists()
    {
        if (WorkbookSession.GetSheetId(TargetSheetName) != IntPtr.Zero)
            return;

        try
        {
            XlCall.Excel(XlCall.xlcWorkbookUnhide, TargetSheetName);
            return;
        }
        catch { }

        XlCall.Excel(XlCall.xlcWorkbookInsert, 1);
        object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
        string fullName = docInfo?.ToString() ?? string.Empty;
        string tempSheetName = fullName.Contains(']') ? fullName.Substring(fullName.LastIndexOf(']') + 1) : fullName;
        XlCall.Excel(XlCall.xlcWorkbookName, tempSheetName, TargetSheetName);
    }
}