namespace SKSB_App.Core.ServerManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookHelper;

public static class ServerLookups
{
    public const bool KeepLookupsSheetVisible = false;
    private const string TargetSheetName = SheetOps.LookupsSheetName;

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
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v)
            .ToArray();

        string listName = namePrefix.EndsWith("List", StringComparison.OrdinalIgnoreCase)
            ? namePrefix
            : $"{namePrefix}List";

        SyncValues(listName, values);
    }

    /// <summary>
    /// Writes raw values to the Lookups sheet using safe active sheet references without UI flickering.
    /// </summary>
    public static void SyncValues(string listName, string[] values)
    {
        if (string.IsNullOrWhiteSpace(listName) || values == null || values.Length == 0) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                XlCall.Excel(XlCall.xlcEcho, false);

                // 1. Remember original active sheet to restore focus cleanly
                object? activeDoc = null;
                try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
                string originalSheet = activeDoc?.ToString() ?? string.Empty;

                // 2. Ensure Lookups sheet exists and is made active for writing
                EnsureLookupsSheetExists();
                MakeSheetVisibleAndActive(TargetSheetName);

                // 3. Scan columns (0, 2, 4, 6...) looking for matching header or first blank slot
                int targetCol = -1;
                int maxColsToScan = 200;

                for (int c = 0; c < maxColsToScan; c += 2)
                {
                    var headerCell = new ExcelReference(0, 0, c, c);
                    object? cellVal = headerCell.GetValue();
                    string existingHeader = cellVal?.ToString()?.Trim() ?? string.Empty;

                    if (string.Equals(existingHeader, listName, StringComparison.OrdinalIgnoreCase))
                    {
                        targetCol = c;
                        break;
                    }

                    if (string.IsNullOrEmpty(existingHeader) && targetCol == -1)
                    {
                        targetCol = c;
                    }
                }

                if (targetCol == -1)
                {
                    targetCol = 0;
                }

                // 4. Wipe rows (0 to 2000) for this column to erase stale entries
                var wipeRef = new ExcelReference(0, 2000, targetCol, targetCol);
                object[,] clearGrid = new object[2001, 1];
                for (int r = 0; r <= 2000; r++) clearGrid[r, 0] = string.Empty;
                wipeRef.SetValue(clearGrid);

                // 5. Write Column Header (Row 0)
                var headerRef = new ExcelReference(0, 0, targetCol, targetCol);
                headerRef.SetValue(new object[,] { { listName } });

                // 6. Write Data Rows (Row 1 to values.Length)
                int rowCount = values.Length;
                var valRef = new ExcelReference(1, rowCount, targetCol, targetCol);
                object[,] cellData = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++)
                {
                    cellData[i, 0] = values[i];
                }
                valRef.SetValue(cellData);

                // 7. Register Excel Defined Name: ='Lookups'!R2C[col]:R[rowCount+1]C[col]
                try { XlCall.Excel(XlCall.xlcDeleteName, listName); } catch { }

                int r1c1StartRow = 2;
                int r1c1EndRow = rowCount + 1;
                int r1c1Col = targetCol + 1;
                string refersTo = $"='{TargetSheetName}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";
                XlCall.Excel(XlCall.xlcDefineName, listName, refersTo, 3);

                // 8. Re-position Lookups behind operational sheets
                SheetOps.MoveAuxiliarySheetsToEnd();

                // 9. Hide Lookups sheet if NOT in Template Edit Mode
                if (!KeepLookupsSheetVisible && !SheetOps.IsTemplateEditMode)
                {
                    try
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { TargetSheetName });
                        XlCall.Excel(XlCall.xlcWorkbookHide);
                    }
                    catch { }
                }

                // 10. Restore focus back to original sheet
                if (!string.IsNullOrWhiteSpace(originalSheet))
                {
                    string cleanOriginal = originalSheet.Contains(']')
                        ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                        : originalSheet;

                    if (!cleanOriginal.Equals(TargetSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                    }
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

    private static void MakeSheetVisibleAndActive(string sheetName)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = xlApp.ActiveWorkbook.Sheets[sheetName];
            if (ws != null)
            {
                ws.Visible = -1; // xlSheetVisible
                ws.Activate();
                return;
            }
        }
        catch { }

        try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { sheetName }); } catch { }
    }

    private static void EnsureLookupsSheetExists()
    {
        // 1. Check if Lookups exists across all sheets (visible or hidden) without GetSheetId
        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                foreach (object item in sheetArray)
                {
                    string raw = item?.ToString() ?? string.Empty;
                    int bIdx = raw.LastIndexOf(']');
                    string clean = bIdx >= 0 ? raw.Substring(bIdx + 1) : raw;

                    if (string.Equals(clean, TargetSheetName, StringComparison.OrdinalIgnoreCase))
                        return; // Sheet already exists
                }
            }
        }
        catch { }

        // 2. Only insert if the sheet does not exist at all
        try
        {
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);

            object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string fullName = docInfo?.ToString() ?? string.Empty;
            string tempSheetName = fullName.Contains(']')
                ? fullName.Substring(fullName.LastIndexOf(']') + 1)
                : fullName;

            XlCall.Excel(XlCall.xlcWorkbookName, tempSheetName, TargetSheetName);
        }
        catch { }
    }
}