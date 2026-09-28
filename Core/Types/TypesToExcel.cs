namespace SKSB_App.Core.Services;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookManager;

public static class TypesToExcel
{
    // Toggle false so Lookups stays hidden from the bottom tab strip
    public const bool KeepLookupsSheetVisible = false;

    public static void Sync(params Type[] enumTypes)
    {
        if (enumTypes == null || enumTypes.Length == 0) return;

        var mappings = new Dictionary<string, string[]>();

        foreach (var type in enumTypes)
        {
            if (!type.IsEnum) continue;

            string listName = $"{type.Name}List";
            string[] values = Enum.GetNames(type)
                                  .Where(n => !string.Equals(n, "None", StringComparison.OrdinalIgnoreCase))
                                  .ToArray();

            mappings[listName] = values;
        }

        SyncDictionary(mappings);
    }

    public static void SyncDictionary(Dictionary<string, string[]> enumMappings)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                XlCall.Excel(XlCall.xlcEcho, false);

                const string targetSheet = "Lookups";

                // 1. Remember currently active sheet to restore focus cleanly
                object activeDoc = null;
                try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
                string originalSheet = activeDoc?.ToString() ?? string.Empty;

                // 2. Ensure Lookups exists WITHOUT spawning rogue sheets
                EnsureLookupsSheetExists(targetSheet);

                // 3. Unhide and Select Lookups so SetValue() executes natively without XlCallException
                try { XlCall.Excel(XlCall.xlcWorkbookUnhide, targetSheet); } catch { }
                try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetSheet }); } catch { }

                int colZero = 0;

                foreach (var (name, values) in enumMappings)
                {
                    int rowCount = values.Length;

                    // Clear previous rows (0 to 100)
                    var wipeRef = new ExcelReference(0, 100, colZero, colZero);
                    object[,] clearGrid = new object[101, 1];
                    for (int r = 0; r <= 100; r++) clearGrid[r, 0] = string.Empty;
                    wipeRef.SetValue(clearGrid);

                    // Write Column Header
                    var headerRef = new ExcelReference(0, 0, colZero, colZero);
                    headerRef.SetValue(new object[,] { { name } });

                    // Write Data Rows and define global Name
                    if (rowCount > 0)
                    {
                        var valRef = new ExcelReference(1, rowCount, colZero, colZero);
                        object[,] cellData = new object[rowCount, 1];
                        for (int i = 0; i < rowCount; i++) cellData[i, 0] = values[i];
                        valRef.SetValue(cellData);

                        try { XlCall.Excel(XlCall.xlcDeleteName, name); } catch { }

                        int r1c1StartRow = 2;
                        int r1c1EndRow = rowCount + 1;
                        int r1c1Col = colZero + 1;
                        string refersTo = $"='{targetSheet}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";

                        XlCall.Excel(XlCall.xlcDefineName, name, refersTo, 3);
                    }

                    colZero += 2; // Spacer column
                }

                // 4. Hide the Lookups sheet if configured
                if (!KeepLookupsSheetVisible)
                {
                    try
                    {
                        // Lookups is active -> 0-arg xlcWorkbookHide hides it cleanly
                        XlCall.Excel(XlCall.xlcWorkbookHide);
                    }
                    catch { }
                }

                // 5. Restore original active sheet focus
                if (!string.IsNullOrWhiteSpace(originalSheet))
                {
                    string cleanOriginal = originalSheet.Contains(']')
                        ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                        : originalSheet;

                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TypesToExcel Error] {ex.Message}");
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, true);
            }
        });
    }

    private static void EnsureLookupsSheetExists(string targetSheet)
    {
        // 1. Direct handle check using C-API textref pointer
        IntPtr existingId = WorkbookSession.GetSheetId(targetSheet);
        if (existingId != IntPtr.Zero)
        {
            return; // Sheet already exists!
        }

        // 2. Scan both dimensions of the workbook table of contents
        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int dim0 = sheetArray.GetLength(0);
                int dim1 = sheetArray.GetLength(1);

                for (int r = 0; r < dim0; r++)
                {
                    for (int c = 0; c < dim1; c++)
                    {
                        string raw = sheetArray[r, c]?.ToString() ?? string.Empty;
                        int bIdx = raw.LastIndexOf(']');
                        string clean = bIdx >= 0 ? raw.Substring(bIdx + 1) : raw;

                        if (string.Equals(clean, targetSheet, StringComparison.OrdinalIgnoreCase))
                        {
                            return; // Sheet exists!
                        }
                    }
                }
            }
        }
        catch { }

        // 3. Genuinely does not exist: insert and rename once
        try
        {
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);

            object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string fullName = docInfo?.ToString() ?? string.Empty;
            string tempName = fullName.Contains(']')
                ? fullName.Substring(fullName.LastIndexOf(']') + 1)
                : fullName;

            XlCall.Excel(XlCall.xlcWorkbookName, tempName, targetSheet);
        }
        catch { }
    }
}