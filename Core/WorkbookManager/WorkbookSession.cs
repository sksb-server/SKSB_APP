namespace SKSB_App.Core.WorkbookManager;

using System;
using ExcelDna.Integration;

/// <summary>
/// Low-level C-API utility operations for workbook selection, sheet purging, and state inspection.
/// </summary>
public static class WorkbookSession
{
    public static bool IsMasterTemplateFile()
    {
        try
        {
            object docNameObj = XlCall.Excel(XlCall.xlfGetDocument, 1);
            string docName = docNameObj?.ToString() ?? string.Empty;

            return docName.IndexOf(".xltx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   docName.IndexOf(".xltm", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch
        {
            return false;
        }
    }

    public static void ShowStatusConfirmation(string message)
    {
        try
        {
            XlCall.Excel(XlCall.xlcMessage, true, $"[SKSB System] {message}");
        }
        catch { }
    }

    public static string ResolveSheetName(string[] candidates, string defaultFallback)
    {
        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int totalSheets = sheetArray.GetLength(0);
                for (int i = 0; i < totalSheets; i++)
                {
                    string rawName = sheetArray[i, 0]?.ToString() ?? string.Empty;
                    int bracketIdx = rawName.LastIndexOf(']');
                    string cleanName = bracketIdx >= 0 ? rawName.Substring(bracketIdx + 1) : rawName;

                    foreach (var candidate in candidates)
                    {
                        if (string.Equals(cleanName, candidate, StringComparison.OrdinalIgnoreCase))
                        {
                            return cleanName;
                        }
                    }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ResolveSheetName] Error: {ex.Message}");
        }

        return defaultFallback;
    }

    public static IntPtr GetSheetId(string sheetName)
    {
        try
        {
            object evaluatedRef = XlCall.Excel(XlCall.xlfTextref, $"'{sheetName}'!R1C1", true);
            if (evaluatedRef is ExcelReference er)
            {
                return er.SheetId;
            }
        }
        catch { }

        return IntPtr.Zero;
    }

    public static void ClearActiveSheetDataRows(int startRow, int endRow, int startCol = 0, int endCol = 15)
    {
        if (endRow < startRow) return;

        try
        {
            var refRange = new ExcelReference(startRow, endRow, startCol, endCol);
            XlCall.Excel(XlCall.xlcSelect, refRange);
            XlCall.Excel(XlCall.xlcClear, 3); // 3 = Clear contents only
            XlCall.Excel(XlCall.xlcSelect, new ExcelReference(0, 0, 0, 0));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[C-API Clear Error] {ex.Message}");
        }
    }

    /// <summary>
    /// Purges all data rows across all sheets in the active workbook without COM.
    /// </summary>
    public static void PurgeAllSheetsData()
    {
        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int totalSheets = sheetArray.GetLength(0);
                for (int i = 1; i <= totalSheets; i++)
                {
                    string sheetName = sheetArray[i - 1, 0]?.ToString() ?? string.Empty;
                    if (!string.IsNullOrEmpty(sheetName))
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { sheetName });
                        ClearActiveSheetDataRows(1, 1000, 0, 15);
                    }
                }
            }
            else
            {
                ClearActiveSheetDataRows(1, 1000, 0, 15);
            }
        }
        finally
        {
            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }
}