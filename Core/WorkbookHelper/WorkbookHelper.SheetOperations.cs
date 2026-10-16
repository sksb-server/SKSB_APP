namespace SKSB_App.Core.WorkbookHelper;

using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Services;

public static class SheetOps
{
    public static bool IsTemplateEditMode { get; set; } = false;
    private static string _lastDetectedFile = string.Empty;

    public const string LookupsSheetName = "Lookups";
    public const string DataSheetName = "Data";

    /// <summary>
    /// Searches the active workbook for the first existing sheet matching a list of candidate aliases.
    /// </summary>
    public static string ResolveSheetName(string[] candidates, string defaultFallback)
    {
        if (candidates == null || candidates.Length == 0)
            return defaultFallback;

        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                foreach (object item in sheetArray)
                {
                    string rawName = item?.ToString() ?? string.Empty;
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
        catch { }

        return defaultFallback;
    }

    /// <summary>
    /// Checks whether the active file is a master .xltx/.xltm template file.
    /// </summary>
    public static bool IsMasterTemplateFile()
    {
        try
        {
            object docNameObj = XlCall.Excel(XlCall.xlfGetDocument, 1);
            string docName = docNameObj?.ToString() ?? string.Empty;

            if (string.IsNullOrEmpty(docName))
                return false;

            int openBracket = docName.IndexOf('[');
            int closeBracket = docName.IndexOf(']');
            string fileName = (openBracket >= 0 && closeBracket > openBracket)
                ? docName.Substring(openBracket + 1, closeBracket - openBracket - 1)
                : docName;

            return fileName.EndsWith(".xltx", StringComparison.OrdinalIgnoreCase) ||
                   fileName.EndsWith(".xltm", StringComparison.OrdinalIgnoreCase);
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

    public static void SetAuxiliarySheetsVisibility(bool visible)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SetSingleSheetVisibility(LookupsSheetName, visible);
            SetSingleSheetVisibility(DataSheetName, visible);
            MoveAuxiliarySheetsToEnd();
        });
    }

    private static void SetSingleSheetVisibility(string sheetName, bool visible)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = null;
            try { ws = xlApp.ActiveWorkbook.Sheets[sheetName]; } catch { }
            if (ws == null) return;

            ws.Visible = visible ? -1 /* xlSheetVisible */ : 0 /* xlSheetHidden */;
        }
        catch { }
    }

    public static void MoveAuxiliarySheetsToEnd()
    {
        MoveSingleSheetToEnd(LookupsSheetName);
        MoveSingleSheetToEnd(DataSheetName);
    }

    private static void MoveSingleSheetToEnd(string targetSheetName)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic sheets = wb.Sheets;
            dynamic targetSheet = null;
            try { targetSheet = sheets[targetSheetName]; } catch { }
            if (targetSheet == null) return;

            dynamic lastSheet = sheets[sheets.Count];
            if (string.Equals((string)targetSheet.Name, (string)lastSheet.Name, StringComparison.OrdinalIgnoreCase))
                return;

            bool wasVisible = (targetSheet.Visible == -1);
            if (!wasVisible) targetSheet.Visible = -1;

            targetSheet.Move(After: sheets[sheets.Count]);
            if (!wasVisible) targetSheet.Visible = 0;
        }
        catch { }
    }
}