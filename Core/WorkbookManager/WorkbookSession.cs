namespace SKSB_App.Core.WorkbookManager;

using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Services;

public static class WorkbookSession
{
    public static bool IsTemplateEditMode { get; set; } = false;
    private static string _lastDetectedFile = string.Empty;

    public const string LookupsSheetName = "Lookups";
    public const string DataSheetName = "Data";

    /// <summary>
    /// Searches the active workbook for the first existing sheet matching a list of candidate aliases.
    /// Iterates across both matrix dimensions to support all Excel C-API versions.
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
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ResolveSheetName Error] {ex.Message}");
        }

        return defaultFallback;
    }

    /// <summary>
    /// Moves both auxiliary sheets ('Lookups' and 'Data') to the very end of the sheet tabs.
    /// </summary>
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
            int totalSheets = sheets.Count;

            dynamic targetSheet = null;
            try { targetSheet = sheets[targetSheetName]; } catch { }
            if (targetSheet == null) return;

            dynamic lastSheet = sheets[totalSheets];
            if (string.Equals((string)targetSheet.Name, (string)lastSheet.Name, StringComparison.OrdinalIgnoreCase))
                return;

            bool wasVisible = (targetSheet.Visible == -1); // -1 = xlSheetVisible
            if (!wasVisible)
            {
                targetSheet.Visible = -1;
            }

            targetSheet.Move(After: sheets[totalSheets]);

            if (!wasVisible)
            {
                targetSheet.Visible = 0; // 0 = xlSheetHidden
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MoveSingleSheetToEnd Error] {ex.Message}");
        }
    }
    /// <summary>
    /// Checks whether the active file is a master .xltx/.xltm template file.
    /// Handles both plain 'Template.xltx' and qualified '[Template.xltx]Sheet1' titles.
    /// </summary>
    public static bool IsMasterTemplateFile()
    {
        try
        {
            // Document type 1 returns '[BookName.xltx]Sheet1' or 'BookName.xltx'
            object docNameObj = XlCall.Excel(XlCall.xlfGetDocument, 1);
            string docName = docNameObj?.ToString() ?? string.Empty;

            if (string.IsNullOrEmpty(docName))
                return false;

            // Extract filename from bracketed document format: [FileName.xltx]SheetName
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

    /// <summary>
    /// Evaluates template mode, updates Lookups & Data visibility, and syncs all Core enums.
    /// </summary>
    public static bool EvaluateTemplateMode(bool showHeadsUp = true)
    {
        try
        {
            object docNameObj = XlCall.Excel(XlCall.xlfGetDocument, 1);
            string docName = docNameObj?.ToString() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(docName))
                return false;

            bool isTemplate = IsMasterTemplateFile();
            IsTemplateEditMode = isTemplate;

            if (isTemplate)
            {
                // 1. Unhide both Lookups & Data sheets and move to end
                SetAuxiliarySheetsVisibility(visible: true);

                // 2. Auto-sync all Core Enums into the Lookups sheet
                TypesToExcel.SyncAllCoreEnums();

                if (showHeadsUp && !string.Equals(_lastDetectedFile, docName, StringComparison.OrdinalIgnoreCase))
                {
                    _lastDetectedFile = docName;
                    ShowStatusConfirmation("Template Design Mode: Lookups & Data unhidden, Enums synced.");

                    MessageBox.Show(
                        $"[Template Detected]\n\n" +
                        $"Document: {docName}\n\n" +
                        $"• Switched to Template Edit Mode automatically.\n" +
                        $"• 'Lookups' & 'Data' sheets are VISIBLE at the back.\n" +
                        $"• All Core Enums have been synchronized.\n" +
                        $"• Double-clicking dropdown cells will open the Template Inspector.",
                        "Template Edit Mode Active",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            else
            {
                _lastDetectedFile = string.Empty;

                // Hide both Lookups & Data sheets in Operational Mode
                SetAuxiliarySheetsVisibility(visible: false);
            }

            return isTemplate;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[EvaluateTemplateMode Error] {ex.Message}");
            IsTemplateEditMode = false;
            return false;
        }
    }

    /// <summary>
    /// Unified visibility controller for both Lookups and Data auxiliary sheets.
    /// </summary>
    public static void SetAuxiliarySheetsVisibility(bool visible)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            // Remember original active sheet to restore focus cleanly
            object activeDoc = null;
            try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
            string originalSheet = activeDoc?.ToString() ?? string.Empty;

            SetSingleSheetVisibility(LookupsSheetName, visible);
            SetSingleSheetVisibility(DataSheetName, visible);

            MoveAuxiliarySheetsToEnd();

            // Restore focus back to original sheet if hiding
            if (!visible && !string.IsNullOrWhiteSpace(originalSheet))
            {
                string cleanOriginal = originalSheet.Contains(']')
                    ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                    : originalSheet;

                if (!cleanOriginal.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase) &&
                    !cleanOriginal.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                }
                else
                {
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { 1 }); } catch { }
                }
            }
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
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SetSingleSheetVisibility Error: {sheetName}] {ex.Message}");
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
        catch { }
    }

    /// <summary>
    /// Purges all operational rows while strictly protecting 'Lookups' and 'Data' sheets.
    /// </summary>
    public static void PurgeAllSheetsData()
    {
        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            foreach (dynamic ws in wb.Sheets)
            {
                string sName = (string)ws.Name;

                // Protect both Lookups and Data from being wiped
                if (sName.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase) ||
                    sName.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                ws.Activate();
                ClearActiveSheetDataRows(1, 1000, 0, 25);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[PurgeAllSheetsData Error] {ex.Message}");
        }
        finally
        {
            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }
}