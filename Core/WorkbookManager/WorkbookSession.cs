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

    /// <summary>
    /// Moves the Lookups sheet to the very end (rightmost position) of the workbook's sheet tabs.
    /// </summary>
    public static void MoveLookupsToEnd()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic sheets = wb.Sheets;
            int totalSheets = sheets.Count;

            dynamic lookups = null;
            try { lookups = sheets[LookupsSheetName]; } catch { }
            if (lookups == null) return;

            // Check if it is already the last sheet
            dynamic lastSheet = sheets[totalSheets];
            if (string.Equals((string)lookups.Name, (string)lastSheet.Name, StringComparison.OrdinalIgnoreCase))
                return;

            // Excel requires a sheet to be visible to move it
            bool wasVisible = (lookups.Visible == -1); // -1 = xlSheetVisible
            if (!wasVisible)
            {
                lookups.Visible = -1;
            }

            // Move after the last sheet in the tab bar
            lookups.Move(After: sheets[totalSheets]);

            // Restore hidden state if it was originally hidden
            if (!wasVisible)
            {
                lookups.Visible = 0; // 0 = xlSheetHidden
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[MoveLookupsToEnd Error] {ex.Message}");
        }
    }

    /// <summary>
    /// Checks whether the active document is a master template (.xltx / .xltm).
    /// </summary>
    public static bool IsMasterTemplateFile()
    {
        try
        {
            object docNameObj = XlCall.Excel(XlCall.xlfGetDocument, 1);
            string docName = docNameObj?.ToString() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(docName))
                return false;

            return docName.IndexOf(".xltx", StringComparison.OrdinalIgnoreCase) >= 0 ||
                   docName.IndexOf(".xltm", StringComparison.OrdinalIgnoreCase) >= 0;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Evaluates template mode, updates Lookups visibility & tab position, and syncs enums.
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
                // 1. Unhide and move Lookups to the end
                SetLookupsVisibility(visible: true);

                // 2. Auto-sync all Core Enums into the Lookups sheet
                TypesToExcel.SyncAllCoreEnums();

                if (showHeadsUp && !string.Equals(_lastDetectedFile, docName, StringComparison.OrdinalIgnoreCase))
                {
                    _lastDetectedFile = docName;
                    ShowStatusConfirmation("Template Design Mode: Lookups unhidden & Enums synced.");

                    MessageBox.Show(
                        $"[Template Detected]\n\n" +
                        $"Document: {docName}\n\n" +
                        $"• Switched to Template Edit Mode automatically.\n" +
                        $"• 'Lookups' sheet is positioned at the back and VISIBLE.\n" +
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

                // Hide Lookups sheet in Operational Mode
                SetLookupsVisibility(visible: false);
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

    public static void SetLookupsVisibility(bool visible)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                if (visible)
                {
                    XlCall.Excel(XlCall.xlcWorkbookUnhide, LookupsSheetName);
                    // Ensure it moves to the very back
                    MoveLookupsToEnd();
                }
                else
                {
                    object activeDoc = null;
                    try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
                    string originalSheet = activeDoc?.ToString() ?? string.Empty;

                    MoveLookupsToEnd();

                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { LookupsSheetName }); } catch { }
                    try { XlCall.Excel(XlCall.xlcWorkbookHide); } catch { }

                    if (!string.IsNullOrWhiteSpace(originalSheet))
                    {
                        string cleanOriginal = originalSheet.Contains(']')
                            ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                            : originalSheet;

                        if (!cleanOriginal.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                        }
                        else
                        {
                            try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { 1 }); } catch { }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SetLookupsVisibility Error] {ex.Message}");
            }
        });
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
            XlCall.Excel(XlCall.xlcClear, 3);
            XlCall.Excel(XlCall.xlcSelect, new ExcelReference(0, 0, 0, 0));
        }
        catch { }
    }

    public static void PurgeAllSheetsData()
    {
        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int totalSheets = sheetArray.GetLength(1);
                for (int i = 0; i < totalSheets; i++)
                {
                    string rawSheet = sheetArray[0, i]?.ToString() ?? string.Empty;
                    int bIdx = rawSheet.LastIndexOf(']');
                    string sheetName = bIdx >= 0 ? rawSheet.Substring(bIdx + 1) : rawSheet;

                    if (!string.IsNullOrEmpty(sheetName) && !sheetName.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { sheetName });
                        ClearActiveSheetDataRows(1, 1000, 0, 20);
                    }
                }
            }
        }
        finally
        {
            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }
}