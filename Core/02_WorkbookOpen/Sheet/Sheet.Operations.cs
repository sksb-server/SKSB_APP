namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;

public static class SheetOperations
{
    public static bool IsTemplateEditMode { get; set; } = false;
    public const string LookupsSheetName = "Lookups";
    public const string DataSheetName = "Data";

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
            using var _ = AppLogger.Scope($"SetAuxiliarySheetsVisibility(visible={visible})");

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
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic? ws = null;
            try { ws = wb.Sheets[sheetName]; } catch { }
            if (ws == null)
            {
                AppLogger.Info($"[SetSingleSheetVisibility] Sheet '{sheetName}' does not exist, skipping.");
                return;
            }

            string currentActive = wb.ActiveSheet?.Name?.ToString() ?? string.Empty;
            AppLogger.Info($"[SetSingleSheetVisibility] Sheet='{sheetName}', CurrentActive='{currentActive}', TargetVisible={visible}");

            if (!visible)
            {
                // If Excel is currently on this sheet, shift focus to a non-auxiliary sheet before hiding
                if (string.Equals(currentActive, sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    foreach (dynamic candidate in wb.Sheets)
                    {
                        string cName = candidate.Name?.ToString() ?? string.Empty;
                        if (!cName.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase) &&
                            !cName.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                        {
                            AppLogger.Info($"[SetSingleSheetVisibility] Shifting focus off '{sheetName}' to '{cName}' before hiding.");
                            candidate.Activate();
                            break;
                        }
                    }
                }
                ws.Visible = 0; /* xlSheetHidden */
            }
            else
            {
                ws.Visible = -1; /* xlSheetVisible */
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"SetSingleSheetVisibility('{sheetName}', visible={visible}) failed", ex);
        }
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
            dynamic? targetSheet = null;
            try { targetSheet = sheets[targetSheetName]; } catch { }
            if (targetSheet == null) return;

            int totalSheets = (int)sheets.Count;
            int currentIdx = (int)targetSheet.Index;

            // If already at the end, do nothing
            if (currentIdx >= totalSheets - 1)
            {
                return;
            }

            dynamic? activeSheetBeforeMove = wb.ActiveSheet;
            string activeName = activeSheetBeforeMove?.Name?.ToString() ?? "Unknown";
            AppLogger.Info($"[MoveSingleSheetToEnd] Moving '{targetSheetName}' (currently index {currentIdx}) to end of {totalSheets} sheets. Active sheet: '{activeName}'");

            bool wasVisible = ((int)targetSheet.Visible == -1);
            if (!wasVisible) targetSheet.Visible = -1;

            targetSheet.Move(After: sheets[totalSheets]);

            if (!wasVisible) targetSheet.Visible = 0;

            if (activeSheetBeforeMove != null)
            {
                try { activeSheetBeforeMove.Activate(); } catch { }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"MoveSingleSheetToEnd('{targetSheetName}') failed", ex);
        }
    }

    public static void PurgeAllSheetsData()
    {
        using var _ = AppLogger.Scope("PurgeAllSheetsData");
        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            foreach (dynamic ws in wb.Sheets)
            {
                string sName = (string)ws.Name;
                if (sName.Equals(LookupsSheetName, StringComparison.OrdinalIgnoreCase) ||
                    sName.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                AppLogger.Info($"Purging operational data on '{sName}'...");
                ws.Activate();
                var refRange = new ExcelReference(1, 1000, 0, 25);
                XlCall.Excel(XlCall.xlcSelect, refRange);
                XlCall.Excel(XlCall.xlcClear, 3); // 3 = Clear contents only
                XlCall.Excel(XlCall.xlcSelect, new ExcelReference(0, 0, 0, 0));
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("PurgeAllSheetsData error", ex);
        }
        finally
        {
            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }
}