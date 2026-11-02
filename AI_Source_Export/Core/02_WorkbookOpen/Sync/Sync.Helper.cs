namespace SKSB_App.Core.WorkbookOpen.Sync;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookOpen.Sheet;

public enum WriteMode
{
    SequentialReset,
    SmartOrganize
}

public class EntityTableBlock
{
    public string EntityName { get; set; } = string.Empty;
    public List<string> ColumnNames { get; set; } = new();
    public List<string> Columns => ColumnNames;
    public object[,] Grid { get; set; } = new object[0, 0];
    public int RowCount => Grid.GetLength(0);
    public int ColCount => Grid.GetLength(1);
}

public static class SyncHelper
{
    public static void SyncEntityBlocks(string targetSheet, List<EntityTableBlock> blocks)
    {
        if (blocks == null || blocks.Count == 0) return;

        using var _ = AppLogger.Scope($"SyncEntityBlocks(targetSheet='{targetSheet}', blocks={blocks.Count})");

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                string activeSheetBefore = GetActiveSheetName(xlApp);
                AppLogger.Info($"[SyncEntityBlocks Start] Requested Target: '{targetSheet}' | Current ActiveSheet: '{activeSheetBefore}'");

                EnsureSheetExists(targetSheet);
                MakeSheetVisibleAndActive(targetSheet);

                string activeSheetAfterSelect = GetActiveSheetName(xlApp);
                AppLogger.Info($"[SyncEntityBlocks Post-Select] ActiveSheet is now: '{activeSheetAfterSelect}'");

                // =========================================================================
                // SAFETY SHIELD: Abort immediately if Excel failed to switch to targetSheet!
                // =========================================================================
                if (!string.Equals(activeSheetAfterSelect, targetSheet, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Error($"CRITICAL SAFETY ABORT: Refusing to write! Target was '{targetSheet}', but Excel is focused on '{activeSheetAfterSelect}'. Writing here would destroy operational data!");
                    return;
                }

                XlCall.Excel(XlCall.xlcEcho, false);

                // Clean columns 0-100 across rows 0-2500
                AppLogger.Info($"Wiping canvas on '{activeSheetAfterSelect}' (Cols 0..100, Rows 0..2500)...");
                var totalWipeRef = new ExcelReference(0, 2500, 0, 100);
                object[,] emptyCanvas = new object[2501, 101];
                for (int r = 0; r <= 2500; r++)
                    for (int c = 0; c <= 100; c++)
                        emptyCanvas[r, c] = string.Empty;
                totalWipeRef.SetValue(emptyCanvas);

                int currentCol = 0;
                foreach (var block in blocks)
                {
                    int rowCount = block.RowCount;
                    int colCount = block.ColCount;
                    if (rowCount == 0 || colCount == 0) continue;

                    AppLogger.Info($"Writing EntityBlock '{block.EntityName}' ({rowCount} rows x {colCount} cols) at start col {currentCol} on sheet '{activeSheetAfterSelect}'");

                    // 1. Write Column Headers (Row 0)
                    object[,] headerGrid = new object[1, colCount];
                    for (int c = 0; c < colCount; c++)
                    {
                        headerGrid[0, c] = block.ColumnNames[c];
                    }
                    var headerRef = new ExcelReference(0, 0, currentCol, currentCol + colCount - 1);
                    headerRef.SetValue(headerGrid);

                    // 2. Write 2D Data Grid (Rows 1..rowCount)
                    var dataRef = new ExcelReference(1, rowCount, currentCol, currentCol + colCount - 1);
                    dataRef.SetValue(block.Grid);

                    // 3. Register Defined Names
                    for (int c = 0; c < colCount; c++)
                    {
                        string listName = block.ColumnNames[c];
                        try { XlCall.Excel(XlCall.xlcDeleteName, listName); } catch { }
                        int r1c1StartRow = 2;
                        int r1c1EndRow = rowCount + 1;
                        int r1c1Col = currentCol + c + 1;
                        string refersTo = $"='{targetSheet}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";
                        XlCall.Excel(XlCall.xlcDefineName, listName, refersTo, 3);
                    }

                    currentCol += colCount + 1;
                }

                SheetOperations.MoveAuxiliarySheetsToEnd();

                if (!SheetOperations.IsTemplateEditMode)
                {
                    try
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetSheet });
                        XlCall.Excel(XlCall.xlcWorkbookHide);
                    }
                    catch { }
                }

                // Return focus to original sheet
                if (!string.IsNullOrWhiteSpace(activeSheetBefore) &&
                    !string.Equals(activeSheetBefore, targetSheet, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Info($"Restoring user view to original sheet: '{activeSheetBefore}'");
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { activeSheetBefore }); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"SyncEntityBlocks on '{targetSheet}' threw exception", ex);
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, true);
            }
        });
    }

    public static void SyncDictionary(
        string targetSheet,
        Dictionary<string, IEnumerable<string>?> lists,
        WriteMode writeMode,
        bool injectPlaceholderIfEmpty = false)
    {
        if (lists == null || lists.Count == 0) return;

        var sanitized = new Dictionary<string, string[]>(StringComparer.OrdinalIgnoreCase);
        foreach (var (rawName, items) in lists)
        {
            string listName = rawName.EndsWith("List", StringComparison.OrdinalIgnoreCase)
                ? rawName
                : $"{rawName}List";

            var distinctValues = (items ?? Enumerable.Empty<string>())
                .Where(v => !string.IsNullOrWhiteSpace(v))
                .Select(v => v.Trim())
                .Where(v => v != "0" && !v.StartsWith("#"))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(v => v)
                .ToArray();

            if (distinctValues.Length > 0)
            {
                sanitized[listName] = distinctValues;
            }
            else if (injectPlaceholderIfEmpty)
            {
                string stub = $"[{listName.Replace("List", "")} Placeholder]";
                sanitized[listName] = new[] { stub };
            }
        }

        if (sanitized.Count > 0)
        {
            SyncNamedLists(targetSheet, sanitized, writeMode);
        }
    }

    public static void SyncNamedLists(
        string targetSheet,
        Dictionary<string, string[]> mappings,
        WriteMode writeMode)
    {
        if (mappings == null || mappings.Count == 0) return;

        using var _ = AppLogger.Scope($"SyncNamedLists(targetSheet='{targetSheet}', keys={mappings.Count}, mode={writeMode})");

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                string activeSheetBefore = GetActiveSheetName(xlApp);
                AppLogger.Info($"[SyncNamedLists Start] Requested Target: '{targetSheet}' | Current ActiveSheet: '{activeSheetBefore}'");

                EnsureSheetExists(targetSheet);
                MakeSheetVisibleAndActive(targetSheet);

                string activeSheetAfterSelect = GetActiveSheetName(xlApp);
                AppLogger.Info($"[SyncNamedLists Post-Select] ActiveSheet is now: '{activeSheetAfterSelect}'");

                // =========================================================================
                // SAFETY SHIELD: Abort immediately if Excel failed to switch to targetSheet!
                // =========================================================================
                if (!string.Equals(activeSheetAfterSelect, targetSheet, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Error($"CRITICAL SAFETY ABORT: Refusing to write! Target was '{targetSheet}', but Excel is focused on '{activeSheetAfterSelect}'. Aborting write to protect '{activeSheetAfterSelect}'!");
                    return;
                }

                XlCall.Excel(XlCall.xlcEcho, false);

                int colZero = 0;
                foreach (var (listName, values) in mappings)
                {
                    int targetCol = colZero;
                    if (writeMode == WriteMode.SmartOrganize)
                    {
                        targetCol = -1;
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
                        if (targetCol == -1) targetCol = 0;
                    }

                    int rowCount = values.Length;
                    AppLogger.Info($"Writing list '{listName}' ({rowCount} items) to Col {targetCol} on '{activeSheetAfterSelect}'");

                    var wipeRef = new ExcelReference(0, 2000, targetCol, targetCol);
                    object[,] clearGrid = new object[2001, 1];
                    for (int r = 0; r <= 2000; r++) clearGrid[r, 0] = string.Empty;
                    wipeRef.SetValue(clearGrid);

                    var headerRef = new ExcelReference(0, 0, targetCol, targetCol);
                    headerRef.SetValue(new object[,] { { listName } });

                    if (rowCount > 0)
                    {
                        var valRef = new ExcelReference(1, rowCount, targetCol, targetCol);
                        object[,] cellData = new object[rowCount, 1];
                        for (int i = 0; i < rowCount; i++)
                        {
                            string v = values[i];
                            if (double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal))
                                cellData[i, 0] = dVal;
                            else
                                cellData[i, 0] = v;
                        }
                        valRef.SetValue(cellData);

                        try { XlCall.Excel(XlCall.xlcDeleteName, listName); } catch { }
                        int r1c1StartRow = 2;
                        int r1c1EndRow = rowCount + 1;
                        int r1c1Col = targetCol + 1;
                        string refersTo = $"='{targetSheet}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";
                        XlCall.Excel(XlCall.xlcDefineName, listName, refersTo, 3);
                    }

                    if (writeMode == WriteMode.SequentialReset)
                    {
                        colZero += 2;
                    }
                }

                SheetOperations.MoveAuxiliarySheetsToEnd();

                // Return focus to original sheet
                if (!string.IsNullOrWhiteSpace(activeSheetBefore) &&
                    !string.Equals(activeSheetBefore, targetSheet, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Info($"Restoring view to: '{activeSheetBefore}'");
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { activeSheetBefore }); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"SyncNamedLists on '{targetSheet}' threw exception", ex);
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
                ws.Visible = -1;
                ws.Activate();
                AppLogger.Info($"MakeSheetVisibleAndActive: Activated '{sheetName}' via COM.");
                return;
            }
        }
        catch { }

        try
        {
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { sheetName });
            AppLogger.Info($"MakeSheetVisibleAndActive: Selected '{sheetName}' via C-API.");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"MakeSheetVisibleAndActive failed for '{sheetName}'", ex);
        }
    }

    private static void EnsureSheetExists(string sheetName)
    {
        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int dim1 = sheetArray.GetLength(1);
                for (int c = 0; c < dim1; c++)
                {
                    string raw = sheetArray[0, c]?.ToString() ?? string.Empty;
                    int bIdx = raw.LastIndexOf(']');
                    string clean = bIdx >= 0 ? raw.Substring(bIdx + 1) : raw;
                    if (string.Equals(clean, sheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
            }
        }
        catch { }

        try
        {
            AppLogger.Info($"Sheet '{sheetName}' not found. Creating new sheet...");
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);
            object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string fullName = docInfo?.ToString() ?? string.Empty;
            string tempName = fullName.Contains(']')
                ? fullName.Substring(fullName.LastIndexOf(']') + 1)
                : fullName;
            XlCall.Excel(XlCall.xlcWorkbookName, tempName, sheetName);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"EnsureSheetExists failed creating '{sheetName}'", ex);
        }
    }

    private static string GetActiveSheetName(dynamic xlApp)
    {
        try
        {
            return xlApp?.ActiveSheet?.Name?.ToString() ?? string.Empty;
        }
        catch
        {
            try
            {
                object doc = XlCall.Excel(XlCall.xlfGetDocument, 76);
                string raw = doc?.ToString() ?? string.Empty;
                return raw.Contains(']') ? raw.Substring(raw.LastIndexOf(']') + 1) : raw;
            }
            catch
            {
                return "Unknown";
            }
        }
    }
}