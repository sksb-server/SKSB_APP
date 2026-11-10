namespace SKSB_App.Core.WorkbookOpen.Sync;

using System;
using System.Collections.Generic;
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
    /// <summary>
    /// Writes entity table blocks directly to targetSheet via COM without activating the sheet or moving the screen.
    /// Performs live verification immediately after write.
    /// </summary>
    public static void SyncEntityBlocks(string targetSheet, List<EntityTableBlock> blocks)
    {
        if (blocks == null || blocks.Count == 0) return;

        using var _ = AppLogger.Scope($"SyncEntityBlocks(targetSheet='{targetSheet}', blocks={blocks.Count})");

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic wb = xlApp.ActiveWorkbook;
                if (wb == null) return;

                string activeSheetBefore = xlApp.ActiveSheet?.Name?.ToString() ?? "Unknown";
                AppLogger.Info($"[SyncEntityBlocks Start] Target: '{targetSheet}' | User ActiveSheet: '{activeSheetBefore}'");

                dynamic ws = EnsureSheetExistsCom(wb, targetSheet);
                if (ws == null)
                {
                    AppLogger.Error($"[SYNC FAILED] Could not locate or create '{targetSheet}'");
                    return;
                }

                // 1. Clear operational area on targetSheet (Cols A..CW, Rows 1..2500) without activating or selecting
                dynamic clearRange = ws.Range[ws.Cells[1, 1], ws.Cells[2500, 100]];
                clearRange.ClearContents();

                int currentCol = 1; // 1-based COM index
                int totalBlocksWritten = 0;
                int totalColumnsWritten = 0;

                foreach (var block in blocks)
                {
                    int rowCount = block.RowCount;
                    int colCount = block.ColCount;
                    if (rowCount == 0 || colCount == 0) continue;

                    AppLogger.Info($"Writing EntityBlock '{block.EntityName}' ({rowCount} rows x {colCount} cols) at Col {currentCol} on '{targetSheet}' (Silent COM)");

                    // A. Write Headers (Row 1)
                    object[,] headerGrid = new object[1, colCount];
                    for (int c = 0; c < colCount; c++)
                    {
                        headerGrid[0, c] = block.ColumnNames[c];
                    }
                    dynamic headerRange = ws.Range[ws.Cells[1, currentCol], ws.Cells[1, currentCol + colCount - 1]];
                    headerRange.Value2 = headerGrid;

                    // B. Write 2D Data Grid (Rows 2..rowCount+1)
                    dynamic dataRange = ws.Range[ws.Cells[2, currentCol], ws.Cells[rowCount + 1, currentCol + colCount - 1]];
                    dataRange.Value2 = block.Grid;

                    // C. Register Defined Names
                    for (int c = 0; c < colCount; c++)
                    {
                        string listName = block.ColumnNames[c];
                        try { wb.Names.Item(listName).Delete(); } catch { }

                        int r1 = 2;
                        int r2 = rowCount + 1;
                        int cPos = currentCol + c;
                        string refersTo = $"='{targetSheet}'!R{r1}C{cPos}:R{r2}C{cPos}";
                        wb.Names.Add(Name: listName, RefersToR1C1: refersTo, Visible: true);
                    }

                    currentCol += colCount + 1;
                    totalBlocksWritten++;
                    totalColumnsWritten += colCount;
                }

                // =========================================================================
                // VERIFICATION: Check whether data was written and Named Ranges are valid
                // =========================================================================
                bool verifySuccess = VerifyEntityBlocks(wb, ws, blocks);
                AppLogger.Info($"[SYNC VERIFICATION RESULT] EntityBlocks on '{targetSheet}': {(verifySuccess ? "SUCCESS" : "FAILED / PARTIAL")} (Blocks={totalBlocksWritten}, Cols={totalColumnsWritten})");

                // Confirm active sheet never shifted
                string activeSheetAfter = xlApp.ActiveSheet?.Name?.ToString() ?? "Unknown";
                if (!string.Equals(activeSheetBefore, activeSheetAfter, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Warn($"[VIEWPORT SHIFT DETECTED] ActiveSheet changed from '{activeSheetBefore}' to '{activeSheetAfter}'. Restoring focus...");
                    try { wb.Sheets[activeSheetBefore].Activate(); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"SyncEntityBlocks on '{targetSheet}' threw exception", ex);
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

    /// <summary>
    /// Writes list mappings directly to targetSheet via COM in the background.
    /// Avoids activating the auxiliary sheet and runs verification immediately afterwards.
    /// </summary>
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
                dynamic wb = xlApp.ActiveWorkbook;
                if (wb == null) return;

                string activeSheetBefore = xlApp.ActiveSheet?.Name?.ToString() ?? "Unknown";
                AppLogger.Info($"[SyncNamedLists Start] Target: '{targetSheet}' | User ActiveSheet: '{activeSheetBefore}'");

                dynamic ws = EnsureSheetExistsCom(wb, targetSheet);
                if (ws == null)
                {
                    AppLogger.Error($"[SYNC FAILED] Could not locate or create '{targetSheet}'");
                    return;
                }

                // Inside SyncHelper.SyncNamedLists (Core\02_WorkbookOpen\Sync\Sync.Helper.cs),
                // replace the foreach loop (lines 166-222) with:
                int maxColsToScan = 200;
                object[,]? headerRowGrid = null;
                if (writeMode == WriteMode.SmartOrganize)
                {
                    try
                    {
                        headerRowGrid = ws.Range[ws.Cells[1, 1], ws.Cells[1, maxColsToScan]].Value2 as object[,];
                    }
                    catch { }
                }

                int colOneBased = 1;
                foreach (var (listName, values) in mappings)
                {
                    int targetCol = colOneBased;

                    if (writeMode == WriteMode.SmartOrganize)
                    {
                        targetCol = 0;
                        for (int c = 1; c <= maxColsToScan; c += 2)
                        {
                            string existingHeader = headerRowGrid != null
                                ? (headerRowGrid[1, c]?.ToString()?.Trim() ?? string.Empty)
                                : (ws.Cells[1, c].Value2?.ToString()?.Trim() ?? string.Empty);

                            if (string.Equals(existingHeader, listName, StringComparison.OrdinalIgnoreCase))
                            {
                                targetCol = c;
                                break;
                            }
                            if (string.IsNullOrEmpty(existingHeader) && targetCol == 0)
                            {
                                targetCol = c;
                            }
                        }
                        if (targetCol == 0) targetCol = 1;
                        if (headerRowGrid != null && targetCol <= maxColsToScan)
                        {
                            headerRowGrid[1, targetCol] = listName;
                        }
                    }

                    int rowCount = values.Length;

                    // Clear target column and write header + values in a single contiguous 2D array write
                    dynamic wipeRange = ws.Range[ws.Cells[1, targetCol], ws.Cells[Math.Max(2000, rowCount + 10), targetCol]];
                    wipeRange.ClearContents();

                    object[,] colPayload = new object[rowCount + 1, 1];
                    colPayload[0, 0] = listName;
                    for (int i = 0; i < rowCount; i++)
                    {
                        string v = values[i];
                        colPayload[i + 1, 0] = double.TryParse(v, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)
                            ? dVal
                            : v;
                    }

                    dynamic writeRange = ws.Range[ws.Cells[1, targetCol], ws.Cells[rowCount + 1, targetCol]];
                    writeRange.Value2 = colPayload;

                    if (rowCount > 0)
                    {
                        try { wb.Names.Item(listName).Delete(); } catch { }
                        string refersTo = $"='{targetSheet}'!R2C{targetCol}:R{rowCount + 1}C{targetCol}";
                        wb.Names.Add(Name: listName, RefersToR1C1: refersTo, Visible: true);
                    }

                    if (writeMode == WriteMode.SequentialReset)
                    {
                        colOneBased += 2;
                    }
                }

                // =========================================================================
                // VERIFICATION: Check whether columns and Defined Names exist
                // =========================================================================
                bool verifySuccess = VerifyNamedLists(wb, ws, mappings);
                AppLogger.Info($"[SYNC VERIFICATION RESULT] NamedLists on '{targetSheet}': {(verifySuccess ? "SUCCESS" : "FAILED / PARTIAL")} (Total Lists={mappings.Count})");

                // Confirm active sheet was not shifted
                string activeSheetAfter = xlApp.ActiveSheet?.Name?.ToString() ?? "Unknown";
                if (!string.Equals(activeSheetBefore, activeSheetAfter, StringComparison.OrdinalIgnoreCase))
                {
                    AppLogger.Warn($"[VIEWPORT SHIFT DETECTED] ActiveSheet changed from '{activeSheetBefore}' to '{activeSheetAfter}'. Restoring focus...");
                    try { wb.Sheets[activeSheetBefore].Activate(); } catch { }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"SyncNamedLists on '{targetSheet}' threw exception", ex);
            }
        });
    }

    private static dynamic EnsureSheetExistsCom(dynamic wb, string sheetName)
    {
        try
        {
            foreach (dynamic ws in wb.Sheets)
            {
                if (string.Equals(ws.Name?.ToString(), sheetName, StringComparison.OrdinalIgnoreCase))
                {
                    return ws;
                }
            }

            AppLogger.Info($"Sheet '{sheetName}' not found. Adding silently at the end...");
            int count = (int)wb.Sheets.Count;
            dynamic newWs = wb.Sheets.Add(After: wb.Sheets[count]);
            newWs.Name = sheetName;
            return newWs;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"EnsureSheetExistsCom('{sheetName}') error", ex);
            return null!;
        }
    }

    private static bool VerifyEntityBlocks(dynamic wb, dynamic ws, List<EntityTableBlock> blocks)
    {
        try
        {
            foreach (var block in blocks)
            {
                if (block.RowCount == 0 || block.ColCount == 0) continue;

                foreach (var colName in block.ColumnNames)
                {
                    dynamic? nameObj = null;
                    try { nameObj = wb.Names.Item(colName); } catch { }
                    if (nameObj == null)
                    {
                        AppLogger.Error($"[VERIFY FAILURE] Defined Name '{colName}' was not found in Workbook Names!");
                        return false;
                    }

                    dynamic range = nameObj.RefersToRange;
                    if (range == null || (int)range.Rows.Count != block.RowCount)
                    {
                        AppLogger.Error($"[VERIFY FAILURE] Name '{colName}' points to invalid range rows: Expected {block.RowCount}");
                        return false;
                    }
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error("VerifyEntityBlocks exception", ex);
            return false;
        }
    }

    private static bool VerifyNamedLists(dynamic wb, dynamic ws, Dictionary<string, string[]> mappings)
    {
        try
        {
            foreach (var (listName, values) in mappings)
            {
                if (values.Length == 0) continue;

                dynamic? nameObj = null;
                try { nameObj = wb.Names.Item(listName); } catch { }
                if (nameObj == null)
                {
                    AppLogger.Error($"[VERIFY FAILURE] Defined Name '{listName}' was not created!");
                    return false;
                }

                dynamic range = nameObj.RefersToRange;
                if (range == null || (int)range.Rows.Count != values.Length)
                {
                    AppLogger.Error($"[VERIFY FAILURE] Defined Name '{listName}' row count mismatch: Expected {values.Length}");
                    return false;
                }
            }
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Error("VerifyNamedLists exception", ex);
            return false;
        }
    }
}