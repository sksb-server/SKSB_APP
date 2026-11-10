namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

public sealed class DropdownOption
{
    public string Display { get; init; } = string.Empty;
    public string Raw { get; init; } = string.Empty;
    public int Section { get; init; } = 0;
    public bool IsValid { get; init; } = true;
}

public static class RuleEvaluator
{
    private static bool _isEvaluating = false;

    /// <summary>
    /// Live self-healing row evaluation: runs on SheetChange and Dropdown selection.
    /// Executes:
    /// 1) Transform on edited cell (e.g. UpperCase plate/VIN, ShortName formatting)
    /// 2) Two-way [ColumnSync] lookups into partner columns
    /// 3) [ColumnDefault] auto-defaults on blank cells (once the row has user input)
    /// 4) [ColumnStateGuard] state enforcement (e.g. forcing "Expired" if date < Today, or reverting "Expired" to "Active" if dates are valid)
    /// </summary>
    /// <summary>
    /// Live self-healing row evaluation: runs on SheetChange and Dropdown selection.
    /// </summary>
    public static List<string> EvaluateAndHealRow(
        dynamic worksheet,
        dynamic table,
        int excelRow,
        int? editedAbsoluteCol = null,
        bool isAlreadyUnprotected = false) // <-- Added parameter to prevent redundant unprotecting
    {
        var notices = new List<string>();
        if (_isEvaluating || worksheet == null || table == null) return notices;

        try
        {
            _isEvaluating = true;
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = worksheet.Parent;

            bool origEvents = true;
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

            try
            {
                string tableName = table.Name?.ToString() ?? string.Empty;
                int headerRow = (int)table.HeaderRowRange.Row;
                int startCol = (int)table.HeaderRowRange.Column;
                int colCount = (int)table.ListColumns.Count;

                if (excelRow <= headerRow) return notices;

                // Ignore anything outside the DataBodyRange (e.g. the Total Row)
                dynamic? bodyRange = table.DataBodyRange;
                if (bodyRange == null) return notices;
                int bodyFirstRow = (int)bodyRange.Row;
                int bodyLastRow = bodyFirstRow + (int)bodyRange.Rows.Count - 1;
                if (excelRow < bodyFirstRow || excelRow > bodyLastRow) return notices;

                var colHeaders = new string[colCount + 1];
                var colRules = new PropertyRuleDescriptor?[colCount + 1];
                var rowValues = new string[colCount + 1];
                var rawObjects = new object?[colCount + 1];

                for (int c = 1; c <= colCount; c++)
                {
                    int absCol = startCol + c - 1;
                    string h = worksheet.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                    colHeaders[c] = h;
                    colRules[c] = RuleRegistry.GetRule(tableName, h);

                    object? valObj = worksheet.Cells[excelRow, absCol].Value2;
                    rawObjects[c] = valObj;
                    rowValues[c] = CleanCellText(valObj);
                }

                bool rowHasUserInput = false;
                for (int c = 1; c <= colCount; c++)
                {
                    int absCol = startCol + c - 1;
                    string f = worksheet.Cells[excelRow, absCol].Formula?.ToString() ?? string.Empty;
                    if (f.StartsWith("=")) continue;

                    string clean = rowValues[c];
                    if (string.IsNullOrWhiteSpace(clean)) continue;
                    if (clean == "0" || clean == "0.00" || clean == "-" || clean == "\u2014" ||
                        clean.Equals("None", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var rule = colRules[c];
                    if (rule?.DefaultRule?.Kind == DefaultKind.Literal &&
                        string.Equals(clean, rule.DefaultRule.LiteralValue, StringComparison.OrdinalIgnoreCase))
                        continue;

                    rowHasUserInput = true;
                    break;
                }

                if (!rowHasUserInput) return notices;

                // Extract the heal logic into an Action
                Action healAction = new Action(() =>
                {
                    // STEP 1: Apply Transform & Explicit Sync from the Edited Column
                    if (editedAbsoluteCol.HasValue)
                    {
                        int relEdited = editedAbsoluteCol.Value - startCol + 1;
                        if (relEdited >= 1 && relEdited <= colCount)
                        {
                            var editedRule = colRules[relEdited];
                            string currentText = rowValues[relEdited];
                            if (editedRule != null && !string.IsNullOrWhiteSpace(currentText))
                            {
                                string transformed = ApplyTextTransform(currentText, editedRule.Transform);
                                if (!string.Equals(transformed, currentText, StringComparison.Ordinal))
                                {
                                    WriteCellSafe(worksheet.Cells[excelRow, editedAbsoluteCol.Value], transformed, editedRule);
                                    rowValues[relEdited] = transformed;
                                    currentText = transformed;
                                }
                                if (editedRule.SyncRules.Count > 0)
                                {
                                    ExecuteSyncRulesForColumn(wb, worksheet, table, excelRow, startCol, colCount, colHeaders, colRules, rowValues, rawObjects, editedRule, currentText);
                                }
                            }
                        }
                    }

                    // STEP 2: Fill Missing Partner Syncs Across the Row
                    for (int c = 1; c <= colCount; c++)
                    {
                        var rule = colRules[c];
                        string val = rowValues[c];
                        if (rule == null || string.IsNullOrWhiteSpace(val) || rule.SyncRules.Count == 0) continue;

                        foreach (var sync in rule.SyncRules)
                        {
                            int targetRelCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
                            if (targetRelCol < 1) continue;

                            var targetRule = colRules[targetRelCol];
                            string currentTargetVal = rowValues[targetRelCol];
                            string? lookedUp = RuleDataCache.LookupValue(wb, val, sync.LookupSource, sync.ResultSource);

                            if (!string.IsNullOrWhiteSpace(lookedUp))
                            {
                                bool isMismatch = !string.IsNullOrWhiteSpace(currentTargetVal) && !RuleDataCache.ValuesMatch(currentTargetVal, lookedUp);
                                if (string.IsNullOrWhiteSpace(currentTargetVal) || (isMismatch && (targetRule?.IsWarningOnly ?? false)))
                                {
                                    int targetAbsCol = startCol + targetRelCol - 1;
                                    WriteCellSafe(worksheet.Cells[excelRow, targetAbsCol], lookedUp, targetRule);
                                    rowValues[targetRelCol] = lookedUp;
                                    rawObjects[targetRelCol] = lookedUp;
                                    if (isMismatch) notices.Add($"Row {excelRow} [{colHeaders[targetRelCol]}]: Auto-corrected '{currentTargetVal}' to '{lookedUp}' based on {colHeaders[c]} '{val}'.");
                                }
                            }
                        }
                    }

                    // STEP 3: Apply [ColumnDefault]
                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        var rule = colRules[c];
                        if (rule?.DefaultRule == null || rule.DefaultRule.Kind == DefaultKind.None) continue;
                        if (!string.IsNullOrWhiteSpace(rowValues[c])) continue;

                        string cellFormula = worksheet.Cells[excelRow, absCol].Formula?.ToString() ?? string.Empty;
                        if (cellFormula.StartsWith("=")) continue;

                        object? defaultVal = ComputeDefaultValue(wb, table, c, colHeaders, rawObjects, rowValues, rule.DefaultRule);
                        if (defaultVal != null)
                        {
                            WriteCellSafe(worksheet.Cells[excelRow, absCol], defaultVal, rule);
                            rawObjects[c] = defaultVal;
                            rowValues[c] = FormatDisplayValue(defaultVal, rule);
                            notices.Add($"Row {excelRow} [{colHeaders[c]}]: Defaulted to '{rowValues[c]}'.");
                        }
                    }

                    // STEP 4: Enforce [ColumnStateGuard]
                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        var rule = colRules[c];
                        if (rule == null || rule.StateGuards.Count == 0) continue;

                        string currentVal = rowValues[c];
                        string? enforcedValue = EvaluateStateGuardsForCell(colHeaders, rawObjects, rowValues, rule, currentVal, out string? guardReason);
                        if (!string.IsNullOrWhiteSpace(enforcedValue) && !RuleDataCache.ValuesMatch(currentVal, enforcedValue))
                        {
                            WriteCellSafe(worksheet.Cells[excelRow, absCol], enforcedValue, rule);
                            rowValues[c] = enforcedValue;
                            rawObjects[c] = enforcedValue;
                            if (!string.IsNullOrWhiteSpace(guardReason))
                            {
                                SheetOperations.ShowStatusConfirmation(guardReason);
                                notices.Add($"Row {excelRow} [{colHeaders[c]}]: {guardReason}");
                            }
                        }
                    }
                });

                // EXECUTION: Use existing bypass if batching, otherwise unprotect single row
                if (isAlreadyUnprotected)
                {
                    healAction();
                }
                else
                {
                    SheetProtections.ExecuteWithBypass((object)worksheet, healAction);
                }
            }
            finally
            {
                try { xlApp.EnableEvents = origEvents; } catch { }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleEvaluator.EvaluateAndHealRow Error] {ex.Message}");
        }
        finally
        {
            _isEvaluating = false;
        }

        return notices;
    }

    /// <summary>
    /// Bulk in-memory table healing: reads headers, formulas, and body values in 3 COM calls total,
    /// evaluates all rules in RAM with O(1) dictionary lookups, and writes back ONLY modified cells.
    /// </summary>
    public static List<string> EvaluateAndHealTable(dynamic worksheet, dynamic table)
    {
        var allNotices = new List<string>();
        if (_isEvaluating || worksheet == null || table == null) return allNotices;

        try
        {
            _isEvaluating = true;
            dynamic? body = table.DataBodyRange;
            if (body == null) return allNotices;

            int startRow = (int)body.Row;
            int rowCount = (int)body.Rows.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            if (rowCount <= 0 || colCount <= 0) return allNotices;

            string tableName = table.Name?.ToString() ?? string.Empty;
            dynamic wb = worksheet.Parent;

            // 1. Bulk-read headers (1 COM call)
            object[,] headerGrid = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, colCount);
            var colHeaders = new string[colCount + 1];
            var colRules = new PropertyRuleDescriptor?[colCount + 1];

            for (int c = 1; c <= colCount; c++)
            {
                string h = headerGrid[1, c]?.ToString()?.Trim() ?? string.Empty;
                colHeaders[c] = h;
                colRules[c] = RuleRegistry.GetRule(tableName, h);
            }

            // 2. Bulk-read archetype formulas on first data row to identify formula columns (1 COM call)
            var isFormulaCol = new bool[colCount + 1];
            try
            {
                dynamic archRange = worksheet.Range[
                    worksheet.Cells[startRow, startCol],
                    worksheet.Cells[startRow, startCol + colCount - 1]
                ];
                object? rawFormulas = archRange.Formula;
                if (rawFormulas is object[,] fMat)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        string f = fMat[1, c]?.ToString() ?? string.Empty;
                        isFormulaCol[c] = f.StartsWith("=");
                    }
                }
                else if (rawFormulas is string singleF)
                {
                    isFormulaCol[1] = singleF.StartsWith("=");
                }
            }
            catch { }

            // 3. Bulk-read entire DataBodyRange into memory (1 COM call)
            object[,] dataGrid = TableMapper.GetRange2DValues(body, rowCount, colCount);

            // 4. Pre-build O(1) lookup dictionaries for all SyncRules on this table
            var lookupDictCache = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, string> GetLookupMap(string lookupSrc, string resultSrc)
            {
                string cacheKey = $"{lookupSrc}||{resultSrc}";
                if (!lookupDictCache.TryGetValue(cacheKey, out var map))
                {
                    map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    var keys = RuleDataCache.GetSourceItems(wb, lookupSrc, preserveDuplicates: true);
                    var vals = RuleDataCache.GetSourceItems(wb, resultSrc, preserveDuplicates: true);
                    int len = Math.Min(keys.Count, vals.Count);
                    for (int i = 0; i < len; i++)
                    {
                        string k = keys[i].Trim();
                        if (!string.IsNullOrEmpty(k) && !map.ContainsKey(k))
                        {
                            map[k] = vals[i];
                            string normKey = k.Replace(" ", "_");
                            if (!map.ContainsKey(normKey)) map[normKey] = vals[i];
                        }
                    }
                    lookupDictCache[cacheKey] = map;
                }
                return map;
            }

            // Pre-compute max AutoIncrement ID per column in memory (0 COM calls)
            var maxAutoIdByCol = new long[colCount + 1];
            for (int c = 1; c <= colCount; c++)
            {
                if (colRules[c]?.DefaultRule?.Kind == DefaultKind.AutoIncrement)
                {
                    long maxId = 0;
                    for (int r = 1; r <= rowCount; r++)
                    {
                        string s = CleanCellText(dataGrid[r, c]);
                        if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long val) && val > maxId)
                            maxId = val;
                    }
                    maxAutoIdByCol[c] = maxId;
                }
            }

            // Track dirty cells to write back only what changed
            var dirtyWrites = new List<(int ExcelRow, int AbsCol, object Value, PropertyRuleDescriptor? Rule)>();

            // 5. Pure In-Memory Row Healing Loop
            for (int r = 1; r <= rowCount; r++)
            {
                int excelRow = startRow + r - 1;
                var rowValues = new string[colCount + 1];
                var rawObjects = new object?[colCount + 1];
                bool rowHasUserInput = false;
                for (int c = 1; c <= colCount; c++)
                {
                    object? valObj = dataGrid[r, c];
                    rawObjects[c] = valObj;
                    string clean = CleanCellText(valObj);
                    rowValues[c] = clean;

                    if (isFormulaCol[c] || string.IsNullOrWhiteSpace(clean)) continue;
                    if (clean == "0" || clean == "0.00" || clean == "-" || clean == "\u2014" ||
                        clean.Equals("None", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var rule = colRules[c];
                    if (rule?.DefaultRule?.Kind == DefaultKind.Literal &&
                        string.Equals(clean, rule.DefaultRule.LiteralValue, StringComparison.OrdinalIgnoreCase))
                        continue;

                    rowHasUserInput = true;
                }

                if (!rowHasUserInput) continue;

                // STEP 0: Apply [ColumnRule(Transform = ...)] in memory
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    string val = rowValues[c];
                    if (isFormulaCol[c] || rule == null || rule.Transform == TransformKind.None || string.IsNullOrWhiteSpace(val))
                        continue;

                    string transformed = ApplyTextTransform(val, rule.Transform);
                    if (!string.Equals(val, transformed, StringComparison.Ordinal))
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, transformed, rule));
                        rowValues[c] = transformed;
                        rawObjects[c] = transformed;
                    }
                }

                // STEP A: Fill Missing Partner Syncs Across the Row
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    string val = rowValues[c];
                    if (rule == null || string.IsNullOrWhiteSpace(val) || rule.SyncRules.Count == 0) continue;

                    foreach (var sync in rule.SyncRules)
                    {
                        int targetRelCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
                        if (targetRelCol < 1 || isFormulaCol[targetRelCol]) continue;

                        var targetRule = colRules[targetRelCol];
                        string currentTargetVal = rowValues[targetRelCol];

                        var map = GetLookupMap(sync.LookupSource, sync.ResultSource);
                        if (!map.TryGetValue(val.Trim(), out string? lookedUp))
                            map.TryGetValue(val.Trim().Replace(" ", "_"), out lookedUp);

                        if (!string.IsNullOrWhiteSpace(lookedUp))
                        {
                            bool isMismatch = !string.IsNullOrWhiteSpace(currentTargetVal) &&
                                              !RuleDataCache.ValuesMatch(currentTargetVal, lookedUp);

                            if (string.IsNullOrWhiteSpace(currentTargetVal) || (isMismatch && (targetRule?.IsWarningOnly ?? false)))
                            {
                                dirtyWrites.Add((excelRow, startCol + targetRelCol - 1, lookedUp, targetRule));
                                rowValues[targetRelCol] = lookedUp;
                                rawObjects[targetRelCol] = lookedUp;
                                if (isMismatch)
                                {
                                    allNotices.Add($"Row {excelRow} [{colHeaders[targetRelCol]}]: Auto-corrected '{currentTargetVal}' to '{lookedUp}' based on {colHeaders[c]} '{val}'.");
                                }
                            }
                        }
                    }
                }

                // STEP B: Apply [ColumnDefault] in memory
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    if (isFormulaCol[c] || rule?.DefaultRule == null || rule.DefaultRule.Kind == DefaultKind.None) continue;
                    if (!string.IsNullOrWhiteSpace(rowValues[c])) continue;

                    object? defaultVal = rule.DefaultRule.Kind == DefaultKind.AutoIncrement
                        ? ++maxAutoIdByCol[c]
                        : ComputeDefaultValue(wb, table, c, colHeaders, rawObjects, rowValues, rule.DefaultRule);

                    if (defaultVal != null)
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, defaultVal, rule));
                        rawObjects[c] = defaultVal;
                        rowValues[c] = FormatDisplayValue(defaultVal, rule);
                        allNotices.Add($"Row {excelRow} [{colHeaders[c]}]: Defaulted to '{rowValues[c]}'.");
                    }
                }

                // STEP C: Enforce [ColumnStateGuard] in memory
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    if (isFormulaCol[c] || rule == null || rule.StateGuards.Count == 0) continue;

                    string currentVal = rowValues[c];
                    string? enforcedValue = EvaluateStateGuardsForCell(colHeaders, rawObjects, rowValues, rule, currentVal, out string? guardReason);

                    if (!string.IsNullOrWhiteSpace(enforcedValue) && !RuleDataCache.ValuesMatch(currentVal, enforcedValue))
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, enforcedValue, rule));
                        rowValues[c] = enforcedValue;
                        rawObjects[c] = enforcedValue;
                        if (!string.IsNullOrWhiteSpace(guardReason))
                        {
                            allNotices.Add($"Row {excelRow} [{colHeaders[c]}]: {guardReason}");
                        }
                    }
                }
            }

            // 6. Write to Excel ONLY if at least one cell was actually healed
            if (dirtyWrites.Count > 0)
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                bool origEvents = true;
                try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

                // Lift any legacy sheet protection; it is never re-applied.
                if (SheetProtections.IsProtected(worksheet)) SheetProtections.Unprotect(worksheet);

                try
                {
                    foreach (var (excelRow, absCol, val, rule) in dirtyWrites)
                    {
                        WriteCellSafe(worksheet.Cells[excelRow, absCol], val, rule);
                    }
                }
                finally
                {
                    try { xlApp.EnableEvents = origEvents; } catch { }
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleEvaluator.EvaluateAndHealTable Error] {ex.Message}");
        }
        finally
        {
            _isEvaluating = false;
        }

        return allNotices;
    }

    /// <summary>
    /// Builds the interactive dropdown menu options for a cell, including section grouping and gray-out validation.
    /// </summary>
    public static List<DropdownOption> GetEvaluatedDropdownOptions(
        dynamic wb,
        dynamic? table,
        int excelRow,
        PropertyRuleDescriptor? rule)
    {
        var options = new List<DropdownOption>();
        if (wb == null || rule == null) return options;

        // 1. If the column has explicit list Sources, load them by section
        if (rule.Sources.Length > 0)
        {
            for (int sec = 0; sec < rule.Sources.Length; sec++)
            {
                var items = RuleDataCache.GetSourceItems(wb, rule.Sources[sec], preserveDuplicates: false);
                foreach (var rawItem in items)
                {
                    string display = ApplyTextTransform(FormatIfDate(rawItem), rule.Transform);
                    options.Add(new DropdownOption
                    {
                        Raw = rawItem,
                        Display = display,
                        Section = sec,
                        IsValid = true
                    });
                }
            }
        }
        // 2. If the column has no list Source, but has a [ColumnDefault] (e.g. AutoIncrement ID or Date offset), show the computed default as a 1-click option
        else if (rule.DefaultRule != null && rule.DefaultRule.Kind != DefaultKind.None && table != null)
        {
            int startCol = (int)table.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            dynamic ws = table.Parent;

            var colHeaders = new string[colCount + 1];
            var rawObjects = new object?[colCount + 1];
            var rowValues = new string[colCount + 1];
            int targetRelCol = 1;

            for (int c = 1; c <= colCount; c++)
            {
                int absCol = startCol + c - 1;
                colHeaders[c] = ws.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                rawObjects[c] = ws.Cells[excelRow, absCol].Value2;
                rowValues[c] = CleanCellText(rawObjects[c]);
                if (RuleRegistry.NormalizeIdentifier(colHeaders[c]).Equals(rule.NormalizedHeader, StringComparison.OrdinalIgnoreCase))
                    targetRelCol = c;
            }

            object? computed = ComputeDefaultValue(wb, table, targetRelCol, colHeaders, rawObjects, rowValues, rule.DefaultRule);
            if (computed != null)
            {
                string disp = FormatDisplayValue(computed, rule);
                options.Add(new DropdownOption { Raw = disp, Display = disp, Section = 0, IsValid = true });
            }
        }

        if (options.Count == 0 || table == null) return options;

        // 3. Evaluate relational gray-out (Partner Sync & StateGuards) on the current row
        try
        {
            dynamic ws = table.Parent;
            int startCol = (int)table.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;

            var colHeaders = new string[colCount + 1];
            var rawObjects = new object?[colCount + 1];
            var rowValues = new string[colCount + 1];

            for (int c = 1; c <= colCount; c++)
            {
                int absCol = startCol + c - 1;
                colHeaders[c] = ws.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                rawObjects[c] = ws.Cells[excelRow, absCol].Value2;
                rowValues[c] = CleanCellText(rawObjects[c]);
            }

            // A) Check if a partner column on this row already has a value that constrains this dropdown
            HashSet<string>? allowedByPartner = null;
            foreach (var sync in rule.SyncRules)
            {
                int partnerCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
                if (partnerCol >= 1 && !string.IsNullOrWhiteSpace(rowValues[partnerCol]))
                {
                    string partnerVal = rowValues[partnerCol];
                    // Find all items in LookupSource whose corresponding ResultSource equals partnerVal
                    var keys = RuleDataCache.GetSourceItems(wb, sync.LookupSource, preserveDuplicates: true);
                    var vals = RuleDataCache.GetSourceItems(wb, sync.ResultSource, preserveDuplicates: true);
                    int len = Math.Min(keys.Count, vals.Count);
                    allowedByPartner ??= new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                    for (int i = 0; i < len; i++)
                    {
                        if (RuleDataCache.ValuesMatch(vals[i], partnerVal))
                        {
                            allowedByPartner.Add(keys[i]);
                        }
                    }
                }
            }

            // B) Check StateGuards on this column
            var forbiddenStates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? forcedState = null;

            foreach (var guard in rule.StateGuards)
            {
                bool condTrue = EvaluateCondition(guard.Condition, guard.ReferenceHeaders, guard.ExpectedValue, colHeaders, rawObjects, rowValues);
                if (condTrue)
                {
                    forcedState = guard.ForcedValueWhenTrue;
                    break;
                }
                else if (guard.DisallowForcedValueWhenFalse)
                {
                    forbiddenStates.Add(guard.ForcedValueWhenTrue);
                }
            }

            // Update IsValid flags on options
            var updated = new List<DropdownOption>(options.Count);
            foreach (var opt in options)
            {
                bool valid = true;

                if (allowedByPartner != null && allowedByPartner.Count > 0)
                {
                    valid = allowedByPartner.Any(a => RuleDataCache.ValuesMatch(a, opt.Raw) || RuleDataCache.ValuesMatch(a, opt.Display));
                }

                if (valid && forcedState != null)
                {
                    valid = RuleDataCache.ValuesMatch(opt.Raw, forcedState);
                }
                else if (valid && forbiddenStates.Count > 0)
                {
                    if (forbiddenStates.Any(f => RuleDataCache.ValuesMatch(opt.Raw, f)))
                        valid = false;
                }

                updated.Add(new DropdownOption
                {
                    Raw = opt.Raw,
                    Display = opt.Display,
                    Section = opt.Section,
                    IsValid = valid
                });
            }

            return updated;
        }
        catch
        {
            return options;
        }
    }

    private static void ExecuteSyncRulesForColumn(
        dynamic wb,
        dynamic worksheet,
        dynamic table,
        int excelRow,
        int startCol,
        int colCount,
        string[] colHeaders,
        PropertyRuleDescriptor?[] colRules,
        string[] rowValues,
        object?[] rawObjects,
        PropertyRuleDescriptor sourceRule,
        string sourceVal)
    {
        var filledTargets = new HashSet<int>();

        foreach (var sync in sourceRule.SyncRules)
        {
            int targetRelCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
            if (targetRelCol < 1 || filledTargets.Contains(targetRelCol))
                continue;

            string? resolved = RuleDataCache.LookupValue(wb, sourceVal, sync.LookupSource, sync.ResultSource);
            if (!string.IsNullOrWhiteSpace(resolved))
            {
                int targetAbsCol = startCol + targetRelCol - 1;
                var targetRule = colRules[targetRelCol];
                string finalVal = targetRule != null
                    ? ApplyTextTransform(resolved, targetRule.Transform)
                    : resolved;

                WriteCellSafe(worksheet.Cells[excelRow, targetAbsCol], finalVal, targetRule);
                rowValues[targetRelCol] = finalVal;
                rawObjects[targetRelCol] = finalVal;
                filledTargets.Add(targetRelCol);
            }
        }
    }

    private static object? ComputeDefaultValue(
        dynamic wb,
        dynamic table,
        int targetRelCol,
        string[] colHeaders,
        object?[] rawObjects,
        string[] rowValues,
        ColumnDefaultAttribute defaultRule)
    {
        switch (defaultRule.Kind)
        {
            case DefaultKind.Literal:
                return defaultRule.LiteralValue;

            case DefaultKind.Today:
                return DateOnly.FromDateTime(DateTime.Today);

            case DefaultKind.DateOffsetDays:
                {
                    DateOnly anchor = DateOnly.FromDateTime(DateTime.Today);
                    if (!string.IsNullOrWhiteSpace(defaultRule.ReferenceHeader))
                    {
                        int refCol = FindColumnByHeader(colHeaders, defaultRule.ReferenceHeader);
                        if (refCol >= 1 && RuleDataCache.TryParseDate(rawObjects[refCol], out DateOnly parsedRef))
                        {
                            anchor = parsedRef;
                        }
                    }
                    return anchor.AddDays(defaultRule.Offset);
                }

            case DefaultKind.DateOffsetMonths:
                {
                    DateOnly anchor = DateOnly.FromDateTime(DateTime.Today);
                    if (!string.IsNullOrWhiteSpace(defaultRule.ReferenceHeader))
                    {
                        int refCol = FindColumnByHeader(colHeaders, defaultRule.ReferenceHeader);
                        if (refCol >= 1 && RuleDataCache.TryParseDate(rawObjects[refCol], out DateOnly parsedRef))
                        {
                            anchor = parsedRef;
                        }
                    }
                    return anchor.AddMonths(defaultRule.Offset);
                }

            case DefaultKind.AutoIncrement:
                {
                    long maxId = 0;
                    try
                    {
                        dynamic? colBody = table.ListColumns[targetRelCol].DataBodyRange;
                        if (colBody != null)
                        {
                            object? rawCol = colBody.Value2;
                            if (rawCol is object[,] mat)
                            {
                                int rStart = mat.GetLowerBound(0);
                                int rEnd = mat.GetUpperBound(0);
                                int cStart = mat.GetLowerBound(1);
                                for (int r = rStart; r <= rEnd; r++)
                                {
                                    string s = CleanCellText(mat[r, cStart]);
                                    if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long val) && val > maxId)
                                        maxId = val;
                                }
                            }
                            else if (rawCol != null)
                            {
                                string s = CleanCellText(rawCol);
                                if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long val) && val > maxId)
                                    maxId = val;
                            }
                        }
                    }
                    catch { }
                    return maxId + 1;
                }

            default:
                return null;
        }
    }

    internal static string? EvaluateStateGuardsForCell(
        string[] colHeaders,
        object?[] rawObjects,
        string[] rowValues,
        PropertyRuleDescriptor rule,
        string currentVal,
        out string? reason)
    {
        reason = null;
        if (rule.StateGuards.Count == 0) return null;

        foreach (var guard in rule.StateGuards)
        {
            if (!string.IsNullOrWhiteSpace(guard.SkipIfCurrentEquals) &&
            RuleDataCache.ValuesMatch(currentVal, guard.SkipIfCurrentEquals))
            {
                return null; // Keeps currentVal exactly as it is (e.g., "Reserved" or "Decommissioned")
            }

            bool condTrue = EvaluateCondition(guard.Condition, guard.ReferenceHeaders, guard.ExpectedValue, colHeaders, rawObjects, rowValues);
            if (condTrue)
            {
                if (!RuleDataCache.ValuesMatch(currentVal, guard.ForcedValueWhenTrue))
                {
                    reason = $"{rule.HeaderName} auto-set to '{guard.ForcedValueWhenTrue}' based on row date/state.";
                    return guard.ForcedValueWhenTrue;
                }
                return guard.ForcedValueWhenTrue;
            }
            else if (guard.DisallowForcedValueWhenFalse && RuleDataCache.ValuesMatch(currentVal, guard.ForcedValueWhenTrue))
            {
                reason = $"'{guard.ForcedValueWhenTrue}' cannot be selected when all dates/conditions are valid. Auto-reverted to '{guard.FallbackValueWhenFalse}'.";
                return guard.FallbackValueWhenFalse;
            }
        }

        if (string.IsNullOrWhiteSpace(currentVal) && rule.DefaultRule?.Kind == DefaultKind.Literal)
        {
            return rule.DefaultRule.LiteralValue;
        }

        return null;
    }

    internal static bool EvaluateCondition(
        StateConditionKind condition,
        string[] referenceHeaders,
        string? expectedValue,
        string[] colHeaders,
        object?[] rawObjects,
        string[] rowValues,
        string? currentCellVal = null)
    {
        if (referenceHeaders == null || referenceHeaders.Length == 0) return false;
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);

        foreach (var refHeader in referenceHeaders)
        {
            int c = FindColumnByHeader(colHeaders, refHeader);
            if (c < 1) continue;

            object? raw = rawObjects[c];
            string str = rowValues[c];

            switch (condition)
            {
                case StateConditionKind.DateExpired:
                    if (RuleDataCache.TryParseDate(raw, out DateOnly d) && d < today)
                        return true;
                    break;

                case StateConditionKind.ColumnNotEmpty:
                    if (!string.IsNullOrWhiteSpace(str) && str != "—" && str != "0")
                        return true;
                    break;

                case StateConditionKind.ColumnEquals:
                    if (RuleDataCache.ValuesMatch(str, expectedValue))
                        return true;
                    break;

                case StateConditionKind.DateGreaterThan:
                    if (RuleDataCache.TryParseDate(currentCellVal, out DateOnly thisDate) &&
                        RuleDataCache.TryParseDate(str, out DateOnly refDate))
                    {
                        if (refDate >= thisDate)
                            return true;
                    }
                    break;
            }
        }

        return false;
    }

    internal static int FindColumnByHeader(string[] colHeaders, string targetHeader)
    {
        if (string.IsNullOrWhiteSpace(targetHeader)) return -1;
        string normTarget = RuleRegistry.NormalizeIdentifier(targetHeader);

        for (int c = 1; c < colHeaders.Length; c++)
        {
            string normCol = RuleRegistry.NormalizeIdentifier(colHeaders[c]);
            if (normCol.Equals(normTarget, StringComparison.OrdinalIgnoreCase) ||
                normCol.TrimEnd('s').Equals(normTarget.TrimEnd('s'), StringComparison.OrdinalIgnoreCase))
            {
                return c;
            }
        }
        return -1;
    }

    public static string ApplyTextTransform(string input, TransformKind transform)
    {
        if (string.IsNullOrWhiteSpace(input)) return string.Empty;
        string trimmed = input.Trim();

        return transform switch
        {
            TransformKind.Trim => trimmed,
            TransformKind.UpperCase => trimmed.ToUpperInvariant(),
            TransformKind.ShortNameFirstLast => FormatShortName(trimmed, maxWords: 2),
            TransformKind.ShortNameThreeWords => FormatShortName(trimmed, maxWords: 3),
            _ => trimmed
        };
    }

    private static string FormatShortName(string fullName, int maxWords)
    {
        var parts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length <= maxWords) return string.Join(" ", parts);
        if (maxWords == 2) return $"{parts[0]} {parts[^1]}";
        return $"{parts[0]} {parts[1]} {parts[^1]}";
    }

    private static void WriteCellSafe(dynamic cell, object value, PropertyRuleDescriptor? rule)
    {
        try
        {
            if (value is DateOnly dOnly)
            {
                cell.Value2 = dOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
                return;
            }
            if (value is DateTime dt)
            {
                cell.Value2 = dt.ToOADate();
                return;
            }
            if (rule != null && rule.IsDateColumn && RuleDataCache.TryParseDate(value, out DateOnly parsedDate))
            {
                cell.Value2 = parsedDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                return;
            }
            if (rule != null && rule.IsNumericColumn &&
                double.TryParse(value.ToString(), NumberStyles.Any, CultureInfo.InvariantCulture, out double dNum))
            {
                cell.Value2 = dNum;
                return;
            }

            cell.Value2 = value.ToString();
        }
        catch { }
    }

    public static string FormatDisplayValue(object? val, PropertyRuleDescriptor? rule)
    {
        if (val == null) return string.Empty;
        if (val is DateOnly d) return d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        if (val is DateTime dt) return dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        if (rule != null && rule.IsDateColumn && RuleDataCache.TryParseDate(val, out DateOnly parsed))
            return parsed.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        return CleanCellText(val);
    }

    private static string FormatIfDate(string input)
    {
        if (RuleDataCache.TryParseDate(input, out DateOnly d) &&
            (input.Contains('/') || input.Contains('-') || (double.TryParse(input, out double oa) && oa > 30000)))
        {
            return d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        }
        return input;
    }

    private static string CleanCellText(object? val)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        if (val is double d && Math.Abs(d - Math.Round(d)) < 0.00001 && d < 1e15)
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (s.StartsWith("#") || s == "—") return string.Empty;
        return s;
    }
}