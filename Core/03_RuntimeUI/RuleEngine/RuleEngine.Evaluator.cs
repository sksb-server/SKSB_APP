namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;

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

    private static readonly HashSet<string> OdometerHeaders = new(StringComparer.OrdinalIgnoreCase)
    {
        "currentodometer", "odometerbefore", "odometerin", "odometerout", "odometerdifference"
    };

    private static string SquashHeaderText(string? h) =>
        new string((h ?? string.Empty).Where(char.IsLetterOrDigit).ToArray()).ToLowerInvariant();

    /// <summary>Formats odometer cells of one row as hours (machinery) or km, based on the row's Vehicle Type.</summary>
    private static void ApplyOdometerUnitFormat(
        dynamic worksheet, int excelRow, int startCol, int colCount, string[] colHeaders, string[] rowValues)
    {
        int typeCol = -1;
        for (int c = 1; c <= colCount; c++)
            if (SquashHeaderText(colHeaders[c]) == "vehicletype") { typeCol = c; break; }
        if (typeCol < 0) return;

        string typeText = (rowValues[typeCol] ?? string.Empty).Trim().Replace(' ', '_').Replace('-', '_');
        bool isHours = Enum.TryParse<SKSB_App.Core.Types.VehicleType>(typeText, true, out var vt)
                       && SKSB_App.Core.Types.VehicleRecord.IsHourMeterType(vt);

        for (int c = 1; c <= colCount; c++)
        {
            string squashed = SquashHeaderText(colHeaders[c]);
            if (!OdometerHeaders.Contains(squashed)) continue;
            bool isDiff = squashed == "odometerdifference";
            string fmt = isHours
                ? (isDiff ? RuleRegistry.OdometerDiffHoursFormat : RuleRegistry.OdometerHoursFormat)
                : (isDiff ? RuleRegistry.OdometerDiffKmFormat : RuleRegistry.OdometerKmFormat);
            try
            {
                dynamic cell = worksheet.Cells[excelRow, startCol + c - 1];
                string? curFmt = cell.NumberFormat?.ToString();
                if (!string.Equals(curFmt, fmt, StringComparison.Ordinal))
                {
                    cell.NumberFormat = fmt;
                }
            }
            catch { }
        }
    }

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
        int? editedAbsoluteCol = null)
    {
        var notices = new List<string>();
        if (_isEvaluating || worksheet == null || table == null) return notices;

        try
        {
            _isEvaluating = true;
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = worksheet!.Parent;

            bool origEvents = true;
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

            try
            {
                string tableName = table!.Name?.ToString() ?? string.Empty;
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
                    rowValues[c] = SheetMapperHelper.CleanCellText(valObj);
                }

                bool rowHasUserInput = false;
                for (int c = 1; c <= colCount; c++)
                {
                    int absCol = startCol + c - 1;
                    var rule = colRules[c];
                    if (rule?.IsLocked == true) continue;

                    string f = worksheet.Cells[excelRow, absCol].Formula?.ToString() ?? string.Empty;
                    if (f.StartsWith("=")) continue;

                    string clean = rowValues[c];
                    if (string.IsNullOrWhiteSpace(clean)) continue;
                    if (clean == "0" || clean == "0.00" || clean == "-" || clean == "\u2014" ||
                        clean.Equals("None", StringComparison.OrdinalIgnoreCase))
                        continue;

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

                            if (editedRule?.RollupRules.Count > 0)
                            {
                                SheetOperations.ShowStatusConfirmation($"Formula evaluated cell [{editedRule.HeaderName}]: {editedRule.DisplayFormula}");
                            }

                            if (editedRule != null && editedRule.RoundDownStep > 0 &&
                                decimal.TryParse(currentText, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal typed) && typed > 0m)
                            {
                                decimal floored = RuleCalculator.FloorToStep(typed, (decimal)editedRule.RoundDownStep);
                                if (floored != typed)
                                {
                                    WriteCellSafe(worksheet.Cells[excelRow, editedAbsoluteCol.Value], (double)floored, editedRule);
                                    rawObjects[relEdited] = (double)floored;
                                    rowValues[relEdited] = SheetMapperHelper.CleanCellText((double)floored);
                                    currentText = rowValues[relEdited];
                                }
                            }
                            if (editedRule != null && string.IsNullOrWhiteSpace(currentText) &&
                                editedRule.IsMultiSelect && !string.IsNullOrWhiteSpace(editedRule.CountTargetHeader))
                            {
                                WriteMultiSelectCount(worksheet, excelRow, startCol, colHeaders, rowValues, editedRule, string.Empty);
                            }
                            if (editedRule != null && !string.IsNullOrWhiteSpace(currentText))
                            {
                                string transformed = ApplyRuleTransform(currentText, editedRule);
                                if (!string.Equals(transformed, currentText, StringComparison.Ordinal))
                                {
                                    WriteCellSafe(worksheet.Cells[excelRow, editedAbsoluteCol.Value], transformed, editedRule);
                                    rowValues[relEdited] = transformed;
                                    currentText = transformed;
                                }
                                if (editedRule.IsMultiSelect && !string.IsNullOrWhiteSpace(editedRule.CountTargetHeader))
                                {
                                    WriteMultiSelectCount(worksheet, excelRow, startCol, colHeaders, rowValues, editedRule, currentText);
                                }
                                if (editedRule.SyncRules.Count > 0)
                                {
                                    ExecuteSyncRulesForColumn(wb, worksheet, excelRow, startCol, colHeaders, colRules, rowValues, rawObjects, editedRule, currentText);
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

                        object? defaultVal = ComputeDefaultValue(table, c, colHeaders, rawObjects, rule.DefaultRule);
                        if (defaultVal != null)
                        {
                            WriteCellSafe(worksheet.Cells[excelRow, absCol], defaultVal, rule);
                            rawObjects[c] = defaultVal;
                            rowValues[c] = FormatDisplayValue(defaultVal, rule);
                            notices.Add($"Row {excelRow} [{colHeaders[c]}]: Defaulted to '{rowValues[c]}'.");
                        }
                    }

                    // STEP 3b: Evaluate [ColumnCalculate] (2 passes so dependent columns settle)
                    if (colRules.Any(r => r?.CalculateRule != null))
                    {
                        var sheetParams = RuleCalculator.ReadSheetParameters((object?)wb, (object?)worksheet, FirstSheetParameters(colRules));
                        Func<string, decimal> resolveParam = n => sheetParams.TryGetValue(n, out var pv) ? pv : 0m;

                        for (int pass = 0; pass < 2; pass++)
                        {
                            for (int c = 1; c <= colCount; c++)
                            {
                                var rule = colRules[c];
                                if (rule?.CalculateRule == null) continue;

                                int absCol = startCol + c - 1;
                                bool isDefaultOnly = rule.CalculateRule.Mode == CalculationMode.DefaultOnly;
                                bool directlyEdited = editedAbsoluteCol == absCol;

                                if (isDefaultOnly &&
                                    !(string.IsNullOrWhiteSpace(rowValues[c]) || rowValues[c] == "0" || rowValues[c] == "0.00"))
                                    continue;

                                // null => "no value" (em-dash); 0 => a genuine zero
                                decimal? computed = RuleCalculator.EvaluateColumnRule(rule, colHeaders, rawObjects, resolveParam);
                                if (isDefaultOnly && (!computed.HasValue || computed.Value == 0m)) continue;

                                bool differs = RuleCalculator.ToNullableDecimal(rawObjects[c]) != computed;
                                if (!differs && !(directlyEdited && !isDefaultOnly)) continue;

                                if (differs)
                                {
                                    object written = computed.HasValue ? (double)computed.Value : NullDash;
                                    WriteCellSafe(worksheet.Cells[excelRow, absCol], written, rule);
                                    rawObjects[c] = written;
                                    rowValues[c] = SheetMapperHelper.CleanCellText(written);
                                }

                                if (directlyEdited && !isDefaultOnly && pass == 0)
                                {
                                    SheetOperations.ShowStatusConfirmation(
                                        $"Formula evaluated cell [{rule.HeaderName}]: {rule.DisplayFormula} = {(computed.HasValue ? computed.Value.ToString(CultureInfo.InvariantCulture) : NullDash)}");
                                }
                            }
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

                    // STEP 5: Unit display for odometer cells (hrs for machinery, km for road vehicles)
                    ApplyOdometerUnitFormat(worksheet, excelRow, startCol, colCount, colHeaders, rowValues);
                });

                healAction();
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
    public static List<ValidationIssue> EvaluateAndHealTable(dynamic worksheet, dynamic table, Func<DateOnly, TimeOnly?, bool>? warningDateFilter = null)
    {
        var allNotices = new List<ValidationIssue>();
        if (_isEvaluating || worksheet == null || table == null) return allNotices;

        try
        {
            _isEvaluating = true;
            dynamic? body = table!.DataBodyRange;
            if (body == null) return allNotices;

            int startRow = (int)body.Row;
            int rowCount = (int)body.Rows.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            if (rowCount <= 0 || colCount <= 0) return allNotices;

            string tableName = table.Name?.ToString() ?? string.Empty;
            string sheetName = string.Empty; try { sheetName = worksheet!.Name?.ToString() ?? string.Empty; } catch { }
            dynamic wb = worksheet!.Parent;

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

            int dateCol = FindColumnByHeader(colHeaders, "Date");
            int timeCol = FindColumnByHeader(colHeaders, "Time");

            // 2. Bulk-read archetype formulas on first data row to identify formula columns (1 COM call)
            var archFormulas = SheetMapperHelper.ReadArchetypeFormulas((object)worksheet, startRow, startCol, colCount);
            var isFormulaCol = new bool[colCount + 1];
            for (int c = 1; c <= colCount; c++)
                isFormulaCol[c] = archFormulas[c].StartsWith("=") && colRules[c]?.HasCSharpFormula != true;

            // 3. Bulk-read entire DataBodyRange into memory (1 COM call)
            object[,] dataGrid = TableMapper.GetRange2DValues(body, rowCount, colCount);

            // 3b. Pre-cache sheet parameters once per table (0 COM calls inside the row loop)
            var sheetParams = RuleCalculator.ReadSheetParameters((object?)wb, (object?)worksheet, FirstSheetParameters(colRules));
            Func<string, decimal> resolveParam = n => sheetParams.TryGetValue(n, out var pv) ? pv : 0m;

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
                        string s = SheetMapperHelper.CleanCellText(dataGrid[r, c]);
                        if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long val) && val > maxId)
                            maxId = val;
                    }
                    maxAutoIdByCol[c] = maxId;
                }
            }

            // Track dirty cells to write back only what changed
            var dirtyWrites = new List<(int ExcelRow, int AbsCol, object Value, PropertyRuleDescriptor? Rule)>();

            // 4b. Cross-table [ColumnRollup] (e.g. Workshops <- Maintenances): one bulk read per source table, all in RAM
            if (colRules.Any(r => r != null && r.RollupRules.Any(x => x.IsCrossTable)))
            {
                var rowValuesGrid = new string[rowCount + 1, colCount + 1];
                var rawGrid = new object?[rowCount + 1, colCount + 1];
                for (int rr = 1; rr <= rowCount; rr++)
                    for (int cc = 1; cc <= colCount; cc++)
                    {
                        rawGrid[rr, cc] = dataGrid[rr, cc];
                        rowValuesGrid[rr, cc] = SheetMapperHelper.CleanCellText(dataGrid[rr, cc]);
                    }

                RuleCalculator.EvaluateTableCrossRollups((object)wb, (object)worksheet, colRules, colHeaders, rowCount, colCount,
                    rowValuesGrid, rawGrid,
                    (r, c, val, rule) =>
                    {
                        dirtyWrites.Add((startRow + r - 1, startCol + c - 1, val, rule));
                        dataGrid[r, c] = val;
                    });
            }

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
                    string clean = SheetMapperHelper.CleanCellText(valObj);
                    rowValues[c] = clean;

                    if (isFormulaCol[c] || string.IsNullOrWhiteSpace(clean)) continue;
                    if (clean == "0" || clean == "0.00" || clean == "-" || clean == "\u2014" ||
                        clean.Equals("None", StringComparison.OrdinalIgnoreCase))
                        continue;

                    var rule = colRules[c];
                    if (rule?.IsLocked == true) continue;
                    if (rule?.DefaultRule?.Kind == DefaultKind.Literal &&
                        string.Equals(clean, rule.DefaultRule.LiteralValue, StringComparison.OrdinalIgnoreCase))
                        continue;

                    rowHasUserInput = true;
                }

                if (!rowHasUserInput) continue;

                bool emitNotices = warningDateFilter == null || dateCol < 1 ||
                    RuleDataCache.MatchesWarningDateFilter(dataGrid[r, dateCol], timeCol >= 1 ? dataGrid[r, timeCol] : null, warningDateFilter);

                // STEP 0: Apply [ColumnRule(Transform = ...)] in memory
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    string val = rowValues[c];
                    if (isFormulaCol[c] || rule == null || rule.Transform == TransformKind.None || string.IsNullOrWhiteSpace(val))
                        continue;

                    string transformed = ApplyRuleTransform(val, rule);
                    if (!string.Equals(val, transformed, StringComparison.Ordinal))
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, transformed, rule));
                        rowValues[c] = transformed;
                        rawObjects[c] = transformed;
                    }
                }

                // STEP 0a: enforce [ColumnRule(RoundDownStep = ...)] in memory
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    if (isFormulaCol[c] || rule == null || rule.RoundDownStep <= 0) continue;
                    if (!decimal.TryParse(rowValues[c], NumberStyles.Any, CultureInfo.InvariantCulture, out decimal typed) || typed <= 0m)
                        continue;

                    decimal floored = RuleCalculator.FloorToStep(typed, (decimal)rule.RoundDownStep);
                    if (floored != typed)
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, (double)floored, rule));
                        rawObjects[c] = (double)floored;
                        rowValues[c] = SheetMapperHelper.CleanCellText((double)floored);
                    }
                }

                // STEP 0b: multi-select count targets (e.g. Workers -> Worker Presence); only when tokens exist,
                // so a manually typed count is kept for rows that have no token list.
                for (int c = 1; c <= colCount; c++)
                {
                    var rule = colRules[c];
                    if (isFormulaCol[c] || rule == null || !rule.IsMultiSelect || string.IsNullOrWhiteSpace(rule.CountTargetHeader)) continue;
                    if (string.IsNullOrWhiteSpace(rowValues[c])) continue;

                    int targetRel = FindColumnByHeader(colHeaders, rule.CountTargetHeader!);
                    if (targetRel < 1 || isFormulaCol[targetRel]) continue;

                    int count = CountMultiSelectTokens(rowValues[c]);
                    if (!string.Equals(rowValues[targetRel], count.ToString(CultureInfo.InvariantCulture), StringComparison.Ordinal))
                    {
                        dirtyWrites.Add((excelRow, startCol + targetRel - 1, (double)count, colRules[targetRel]));
                        rowValues[targetRel] = count.ToString(CultureInfo.InvariantCulture);
                        rawObjects[targetRel] = (double)count;
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
                                if (isMismatch && emitNotices)
                                {
                                    allNotices.Add(new ValidationIssue(ValidationSeverity.Warning, sheetName, tableName, excelRow, startCol + targetRelCol - 1, colHeaders[targetRelCol], $"Auto-corrected '{currentTargetVal}' to '{lookedUp}' based on {colHeaders[c]} '{val}'."));
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
                        : ComputeDefaultValue(table, c, colHeaders, rawObjects, rule.DefaultRule);

                    if (defaultVal != null)
                    {
                        dirtyWrites.Add((excelRow, startCol + c - 1, defaultVal, rule));
                        rawObjects[c] = defaultVal;
                        rowValues[c] = FormatDisplayValue(defaultVal, rule);
                        if (emitNotices)
                            allNotices.Add(new ValidationIssue(ValidationSeverity.Warning, sheetName, tableName, excelRow, startCol + c - 1, colHeaders[c], $"Defaulted to '{rowValues[c]}'."));
                    }
                }

                // STEP B2: Evaluate [ColumnCalculate] in memory (2 passes so dependent columns settle)
                for (int pass = 0; pass < 2; pass++)
                {
                    for (int c = 1; c <= colCount; c++)
                    {
                        var rule = colRules[c];
                        if (rule?.CalculateRule == null) continue;

                        bool isDefaultOnly = rule.CalculateRule.Mode == CalculationMode.DefaultOnly;
                        if (isDefaultOnly &&
                            !(string.IsNullOrWhiteSpace(rowValues[c]) || rowValues[c] == "0" || rowValues[c] == "0.00"))
                            continue;

                        decimal? computed = RuleCalculator.EvaluateColumnRule(rule, colHeaders, rawObjects, resolveParam);
                        if (isDefaultOnly && (!computed.HasValue || computed.Value == 0m)) continue;
                        if (RuleCalculator.ToNullableDecimal(rawObjects[c]) == computed) continue;

                        object written = computed.HasValue ? (double)computed.Value : NullDash;
                        dirtyWrites.Add((excelRow, startCol + c - 1, written, rule));
                        rawObjects[c] = written;
                        rowValues[c] = SheetMapperHelper.CleanCellText(written);
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
                        if (emitNotices && !string.IsNullOrWhiteSpace(guardReason))
                        {
                            allNotices.Add(new ValidationIssue(ValidationSeverity.Warning, sheetName, tableName, excelRow, startCol + c - 1, colHeaders[c], guardReason));
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
            AppLogger.Error($"[RuleEvaluator.EvaluateAndHealTable] Failed healing '{table?.Name}'", ex);
        }
        finally
        {
            _isEvaluating = false;
        }

        return allNotices;
    }

    /// <summary>Maps a Vehicle Plate / Staff ID source to its index-aligned companion list (Type / Name) for enriched dropdown labels.</summary>
    private static Dictionary<string, string>? BuildCompanionLabelMap(dynamic wb, string sourceName, out bool isStaffId)
    {
        isStaffId = false;
        string key = (sourceName ?? string.Empty).Trim();
        string[] companionSources;
        if (key.Equals("AllVehiclesTable[Vehicle Plate]", StringComparison.OrdinalIgnoreCase))
            companionSources = new[] { "AllVehiclesTable[Vehicle Type]" };
        else if (key.Equals("VehiclePlateList", StringComparison.OrdinalIgnoreCase))
            companionSources = new[] { "VehiclePlateTypeList" };
        else if (key.Equals("AllStaffsTable[Staff ID]", StringComparison.OrdinalIgnoreCase))
        { isStaffId = true; companionSources = new[] { "AllStaffsTable[Staff Name]" }; }
        else if (key.Equals("StaffIDList", StringComparison.OrdinalIgnoreCase))
        { isStaffId = true; companionSources = new[] { "StaffShortNameList", "StaffNameList" }; }
        else if (key.Equals("AllFieldStaffIDList", StringComparison.OrdinalIgnoreCase))
        { isStaffId = true; companionSources = new[] { "AllFieldStaffNameList" }; }
        else return null;

        try
        {
            var keys = RuleDataCache.GetSourceItems(wb, sourceName!, preserveDuplicates: true);
            foreach (var comp in companionSources)
            {
                var vals = RuleDataCache.GetSourceItems(wb, comp, preserveDuplicates: true);
                if (vals.Count == 0) continue;
                var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                int len = Math.Min(keys.Count, vals.Count);
                for (int i = 0; i < len; i++)
                {
                    string v = (vals[i] ?? string.Empty).Trim();
                    if (v.Length == 0 || v.Equals("None", StringComparison.OrdinalIgnoreCase)) continue;
                    if (!string.IsNullOrWhiteSpace(keys[i])) map.TryAdd(keys[i].Trim(), v);
                }
                if (map.Count > 0) return map;
            }
        }
        catch { }
        return null;
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
        if (rule!.Sources.Length > 0)
        {
            for (int sec = 0; sec < rule.Sources.Length; sec++)
            {
                var items = RuleDataCache.GetSourceItems(wb, rule.Sources[sec], preserveDuplicates: false);
                var companions = BuildCompanionLabelMap(wb, rule.Sources[sec], out bool isStaffId);
                foreach (var rawItem in items)
                {
                    string transformed = ApplyTextTransform(FormatIfDate(rawItem), rule.Transform);
                    string display = transformed;
                    string raw = rawItem;
                    string? label = null;
                    if (companions != null &&
                        (companions.TryGetValue(rawItem, out label) || companions.TryGetValue(transformed, out label)) &&
                        !string.IsNullOrWhiteSpace(label))
                    {
                        raw = transformed;
                        display = isStaffId && long.TryParse(transformed, out long sid)
                            ? $"{sid:D4} — {label}"
                            : $"{transformed} — {label}";
                    }
                    options.Add(new DropdownOption
                    {
                        Raw = raw,
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
            int startCol = (int)table!.HeaderRowRange.Column;
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
                rowValues[c] = SheetMapperHelper.CleanCellText(rawObjects[c]);
                if (RuleRegistry.NormalizeIdentifier(colHeaders[c]).Equals(rule.NormalizedHeader, StringComparison.OrdinalIgnoreCase))
                    targetRelCol = c;
            }

            object? computed = ComputeDefaultValue(table, targetRelCol, colHeaders, rawObjects, rule.DefaultRule);
            if (computed != null)
            {
                string disp = FormatDisplayValue(computed, rule);
                options.Add(new DropdownOption { Raw = disp, Display = disp, Section = 0, IsValid = true });
            }
        }

        // 3. Soft-calculated column ([ColumnCalculate] DefaultOnly): offer the computed value as a 1-click option
        else if (rule.CalculateRule != null && rule.CalculateRule.Mode == CalculationMode.DefaultOnly && table != null)
        {
            int startCol = (int)table!.HeaderRowRange.Column;
            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            dynamic ws = table.Parent;

            var colHeaders = new string[colCount + 1];
            var rawObjects = new object?[colCount + 1];
            for (int c = 1; c <= colCount; c++)
            {
                int absCol = startCol + c - 1;
                colHeaders[c] = ws.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                rawObjects[c] = ws.Cells[excelRow, absCol].Value2;
            }

            var sheetParams = RuleCalculator.ReadSheetParameters((object?)wb, (object?)ws, rule.SheetParameters);
            decimal? computed = RuleCalculator.EvaluateColumnRule(
                rule, colHeaders, rawObjects, n => sheetParams.TryGetValue(n, out var pv) ? pv : 0m);

            if (computed.HasValue)
            {
                string disp = computed.Value.ToString("0.############", CultureInfo.InvariantCulture);
                options.Add(new DropdownOption { Raw = disp, Display = disp, Section = 0, IsValid = true });
            }
        }

        if (options.Count == 0 || table == null) return options;

        // 3. Evaluate relational gray-out (Partner Sync & StateGuards) on the current row
        try
        {
            dynamic ws = table!.Parent;
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
                rowValues[c] = SheetMapperHelper.CleanCellText(rawObjects[c]);
            }

            // A) Check if a partner column on this row already has a value that constrains this dropdown
            var allowedByPartnerCol = new Dictionary<int, HashSet<string>>();
            foreach (var sync in rule.SyncRules)
            {
                if (sync.OnlyWhenTargetBlank) continue;
                int partnerCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
                if (partnerCol >= 1 && !string.IsNullOrWhiteSpace(rowValues[partnerCol]))
                {
                    string partnerVal = rowValues[partnerCol];
                    // Find all items in LookupSource whose corresponding ResultSource equals partnerVal
                    var keys = RuleDataCache.GetSourceItems(wb, sync.LookupSource, preserveDuplicates: true);
                    var vals = RuleDataCache.GetSourceItems(wb, sync.ResultSource, preserveDuplicates: true);
                    int len = Math.Min(keys.Count, vals.Count);
                    if (len == 0) continue;
                    if (!allowedByPartnerCol.TryGetValue(partnerCol, out var colAllowed))
                    {
                        colAllowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        allowedByPartnerCol[partnerCol] = colAllowed;
                    }

                    for (int i = 0; i < len; i++)
                    {
                        if (RuleDataCache.ValuesMatch(vals[i], partnerVal))
                        {
                            colAllowed.Add(keys[i]);
                        }
                    }
                }
            }

            // B) Check StateGuards on this column
            var forbiddenStates = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            string? forcedState = null;
            string? forcedSkipList = null;

            foreach (var guard in rule.StateGuards)
            {
                bool condTrue = EvaluateCondition(guard.Condition, guard.ReferenceHeaders, guard.ExpectedValue, colHeaders, rawObjects, rowValues);
                if (condTrue)
                {
                    forcedState = guard.ForcedValueWhenTrue;
                    // Comma-separated SkipIfCurrentEquals values stay selectable (they are preserved by EvaluateStateGuardsForCell)
                    forcedSkipList = guard.SkipIfCurrentEquals;
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

                if (allowedByPartnerCol.Count > 0)
                {
                    valid = allowedByPartnerCol.Values.All(allowed =>
                        allowed.Any(a => RuleDataCache.ValuesMatch(a, opt.Raw) || RuleDataCache.ValuesMatch(a, opt.Display)));
                }

                if (valid && forcedState != null)
                {
                    valid = RuleDataCache.ValuesMatch(opt.Raw, forcedState) || MatchesSkipList(opt.Raw, forcedSkipList);
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
        int excelRow,
        int startCol,
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

            if (sync.OnlyWhenTargetBlank && !string.IsNullOrWhiteSpace(rowValues[targetRelCol]))
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
        dynamic table,
        int targetRelCol,
        string[] colHeaders,
        object?[] rawObjects,
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
                                    string s = SheetMapperHelper.CleanCellText(mat[r, cStart]);
                                    if (long.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out long val) && val > maxId)
                                        maxId = val;
                                }
                            }
                            else if (rawCol != null)
                            {
                                string s = SheetMapperHelper.CleanCellText(rawCol);
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
            if (MatchesSkipList(currentVal, guard.SkipIfCurrentEquals))
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

    /// <summary>
    /// True if <paramref name="value"/> matches any comma-separated token of <paramref name="skipList"/>
    /// (e.g. <c>SkipIfCurrentEquals = "In_Shop,Reserved,Suspended"</c>).
    /// </summary>
    internal static bool MatchesSkipList(string? value, string? skipList)
    {
        if (string.IsNullOrWhiteSpace(skipList) || string.IsNullOrWhiteSpace(value)) return false;
        foreach (var token in skipList.Split(','))
        {
            string t = token.Trim();
            if (t.Length > 0 && RuleDataCache.ValuesMatch(value, t)) return true;
        }
        return false;
    }

    private static bool TryParseDecimalCell(object? raw, string? str, out decimal value)
    {
        value = 0m;
        switch (raw)
        {
            case double dbl: value = (decimal)dbl; return true;
            case decimal dec: value = dec; return true;
            case float f: value = (decimal)f; return true;
            case int i: value = i; return true;
            case long l: value = l; return true;
            case short s: value = s; return true;
        }

        string text = (str ?? raw?.ToString() ?? string.Empty).Trim();
        if (text.Length == 0 || text == "—") return false;
        return decimal.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out value);
    }

    internal static bool EvaluateCondition(
        StateConditionKind condition,
        string[] referenceHeaders,
        string? expectedValue,
        string[] colHeaders,
        object?[] rawObjects,
        string[] rowValues)
    {
        if (referenceHeaders == null || referenceHeaders.Length == 0) return false;
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);

        if (condition == StateConditionKind.DateExpiredOrNumberReached)
        {
            // [0] = date column, [1] = current numeric column, [2] = target numeric column
            if (referenceHeaders.Length < 3) return false;

            int cDate = FindColumnByHeader(colHeaders, referenceHeaders[0]);
            if (cDate >= 1 && RuleDataCache.TryParseDate(rawObjects[cDate], out DateOnly due) && due.Year > 1900 && due < today)
                return true;

            int cCur = FindColumnByHeader(colHeaders, referenceHeaders[1]);
            int cTgt = FindColumnByHeader(colHeaders, referenceHeaders[2]);
            if (cCur >= 1 && cTgt >= 1 &&
                TryParseDecimalCell(rawObjects[cCur], rowValues[cCur], out decimal current) &&
                TryParseDecimalCell(rawObjects[cTgt], rowValues[cTgt], out decimal target) &&
                target > 0m && current >= target)
                return true;

            return false;
        }

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

                case StateConditionKind.DateReached:
                    if (RuleDataCache.TryParseDate(raw, out DateOnly dr) && dr.Year > 1900 && dr <= today)
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

    /// <summary>
    /// Applies the rule's transform. Multi-select values ("A, B") are transformed token by token and
    /// re-joined with ", " so the separator can never be mangled (e.g. by UpperCase or ShortName rules).
    /// </summary>
    public static string ApplyRuleTransform(string input, PropertyRuleDescriptor rule)
    {
        if (!rule.IsMultiSelect) return ApplyTextTransform(input, rule.Transform);

        var tokens = SplitMultiSelectTokens(input)
            .Select(t => ApplyTextTransform(t, rule.Transform))
            .Where(t => t.Length > 0);
        return string.Join(", ", tokens);
    }

    /// <summary>Splits a comma separated multi-select cell into distinct, trimmed, non-empty tokens.</summary>
    public static List<string> SplitMultiSelectTokens(string? text) =>
        (text ?? string.Empty)
            .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(t => t.Trim())
            .Where(t => t.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    public static int CountMultiSelectTokens(string? text) => SplitMultiSelectTokens(text).Count;

    /// <summary>The [SheetParameter] declarations shared by the table's columns (identical on every rule of one table).</summary>
    private static IReadOnlyList<SheetParameterAttribute> FirstSheetParameters(PropertyRuleDescriptor?[] colRules) =>
        colRules.FirstOrDefault(r => r != null && r.SheetParameters.Count > 0)?.SheetParameters
        ?? Array.Empty<SheetParameterAttribute>();

    private static void WriteMultiSelectCount(
        dynamic worksheet, int excelRow, int startCol, string[] colHeaders, string[] rowValues,
        PropertyRuleDescriptor rule, string multiText)
    {
        int targetRel = FindColumnByHeader(colHeaders, rule.CountTargetHeader!);
        if (targetRel < 1) return;

        int count = CountMultiSelectTokens(multiText);
        string countText = count.ToString(CultureInfo.InvariantCulture);
        if (string.Equals(rowValues[targetRel], countText, StringComparison.Ordinal)) return;

        try { worksheet.Cells[excelRow, startCol + targetRel - 1].Value2 = (double)count; } catch { }
        rowValues[targetRel] = countText;
    }

    private static readonly HashSet<string> ChineseSurnames = new(StringComparer.OrdinalIgnoreCase)
    {
        "ANG", "AW", "BEH", "BOH", "BOO", "BOON", "CHAI", "CHAM", "CHAN", "CHANG", "CHAU", "CHEA",
        "CHEAH", "CHEE", "CHEN", "CHENG", "CHEONG", "CHEW", "CHIA", "CHIAH", "CHIAM", "CHIANG", "CHIEW", "CHIN",
        "CHING", "CHIONG", "CHNG", "CHO", "CHOA", "CHOH", "CHOK", "CHONG", "CHOO", "CHOOI", "CHOW", "CHOY",
        "CHU", "CHUA", "CHUAH", "CHUNG", "EE", "ENG", "EOW", "EU", "FAN", "FANG", "FOH", "FOK",
        "FONG", "FOO", "FOONG", "FU", "FUNG", "GAN", "GEE", "GOH", "GUAN", "GUI", "HAN", "HAU",
        "HEE", "HENG", "HEW", "HIA", "HII", "HING", "HO", "HOH", "HON", "HONG", "HOO", "HOOI",
        "HOR", "HOU", "HOW", "HSU", "HU", "HUA", "HUANG", "HUI", "HUNG", "HWANG", "JIN", "KAM",
        "KANG", "KAO", "KEE", "KEK", "KEONG", "KEW", "KHAW", "KHOO", "KHOR", "KIEW", "KIU", "KO",
        "KOAY", "KOH", "KOK", "KONG", "KOO", "KOON", "KOR", "KOW", "KU", "KUA", "KUAH", "KUAN",
        "KUANG", "KUEK", "KUNG", "KUO", "KUOK", "KWA", "KWAN", "KWOK", "KWONG", "LAI", "LAM", "LAU",
        "LAW", "LEE", "LEI", "LENG", "LEO", "LEONG", "LEOW", "LEW", "LI", "LIAN", "LIANG", "LIAO",
        "LIAU", "LIAW", "LIEW", "LIM", "LIN", "LING", "LIOW", "LIU", "LO", "LOH", "LOI", "LOK",
        "LOKE", "LONG", "LOO", "LOOI", "LOONG", "LOU", "LOW", "LOY", "LU", "LUI", "LUK", "LUM",
        "LUNG", "LYE", "MA", "MAH", "MAK", "MOK", "MOO", "MUI", "NEO", "NG", "NGAI", "NGAN",
        "NGEOW", "NGO", "NGOH", "NGOI", "NGUI", "NI", "OOI", "ONG", "OW", "PAH", "PAN", "PANG",
        "PAU", "PEE", "PEH", "PEI", "PENG", "PHAN", "PHANG", "PHEE", "PHENG", "PHUA", "PHUAH", "PNG",
        "POH", "PONG", "POO", "POON", "PUA", "PUAH", "PUN", "QUAH", "QUEK", "QUOK", "SAW", "SEAH",
        "SEE", "SEET", "SENG", "SEOW", "SHAM", "SHAO", "SHEN", "SHEW", "SHIA", "SHIH", "SHIM", "SHIU",
        "SHUM", "SIA", "SIAH", "SIEW", "SII", "SIM", "SIN", "SING", "SIO", "SIOW", "SIT", "SIU",
        "SO", "SOH", "SONG", "SOO", "SOON", "SOONG", "SU", "SUEN", "SUM", "SUN", "SUNG", "SZE",
        "TAI", "TAM", "TAN", "TANG", "TAO", "TAY", "TEE", "TEH", "TENG", "TEO", "TEOH", "TEOW",
        "TEW", "THAM", "THEAN", "THENG", "THIAN", "THIEN", "THIO", "THNG", "THO", "THONG", "THOO", "THUM",
        "TIAN", "TIEW", "TIN", "TING", "TIONG", "TOH", "TONG", "TOO", "TSAI", "TSANG", "TSE", "TSEN",
        "TSENG", "TSOI", "TU", "TUNG", "UNG", "VOO", "VOON", "WAI", "WAN", "WANG", "WEE", "WEI",
        "WEN", "WONG", "WOO", "WOON", "WU", "XIE", "XU", "YAN", "YANG", "YAO", "YAP", "YAU",
        "YEE", "YEO", "YEOH", "YEOW", "YEW", "YI", "YIM", "YIN", "YING", "YIP", "YONG", "YOO",
        "YOON", "YOONG", "YOU", "YOW", "YU", "YUAN", "YUE", "YUEN", "YUNG", "ZENG", "ZHANG", "ZHAO",
        "ZHENG", "ZHOU", "ZHU"
    };

    private static readonly HashSet<string> PatronymicConnectors = new(StringComparer.OrdinalIgnoreCase)
    {
        "BIN", "BINTI", "BT", "BT.", "B", "B.", "ANAK", "AK", "A/L", "A/P", "S/O", "D/O"
    };

    private static readonly HashSet<string> MalayGivenPrefixes = new(StringComparer.OrdinalIgnoreCase)
    {
        "MOHD", "MOHD.", "MOHAMAD", "MOHAMMAD", "MUHAMAD", "MUHAMMAD", "ABDUL", "ABD", "ABD.", "NUR", "NOR", "SITI", "AHMAD"
    };

    /// <summary>True for 3-word Malaysian Chinese names (surname + 2-part given name) that should keep all words.</summary>
    private static bool IsChineseThreeWordName(string[] parts)
    {
        if (parts.Length != 3 || !ChineseSurnames.Contains(parts[0])) return false;
        if (parts.Any(p => PatronymicConnectors.Contains(p))) return false;
        if (parts[1].StartsWith("B.", StringComparison.OrdinalIgnoreCase) ||
            parts[1].StartsWith("BT.", StringComparison.OrdinalIgnoreCase)) return false;
        return !MalayGivenPrefixes.Contains(parts[1]);
    }

    private static string FormatShortName(string fullName, int maxWords)
    {
        var parts = fullName.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
        if (parts.Length <= maxWords) return string.Join(" ", parts);
        if (maxWords == 2 && IsChineseThreeWordName(parts)) return string.Join(" ", parts);
        if (maxWords == 2) return $"{parts[0]} {parts[^1]}";
        return $"{parts[0]} {parts[1]} {parts[^1]}";
    }

    /// <summary>The text written for a null value in a dash-capable column (colour scales / icon sets ignore text).</summary>
    internal const string NullDash = "—";

    private static bool IsDashColumn(PropertyRuleDescriptor? rule) =>
        rule != null && rule.EffectiveZeroBehavior == ZeroBehavior.Dash &&
        rule.EffectiveFormatKind is not (ColumnFormatKind.Text or ColumnFormatKind.General or ColumnFormatKind.None);

    private static void WriteCellSafe(dynamic cell, object? value, PropertyRuleDescriptor? rule)
    {
        try
        {
            // null (or the dash text) => em-dash on dash columns, an empty cell elsewhere
            if (value == null || (value is string dashText && dashText == NullDash))
            {
                cell.Value2 = IsDashColumn(rule) ? NullDash : null;
                return;
            }

            if (value is double || value is decimal)
            {
                double num = Convert.ToDouble(value, CultureInfo.InvariantCulture);
                // A genuine 0 stays numeric 0 (shown as 0 / RM 0.00 / 0 km); only Blank columns clear it.
                cell.Value2 = Math.Abs(num) < 0.00001 && rule?.EffectiveZeroBehavior == ZeroBehavior.Blank ? (object?)null : num;
                return;
            }
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
        return SheetMapperHelper.CleanCellText(val);
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

}