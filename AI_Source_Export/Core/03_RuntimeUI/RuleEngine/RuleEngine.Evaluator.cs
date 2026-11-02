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

                // Map table columns: RelativeIndex (1..colCount) <-> Header & Rule
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

                // Check if row has any user input at all; if completely blank, do not auto-default yet
                bool rowHasUserInput = false;
                for (int c = 1; c <= colCount; c++)
                {
                    int absCol = startCol + c - 1;
                    string f = worksheet.Cells[excelRow, absCol].Formula?.ToString() ?? string.Empty;
                    if (f.StartsWith("=")) continue;
                    if (!string.IsNullOrWhiteSpace(rowValues[c]))
                    {
                        rowHasUserInput = true;
                        break;
                    }
                }

                if (!rowHasUserInput) return notices;

                SheetProtections.ExecuteWithBypass((object)worksheet, new Action(() =>
                {
                    // =========================================================================
                    // STEP 1: Apply Transform & Explicit Sync from the Edited Column
                    // =========================================================================
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

                                // Execute [ColumnSync] rules for the edited column
                                if (editedRule.SyncRules.Count > 0)
                                {
                                    ExecuteSyncRulesForColumn(wb, worksheet, table, excelRow, startCol, colCount, colHeaders, colRules, rowValues, rawObjects, editedRule, currentText);
                                }
                            }
                        }
                    }

                    // =========================================================================
                    // STEP 2: Fill Missing Partner Syncs Across the Row (If One Side is Blank)
                    // =========================================================================
                    for (int c = 1; c <= colCount; c++)
                    {
                        var rule = colRules[c];
                        string val = rowValues[c];
                        if (rule == null || string.IsNullOrWhiteSpace(val) || rule.SyncRules.Count == 0)
                            continue;

                        foreach (var sync in rule.SyncRules)
                        {
                            int targetRelCol = FindColumnByHeader(colHeaders, sync.TargetHeader);
                            if (targetRelCol >= 1 && string.IsNullOrWhiteSpace(rowValues[targetRelCol]))
                            {
                                string? lookedUp = RuleDataCache.LookupValue(wb, val, sync.LookupSource, sync.ResultSource);
                                if (!string.IsNullOrWhiteSpace(lookedUp))
                                {
                                    int targetAbsCol = startCol + targetRelCol - 1;
                                    WriteCellSafe(worksheet.Cells[excelRow, targetAbsCol], lookedUp, colRules[targetRelCol]);
                                    rowValues[targetRelCol] = lookedUp;
                                    rawObjects[targetRelCol] = lookedUp;
                                }
                            }
                        }
                    }

                    // =========================================================================
                    // STEP 3: Apply [ColumnDefault] to Empty Cells in Active Row
                    // =========================================================================
                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        var rule = colRules[c];
                        if (rule?.DefaultRule == null || rule.DefaultRule.Kind == DefaultKind.None)
                            continue;

                        if (!string.IsNullOrWhiteSpace(rowValues[c]))
                            continue;

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

                    // =========================================================================
                    // STEP 4: Enforce Self-Healing [ColumnStateGuard] Rules (e.g. Status)
                    // =========================================================================
                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        var rule = colRules[c];
                        if (rule == null || rule.StateGuards.Count == 0)
                            continue;

                        string currentVal = rowValues[c];
                        string? enforcedValue = EvaluateStateGuardsForCell(colHeaders, rawObjects, rowValues, rule, currentVal, out string? guardReason);

                        if (!string.IsNullOrWhiteSpace(enforcedValue) &&
                            !RuleDataCache.ValuesMatch(currentVal, enforcedValue))
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
                }));
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
    /// Runs EvaluateAndHealRow across every populated row in a table (used prior to Save/Extraction).
    /// </summary>
    public static List<string> EvaluateAndHealTable(dynamic worksheet, dynamic table)
    {
        var allNotices = new List<string>();
        if (worksheet == null || table == null) return allNotices;

        try
        {
            dynamic? body = table.DataBodyRange;
            if (body == null) return allNotices;

            int startRow = (int)body.Row;
            int rowCount = (int)body.Rows.Count;

            for (int r = 0; r < rowCount; r++)
            {
                var rowNotices = EvaluateAndHealRow(worksheet, table, startRow + r, null);
                allNotices.AddRange(rowNotices);
            }
        }
        catch { }

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
        string[] rowValues)
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