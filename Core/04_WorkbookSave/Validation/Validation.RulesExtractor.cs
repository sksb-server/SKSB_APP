namespace SKSB_App.Core.WorkbookSave.Validation;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

public record ColumnRule(
    string Header,
    bool IsRequired,
    bool IsFormula,
    PropertyRuleDescriptor? CSharpRule = null
);

public class SheetRulesCache
{
    public List<ColumnRule> ColumnRules { get; set; } = new();
}

public static class RulesExtractor
{
    private static readonly Dictionary<string, SheetRulesCache> _cachedRules = new(StringComparer.OrdinalIgnoreCase);

    public static SheetRulesCache InitializeRulesForSheet(object worksheetObj, int headerRow, int startCol, int columnSpan)
    {
        RuleRegistry.EnsureInitialized();
        dynamic worksheet = worksheetObj;

        string sheetName = "Sheet";
        try { sheetName = worksheet.Name?.ToString() ?? "Sheet"; } catch { }

        string tableName = string.Empty;
        try
        {
            dynamic? tbl = worksheet.Cells[headerRow, startCol].ListObject;
            tableName = tbl?.Name?.ToString() ?? string.Empty;
        }
        catch { }

        string cacheKey = $"{sheetName}_{tableName}_Col{startCol}_Span{columnSpan}";
        if (_cachedRules.TryGetValue(cacheKey, out var existingCache))
        {
            return existingCache;
        }

        AppLogger.Info($"[RulesExtractor] Initializing rule cache for '{tableName}' on '{sheetName}'...");
        var cache = new SheetRulesCache();

        // Bulk-read header row and archetype formulas in 2 COM calls instead of 3 * columnSpan calls
        int endCol = startCol + columnSpan - 1;
        object[,] headerGrid = TableMapper.GetRange2DValues(
            worksheet.Range[worksheet.Cells[headerRow, startCol], worksheet.Cells[headerRow, endCol]],
            1, columnSpan
        );

        var archFormulas = SheetMapperHelper.ReadArchetypeFormulas((object)worksheet, headerRow + 1, startCol, columnSpan);

        for (int c = 0; c < columnSpan; c++)
        {
            int colIndex = startCol + c;
            string headerText = headerGrid[1, c + 1]?.ToString()?.Trim() ?? $"Col_{colIndex}";
            var csharpRule = RuleRegistry.GetRule(tableName, headerText);
            bool isFormula = archFormulas[c + 1].StartsWith("=") || (csharpRule?.IsStrictlyCalculated ?? false);
            // Only query legacy COM metadata if no C# attribute rule governs this column
            var legacyConfig = (csharpRule != null || isFormula)
                ? new Config()
                : Metadata.ReadConfig(worksheet.Cells[headerRow + 1, colIndex]);

            bool isRequired = false;

            if (!isFormula)
            {
                if (csharpRule != null)
                {
                    isRequired = csharpRule.IsRequired && !csharpRule.IsWarningOnly;
                }
                else if (!legacyConfig.IsEmpty)
                {
                    isRequired = !legacyConfig.IsMultiSelect && !legacyConfig.IsWarningOnly;
                }
            }

            cache.ColumnRules.Add(new ColumnRule(
                Header: headerText,
                IsRequired: isRequired,
                IsFormula: isFormula,
                CSharpRule: csharpRule
            ));
        }

        _cachedRules[cacheKey] = cache;
        return cache;
    }

    public static (List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ValidateAllWorkbookTables(
        object wbObj,
        params string[] excludeTableNames)
        => ValidateAllWorkbookTables(wbObj, null, excludeTableNames);

    public static (List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ValidateAllWorkbookTables(
        object wbObj,
        IReadOnlyDictionary<string, Func<DateOnly, TimeOnly?, bool>>? warningDateFilters,
        params string[] excludeTableNames)
        => ValidateAllWorkbookTables(wbObj, warningDateFilters, null, excludeTableNames);

    /// <param name="skipHealAndDedupeTables">
    /// Tables the caller has already healed + deduplicated this save pass (e.g. via <c>TableMapper.ExtractFromTable</c>);
    /// they are only validated, skipping the redundant <c>EvaluateAndHealTable</c> / <c>CleanTableDuplicateRows</c> pass.
    /// </param>
    public static (List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ValidateAllWorkbookTables(
        object wbObj,
        IReadOnlyDictionary<string, Func<DateOnly, TimeOnly?, bool>>? warningDateFilters,
        IReadOnlySet<string>? skipHealAndDedupeTables,
        params string[] excludeTableNames)
    {
        using var _ = AppLogger.Scope("ValidateAllWorkbookTables");
        var errors = new List<ValidationIssue>();
        var warnings = new List<ValidationIssue>();
        if (wbObj == null) return (errors, warnings);

        dynamic wb = wbObj;
        var excludeSet = new HashSet<string>(excludeTableNames ?? Array.Empty<string>(), StringComparer.OrdinalIgnoreCase);

        // Shared source cache across all tables during this validation pass
        var sharedSourceCache = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);

        foreach (dynamic ws in wb.Sheets)
        {
            string sName = ws.Name?.ToString() ?? string.Empty;
            if (sName.Equals("Lookups", StringComparison.OrdinalIgnoreCase) ||
                sName.Equals("Data", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            int tableCount = 0;
            try { tableCount = (int)ws.ListObjects.Count; } catch { }
            if (tableCount == 0) continue;

            for (int t = 1; t <= tableCount; t++)
            {
                dynamic table = ws.ListObjects[t];
                string tableName = table.Name?.ToString() ?? string.Empty;
                if (excludeSet.Contains(tableName)) continue;

                dynamic? bodyRange = table.DataBodyRange;
                if (bodyRange == null) continue;

                int rowCount = (int)bodyRange.Rows.Count;
                int startRow = (int)bodyRange.Row;
                int startCol = (int)table.HeaderRowRange.Column;
                int colSpan = (int)table.ListColumns.Count;

                LoadingOverlay.Report($"Validating table: {tableName} ({rowCount} row(s))...", $"Worksheet: {sName}");
                var (tblErrors, tblWarnings) = EvaluateAndExtractInternal(
                    worksheet: (object)ws,
                    tableObj: (object)table,
                    bodyRangeObj: (object)bodyRange,
                    startRow: startRow,
                    maxRows: rowCount,
                    startCol: startCol,
                    columnSpan: colSpan,
                    sharedSourceCache: sharedSourceCache,
                    rowWarningFilter: warningDateFilters != null && warningDateFilters.TryGetValue(tableName, out var filter) ? filter : null,
                    skipHealAndDedupe: skipHealAndDedupeTables != null && skipHealAndDedupeTables.Contains(tableName)
                );

                errors.AddRange(tblErrors);
                warnings.AddRange(tblWarnings);
            }
        }

        return (errors, warnings);
    }

    private static (List<ValidationIssue> Errors, List<ValidationIssue> Warnings) EvaluateAndExtractInternal(
        object worksheet,
        object? tableObj,
        object? bodyRangeObj,
        int startRow,
        int maxRows,
        int startCol,
        int columnSpan,
        Dictionary<string, List<string>> sharedSourceCache,
        Func<DateOnly, TimeOnly?, bool>? rowWarningFilter = null,
        bool skipHealAndDedupe = false)
    {
        var errors = new List<ValidationIssue>();
        var warnings = new List<ValidationIssue>();

        dynamic wsDynamic = worksheet;
        dynamic wb = wsDynamic.Parent;
        string sheetName = wsDynamic.Name?.ToString() ?? "Sheet";
        dynamic? table = tableObj;
        string tableName = table?.Name?.ToString() ?? "Table";

        if (table != null && !skipHealAndDedupe)
        {
            // 1. In-memory cell healing
            warnings.AddRange(RuleEvaluator.EvaluateAndHealTable(wsDynamic, table, rowWarningFilter));

            // 2. VALIDATION DEDUPLICATION: Detect, Warn, Delete, and Skip
            var dedupeWarnings = new List<string>();
            int dupesRemoved = TableRowManager.CleanTableDuplicateRows(table, dedupeWarnings, warningDateFilter: rowWarningFilter);

            if (dupesRemoved > 0)
            {
                foreach (var w in dedupeWarnings)
                {
                    warnings.Add(ValidationIssue.FromLegacyMessage(ValidationSeverity.Warning, w, sheetName, tableName));
                }

                // Refresh DataBodyRange so subsequent array read skips the deleted rows
                dynamic? refreshedBody = table.DataBodyRange;
                if (refreshedBody == null)
                {
                    return (errors, warnings);
                }

                bodyRangeObj = refreshedBody;
                startRow = (int)refreshedBody.Row;
                maxRows = (int)refreshedBody.Rows.Count;
            }
        }

        var rules = InitializeRulesForSheet(worksheet, startRow - 1, startCol, columnSpan);

        object[,] grid;
        try
        {
            dynamic range = bodyRangeObj ?? wsDynamic.Range[
                wsDynamic.Cells[startRow, startCol],
                wsDynamic.Cells[startRow + maxRows - 1, startCol + columnSpan - 1]
            ];
            grid = TableMapper.GetRange2DValues(range, maxRows, columnSpan);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[Validation] Failed reading 2D Array for '{tableName}'", ex);
            return (errors, warnings);
        }

        int rowCount = Math.Min(maxRows, grid.GetLength(0));

        List<string> GetCachedSource(string src, bool preserveDuplicates)
        {
            if (string.IsNullOrWhiteSpace(src)) return new List<string>();
            string key = $"{src}|{preserveDuplicates}";
            if (!sharedSourceCache.TryGetValue(key, out var list))
            {
                list = RuleDataCache.GetSourceItems(wb, src, preserveDuplicates);
                sharedSourceCache[key] = list;
            }
            return list;
        }

        // Pre-build O(1) HashSet per column for allowed picklist options
        var allowedSetByCol = new HashSet<string>?[columnSpan];
        var hasTableSourceByCol = new bool[columnSpan];
        for (int c = 0; c < columnSpan; c++)
        {
            var cRule = rules.ColumnRules[c].CSharpRule;
            if (cRule != null && cRule.Sources.Length > 0 && !cRule.IsWarningOnly)
            {
                hasTableSourceByCol[c] = cRule.Sources.Any(s => !string.IsNullOrWhiteSpace(s) && s.Contains('['));
                var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (var src in cRule.Sources)
                {
                    foreach (var item in GetCachedSource(src, preserveDuplicates: false))
                    {
                        AddNormalizedVariantsToSet(set, item);
                    }
                }
                allowedSetByCol[c] = set;
            }
        }

        // Pre-build O(1) lookup maps for SyncRules: Key -> Allowed Partner Values
        var syncLookupCache = new Dictionary<string, Dictionary<string, HashSet<string>>>(StringComparer.OrdinalIgnoreCase);
        Dictionary<string, HashSet<string>> GetSyncPairMap(string lookupSrc, string resultSrc)
        {
            string cacheKey = $"{lookupSrc}>>{resultSrc}";
            if (!syncLookupCache.TryGetValue(cacheKey, out var map))
            {
                map = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);
                var keys = GetCachedSource(lookupSrc, preserveDuplicates: true);
                var vals = GetCachedSource(resultSrc, preserveDuplicates: true);
                int count = Math.Min(keys.Count, vals.Count);
                for (int i = 0; i < count; i++)
                {
                    string k = NormalizeLookupToken(keys[i]);
                    if (string.IsNullOrEmpty(k)) continue;
                    if (!map.TryGetValue(k, out var partnerSet))
                    {
                        partnerSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        map[k] = partnerSet;
                    }
                    AddNormalizedVariantsToSet(partnerSet, vals[i]);
                }
                syncLookupCache[cacheKey] = map;
            }
            return map;
        }

        var colHeadersOneBased = new string[columnSpan + 1];
        for (int c = 0; c < columnSpan; c++)
        {
            colHeadersOneBased[c + 1] = rules.ColumnRules[c].Header;
        }

        int dateCol = RuleEvaluator.FindColumnByHeader(colHeadersOneBased, "Date");
        int timeCol = RuleEvaluator.FindColumnByHeader(colHeadersOneBased, "Time");

        for (int r = 1; r <= rowCount; r++)
        {
            int excelRow = startRow + r - 1;
            bool hasMeaningfulData = false;
            var rowStringsOneBased = new string[columnSpan + 1];
            var rowObjectsOneBased = new object?[columnSpan + 1];

            for (int c = 0; c < columnSpan; c++)
            {
                var rule = rules.ColumnRules[c];
                object? cellObj = grid[r, c + 1];
                rowObjectsOneBased[c + 1] = cellObj;
                string valStr = SheetMapperHelper.CleanString(cellObj);
                rowStringsOneBased[c + 1] = valStr;

                if (rule.IsFormula) continue;

                if (!string.IsNullOrWhiteSpace(valStr) &&
                    valStr != "0" && valStr != "0.00" && valStr != "-" && valStr != " " &&
                    !valStr.Equals("None", StringComparison.OrdinalIgnoreCase) &&
                    !valStr.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    hasMeaningfulData = true;
                }
            }

            if (!hasMeaningfulData) continue;


            for (int c = 0; c < columnSpan; c++)
            {
                var rule = rules.ColumnRules[c];
                var cRule = rule.CSharpRule;
                object? rawVal = rowObjectsOneBased[c + 1];
                string strVal = rowStringsOneBased[c + 1];
                int excelCol = startCol + c;
                ValidationIssue CellIssue(ValidationSeverity severity, string message) =>
                    new(severity, sheetName, tableName, excelRow, excelCol, rule.Header, message);

                if (rule.IsFormula) continue;

                if (string.IsNullOrWhiteSpace(strVal))
                {
                    if (rule.IsRequired)
                    {
                        errors.Add(CellIssue(ValidationSeverity.Error, $"Field '{rule.Header}' is required and cannot be empty."));
                    }
                    continue;
                }

                if (cRule == null) continue;

                // "—" / "-" is the engine's display placeholder for null on numeric columns.
                if (cRule.IsNumericColumn && (strVal is "—" or "–" or "-" or "−" or "‒" or "―" or "--" or "—-")) continue;

                if (cRule.IsNumericColumn)
                {
                    if (!double.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double dNum))
                    {
                        errors.Add(CellIssue(ValidationSeverity.Error, $"Value '{strVal}' in '{rule.Header}' must be a valid number."));
                        continue;
                    }
                    if (cRule.AllowNonNegativeZero && dNum < -0.0001)
                    {
                        errors.Add(CellIssue(ValidationSeverity.Error, $"Value '{strVal}' in '{rule.Header}' cannot be negative."));
                        continue;
                    }
                }

                if (!string.IsNullOrWhiteSpace(cRule.DateMustBeGreaterThanHeader))
                {
                    int refCol = RuleEvaluator.FindColumnByHeader(colHeadersOneBased, cRule.DateMustBeGreaterThanHeader);
                    if (refCol >= 1)
                    {
                        string refValStr = rowStringsOneBased[refCol];
                        if (RuleDataCache.TryParseDate(strVal, out DateOnly thisDate) &&
                            RuleDataCache.TryParseDate(refValStr, out DateOnly refDate))
                        {
                            if (cRule.DateMustBeGreaterThanAllowEqual)
                            {
                                if (thisDate < refDate)
                                {
                                    errors.Add(CellIssue(ValidationSeverity.Error, $"'{rule.Header}' ({thisDate:dd/MM/yyyy}) must be on or after '{colHeadersOneBased[refCol]}' ({refDate:dd/MM/yyyy})."));
                                    continue;
                                }
                            }
                            else if (thisDate <= refDate)
                            {
                                errors.Add(CellIssue(ValidationSeverity.Error, $"'{rule.Header}' ({thisDate:dd/MM/yyyy}) must be strictly later than '{colHeadersOneBased[refCol]}' ({refDate:dd/MM/yyyy})."));
                                continue;
                            }
                        }
                    }
                }

                // O(1) HashSet Picklist Lookup
                var allowedSet = allowedSetByCol[c];
                if (allowedSet != null && (allowedSet.Count > 0 || hasTableSourceByCol[c]) && !cRule.IsWarningOnly)
                {
                    string emptyNote = allowedSet.Count == 0 ? " (source table is empty)" : string.Empty;
                    if (cRule.IsMultiSelect)
                    {
                        // Every comma separated token is checked on its own against the allowed set.
                        var tokens = RuleEvaluator.SplitMultiSelectTokens(strVal);
                        foreach (var token in tokens)
                        {
                            if (!IsTokenInNormalizedSet(allowedSet, token))
                            {
                                errors.Add(CellIssue(ValidationSeverity.Error, $"'{token}' is not a valid option for '{rule.Header}'{emptyNote}."));
                            }
                        }
                    }
                    else
                    {
                        if (!IsTokenInNormalizedSet(allowedSet, strVal))
                        {
                            string displayVal = RuleEvaluator.FormatDisplayValue(rawVal, cRule);
                            errors.Add(CellIssue(ValidationSeverity.Error, $"'{displayVal}' is not a valid option for '{rule.Header}'{emptyNote}."));
                        }
                    }
                }

                // O(1) Dictionary Sync Partner Validation
                if (cRule.SyncRules.Count > 0)
                {
                    var syncGroups = cRule.SyncRules.Where(s => !s.OnlyWhenTargetBlank).GroupBy(s => RuleRegistry.NormalizeIdentifier(s.TargetHeader));
                    string normKey = NormalizeLookupToken(strVal);

                    foreach (var group in syncGroups)
                    {
                        int partnerCol = RuleEvaluator.FindColumnByHeader(colHeadersOneBased, group.First().TargetHeader);
                        if (partnerCol < 1) continue;

                        string partnerVal = rowStringsOneBased[partnerCol];
                        if (string.IsNullOrWhiteSpace(partnerVal)) continue;

                        bool anyLookupFound = false;
                        bool anyLookupMatched = false;

                        foreach (var sync in group)
                        {
                            var pairMap = GetSyncPairMap(sync.LookupSource, sync.ResultSource);
                            if (pairMap.TryGetValue(normKey, out var validPartners) && validPartners.Count > 0)
                            {
                                anyLookupFound = true;
                                if (IsTokenInNormalizedSet(validPartners, partnerVal))
                                {
                                    anyLookupMatched = true;
                                    break;
                                }
                            }
                        }

                        if (anyLookupFound && !anyLookupMatched)
                        {
                            string conflictMsg = $"'{strVal}' in '{rule.Header}' conflicts with '{partnerVal}' in '{colHeadersOneBased[partnerCol]}'.";
                            if (cRule.IsWarningOnly)
                            {
                                if (rowWarningFilter == null || dateCol < 1 ||
                                    RuleDataCache.MatchesWarningDateFilter(rowObjectsOneBased[dateCol], timeCol >= 1 ? rowObjectsOneBased[timeCol] : null, rowWarningFilter))
                                    warnings.Add(CellIssue(ValidationSeverity.Warning, conflictMsg));
                            }
                            else
                            {
                                errors.Add(CellIssue(ValidationSeverity.Error, conflictMsg));
                            }
                        }
                    }
                }
            }
        }

        return (errors, warnings);
    }

    private static string NormalizeLookupToken(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return string.Empty;
        string trimmed = s.Trim();
        if (double.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) &&
            Math.Abs(d - Math.Round(d)) < 0.0001 && d < 1e15)
        {
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        }
        return trimmed.Replace(" ", "_");
    }

    private static void AddNormalizedVariantsToSet(HashSet<string> set, string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return;
        string trimmed = raw.Trim();
        set.Add(trimmed);
        set.Add(trimmed.Replace(" ", "_"));
        set.Add(trimmed.Replace("_", " "));
        if (double.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) &&
            Math.Abs(d - Math.Round(d)) < 0.0001 && d < 1e15)
        {
            set.Add(((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture));
        }
        if (RuleDataCache.TryParseDate(trimmed, out DateOnly dt))
        {
            set.Add(dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture));
        }
    }

    private static bool IsTokenInNormalizedSet(HashSet<string> set, string? token)
    {
        if (string.IsNullOrWhiteSpace(token)) return false;
        string trimmed = token.Trim();
        if (set.Contains(trimmed) || set.Contains(trimmed.Replace(" ", "_")))
            return true;

        if (double.TryParse(trimmed, NumberStyles.Any, CultureInfo.InvariantCulture, out double d) &&
            Math.Abs(d - Math.Round(d)) < 0.0001 && d < 1e15)
        {
            if (set.Contains(((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture)))
                return true;
        }

        if (RuleDataCache.TryParseDate(trimmed, out DateOnly dt))
        {
            if (set.Contains(dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)))
                return true;
        }

        return false;
    }

}