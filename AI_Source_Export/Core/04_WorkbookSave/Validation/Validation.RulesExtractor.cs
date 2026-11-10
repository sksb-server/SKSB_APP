namespace SKSB_App.Core.WorkbookSave.Validation;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.Linq;

public record ColumnRule(
    int ColumnIndex,
    string Header,
    bool IsRequired,
    int AlertStyle, // 1 = Stop (Error), 2 = Warning
    string ErrorMessage,
    Config DropDownConfig,
    bool IsFormula,
    PropertyRuleDescriptor? CSharpRule = null
);

public class SheetRulesCache
{
    public string CacheKey { get; set; } = string.Empty;
    public List<ColumnRule> ColumnRules { get; set; } = new();
    public DateTime CachedAt { get; set; } = DateTime.UtcNow;
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
        var cache = new SheetRulesCache { CacheKey = cacheKey };

        // Bulk-read header row and archetype formulas in 2 COM calls instead of 3 * columnSpan calls
        int endCol = startCol + columnSpan - 1;
        object[,] headerGrid = TableMapper.GetRange2DValues(
            worksheet.Range[worksheet.Cells[headerRow, startCol], worksheet.Cells[headerRow, endCol]],
            1, columnSpan
        );

        var archFormulas = new string[columnSpan + 1];
        try
        {
            dynamic archRange = worksheet.Range[worksheet.Cells[headerRow + 1, startCol], worksheet.Cells[headerRow + 1, endCol]];
            object? rawFormulas = archRange.Formula;
            if (rawFormulas is object[,] fMat)
            {
                for (int c = 1; c <= columnSpan; c++)
                    archFormulas[c] = fMat[1, c]?.ToString() ?? string.Empty;
            }
            else if (rawFormulas is string singleF)
            {
                archFormulas[1] = singleF;
            }
        }
        catch { }

        for (int c = 0; c < columnSpan; c++)
        {
            int colIndex = startCol + c;
            string headerText = headerGrid[1, c + 1]?.ToString()?.Trim() ?? $"Col_{colIndex}";
            bool isFormula = archFormulas[c + 1].StartsWith("=");

            var csharpRule = RuleRegistry.GetRule(tableName, headerText);
            // Only query legacy COM metadata if no C# attribute rule governs this column
            var legacyConfig = (csharpRule != null || isFormula)
                ? new Config()
                : Metadata.ReadConfig(worksheet.Cells[headerRow + 1, colIndex]);

            bool isRequired = false;
            int alertStyle = 1;
            string errorMsg = string.Empty;

            if (!isFormula)
            {
                if (csharpRule != null)
                {
                    isRequired = csharpRule.IsRequired && !csharpRule.IsWarningOnly;
                    alertStyle = csharpRule.IsWarningOnly ? 2 : 1;
                    errorMsg = $"Field '{headerText}' has an invalid or missing value.";
                }
                else if (!legacyConfig.IsEmpty)
                {
                    isRequired = !legacyConfig.IsMultiSelect && !legacyConfig.IsWarningOnly;
                    alertStyle = legacyConfig.IsWarningOnly ? 2 : 1;
                }
            }

            cache.ColumnRules.Add(new ColumnRule(
                ColumnIndex: colIndex,
                Header: headerText,
                IsRequired: isRequired,
                AlertStyle: alertStyle,
                ErrorMessage: errorMsg,
                DropDownConfig: legacyConfig,
                IsFormula: isFormula,
                CSharpRule: csharpRule
            ));
        }

        _cachedRules[cacheKey] = cache;
        return cache;
    }

    public static (List<string> Errors, List<string> Warnings) ValidateAllWorkbookTables(
        object wbObj,
        params string[] excludeTableNames)
    {
        using var _ = AppLogger.Scope("ValidateAllWorkbookTables");
        var errors = new List<string>();
        var warnings = new List<string>();
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

                var (_, tblErrors, tblWarnings) = EvaluateAndExtractInternal(
                    worksheet: (object)ws,
                    tableObj: (object)table,
                    bodyRangeObj: (object)bodyRange,
                    startRow: startRow,
                    maxRows: rowCount,
                    startCol: startCol,
                    columnSpan: colSpan,
                    sharedSourceCache: sharedSourceCache
                );

                errors.AddRange(tblErrors);
                warnings.AddRange(tblWarnings);
            }
        }

        return (errors, warnings);
    }

    public static (List<Dictionary<string, object?>> Rows, List<string> Errors, List<string> Warnings) EvaluateAndExtract(
        object worksheet,
        int startRow,
        int maxRows,
        int startCol,
        int columnSpan,
        int driverColIndex = 1)
    {
        if (worksheet == null || maxRows <= 0 || columnSpan <= 0)
            return (new(), new(), new());

        dynamic wsDynamic = worksheet;
        dynamic? table = null;
        dynamic? bodyRange = null;
        try
        {
            table = wsDynamic.Cells[startRow, startCol].ListObject;
            bodyRange = table?.DataBodyRange;
        }
        catch { }

        return EvaluateAndExtractInternal(
            worksheet,
            table,
            bodyRange,
            startRow,
            maxRows,
            startCol,
            columnSpan,
            new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase)
        );
    }

    private static (List<Dictionary<string, object?>> Rows, List<string> Errors, List<string> Warnings) EvaluateAndExtractInternal(
        object worksheet,
        object? tableObj,
        object? bodyRangeObj,
        int startRow,
        int maxRows,
        int startCol,
        int columnSpan,
        Dictionary<string, List<string>> sharedSourceCache)
    {
        var rows = new List<Dictionary<string, object?>>();
        var errors = new List<string>();
        var warnings = new List<string>();

        dynamic wsDynamic = worksheet;
        dynamic wb = wsDynamic.Parent;
        string sheetName = wsDynamic.Name?.ToString() ?? "Sheet";
        dynamic? table = tableObj;
        string tableName = table?.Name?.ToString() ?? "Table";

        if (table != null)
        {
            // 1. In-memory cell healing
            var healNotices = RuleEvaluator.EvaluateAndHealTable(wsDynamic, table);
            foreach (var notice in healNotices)
            {
                warnings.Add($"  [{sheetName}] {notice}");
            }

            // 2. VALIDATION DEDUPLICATION: Detect, Warn, Delete, and Skip
            var dedupeWarnings = new List<string>();
            int dupesRemoved = 0;

            SheetProtections.ExecuteWithBypass((object)wsDynamic, new Action(() =>
            {
                dupesRemoved = TableRowManager.CleanTableDuplicateRows(table, dedupeWarnings);
            }));

            if (dupesRemoved > 0)
            {
                foreach (var w in dedupeWarnings)
                {
                    warnings.Add($"  [{sheetName}] {w}");
                }

                // Refresh DataBodyRange so subsequent array read skips the deleted rows
                dynamic? refreshedBody = table.DataBodyRange;
                if (refreshedBody == null)
                {
                    return (rows, errors, warnings);
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
            return (rows, errors, warnings);
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
        for (int c = 0; c < columnSpan; c++)
        {
            var cRule = rules.ColumnRules[c].CSharpRule;
            if (cRule != null && cRule.Sources.Length > 0 && !cRule.IsWarningOnly)
            {
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
                string valStr = CleanString(cellObj);
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

            var rowDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            bool rowHasError = false;

            for (int c = 0; c < columnSpan; c++)
            {
                var rule = rules.ColumnRules[c];
                var cRule = rule.CSharpRule;
                object? rawVal = rowObjectsOneBased[c + 1];
                string strVal = rowStringsOneBased[c + 1];

                rowDict[rule.Header] = rawVal;
                if (rule.IsFormula) continue;

                if (string.IsNullOrWhiteSpace(strVal))
                {
                    if (rule.IsRequired)
                    {
                        errors.Add($"  Row {excelRow} [{sheetName}]: Field '{rule.Header}' is required and cannot be empty.");
                        rowHasError = true;
                    }
                    continue;
                }

                if (cRule == null) continue;

                if (cRule.IsNumericColumn)
                {
                    if (!double.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double dNum))
                    {
                        errors.Add($"  Row {excelRow} [{sheetName}]: Value '{strVal}' in '{rule.Header}' must be a valid number.");
                        rowHasError = true;
                        continue;
                    }
                    if (cRule.AllowNonNegativeZero && dNum < -0.0001)
                    {
                        errors.Add($"  Row {excelRow} [{sheetName}]: Value '{strVal}' in '{rule.Header}' cannot be negative.");
                        rowHasError = true;
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
                            if (thisDate <= refDate)
                            {
                                errors.Add($"  Row {excelRow} [{sheetName}]: '{rule.Header}' ({thisDate:dd/MM/yyyy}) must be strictly later than '{colHeadersOneBased[refCol]}' ({refDate:dd/MM/yyyy}).");
                                rowHasError = true;
                                continue;
                            }
                        }
                    }
                }

                // O(1) HashSet Picklist Lookup
                var allowedSet = allowedSetByCol[c];
                if (allowedSet != null && allowedSet.Count > 0 && !cRule.IsWarningOnly)
                {
                    if (cRule.IsMultiSelect)
                    {
                        var tokens = strVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim());
                        foreach (var token in tokens)
                        {
                            if (!IsTokenInNormalizedSet(allowedSet, token))
                            {
                                errors.Add($"  Row {excelRow} [{sheetName}]: '{token}' is not a valid option for '{rule.Header}'.");
                                rowHasError = true;
                            }
                        }
                    }
                    else
                    {
                        if (!IsTokenInNormalizedSet(allowedSet, strVal))
                        {
                            string displayVal = RuleEvaluator.FormatDisplayValue(rawVal, cRule);
                            errors.Add($"  Row {excelRow} [{sheetName}]: '{displayVal}' is not a valid option for '{rule.Header}'.");
                            rowHasError = true;
                        }
                    }
                }

                // O(1) Dictionary Sync Partner Validation
                if (cRule.SyncRules.Count > 0)
                {
                    var syncGroups = cRule.SyncRules.GroupBy(s => RuleRegistry.NormalizeIdentifier(s.TargetHeader));
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
                            string conflictMsg = $"  Row {excelRow} [{sheetName}]: '{strVal}' in '{rule.Header}' conflicts with '{partnerVal}' in '{colHeadersOneBased[partnerCol]}'.";
                            if (cRule.IsWarningOnly)
                                warnings.Add(conflictMsg);
                            else
                            {
                                errors.Add(conflictMsg);
                                rowHasError = true;
                            }
                        }
                    }
                }
            }

            if (!rowHasError)
            {
                rows.Add(rowDict);
            }
        }

        return (rows, errors, warnings);
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

    private static string CleanString(object? val)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        if (val is double d && Math.Abs(d - Math.Round(d)) < 0.00001 && d < 1e15)
            return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (s.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        if (long.TryParse(s, out long num) && num < -2000000000) return string.Empty;
        return s;
    }
}