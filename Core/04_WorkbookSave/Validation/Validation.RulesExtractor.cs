namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookSave.SheetMapper;

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

        for (int c = 0; c < columnSpan; c++)
        {
            int colIndex = startCol + c;
            string headerText = $"Col_{colIndex}";

            try
            {
                dynamic headerCell = worksheet.Cells[headerRow, colIndex];
                headerText = headerCell.Value2?.ToString()?.Trim() ?? $"Col_{colIndex}";
            }
            catch { }

            dynamic archetypeCell = worksheet.Cells[headerRow + 1, colIndex];
            string archFormula = string.Empty;
            try { archFormula = archetypeCell.Formula2?.ToString() ?? string.Empty; } catch { }
            if (string.IsNullOrWhiteSpace(archFormula))
            {
                try { archFormula = archetypeCell.Formula?.ToString() ?? string.Empty; } catch { }
            }

            bool isFormula = archFormula.StartsWith("=");
            var csharpRule = RuleRegistry.GetRule(tableName, headerText);
            var legacyConfig = Metadata.ReadConfig(archetypeCell);

            bool isRequired = false;
            int alertStyle = 1;
            string errorMsg = string.Empty;

            if (isFormula)
            {
                isRequired = false;
            }
            else if (csharpRule != null)
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

                try { if ((bool)table.ShowTotals) table.ShowTotals = false; } catch { }

                dynamic bodyRange = table.DataBodyRange;
                if (bodyRange == null) continue;

                int rowCount = (int)bodyRange.Rows.Count;
                int startRow = (int)bodyRange.Row;
                int startCol = (int)table.HeaderRowRange.Column;
                int colSpan = (int)table.ListColumns.Count;

                var (_, tblErrors, tblWarnings) = EvaluateAndExtract(
                    worksheet: (object)ws,
                    startRow: startRow,
                    maxRows: rowCount,
                    startCol: startCol,
                    columnSpan: colSpan,
                    driverColIndex: 1
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
        var rows = new List<Dictionary<string, object?>>();
        var errors = new List<string>();
        var warnings = new List<string>();

        if (worksheet == null || maxRows <= 0 || columnSpan <= 0)
            return (rows, errors, warnings);

        dynamic wsDynamic = worksheet;
        dynamic wb = wsDynamic.Parent;
        string sheetName = "Sheet";
        try { sheetName = wsDynamic.Name?.ToString() ?? "Sheet"; } catch { }

        dynamic? table = null;
        string tableName = "Table";
        try
        {
            table = wsDynamic.Cells[startRow, startCol].ListObject;
            tableName = table?.Name?.ToString() ?? "Table";
        }
        catch { }

        using var _ = AppLogger.Scope($"EvaluateAndExtract('{tableName}', maxRows={maxRows})");

        if (table != null)
        {
            AppLogger.Info($"[Validation] Running Pre-Save Heal on '{tableName}'...");
            var healNotices = RuleEvaluator.EvaluateAndHealTable(wsDynamic, table);
            foreach (var notice in healNotices)
            {
                warnings.Add($"  [{sheetName}] {notice}");
            }
        }

        var rules = InitializeRulesForSheet((object)wsDynamic, startRow - 1, startCol, columnSpan);

        int endRow = startRow + maxRows - 1;
        int endCol = startCol + columnSpan - 1;

        object[,] grid;
        try
        {
            AppLogger.Info($"[Validation] Reading 2D Array for '{tableName}' (Rows {startRow} to {endRow})...");
            dynamic range = wsDynamic.Range[wsDynamic.Cells[startRow, startCol], wsDynamic.Cells[endRow, endCol]];
            grid = TableMapper.GetRange2DValues(range, maxRows, columnSpan);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[Validation] Failed reading 2D Array for '{tableName}'", ex);
            return (rows, errors, warnings);
        }

        int rowCount = Math.Min(maxRows, grid.GetLength(0));
        AppLogger.Info($"[Validation] Starting in-memory validation loop for {rowCount} rows...");

        var allowedSetByCol = new List<string>?[columnSpan];
        for (int c = 0; c < columnSpan; c++)
        {
            var cRule = rules.ColumnRules[c].CSharpRule;
            if (cRule != null && cRule.Sources.Length > 0 && !cRule.IsWarningOnly)
            {
                var combined = new List<string>();
                foreach (var src in cRule.Sources)
                {
                    combined.AddRange(RuleDataCache.GetSourceItems(wb, src, preserveDuplicates: false));
                }
                allowedSetByCol[c] = combined;
            }
        }

        var colHeadersOneBased = new string[columnSpan + 1];
        for (int c = 0; c < columnSpan; c++)
        {
            colHeadersOneBased[c + 1] = rules.ColumnRules[c].Header;
        }

        int meaningfulRowsFound = 0;

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
            meaningfulRowsFound++;

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

                // 1. Required Field Check
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

                // 2. Numeric Non-Negative Check
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

                // 3. Allowed Picklist Check
                var allowedItems = allowedSetByCol[c];
                if (allowedItems != null && allowedItems.Count > 0 && !cRule.IsWarningOnly)
                {
                    if (cRule.IsMultiSelect)
                    {
                        var tokens = strVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries).Select(t => t.Trim());
                        foreach (var token in tokens)
                        {
                            if (!allowedItems.Any(a => RuleDataCache.ValuesMatch(a, token)))
                            {
                                errors.Add($"  Row {excelRow} [{sheetName}]: '{token}' is not a valid option for '{rule.Header}'.");
                                rowHasError = true;
                            }
                        }
                    }
                    else
                    {
                        bool matched = allowedItems.Any(a => RuleDataCache.ValuesMatch(a, strVal));
                        if (!matched)
                        {
                            string displayVal = RuleEvaluator.FormatDisplayValue(rawVal, cRule);
                            errors.Add($"  Row {excelRow} [{sheetName}]: '{displayVal}' is not a valid option for '{rule.Header}'.");
                            rowHasError = true;
                        }
                    }
                }

                // 4. Relational Partner Sync Conflict Check
                if (cRule.SyncRules.Count > 0 && !cRule.IsWarningOnly)
                {
                    var syncGroups = cRule.SyncRules.GroupBy(s => RuleRegistry.NormalizeIdentifier(s.TargetHeader));
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
                            string? expectedPartner = RuleDataCache.LookupValue(wb, strVal, sync.LookupSource, sync.ResultSource);
                            if (!string.IsNullOrWhiteSpace(expectedPartner))
                            {
                                anyLookupFound = true;
                                if (RuleDataCache.ValuesMatch(expectedPartner, partnerVal))
                                {
                                    anyLookupMatched = true;
                                    break;
                                }
                            }
                        }

                        if (anyLookupFound && !anyLookupMatched)
                        {
                            errors.Add($"  Row {excelRow} [{sheetName}]: '{strVal}' in '{rule.Header}' conflicts with '{partnerVal}' in '{colHeadersOneBased[partnerCol]}'.");
                            rowHasError = true;
                        }
                    }
                }
            }

            if (!rowHasError)
            {
                rows.Add(rowDict);
            }
        }

        AppLogger.Info($"[Validation] Validation complete. Meaningful rows found: {meaningfulRowsFound}");
        return (rows, errors, warnings);
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