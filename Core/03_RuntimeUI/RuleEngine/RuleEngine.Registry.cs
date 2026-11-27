namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;

public sealed class PropertyRuleDescriptor
{
    public string TableName { get; init; } = string.Empty;
    public string HeaderName { get; init; } = string.Empty;
    public string NormalizedHeader { get; init; } = string.Empty;
    public PropertyInfo Property { get; init; } = null!;
    public Type PropertyType { get; init; } = typeof(string);

    public string[] Sources { get; init; } = Array.Empty<string>();
    public bool IsRequired { get; init; }
    public bool IsWarningOnly { get; init; }
    public bool IsReadOnly { get; init; }
    public bool IsMultiSelect { get; init; }
    public string? CountTargetHeader { get; init; }
    public bool AllowNonNegativeZero { get; init; } = true;
    public TransformKind Transform { get; init; } = TransformKind.None;

    public List<ColumnSyncAttribute> SyncRules { get; init; } = new();
    public ColumnDefaultAttribute? DefaultRule { get; init; }
    public List<ColumnStateGuardAttribute> StateGuards { get; init; } = new();
    public string? DateMustBeGreaterThanHeader { get; init; }
    public bool DateMustBeGreaterThanAllowEqual { get; init; }

    public double RoundDownStep { get; init; }

    public ColumnFormatKind EffectiveFormatKind { get; init; }
    public ZeroBehavior EffectiveZeroBehavior { get; init; }
    public string? ExcelNumberFormat { get; init; }

    public ColumnCalculateAttribute? CalculateRule { get; init; }
    public IReadOnlyList<ColumnRollupAttribute> RollupRules { get; init; } = Array.Empty<ColumnRollupAttribute>();
    public IReadOnlyList<SheetParameterAttribute> SheetParameters { get; init; } = Array.Empty<SheetParameterAttribute>();

    /// <summary>True when C# always owns the cell value (strict calculation or parent rollup).</summary>
    public bool IsStrictlyCalculated =>
        (CalculateRule != null && CalculateRule.Mode == CalculationMode.Always) ||
        RollupRules.Any(r => r.Mode == CalculationMode.Always);

    /// <summary>True when users may not edit the column (strictly calculated or declared read-only).</summary>
    public bool IsLocked => IsStrictlyCalculated || IsReadOnly;

    /// <summary>True when C# declares a formula/rollup for the column (any mode); legacy Excel formulas are stripped.</summary>
    public bool HasCSharpFormula => CalculateRule != null || RollupRules.Count > 0;

    public string? DisplayFormula =>
        CalculateRule != null ? $"= {CalculateRule.Expression}" :
        RollupRules.Count > 0 ? "= " + string.Join(" + ", RollupRules.Select(r =>
            $"{r.Kind.ToString().ToUpperInvariant()}({r.CollectionOrTable}[{string.Join("+", r.SourceProperties)}])")) :
        null;

    public bool HasDropdown =>
        !IsLocked &&
        (Sources.Length > 0 ||
         (DefaultRule != null && DefaultRule.Kind != DefaultKind.None) ||
         (CalculateRule != null && CalculateRule.Mode == CalculationMode.DefaultOnly));

    public bool IsDateColumn =>
        (Nullable.GetUnderlyingType(PropertyType) ?? PropertyType) == typeof(DateOnly) ||
        (Nullable.GetUnderlyingType(PropertyType) ?? PropertyType) == typeof(DateTime);

    public bool IsNumericColumn
    {
        get
        {
            Type t = Nullable.GetUnderlyingType(PropertyType) ?? PropertyType;
            return t == typeof(int) || t == typeof(long) || t == typeof(decimal) || t == typeof(double);
        }
    }
}

public static partial class RuleRegistry
{
    private static bool _isInitialized = false;
    private static readonly object _initLock = new();

    // Key: NormalizedTableName -> (NormalizedHeader -> PropertyRuleDescriptor)
    private static readonly Dictionary<string, Dictionary<string, PropertyRuleDescriptor>> _tableMap =
        new(StringComparer.OrdinalIgnoreCase);

    public static void EnsureInitialized()
    {
        if (_isInitialized) return;

        lock (_initLock)
        {
            if (_isInitialized) return;

            var coreAssembly = typeof(ExcelTableAttribute).Assembly;
            var types = coreAssembly.GetTypes()
                .Where(t => t.IsClass && !t.IsAbstract && t.GetCustomAttributes<ExcelTableAttribute>(true).Any())
                .ToList();

            foreach (var rootType in types)
            {
                var sheetParams = CollectSheetParameters(rootType);
                var tableAttrs = rootType.GetCustomAttributes<ExcelTableAttribute>(true);
                foreach (var tableAttr in tableAttrs)
                {
                    string normTable = NormalizeIdentifier(tableAttr.TableName);
                    if (!_tableMap.TryGetValue(normTable, out var colDict))
                    {
                        colDict = new Dictionary<string, PropertyRuleDescriptor>(StringComparer.OrdinalIgnoreCase);
                        _tableMap[normTable] = colDict;
                    }

                    ScanTypeProperties(rootType, tableAttr.TableName, colDict, sheetParams);
                }
            }

            _isInitialized = true;
        }
    }

    /// <summary>True for multi-select cell types: List&lt;string&gt; or List&lt;TEnum&gt; ("A, B" in one cell).</summary>
    internal static bool IsMultiValueListType(Type type) =>
        type == typeof(List<string>) ||
        (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(List<>) && type.GetGenericArguments()[0].IsEnum);

    /// <summary>Collects [SheetParameter] from the root type and every [DataDetail] child type.</summary>
    private static List<SheetParameterAttribute> CollectSheetParameters(Type rootType)
    {
        var list = new List<SheetParameterAttribute>(rootType.GetCustomAttributes<SheetParameterAttribute>(true));
        foreach (var prop in rootType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetCustomAttribute<DataDetailAttribute>() == null) continue;
            foreach (var a in prop.PropertyType.GetCustomAttributes<SheetParameterAttribute>(true))
                if (!list.Any(x => x.Name.Equals(a.Name, StringComparison.OrdinalIgnoreCase))) list.Add(a);
        }
        return list;
    }

    private static void ScanTypeProperties(
    Type targetType,
    string tableName,
    Dictionary<string, PropertyRuleDescriptor> colDict,
    IReadOnlyList<SheetParameterAttribute> sheetParams)
    {
        foreach (var prop in targetType.GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (prop.GetIndexParameters().Length > 0) continue;

            // Recurse into [DataDetail] child objects (e.g., VehicleDetail, StaffDetails, ProjectDetails)
            if (prop.GetCustomAttribute<DataDetailAttribute>() != null)
            {
                ScanTypeProperties(prop.PropertyType, tableName, colDict, sheetParams);
                continue;
            }

            // Skip child entity collections (e.g., List<Maintenance>), but KEEP List<string> (e.g., Workers/Vehicles multi-select)
            if (typeof(IEnumerable).IsAssignableFrom(prop.PropertyType) &&
                prop.PropertyType != typeof(string) &&
                !IsMultiValueListType(prop.PropertyType) &&
                prop.PropertyType != typeof(string[]))
            {
                continue;
            }

            var excelColAttr = prop.GetCustomAttribute<ExcelColumnAttribute>();
            var ruleCandidates = prop.GetCustomAttributes<ColumnRuleAttribute>().ToList();
            var ruleAttr =
                ruleCandidates.FirstOrDefault(r => !string.IsNullOrWhiteSpace(r.TableName) &&
                                                   string.Equals(r.TableName, tableName, StringComparison.OrdinalIgnoreCase))
                ?? ruleCandidates.FirstOrDefault(r => string.IsNullOrWhiteSpace(r.TableName));
            var defaultCandidates = prop.GetCustomAttributes<ColumnDefaultAttribute>().ToList();
            var defaultAttr =
                defaultCandidates.FirstOrDefault(d => !string.IsNullOrWhiteSpace(d.TableName) &&
                                                      string.Equals(d.TableName, tableName, StringComparison.OrdinalIgnoreCase))
                ?? defaultCandidates.FirstOrDefault(d => string.IsNullOrWhiteSpace(d.TableName));
            var dateGreaterAttr = prop.GetCustomAttribute<ColumnDateGreaterThanAttribute>();
            var formatAttr = prop.GetCustomAttribute<ColumnFormatAttribute>();
            var calcAttr = prop.GetCustomAttribute<ColumnCalculateAttribute>();
            var rollupAttrs = prop.GetCustomAttributes<ColumnRollupAttribute>().ToList();
            var syncAttrs = prop.GetCustomAttributes<ColumnSyncAttribute>()
                                .Where(s => string.IsNullOrWhiteSpace(s.TableName) ||
                                            string.Equals(s.TableName, tableName, StringComparison.OrdinalIgnoreCase))
                                .ToList();
            var guardAttrs = prop.GetCustomAttributes<ColumnStateGuardAttribute>()
                                 .OrderBy(g => g.Priority)
                                 .ToList();

            // If no attributes exist at all, skip the property
            if (excelColAttr == null && ruleAttr == null && formatAttr == null && defaultAttr == null && syncAttrs.Count == 0 && guardAttrs.Count == 0 &&
                calcAttr == null && rollupAttrs.Count == 0)
                continue;

            // Prefer ExcelColumn header; fallback to the property name
            string headerName = excelColAttr != null && !string.IsNullOrWhiteSpace(excelColAttr.HeaderName)
                ? excelColAttr.HeaderName
                : prop.Name;

            string normHeader = NormalizeIdentifier(headerName);

            // If there's no ExcelColumn attribute, it is a UI-only column (never required to save)
            bool isMappedToStorage = excelColAttr != null;
            bool isStrictlyCalculated = (calcAttr != null && calcAttr.Mode == CalculationMode.Always) || rollupAttrs.Any(r => r.Mode == CalculationMode.Always);
            bool isRequired = isMappedToStorage && (ruleAttr?.IsRequired ?? excelColAttr!.IsRequired);
            bool isWarningOnly = !isMappedToStorage || (ruleAttr?.IsWarningOnly ?? false);
            if (isStrictlyCalculated)
            {
                isRequired = false;
                isWarningOnly = true;
            }

            var fmt = ResolveEffectiveFormat(prop, headerName, formatAttr, defaultAttr);

            var descriptor = new PropertyRuleDescriptor
            {
                TableName = tableName,
                HeaderName = headerName,
                NormalizedHeader = normHeader,
                Property = prop,
                PropertyType = prop.PropertyType,
                Sources = ruleAttr?.Sources ?? Array.Empty<string>(),
                IsRequired = isRequired,
                IsWarningOnly = isWarningOnly,
                IsReadOnly = ruleAttr?.IsReadOnly ?? false,
                IsMultiSelect = ruleAttr?.IsMultiSelect ?? false,
                CountTargetHeader = ruleAttr?.CountTargetHeader,
                AllowNonNegativeZero = ruleAttr?.AllowNonNegativeZero ?? true,
                Transform = ruleAttr?.Transform ?? TransformKind.None,
                SyncRules = syncAttrs,
                DefaultRule = defaultAttr,
                StateGuards = guardAttrs,
                DateMustBeGreaterThanHeader = dateGreaterAttr?.ReferenceHeader,
                DateMustBeGreaterThanAllowEqual = dateGreaterAttr?.AllowEqual ?? false,
                RoundDownStep = ruleAttr?.RoundDownStep ?? 0.0,
                EffectiveFormatKind = fmt.Kind,
                EffectiveZeroBehavior = fmt.Zero,
                ExcelNumberFormat = fmt.ExcelNumberFormat,
                CalculateRule = calcAttr,
                RollupRules = rollupAttrs,
                SheetParameters = sheetParams
            };

            colDict[normHeader] = descriptor;
        }
    }

    public static PropertyRuleDescriptor? GetRule(string tableName, string columnHeader)
    {
        EnsureInitialized();
        if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(columnHeader))
            return null;

        if (!TryResolveTableDictionary(tableName, out var colDict))
            return null;

        return ResolveHeaderInDictionary(colDict, columnHeader);
    }

    public static PropertyRuleDescriptor? GetRuleForCell(dynamic singleCell)
    {
        if (singleCell == null) return null;
        try
        {
            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }
            if (table == null) return null;

            string tableName = table.Name?.ToString() ?? string.Empty;
            int headerRow = (int)table.HeaderRowRange.Row;
            int col = (int)singleCell.Column;
            dynamic sheet = singleCell.Worksheet;

            string colHeader = sheet.Cells[headerRow, col].Value2?.ToString()?.Trim() ?? string.Empty;
            return GetRule(tableName, colHeader);
        }
        catch
        {
            return null;
        }
    }

    /// <summary>
    /// Auto-locks a C#-calculated column without sheet protection: a custom '=FALSE' Data Validation (Stop alert)
    /// blocks interactive typing while C# COM writes (.Value2) pass through. Paste / Delete are caught by the
    /// SheetChange guard in EventsAutoOpen. No on-cell input message is shown (ShowInput = false);
    /// the calculation logic lives solely in the table header hover Note.
    /// </summary>
    internal static void ApplyCalculatedColumnLock(dynamic targetRange, PropertyRuleDescriptor? rule)
    {
        if (targetRange == null || rule == null || !rule!.IsLocked) return;
        try
        {
            try { targetRange!.Validation.Delete(); } catch { }

            // Type 7 = xlValidateCustom, AlertStyle 1 = xlValidAlertStop, Operator 1 = xlBetween
            targetRange!.Validation.Add(7, 1, 1, "=FALSE");
            targetRange.Validation.IgnoreBlank = false;

            string errorTitle, error;
            if (rule!.IsStrictlyCalculated)
            {
                string formulaText = rule.DisplayFormula ?? "Calculated automatically by SKSB.";
                if (formulaText.Length > 220) formulaText = formulaText.Substring(0, 220);
                errorTitle = "Locked Calculated Column";
                error = $"'{rule.HeaderName}' is automatically calculated ({formulaText}) and cannot be edited manually.";
            }
            else
            {
                errorTitle = "Read-Only Column";
                error = $"'{rule.HeaderName}' is managed automatically by SKSB and cannot be edited manually.";
            }

            try { targetRange.Validation.ShowInput = false; } catch { }

            targetRange.Validation.ErrorTitle = Truncate(errorTitle, 32);
            targetRange.Validation.ErrorMessage = Truncate(error, 255);
            targetRange.Validation.ShowError = true;
        }
        catch { }
    }

    /// <summary>
    /// Automatically applies the "DropDown" named style and clears legacy native validation
    /// on all table columns governed by C# [ColumnRule] or [ColumnDefault] attributes,
    /// while also disabling legacy ShowTotals rows across all tables.
    /// </summary>
    public static void ApplyDropdownStylesToWorkbook(dynamic wb)
    {
        if (wb == null) return;
        EnsureInitialized();

        dynamic xlApp = wb.Application;
        bool origUpdating = true;
        bool origEvents = true;

        try
        {
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

            dynamic? activeSheet = null;
            try { activeSheet = wb.ActiveSheet; } catch { }

            foreach (dynamic ws in wb.Sheets)
            {
                string sName = ws.Name?.ToString() ?? string.Empty;
                if (sName.Equals(SheetOperations.LookupsSheetName, StringComparison.OrdinalIgnoreCase) ||
                    sName.Equals(SheetOperations.DataSheetName, StringComparison.OrdinalIgnoreCase))
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
                    LoadingOverlay.Report("Applying table rules & column tooltips...", $"Styling {tableName}...");

                    int headerRow = (int)table.HeaderRowRange.Row;
                    int startCol = (int)table.HeaderRowRange.Column;
                    int colCount = (int)table.ListColumns.Count;
                    dynamic? bodyRange = table.DataBodyRange;

                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        string headerText = ws.Cells[headerRow, absCol].Value2?.ToString()?.Trim() ?? string.Empty;
                        var rule = GetRule(tableName, headerText);
                        if (rule == null) continue;

                        dynamic headerCell = ws.Cells[headerRow, absCol];
                        dynamic archCell = ws.Cells[headerRow + 1, absCol];
                        string archFormula = archCell.Formula?.ToString() ?? string.Empty;

                        if (rule.IsStrictlyCalculated)
                        {
                            try
                            {
                                dynamic calcRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                                if (archFormula.StartsWith("="))
                                {
                                    // Strip the legacy Excel formula; C# owns this column now.
                                    object? existingVal = calcRange.Value2;
                                    calcRange.Value2 = existingVal;
                                }
                                ApplyColumnNumberFormat(table, c, archCell, rule);
                                ApplyCalculatedColumnLock(calcRange, rule);
                                ApplyColumnHeaderNote(headerCell, rule);
                            }
                            catch { }
                            continue;
                        }

                        if (rule.IsReadOnly)
                        {
                            try
                            {
                                dynamic lockRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                                ApplyColumnNumberFormat(table, c, archCell, rule);
                                ApplyCalculatedColumnLock(lockRange, rule);
                                ApplyColumnHeaderNote(headerCell, rule);
                            }
                            catch { }
                            continue;
                        }

                        if (!rule.IsStrictlyCalculated && rule.HasCSharpFormula && archFormula.StartsWith("="))
                        {
                            // Soft (DefaultOnly) C# formula: drop the legacy Excel formula but keep the values.
                            try
                            {
                                dynamic softRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                                object? softVal = softRange.Value2;
                                softRange.Value2 = softVal;
                                archFormula = string.Empty;
                            }
                            catch { }
                        }

                        if (archFormula.StartsWith("="))
                        {
                            // Legacy Excel formula column: note on header only, never touch its Data Validation.
                            ApplyColumnHeaderNote(headerCell, rule, archFormula);
                            continue;
                        }

                        // Inside RuleRegistry.ApplyDropdownStylesToWorkbook:
                        if (rule.HasDropdown)
                        {
                            try
                            {
                                dynamic targetRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                                if (targetRange != null)
                                {
                                    try { targetRange.Validation.Delete(); } catch { }
                                }
                            }
                            catch { }
                        }

                        ApplyColumnNumberFormat(table, c, archCell, rule);

                        try
                        {
                            dynamic tooltipRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                            ApplyColumnTooltips(headerCell, tooltipRange, rule);
                        }
                        catch { }
                    }
                }

            }

            // Restore original sheet view so Excel doesn't jump
            if (activeSheet != null)
            {
                try { activeSheet.Activate(); } catch { }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleRegistry.ApplyDropdownStylesToWorkbook Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    /// <summary>
    /// Returns the headers of sibling columns whose [ColumnSync] auto-fills the given column (excluding itself).
    /// </summary>
    internal static List<string> GetSyncSourcesForColumn(string tableName, string headerName)
    {
        EnsureInitialized();
        if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(headerName))
            return new List<string>();

        if (!TryResolveTableDictionary(tableName, out var colDict))
            return new List<string>();

        string normHeader = NormalizeIdentifier(headerName);
        string strippedHeader = normHeader.TrimEnd('s');

        return colDict.Values
            .Where(d => !d.NormalizedHeader.Equals(normHeader, StringComparison.OrdinalIgnoreCase))
            .Where(d => d.SyncRules.Any(s =>
            {
                string normTarget = NormalizeIdentifier(s.TargetHeader ?? string.Empty);
                return normTarget.Length > 0 &&
                       (normTarget.Equals(normHeader, StringComparison.OrdinalIgnoreCase) ||
                        normTarget.TrimEnd('s').Equals(strippedHeader, StringComparison.OrdinalIgnoreCase));
            }))
            .Select(d => d.HeaderName)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static readonly Regex BracketTokenRegex = new(@"\[([^\[\]]+)\]", RegexOptions.Compiled);

    /// <summary>
    /// Returns the sibling columns that react when the given input column changes: its own [ColumnSync] targets and
    /// count target, plus every column whose dropdown sources, state guards, date rule or calculation reference it.
    /// </summary>
    internal static List<string> GetAffectedColumnsForInput(string tableName, string headerName)
    {
        EnsureInitialized();
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(headerName))
            return result;

        if (!TryResolveTableDictionary(tableName, out var colDict))
            return result;

        string normHeader = NormalizeIdentifier(headerName);
        string stripped = normHeader.TrimEnd('s');

        bool MatchesHeader(string? candidate)
        {
            if (string.IsNullOrWhiteSpace(candidate)) return false;
            string norm = NormalizeIdentifier(candidate);
            return norm.Length > 0 &&
                   (norm.Equals(normHeader, StringComparison.OrdinalIgnoreCase) ||
                    norm.TrimEnd('s').Equals(stripped, StringComparison.OrdinalIgnoreCase));
        }

        string Canonical(string raw) => ResolveHeaderInDictionary(colDict, raw)?.HeaderName ?? raw;

        // a) The source column's own outgoing rules
        var self = ResolveHeaderInDictionary(colDict, headerName);
        if (self != null)
        {
            foreach (var sync in self.SyncRules)
            {
                if (!string.IsNullOrWhiteSpace(sync.TargetHeader))
                    result.Add(Canonical(sync.TargetHeader));
            }
            if (!string.IsNullOrWhiteSpace(self.CountTargetHeader))
                result.Add(Canonical(self.CountTargetHeader!));
        }

        // b) Sibling columns that depend on this column
        foreach (var d in colDict.Values)
        {
            if (d.NormalizedHeader.Equals(normHeader, StringComparison.OrdinalIgnoreCase)) continue;

            bool depends = false;

            if (d.Sources.Any(src =>
                {
                    if (MatchesHeader(src)) return true;
                    int open = src.IndexOf('[');
                    int close = open >= 0 ? src.IndexOf(']', open + 1) : -1;
                    return close > open + 1 && MatchesHeader(src.Substring(open + 1, close - open - 1));
                }))
            {
                depends = true;
            }

            if (!depends && d.StateGuards.Any(g => g.ReferenceHeaders.Any(MatchesHeader)))
                depends = true;

            if (!depends && MatchesHeader(d.DateMustBeGreaterThanHeader))
                depends = true;

            if (!depends && d.CalculateRule != null)
            {
                string expr = RuleCalculator.GetColumnExpressionOverride(tableName, d.HeaderName) ?? d.CalculateRule.Expression ?? string.Empty;
                if (BracketTokenRegex.Matches(expr).Any(m => MatchesHeader(m.Groups[1].Value)))
                    depends = true;
            }

            if (depends) result.Add(d.HeaderName);
        }

        return result
            .Where(h => !MatchesHeader(h))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    /// <summary>Returns the sibling columns referenced (as [Column] tokens) by the column's C# calculation expression.</summary>
    internal static List<string> GetCalculationSources(string tableName, string headerName)
    {
        EnsureInitialized();
        if (string.IsNullOrWhiteSpace(tableName) || string.IsNullOrWhiteSpace(headerName))
            return new List<string>();

        if (!TryResolveTableDictionary(tableName, out var colDict))
            return new List<string>();

        var rule = ResolveHeaderInDictionary(colDict, headerName);
        if (rule?.CalculateRule == null)
            return new List<string>();

        string expr = RuleCalculator.GetColumnExpressionOverride(tableName, rule.HeaderName) ?? rule.CalculateRule.Expression ?? string.Empty;
        string normSelf = NormalizeIdentifier(rule.HeaderName);

        return BracketTokenRegex.Matches(expr)
            .Select(m => m.Groups[1].Value.Trim())
            .Where(raw => raw.Length > 0 && !NormalizeIdentifier(raw).Equals(normSelf, StringComparison.OrdinalIgnoreCase))
            .Select(raw => ResolveHeaderInDictionary(colDict, raw)?.HeaderName ?? raw)
            .Where(h => !NormalizeIdentifier(h).Equals(normSelf, StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    internal static string NormalizeIdentifier(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        string clean = raw.Trim()
            .Replace("'s", "", StringComparison.OrdinalIgnoreCase)
            .Replace("’s", "", StringComparison.OrdinalIgnoreCase)
            .Replace(" s ", " ", StringComparison.OrdinalIgnoreCase);

        // Common header alias normalization across templates
        clean = Regex.Replace(clean, @"[\s_\-\.'`]+", string.Empty).ToLowerInvariant();
        if (clean == "workerid") return "staffid";
        if (clean == "workername") return "staffname";
        if (clean == "cost" || clean == "transactionamount") return "amount";
        if (clean == "insuracerenewalid") return "insurancerenewalid";
        if (clean == "recieptnumber") return "receiptnumber";
        if (clean == "cardnumber") return "fuelcardnumber";
        if (clean == "othours") return "overtimehours";
        if (clean == "additionalothours") return "additionalovertimehours";
        if (clean == "otshifts") return "overtimeshifts";
        return clean;
    }

    private static bool TryResolveTableDictionary(
        string tableName,
        out Dictionary<string, PropertyRuleDescriptor> colDict)
    {
        string norm = NormalizeIdentifier(tableName);
        if (_tableMap.TryGetValue(norm, out colDict!))
            return true;

        string stripped = norm.TrimEnd('s');
        foreach (var (k, v) in _tableMap)
        {
            if (k.TrimEnd('s').Equals(stripped, StringComparison.OrdinalIgnoreCase))
            {
                colDict = v;
                return true;
            }
        }

        colDict = null!;
        return false;
    }

    private static PropertyRuleDescriptor? ResolveHeaderInDictionary(
        Dictionary<string, PropertyRuleDescriptor> colDict,
        string header)
    {
        string norm = NormalizeIdentifier(header);
        if (colDict.TryGetValue(norm, out var exact))
            return exact;

        string stripped = norm.TrimEnd('s');
        foreach (var (k, v) in colDict)
        {
            if (k.TrimEnd('s').Equals(stripped, StringComparison.OrdinalIgnoreCase))
                return v;
        }

        // Tolerant fallback for ID/Cost/Contact variations
        if (norm.EndsWith("id") && colDict.Values.Count(d => d.NormalizedHeader.EndsWith("id") && d.DefaultRule?.Kind == DefaultKind.AutoIncrement) == 1)
        {
            return colDict.Values.First(d => d.NormalizedHeader.EndsWith("id") && d.DefaultRule?.Kind == DefaultKind.AutoIncrement);
        }

        if (norm.EndsWith("cost") && colDict.TryGetValue(norm.Replace("cost", ""), out var costMatch))
            return costMatch;

        if (!norm.EndsWith("cost") && colDict.TryGetValue(norm + "cost", out var addCostMatch))
            return addCostMatch;

        return null;
    }
}