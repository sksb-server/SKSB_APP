namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using SKSB_App.Core.RuntimeUI.DropDown;
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
    public bool IsMultiSelect { get; init; }
    public string? CountTargetHeader { get; init; }
    public bool AllowNonNegativeZero { get; init; } = true;
    public TransformKind Transform { get; init; } = TransformKind.None;

    public List<ColumnSyncAttribute> SyncRules { get; init; } = new();
    public ColumnDefaultAttribute? DefaultRule { get; init; }
    public List<ColumnStateGuardAttribute> StateGuards { get; init; } = new();
    public string? DateMustBeGreaterThanHeader { get; init; }

    public double RoundDownStep { get; init; }

    public ColumnFormatAttribute? FormatRule { get; init; }
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

    /// <summary>True when C# declares a formula/rollup for the column (any mode); legacy Excel formulas are stripped.</summary>
    public bool HasCSharpFormula => CalculateRule != null || RollupRules.Count > 0;

    public string? DisplayFormula =>
        CalculateRule != null ? $"= {CalculateRule.Expression}" :
        RollupRules.Count > 0 ? "= " + string.Join(" + ", RollupRules.Select(r =>
            $"{r.Kind.ToString().ToUpperInvariant()}({r.CollectionOrTable}[{string.Join("+", r.SourceProperties)}])")) :
        null;

    public bool HasDropdown =>
        Sources.Length > 0 ||
        (DefaultRule != null && DefaultRule.Kind != DefaultKind.None) ||
        (CalculateRule != null && CalculateRule.Mode == CalculationMode.DefaultOnly);

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
            var ruleAttr = prop.GetCustomAttribute<ColumnRuleAttribute>();
            var defaultAttr = prop.GetCustomAttribute<ColumnDefaultAttribute>();
            var dateGreaterAttr = prop.GetCustomAttribute<ColumnDateGreaterThanAttribute>();
            var formatAttr = prop.GetCustomAttribute<ColumnFormatAttribute>();
            var calcAttr = prop.GetCustomAttribute<ColumnCalculateAttribute>();
            var rollupAttrs = prop.GetCustomAttributes<ColumnRollupAttribute>().ToList();
            var syncAttrs = prop.GetCustomAttributes<ColumnSyncAttribute>().ToList();
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
                IsMultiSelect = ruleAttr?.IsMultiSelect ?? false,
                CountTargetHeader = ruleAttr?.CountTargetHeader,
                AllowNonNegativeZero = ruleAttr?.AllowNonNegativeZero ?? true,
                Transform = ruleAttr?.Transform ?? TransformKind.None,
                SyncRules = syncAttrs,
                DefaultRule = defaultAttr,
                StateGuards = guardAttrs,
                DateMustBeGreaterThanHeader = dateGreaterAttr?.ReferenceHeader,
                RoundDownStep = ruleAttr?.RoundDownStep ?? 0.0,
                FormatRule = formatAttr,
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

    public static IReadOnlyList<PropertyRuleDescriptor> GetTableRules(string tableName)
    {
        EnsureInitialized();
        if (string.IsNullOrWhiteSpace(tableName)) return Array.Empty<PropertyRuleDescriptor>();

        if (TryResolveTableDictionary(tableName, out var colDict))
        {
            return colDict.Values.ToList();
        }

        return Array.Empty<PropertyRuleDescriptor>();
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

    public static bool HasDropdown(string tableName, string columnHeader)
    {
        var rule = GetRule(tableName, columnHeader);
        return rule != null && rule.HasDropdown;
    }

    /// <summary>
    /// Auto-locks a C#-calculated column without sheet protection: a custom '=FALSE' Data Validation (Stop alert)
    /// blocks interactive typing while C# COM writes (.Value2) pass through. Paste / Delete are caught by the
    /// SheetChange guard in EventsAutoOpen. Also shows the formula as the cell's input message.
    /// </summary>
    internal static void ApplyCalculatedColumnLock(dynamic targetRange, PropertyRuleDescriptor? rule)
    {
        if (targetRange == null || rule == null || !rule.IsStrictlyCalculated) return;
        try
        {
            try { targetRange.Validation.Delete(); } catch { }

            // Type 7 = xlValidateCustom, AlertStyle 1 = xlValidAlertStop, Operator 1 = xlBetween
            targetRange.Validation.Add(7, 1, 1, "=FALSE");
            targetRange.Validation.IgnoreBlank = false;

            string formulaText = rule.DisplayFormula ?? "Calculated automatically by SKSB.";
            if (formulaText.Length > 220) formulaText = formulaText.Substring(0, 220);

            targetRange.Validation.InputTitle = "Formula Evaluated Cell (Locked)";
            targetRange.Validation.InputMessage = formulaText;
            targetRange.Validation.ShowInput = true;

            targetRange.Validation.ErrorTitle = "Locked Calculated Column";
            string error = $"'{rule.HeaderName}' is automatically calculated ({formulaText}) and cannot be edited manually.";
            targetRange.Validation.ErrorMessage = error.Length > 255 ? error.Substring(0, 255) : error;
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

            Commands.EnsureNamedStyleExists(xlApp, Commands.DropDownStyleName);

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

                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
                    for (int t = 1; t <= tableCount; t++)
                    {
                        dynamic table = ws.ListObjects[t];
                        string tableName = table.Name?.ToString() ?? string.Empty;

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

                            if (archFormula.StartsWith("=")) continue;

                            // Inside RuleRegistry.ApplyDropdownStylesToWorkbook:
                            if (rule.HasDropdown)
                            {
                                try
                                {
                                    // Pre-read the template's original format & style on the archetype cell
                                    string origStyle = archCell.Style?.Name?.ToString() ?? string.Empty;
                                    string origFmt = archCell.NumberFormat?.ToString() ?? string.Empty;
                                    dynamic? origFmtLocal = null;
                                    try { origFmtLocal = archCell.NumberFormatLocal; } catch { }

                                    dynamic targetRange = (bodyRange != null) ? table.ListColumns[c].DataBodyRange : archCell;
                                    if (targetRange != null)
                                    {
                                        // Only set DropDown style if the cell doesn't already have a custom unlocked style
                                        if (string.IsNullOrEmpty(origStyle) || origStyle.Equals("Normal", StringComparison.OrdinalIgnoreCase))
                                        {
                                            targetRange.Style = Commands.DropDownStyleName;
                                        }
                                        try { targetRange.Validation.Delete(); } catch { }

                                        // Restore the template's cell-level NumberFormat (e.g. Date format) if the style reset it
                                        if (!string.IsNullOrWhiteSpace(origFmt) && !origFmt.Equals("General", StringComparison.OrdinalIgnoreCase))
                                        {
                                            try { targetRange.NumberFormat = origFmt; } catch { }
                                            try { if (origFmtLocal != null) targetRange.NumberFormatLocal = origFmtLocal; } catch { }
                                        }
                                    }
                                }
                                catch { }
                            }

                            ApplyColumnNumberFormat(table, c, archCell, rule);
                        }
                    }
                }));

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