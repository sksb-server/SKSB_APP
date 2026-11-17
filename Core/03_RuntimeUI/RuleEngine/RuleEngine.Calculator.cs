namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.SheetMapper;

// =========================================================================
// 1. WAGE PROVIDER (effective-dated worker wage resolution)
// =========================================================================

/// <summary>
/// Resolves the hourly rate used by WAGE_RATE(...). Until a Staff Manager wages sheet registers per-worker
/// schedules, every lookup returns the sheet's flat @WorkerHourlyRate.
/// </summary>
public static class WageProvider
{
    private static readonly ConcurrentDictionary<string, SortedDictionary<DateTime, double>> _schedules =
        new(StringComparer.OrdinalIgnoreCase);

    public static void RegisterWorkerWages(string workerIdentifier, IDictionary<DateTime, double> dateToHourlyWage)
    {
        if (string.IsNullOrWhiteSpace(workerIdentifier) || dateToHourlyWage == null) return;
        _schedules[workerIdentifier.Trim()] = new SortedDictionary<DateTime, double>(dateToHourlyWage);
    }

    public static void ClearSchedules() => _schedules.Clear();

    public static decimal ResolveRate(IReadOnlyList<string>? workers, DateOnly date, decimal fallbackRate)
    {
        if (workers == null || workers.Count == 0 || _schedules.IsEmpty) return fallbackRate;

        DateTime asOf = date.ToDateTime(TimeOnly.MinValue);
        decimal total = 0m;
        int counted = 0;

        foreach (string worker in workers)
        {
            decimal rate = fallbackRate;
            if (!string.IsNullOrWhiteSpace(worker) && _schedules.TryGetValue(worker.Trim(), out var schedule))
            {
                foreach (var (effective, wage) in schedule)
                {
                    if (effective > asOf) break;
                    rate = (decimal)wage;
                }
            }
            total += rate;
            counted++;
        }

        return counted == 0 ? fallbackRate : total / counted;
    }
}

// =========================================================================
// 2. RULE CALCULATOR (expression engine + object-level rollups)
// =========================================================================

public static class RuleCalculator
{
    // =========================================================================
    // 2.1 STEP ROUNDING
    // =========================================================================

    /// <summary>Rounds down to a multiple of <paramref name="step"/> (1.8 -> 1.5, 1.4 -> 1.0 for step 0.5).</summary>
    public static decimal FloorToStep(decimal value, decimal step)
    {
        if (step <= 0m) return value;
        return Math.Floor((value / step) + 0.0000001m) * step;
    }

    // =========================================================================
    // 2.2 SHEET PARAMETERS
    // =========================================================================

    /// <summary>
    /// Reads a [SheetParameter] value: the primary cell on <paramref name="worksheet"/>, then the fallback
    /// sheet/cell, then the attribute default. Non-positive cell values count as "not set".
    /// </summary>
    public static decimal ReadSheetParameter(
        object? wb, object? worksheet, IReadOnlyList<SheetParameterAttribute> parameters, string paramName)
    {
        var p = parameters?.FirstOrDefault(x => x.Name.Equals(paramName, StringComparison.OrdinalIgnoreCase));
        if (p == null) return 0m;

        try
        {
            if (worksheet != null)
            {
                dynamic ws = worksheet;
                decimal v = ToDecimal((object?)ws.Range[p.CellAddress].Value2);
                if (v > 0m) return v;
            }
        }
        catch { }

        try
        {
            if (wb != null && !string.IsNullOrWhiteSpace(p.FallbackSheet) && !string.IsNullOrWhiteSpace(p.FallbackAddress))
            {
                dynamic book = wb;
                dynamic fb = book.Worksheets[p.FallbackSheet];
                decimal v = ToDecimal((object?)fb.Range[p.FallbackAddress].Value2);
                if (v > 0m) return v;
            }
        }
        catch { }

        return (decimal)p.DefaultValue;
    }

    /// <summary>Reads every declared parameter once (a handful of COM calls per table, none per row).</summary>
    public static Dictionary<string, decimal> ReadSheetParameters(
        object? wb, object? worksheet, IReadOnlyList<SheetParameterAttribute> parameters)
    {
        var result = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        if (parameters == null) return result;
        foreach (var p in parameters)
            result[p.Name] = ReadSheetParameter(wb, worksheet, parameters, p.Name);
        return result;
    }

    // =========================================================================
    // 2.3 EXPRESSION EVALUATION (Excel row or C# object)
    // =========================================================================

    private static readonly ConcurrentDictionary<string, ExprNode> _astCache = new(StringComparer.Ordinal);

    /// <summary>
    /// Null-aware evaluation of an Excel row. null means "no value" (shown as an em-dash) and propagates:
    /// a difference / product / quotient with a null operand is null, a sum is null only when every operand is null,
    /// and a genuine 0 stays 0.
    /// </summary>
    public static decimal? EvaluateRowExpressionNullable(
        string expression, string[] colHeaders, object?[] rawObjects, Func<string, decimal> parameterResolver)
    {
        return EvaluateRowExpressionNullable(expression,
            name =>
            {
                int c = RuleEvaluator.FindColumnByHeader(colHeaders, name);
                return c >= 1 && c < rawObjects.Length ? rawObjects[c] : null;
            },
            parameterResolver);
    }

    public static decimal? EvaluateRowExpressionNullable(
        string expression, Func<string, object?> columnResolver, Func<string, decimal> parameterResolver)
    {
        if (string.IsNullOrWhiteSpace(expression)) return null;
        try
        {
            var ast = _astCache.GetOrAdd(expression, e => new Parser(e).ParseAll());
            return ToNullableDecimal(ast.Eval(new EvalContext(columnResolver, parameterResolver)));
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RuleCalculator] '{expression}' failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>Evaluates a rule's expression, then applies DecimalPlaces and RoundDownStep. null => no value.</summary>
    public static decimal? EvaluateColumnRule(
        PropertyRuleDescriptor rule, string[] colHeaders, object?[] rawObjects, Func<string, decimal> parameterResolver)
    {
        var calc = rule.CalculateRule;
        if (calc == null) return null;
        decimal? v = EvaluateRowExpressionNullable(calc.Expression, colHeaders, rawObjects, parameterResolver);
        return v.HasValue ? FinalizeValue(v.Value, calc.DecimalPlaces, rule.RoundDownStep) : null;
    }

    private static decimal FinalizeValue(decimal v, int decimalPlaces, double roundDownStep)
    {
        if (roundDownStep > 0) v = FloorToStep(v, (decimal)roundDownStep);
        return Math.Round(v, Math.Clamp(decimalPlaces, 0, 15), MidpointRounding.AwayFromZero);
    }

    /// <summary>A cell / property value as a nullable number: blank, the em-dash, COM errors and nulls are null; 0 stays 0.</summary>
    internal static decimal? ToNullableDecimal(object? v)
    {
        switch (v)
        {
            case null: return null;
            case string s:
                {
                    string t = s.Trim();
                    if (t.Length == 0 || t == "\u2014" || t == "-" || t.StartsWith("#")) return null;
                    return decimal.TryParse(t, NumberStyles.Any, CultureInfo.InvariantCulture, out var r) ? r : null;
                }
            case DateOnly or DateTime: return ToDecimal(v);
            case bool: return ToDecimal(v);
            default:
                if (SheetMapperHelper.IsComError(v)) return null;
                return ToDecimal(v);
        }
    }

    // =========================================================================
    // 2.4 OBJECT-LEVEL CALCULATIONS & ROLLUPS
    // =========================================================================

    private sealed record RoundProp(PropertyInfo Property, double Step);
    private sealed record CalcProp(PropertyInfo Property, ColumnCalculateAttribute Calc, double Step);

    private sealed class TypePlan
    {
        public List<RoundProp> Rounding { get; } = new();
        public List<CalcProp> Calcs { get; } = new();
    }

    private static readonly ConcurrentDictionary<Type, TypePlan> _plans = new();
    private static readonly ConcurrentDictionary<Type, Dictionary<string, PropertyInfo>> _headerMaps = new();

    private static TypePlan GetPlan(Type type) =>
        _plans.GetOrAdd(type, t =>
        {
            var plan = new TypePlan();
            foreach (var p in t.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!p.CanWrite || p.GetIndexParameters().Length > 0) continue;
                double step = p.GetCustomAttribute<ColumnRuleAttribute>()?.RoundDownStep ?? 0.0;
                if (step > 0) plan.Rounding.Add(new RoundProp(p, step));
                var calc = p.GetCustomAttribute<ColumnCalculateAttribute>();
                if (calc != null) plan.Calcs.Add(new CalcProp(p, calc, step));
            }
            return plan;
        });

    private static Dictionary<string, PropertyInfo> GetHeaderMap(Type t) =>
        _headerMaps.GetOrAdd(t, type =>
        {
            var map = new Dictionary<string, PropertyInfo>(StringComparer.OrdinalIgnoreCase);
            foreach (var p in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (p.GetIndexParameters().Length > 0) continue;
                var col = p.GetCustomAttribute<ExcelColumnAttribute>();
                if (col != null && !string.IsNullOrWhiteSpace(col.HeaderName))
                    map[RuleRegistry.NormalizeIdentifier(col.HeaderName)] = p;
                map.TryAdd(RuleRegistry.NormalizeIdentifier(p.Name), p);
            }
            return map;
        });

    private static bool IsChildCollection(PropertyInfo p, out Type? elementType)
    {
        elementType = null;
        if (p.PropertyType == typeof(string) || !typeof(IEnumerable).IsAssignableFrom(p.PropertyType)) return false;
        if (!p.PropertyType.IsGenericType) return false;
        var arg = p.PropertyType.GetGenericArguments()[0];
        if (!arg.IsClass || arg == typeof(string)) return false;
        elementType = arg;
        return true;
    }

    private static IEnumerable<object> DataDetailChildren(object record)
    {
        foreach (var p in record.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
        {
            if (p.GetIndexParameters().Length > 0 || p.GetCustomAttribute<DataDetailAttribute>() == null) continue;
            var child = p.GetValue(record);
            if (child != null) yield return child;
        }
    }

    private static Func<string, decimal> BuildParameterResolver(
        IReadOnlyDictionary<string, decimal>? parameters, IEnumerable<Type> types)
    {
        var defaults = new Dictionary<string, decimal>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in types)
            foreach (var a in t.GetCustomAttributes<SheetParameterAttribute>(true))
                defaults.TryAdd(a.Name, (decimal)a.DefaultValue);

        return name =>
        {
            if (parameters != null && parameters.TryGetValue(name, out var v) && v > 0m) return v;
            return defaults.TryGetValue(name, out var d) ? d : 0m;
        };
    }

    private static Func<string, object?> BuildObjectResolver(IReadOnlyList<object> targets) =>
        name =>
        {
            string norm = RuleRegistry.NormalizeIdentifier(name);
            foreach (var target in targets)
                if (GetHeaderMap(target.GetType()).TryGetValue(norm, out var p)) return p.GetValue(target);
            return null;
        };

    /// <summary>
    /// Enforces RoundDownStep and evaluates [ColumnCalculate] (two passes so dependent columns settle)
    /// on a single child row object.
    /// </summary>
    private static void ApplyRowCalculationsOnObject(object item, IReadOnlyDictionary<string, decimal>? parameters = null)
    {
        ApplyCalculations(item, new List<object> { item }, parameters, includeRounding: true);
    }

    private static void ApplyCalculations(
        object owner, IReadOnlyList<object> targets, IReadOnlyDictionary<string, decimal>? parameters, bool includeRounding)
    {
        var plan = GetPlan(owner.GetType());

        if (includeRounding)
        {
            foreach (var r in plan.Rounding)
                if (ToNullableDecimal(r.Property.GetValue(owner)) is decimal current)
                    SetNumeric(owner, r.Property, FloorToStep(current, (decimal)r.Step));
        }

        if (plan.Calcs.Count == 0) return;

        var resolveParam = BuildParameterResolver(parameters, targets.Select(t => t.GetType()));
        var resolveColumn = BuildObjectResolver(targets);

        for (int pass = 0; pass < 2; pass++)
        {
            foreach (var cp in plan.Calcs)
            {
                if (cp.Calc.Mode == CalculationMode.DefaultOnly)
                {
                    object? cur = cp.Property.GetValue(owner);
                    if (cur is string s ? !string.IsNullOrWhiteSpace(s) : ToDecimal(cur) != 0m) continue;
                }

                decimal? raw = EvaluateRowExpressionNullable(cp.Calc.Expression, resolveColumn, resolveParam);
                decimal? v = raw.HasValue ? FinalizeValue(raw.Value, cp.Calc.DecimalPlaces, cp.Step) : null;
                if (cp.Calc.Mode == CalculationMode.DefaultOnly && (!v.HasValue || v.Value == 0m)) continue;
                SetNumeric(owner, cp.Property, v); // null => nullable property cleared (shown as the em-dash)
            }
        }
    }

    /// <summary>
    /// Applies child row calculations, parent rollups and parent calculations to <paramref name="record"/>
    /// (and its [DataDetail] members), in that order. Rollups with a DateProperty honour the date filter.
    /// </summary>
    public static void ApplyRecordCalculations<TRecord>(
        TRecord record,
        IReadOnlyDictionary<string, decimal>? parameters = null,
        IReadOnlyDictionary<string, decimal>? fallbacks = null,
        DateOnly? filterAnchorDate = null,
        DateFilterScope filterScope = DateFilterScope.All_Time)
    {
        if (record == null) return;
        object root = record;
        var rootProps = root.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance);
        DateOnly anchor = filterAnchorDate ?? DateOnly.FromDateTime(DateTime.Today);

        // Step A: child entity collections (e.g. Maintenances, OdometerLogs)
        var seenCollections = new HashSet<object>(ReferenceEqualityComparer.Instance);
        foreach (var cp in rootProps)
        {
            if (!IsChildCollection(cp, out _) || cp.GetValue(root) is not IEnumerable items) continue;
            if (!seenCollections.Add(items)) continue; // alias properties share one list
            foreach (var child in items)
                if (child != null) ApplyRowCalculationsOnObject(child, parameters);
        }

        var owners = new List<object> { root };
        owners.AddRange(DataDetailChildren(root));

        // Step B: rollups
        foreach (var owner in owners)
        {
            foreach (var p in owner.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!p.CanWrite) continue;
                var rollups = p.GetCustomAttributes<ColumnRollupAttribute>().Where(r => !r.IsCrossTable).ToList();
                if (rollups.Count == 0) continue;

                if (rollups.All(r => r.Mode == CalculationMode.DefaultOnly))
                {
                    object? cur = p.GetValue(owner);
                    if (cur is string s ? !string.IsNullOrWhiteSpace(s) : (cur != null && ToDecimal(cur) != 0m)) continue;
                }

                ApplyRollupsToProperty(owner, p, rollups, root, rootProps, fallbacks, anchor, filterScope);
            }
        }

        // Step B2: sentinel dates (0001-01-01 / 1900) loaded from the database or a JSON snapshot are not dates
        foreach (var owner in owners)
        {
            foreach (var p in owner.GetType().GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (p.PropertyType != typeof(DateOnly?) || !p.CanWrite) continue;
                if (p.GetValue(owner) is DateOnly d && d.Year <= 1900) p.SetValue(owner, null);
            }
        }

        // Step C: parent calculations (2 passes inside: Actual Project Cost settles before Profit / Margin)
        foreach (var owner in owners)
            ApplyCalculations(owner, owners, parameters, includeRounding: false);
    }

    private static void ApplyRollupsToProperty(
        object owner, PropertyInfo target, List<ColumnRollupAttribute> rollups, object root, PropertyInfo[] rootProps,
        IReadOnlyDictionary<string, decimal>? fallbacks, DateOnly anchor, DateFilterScope scope)
    {
        Type targetType = Nullable.GetUnderlyingType(target.PropertyType) ?? target.PropertyType;
        decimal total = 0m;
        bool anyData = false;
        var joined = new List<string>();
        bool anyDate = false;
        DateOnly bestDate = default;
        bool anyItems = false;
        string? fallbackKey = rollups.Select(r => r.FallbackKey).FirstOrDefault(k => !string.IsNullOrWhiteSpace(k));

        foreach (var rollup in rollups)
        {
            var collectionProp = rootProps.FirstOrDefault(x => x.Name == rollup.CollectionOrTable);
            if (collectionProp == null || !IsChildCollection(collectionProp, out var elementType)) continue;

            var items = (collectionProp.GetValue(root) as IEnumerable)?.Cast<object>().Where(o => o != null).ToList()
                        ?? new List<object>();

            if (!string.IsNullOrWhiteSpace(rollup.DateProperty) && scope != DateFilterScope.All_Time && scope != DateFilterScope.None)
            {
                var dateProp = elementType!.GetProperty(rollup.DateProperty!);
                if (dateProp != null)
                    items = items.Where(i => !TryGetDate(dateProp.GetValue(i), out var d) || scope.IsWithinScope(d, anchor)).ToList();
            }

            var sources = rollup.SourceProperties
                .Select(n => elementType!.GetProperty(n))
                .Where(sp => sp != null).Select(sp => sp!).ToList();

            // Date targets: earliest / latest of the first source property
            if (targetType == typeof(DateOnly))
            {
                if (sources.Count == 0) continue;
                foreach (var raw in items.Select(i => sources[0].GetValue(i)))
                {
                    if (!TryGetDate(raw, out var d)) continue;
                    if (!anyDate || (rollup.Kind == RollupKind.Min ? d < bestDate : d > bestDate)) bestDate = d;
                    anyDate = true;
                }
                continue;
            }

            if (rollup.Kind == RollupKind.DistinctJoin)
            {
                if (sources.Count == 0) continue;
                foreach (var v in items.Select(i => sources[0].GetValue(i)?.ToString()?.Trim()))
                    if (!string.IsNullOrWhiteSpace(v) && !joined.Contains(v!, StringComparer.OrdinalIgnoreCase)) joined.Add(v!);
                continue;
            }

            decimal? part = RollupNumber(items, sources, rollup.Kind);
            if (part.HasValue) { total += part.Value; anyData = true; }
            anyItems |= items.Count > 0;
        }

        if (targetType == typeof(DateOnly))
        {
            if (anyDate)
            {
                target.SetValue(owner, bestDate);
            }
            else if (Nullable.GetUnderlyingType(target.PropertyType) != null)
            {
                // No child row carries a valid date: clear stale / 0001-01-01 sentinel values (never 00/01/1900).
                target.SetValue(owner, null);
            }
            return;
        }

        if (targetType == typeof(string))
        {
            if (rollups.Any(r => r.Kind == RollupKind.DistinctJoin)) target.SetValue(owner, string.Join(", ", joined));
            return;
        }

        // A non-nullable target (persisted column) keeps the sheet / database value when there is nothing to roll up.
        bool targetNullable = Nullable.GetUnderlyingType(target.PropertyType) != null;
        if (!anyData && !targetNullable) return;

        // No matching child rows (or only null source values) => null, shown as the em-dash. A genuine 0 stays 0.
        decimal? result = anyData ? total : null;
        if ((!result.HasValue || result.Value == 0m) && !string.IsNullOrWhiteSpace(fallbackKey) &&
            fallbacks != null && fallbacks.TryGetValue(fallbackKey!, out var fb))
            result = fb;

        SetNumeric(owner, target, result);
    }

    /// <summary>Aggregates child rows; null when there are no rows or every source value is null.</summary>
    private static decimal? RollupNumber(List<object> items, List<PropertyInfo> sources, RollupKind kind)
    {
        if (items.Count == 0) return null;

        static bool Present(object? v) => v != null && !(v is string s && string.IsNullOrWhiteSpace(s));

        if (kind == RollupKind.Count)
        {
            if (sources.Count == 0) return items.Count;
            int c = items.Count(i => sources.Any(sp => Present(sp.GetValue(i))));
            return c == 0 ? null : c;
        }

        if (sources.Count == 0) return null;

        if (kind == RollupKind.DistinctCount)
        {
            int d = items.Select(i => sources[0].GetValue(i)).Where(Present)
                         .Select(v => v!.ToString()!).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            return d == 0 ? null : d;
        }

        // One value per row (sum of its non-null sources); rows without any value do not take part.
        var perItem = items
            .Select(i => sources.Select(sp => ToNullableDecimal(sp.GetValue(i))).Where(x => x.HasValue).Select(x => x!.Value).ToList())
            .Where(values => values.Count > 0)
            .Select(values => values.Sum())
            .ToList();
        if (perItem.Count == 0) return null;

        return kind switch
        {
            RollupKind.Sum => perItem.Sum(),
            RollupKind.Min => perItem.Min(),
            RollupKind.Max or RollupKind.PriorMax => perItem.Max(),
            RollupKind.Average => perItem.Average(),
            _ => null
        };
    }

    private static bool TryGetDate(object? raw, out DateOnly date)
    {
        switch (raw)
        {
            case DateOnly d: date = d; return d.Year > 1900;           // 0001-01-01 / 00-01-1900 are sentinels
            case DateTime dt: date = DateOnly.FromDateTime(dt); return dt.Year > 1900;
            default: date = default; return false;
        }
    }

    /// <summary>Sets a numeric property. null clears a nullable property (em-dash) and zeroes a non-nullable one.</summary>
    private static void SetNumeric(object target, PropertyInfo p, decimal? value)
    {
        Type t = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;
        bool nullable = Nullable.GetUnderlyingType(p.PropertyType) != null;
        try
        {
            if (!value.HasValue)
            {
                if (nullable) p.SetValue(target, null);
                else SetNumeric(target, p, 0m);
                return;
            }

            decimal v = value.Value;
            object? boxed =
                t == typeof(decimal) ? v :
                t == typeof(double) ? (double)v :
                t == typeof(int) ? (int)Math.Round(v, MidpointRounding.AwayFromZero) :
                t == typeof(long) ? (long)Math.Round(v, MidpointRounding.AwayFromZero) :
                t == typeof(float) ? (float)v :
                null;
            if (boxed != null) p.SetValue(target, boxed);
        }
        catch { }
    }

    // =========================================================================
    // 2.4b CROSS-TABLE WORKBOOK ROLLUPS (WorkshopsTable, InsuranceCompaniesTable, OdometersTable)
    // =========================================================================

    private sealed record SourceTable(string[] Headers, object?[,]? Data, int Rows);

    /// <summary>
    /// Evaluates cross-table [ColumnRollup] rules (ForeignKeyHeader == LocalKeyHeader) for every row of a table,
    /// entirely in RAM. Each source table is read once with a single bulk COM call. Changed cells are reported via
    /// <paramref name="queueWrite"/>(row, col, value, rule); all indices are 1-based within the table body.
    /// </summary>
    public static void EvaluateTableCrossRollups(
        object wbObj,
        object worksheetObj,
        PropertyRuleDescriptor?[] colRules,
        string[] colHeaders,
        int rowCount,
        int colCount,
        string[,] rowValuesGrid,
        object?[,] rawObjectsGrid,
        Action<int, int, object, PropertyRuleDescriptor?> queueWrite)
    {
        if (rowCount <= 0 || colCount <= 0) return;

        var sourceCache = new Dictionary<string, SourceTable>(StringComparer.OrdinalIgnoreCase);
        (DateOnly Anchor, DateFilterScope Scope)? filter = null;

        for (int c = 1; c <= colCount; c++)
        {
            var rule = colRules[c];
            var rollups = rule?.RollupRules.Where(r => r.IsCrossTable).ToList();
            if (rule == null || rollups == null || rollups.Count == 0) continue;

            int localCol = RuleEvaluator.FindColumnByHeader(colHeaders, rollups[0].LocalKeyHeader!);
            if (localCol < 1) continue;

            bool isText = (Nullable.GetUnderlyingType(rule.PropertyType) ?? rule.PropertyType) == typeof(string);
            bool allDefaultOnly = rollups.All(x => x.Mode == CalculationMode.DefaultOnly);

            for (int r = 1; r <= rowCount; r++)
            {
                string current = rowValuesGrid[r, c] ?? string.Empty;
                if (allDefaultOnly && !(string.IsNullOrWhiteSpace(current) || current == "0" || current == "0.00")) continue;

                string key = KeyText(rawObjectsGrid[r, localCol]);
                if (key.Length == 0) continue;

                decimal total = 0m;
                bool anyData = false;
                var joined = new List<string>();

                foreach (var rollup in rollups)
                {
                    var src = LoadSource(rollup.CollectionOrTable);
                    var data = src.Data;
                    if (data == null || src.Rows == 0) continue;

                    int fkCol = RuleEvaluator.FindColumnByHeader(src.Headers, rollup.ForeignKeyHeader!);
                    if (fkCol < 1) continue;
                    var srcCols = rollup.SourceProperties
                        .Select(h => RuleEvaluator.FindColumnByHeader(src.Headers, h)).Where(i => i >= 1).ToList();
                    int dateCol = string.IsNullOrWhiteSpace(rollup.DateProperty)
                        ? -1 : RuleEvaluator.FindColumnByHeader(src.Headers, rollup.DateProperty!);

                    bool prior = rollup.Kind == RollupKind.PriorMax;
                    bool selfTable = string.Equals(rollup.CollectionOrTable, rule.TableName, StringComparison.OrdinalIgnoreCase);
                    DateOnly anchor = default;
                    DateFilterScope scope = DateFilterScope.All_Time;
                    if (!prior && dateCol >= 1) (anchor, scope) = ReadFilter();

                    DateOnly? ownDate = null;
                    if (prior && !string.IsNullOrWhiteSpace(rollup.DateProperty))
                    {
                        int ownDateCol = RuleEvaluator.FindColumnByHeader(colHeaders, rollup.DateProperty!);
                        if (ownDateCol >= 1 && RuleDataCache.TryParseDate(rawObjectsGrid[r, ownDateCol], out var od)) ownDate = od;
                    }

                    var matches = new List<int>();
                    for (int s = 1; s <= src.Rows; s++)
                    {
                        if (!string.Equals(KeyText(data[s, fkCol]), key, StringComparison.OrdinalIgnoreCase)) continue;

                        if (prior)
                        {
                            // Strictly earlier by date, or same date and earlier in the table.
                            if (selfTable && s == r) continue;
                            bool hasDates = ownDate.HasValue && dateCol >= 1 && RuleDataCache.TryParseDate(data[s, dateCol], out _);
                            if (hasDates)
                            {
                                RuleDataCache.TryParseDate(data[s, dateCol], out var sd);
                                if (sd > ownDate!.Value) continue;
                                if (sd == ownDate.Value && selfTable && s > r) continue;
                            }
                            else if (selfTable && s > r) continue;
                        }
                        else if (dateCol >= 1 && RuleDataCache.TryParseDate(data[s, dateCol], out var rowDate) &&
                                 !scope.IsWithinScope(rowDate, anchor))
                        {
                            continue;
                        }

                        matches.Add(s);
                    }

                    // Per matching row: the sum of its non-null source cells (rows without a value are ignored).
                    decimal? RowValue(int srcRow)
                    {
                        decimal sum = 0m;
                        bool any = false;
                        foreach (int sc in srcCols)
                            if (ToNullableDecimal(data[srcRow, sc]) is decimal cell) { sum += cell; any = true; }
                        return any ? sum : null;
                    }

                    switch (rollup.Kind)
                    {
                        case RollupKind.Count:
                            {
                                int n = srcCols.Count == 0
                                    ? matches.Count
                                    : matches.Count(s => KeyText(data[s, srcCols[0]]).Length > 0);
                                if (n > 0) { total += n; anyData = true; }
                                break;
                            }
                        case RollupKind.DistinctCount:
                            if (srcCols.Count > 0)
                            {
                                int n = matches.Select(s => KeyText(data[s, srcCols[0]])).Where(v => v.Length > 0)
                                               .Distinct(StringComparer.OrdinalIgnoreCase).Count();
                                if (n > 0) { total += n; anyData = true; }
                            }
                            break;
                        case RollupKind.DistinctJoin:
                            if (srcCols.Count > 0)
                                foreach (var v in matches.Select(s => KeyText(data[s, srcCols[0]])))
                                    if (v.Length > 0 && !joined.Contains(v, StringComparer.OrdinalIgnoreCase)) joined.Add(v);
                            break;
                        default:
                            {
                                var values = matches.Select(RowValue).Where(x => x.HasValue).Select(x => x!.Value).ToList();
                                if (values.Count == 0) break;
                                anyData = true;
                                total += rollup.Kind switch
                                {
                                    RollupKind.Sum => values.Sum(),
                                    RollupKind.Min => values.Min(),
                                    RollupKind.Average => values.Average(),
                                    _ => values.Max() // Max, PriorMax
                                };
                                break;
                            }
                    }
                }

                if (isText)
                {
                    string text = string.Join(", ", joined);
                    if (!string.Equals(text, current, StringComparison.Ordinal)) queueWrite(r, c, text, rule);
                }
                else
                {
                    // No matching rows => null (em-dash); matching rows summing to 0 => a real 0.
                    decimal? desired = anyData ? FinalizeValue(total, 2, 0) : null;
                    if (ToNullableDecimal(rawObjectsGrid[r, c]) != desired)
                        queueWrite(r, c, desired.HasValue ? (double)desired.Value : RuleEvaluator.NullDash, rule);
                }
            }
        }

        // ---- local helpers -------------------------------------------------
        SourceTable LoadSource(string tableName)
        {
            if (sourceCache.TryGetValue(tableName, out var cached)) return cached;
            var result = new SourceTable(Array.Empty<string>(), null, 0);
            try
            {
                dynamic? tbl = RuleDataCache.FindTableInWorkbook((object)wbObj, tableName);
                if (tbl != null)
                {
                    int cols = (int)tbl.ListColumns.Count;
                    object[,] hdr = TableMapper.GetRange2DValues(tbl.HeaderRowRange, 1, cols);
                    var headers = new string[cols + 1];
                    for (int h = 1; h <= cols; h++) headers[h] = hdr[1, h]?.ToString()?.Trim() ?? string.Empty;

                    dynamic? body = tbl.DataBodyRange;
                    if (body != null)
                    {
                        int rows = (int)body.Rows.Count;
                        result = new SourceTable(headers, To1Based(TableMapper.GetRange2DValues(body, rows, cols)), rows);
                    }
                    else
                    {
                        result = new SourceTable(headers, null, 0);
                    }
                }
            }
            catch { }
            sourceCache[tableName] = result;
            return result;
        }

        (DateOnly, DateFilterScope) ReadFilter()
        {
            if (filter.HasValue) return filter.Value;
            DateOnly anchor = DateOnly.FromDateTime(DateTime.Today);
            DateFilterScope scope = DateFilterScope.All_Time;
            try
            {
                dynamic ws = worksheetObj;
                if (RuleDataCache.TryParseDate((object?)ws.Range["B2"].Value2, out var a)) anchor = a;
                scope = DateFilterHelper.ParseScope(((object?)ws.Range["C2"].Value2)?.ToString());
            }
            catch { }
            filter = (anchor, scope);
            return filter.Value;
        }
    }

    private static string KeyText(object? v)
    {
        if (v == null) return string.Empty;
        if (v is double d && d == Math.Floor(d) && Math.Abs(d) < 1e15) return ((long)d).ToString(CultureInfo.InvariantCulture);
        return v.ToString()?.Trim() ?? string.Empty;
    }

    private static object?[,] To1Based(object[,] src)
    {
        int r0 = src.GetLowerBound(0), r1 = src.GetUpperBound(0), c0 = src.GetLowerBound(1), c1 = src.GetUpperBound(1);
        var dst = new object?[r1 - r0 + 2, c1 - c0 + 2];
        for (int i = r0; i <= r1; i++)
            for (int j = c0; j <= c1; j++)
                dst[i - r0 + 1, j - c0 + 1] = src[i, j];
        return dst;
    }

    // =========================================================================
    // 2.5 VALUE CONVERSION HELPERS
    // =========================================================================

    internal static decimal ToDecimal(object? v)
    {
        switch (v)
        {
            case null: return 0m;
            case decimal d: return d;
            case double db: return double.IsFinite(db) ? (decimal)db : 0m;
            case float f: return float.IsFinite(f) ? (decimal)f : 0m;
            case int i: return i;
            case long l: return l;
            case bool b: return b ? 1m : 0m;
            case DateOnly dOnly: return (decimal)dOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
            case DateTime dTime: return (decimal)dTime.ToOADate();
            case ICollection col when v is not string: return col.Count;
            default:
                string s = v.ToString()?.Trim() ?? string.Empty;
                if (s.Length == 0 || s.StartsWith("#")) return 0m;
                return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var r) ? r : 0m;
        }
    }

    internal static bool ToBool(object? v)
    {
        switch (v)
        {
            case null: return false;
            case bool b: return b;
            case string s:
                s = s.Trim();
                return s.Equals("true", StringComparison.OrdinalIgnoreCase) ||
                       s.Equals("yes", StringComparison.OrdinalIgnoreCase) ||
                       s.Equals("y", StringComparison.OrdinalIgnoreCase) ||
                       (decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var n) && n != 0m);
            default: return ToDecimal(v) != 0m;
        }
    }

    /// <summary>Converts a DateOnly / DateTime / OADate number / date text into a DateOnly.</summary>
    private static bool TryToDate(object? v, out DateOnly date)
    {
        switch (v)
        {
            case DateOnly d: date = d; return true;
            case DateTime dt: date = DateOnly.FromDateTime(dt); return true;
            case decimal dec when dec >= 15000m && dec <= 80000m:
                date = DateOnly.FromDateTime(DateTime.FromOADate((double)dec)); return true;
            case double db when db >= 15000 && db <= 80000:
                date = DateOnly.FromDateTime(DateTime.FromOADate(db)); return true;
            default: return RuleDataCache.TryParseDate(v, out date);
        }
    }

    private static bool TryParseTime(object? v, out TimeSpan time)
    {
        time = default;
        if (v == null) return false;

        if (v is double or decimal or float or int)
        {
            double d = Convert.ToDouble(v, CultureInfo.InvariantCulture);
            if (d < 0 || d > 1) return false;
            time = TimeSpan.FromDays(d);
            return true;
        }

        string s = v.ToString()?.Trim() ?? string.Empty;
        if (s.Length == 0) return false;
        if (TimeOnly.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out var t))
        {
            time = t.ToTimeSpan();
            return true;
        }
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.InvariantCulture, out double frac) && frac >= 0 && frac <= 1)
        {
            time = TimeSpan.FromDays(frac);
            return true;
        }
        return false;
    }

    private static decimal? HoursBetween(object? start, object? end)
    {
        if (!TryParseTime(start, out var s) || !TryParseTime(end, out var e)) return null;
        double hours = (e - s).TotalHours;
        if (hours < 0) hours += 24.0; // overnight wrap
        return FloorToStep((decimal)Math.Round(hours, 6), 0.5m);
    }

    // =========================================================================
    // 2.6 EXPRESSION PARSER & AST
    // =========================================================================

    private sealed class EvalContext
    {
        public Func<string, object?> Column { get; }
        public Func<string, decimal> Parameter { get; }
        public EvalContext(Func<string, object?> column, Func<string, decimal> parameter)
        {
            Column = column;
            Parameter = parameter;
        }
    }

    private abstract class ExprNode
    {
        public abstract object? Eval(EvalContext ctx);
    }

    private sealed class ConstNode(object? value) : ExprNode
    {
        public override object? Eval(EvalContext ctx) => value;
    }

    private sealed class ColumnNode(string name) : ExprNode
    {
        public override object? Eval(EvalContext ctx)
        {
            object? v = ctx.Column(name);
            // blank cells and the em-dash text are "no value"
            if (v is string s && (s.Trim().Length == 0 || s.Trim() == "\u2014")) return null;
            return v;
        }
    }

    private sealed class ParamNode(string name) : ExprNode
    {
        public override object? Eval(EvalContext ctx) => ctx.Parameter(name);
    }

    private sealed class UnaryNode(char op, ExprNode operand) : ExprNode
    {
        public override object? Eval(EvalContext ctx)
        {
            if (ToNullableDecimal(operand.Eval(ctx)) is not decimal v) return null;
            return op == '-' ? -v : v;
        }
    }

    private sealed class BinaryNode(string op, ExprNode left, ExprNode right) : ExprNode
    {
        public override object? Eval(EvalContext ctx)
        {
            decimal? na = ToNullableDecimal(left.Eval(ctx));
            decimal? nb = ToNullableDecimal(right.Eval(ctx));

            switch (op)
            {
                // Add: null only when every operand is null; otherwise a null operand counts as 0.
                case "+":
                    if (!na.HasValue && !nb.HasValue) return null;
                    return (na ?? 0m) + (nb ?? 0m);

                // Difference / product / quotient: any null operand makes the result null.
                case "-": return na.HasValue && nb.HasValue ? na.Value - nb.Value : null;
                case "*": return na.HasValue && nb.HasValue ? na.Value * nb.Value : null;
                case "/":
                    if (!na.HasValue || !nb.HasValue) return null;
                    return nb.Value == 0m ? 0m : na.Value / nb.Value;
            }

            // Comparisons treat "no value" as 0 (so [Sold Price] > 0 is false when unset).
            decimal a = na ?? 0m, b = nb ?? 0m;
            switch (op)
            {
                case ">": return a > b;
                case "<": return a < b;
                case ">=": return a >= b;
                case "<=": return a <= b;
                case "==": return a == b;
                case "!=": return a != b;
                default: throw new InvalidOperationException($"Unknown operator '{op}'");
            }
        }
    }

    private sealed class CallNode(string name, List<ExprNode> args) : ExprNode
    {
        public override object? Eval(EvalContext ctx)
        {
            switch (name)
            {
                case "IF":
                    Require(3);
                    return ToBool(args[0].Eval(ctx)) ? args[1].Eval(ctx) : args[2].Eval(ctx);

                case "ROUND":
                    Require(1, 2);
                    int digits = args.Count > 1 ? (int)ToDecimal(args[1].Eval(ctx)) : 0;
                    return ToNullableDecimal(args[0].Eval(ctx)) is decimal rv
                        ? Math.Round(rv, Math.Clamp(digits, 0, 15), MidpointRounding.AwayFromZero)
                        : null;

                case "FLOOR":
                    Require(1, 2);
                    decimal step = args.Count > 1 ? ToDecimal(args[1].Eval(ctx)) : 0.5m;
                    return ToNullableDecimal(args[0].Eval(ctx)) is decimal fv ? FloorToStep(fv, step) : null;

                case "MAX":
                    Require(1, int.MaxValue);
                    {
                        var present = args.Select(a => ToNullableDecimal(a.Eval(ctx))).Where(x => x.HasValue).Select(x => x!.Value).ToList();
                        return present.Count == 0 ? null : present.Max();
                    }

                case "MIN":
                    Require(1, int.MaxValue);
                    {
                        var present = args.Select(a => ToNullableDecimal(a.Eval(ctx))).Where(x => x.HasValue).Select(x => x!.Value).ToList();
                        return present.Count == 0 ? null : present.Min();
                    }

                case "COUNT":
                    {
                        Require(1);
                        object? v = args[0].Eval(ctx);
                        return v switch
                        {
                            null => 0m,
                            string s => (decimal)RuleEvaluator.CountMultiSelectTokens(s),
                            ICollection c => (decimal)c.Count,
                            _ => string.IsNullOrWhiteSpace(v.ToString()) ? 0m : 1m
                        };
                    }

                case "TODAY":
                    Require(0);
                    return (decimal)DateTime.Today.ToOADate();

                case "YEARS_BETWEEN":
                    {
                        Require(2);
                        if (!TryToDate(args[0].Eval(ctx), out var from) || !TryToDate(args[1].Eval(ctx), out var to)) return null;
                        int years = to.Year - from.Year;
                        if (to < from.AddYears(years)) years--; // whole elapsed years, like DATEDIF(..., "Y")
                        return (decimal)Math.Max(years, 0);
                    }

                case "HOURS":
                    Require(2);
                    return HoursBetween(args[0].Eval(ctx), args[1].Eval(ctx));

                case "WAGE_RATE":
                    {
                        Require(3);
                        object? w = args[0].Eval(ctx);
                        IReadOnlyList<string> workers = w switch
                        {
                            null => Array.Empty<string>(),
                            string s => RuleEvaluator.SplitMultiSelectTokens(s),
                            IEnumerable<string> list => list.ToList(),
                            _ => RuleEvaluator.SplitMultiSelectTokens(w.ToString())
                        };
                        object? dateRaw = args[1].Eval(ctx);
                        decimal fallback = ToDecimal(args[2].Eval(ctx));
                        if (!TryToDate(dateRaw, out DateOnly date)) return fallback;
                        return WageProvider.ResolveRate(workers, date, fallback);
                    }

                default:
                    throw new InvalidOperationException($"Unknown function '{name}'");
            }
        }

        private void Require(int min, int max = -1)
        {
            if (max < 0) max = min;
            if (args.Count < min || args.Count > max)
                throw new InvalidOperationException($"{name} expects {min}{(max != min ? $"..{max}" : "")} argument(s), got {args.Count}");
        }
    }

    private sealed class Parser
    {
        private readonly string _s;
        private int _i;

        public Parser(string s) { _s = s; }

        public ExprNode ParseAll()
        {
            var node = ParseComparison();
            SkipWs();
            if (_i < _s.Length) throw new FormatException($"Unexpected '{_s[_i]}' at {_i}");
            return node;
        }

        private void SkipWs() { while (_i < _s.Length && char.IsWhiteSpace(_s[_i])) _i++; }

        private bool Match(string tok)
        {
            SkipWs();
            if (string.CompareOrdinal(_s, _i, tok, 0, tok.Length) != 0) return false;
            _i += tok.Length;
            return true;
        }

        private ExprNode ParseComparison()
        {
            var left = ParseAdditive();
            while (true)
            {
                string? op = null;
                foreach (var cand in new[] { ">=", "<=", "==", "!=", ">", "<", "=" })
                {
                    SkipWs();
                    if (string.CompareOrdinal(_s, _i, cand, 0, cand.Length) == 0) { op = cand; break; }
                }
                if (op == null) return left;
                _i += op.Length;
                left = new BinaryNode(op == "=" ? "==" : op, left, ParseAdditive());
            }
        }

        private ExprNode ParseAdditive()
        {
            var left = ParseTerm();
            while (true)
            {
                if (Match("+")) left = new BinaryNode("+", left, ParseTerm());
                else if (Match("-")) left = new BinaryNode("-", left, ParseTerm());
                else return left;
            }
        }

        private ExprNode ParseTerm()
        {
            var left = ParseUnary();
            while (true)
            {
                if (Match("*")) left = new BinaryNode("*", left, ParseUnary());
                else if (Match("/")) left = new BinaryNode("/", left, ParseUnary());
                else return left;
            }
        }

        private ExprNode ParseUnary()
        {
            if (Match("-")) return new UnaryNode('-', ParseUnary());
            if (Match("+")) return new UnaryNode('+', ParseUnary());
            return ParsePrimary();
        }

        private ExprNode ParsePrimary()
        {
            SkipWs();
            if (_i >= _s.Length) throw new FormatException("Unexpected end of expression");
            char c = _s[_i];

            if (c == '(')
            {
                _i++;
                var inner = ParseComparison();
                if (!Match(")")) throw new FormatException("Missing ')'");
                return inner;
            }

            if (c == '[')
            {
                int end = _s.IndexOf(']', _i + 1);
                if (end < 0) throw new FormatException("Missing ']'");
                string name = _s.Substring(_i + 1, end - _i - 1).Trim();
                _i = end + 1;
                return new ColumnNode(name);
            }

            if (c == '@')
            {
                int start = ++_i;
                while (_i < _s.Length && (char.IsLetterOrDigit(_s[_i]) || _s[_i] == '_')) _i++;
                return new ParamNode(_s.Substring(start, _i - start));
            }

            if (char.IsDigit(c) || c == '.')
            {
                int start = _i;
                while (_i < _s.Length && (char.IsDigit(_s[_i]) || _s[_i] == '.')) _i++;
                return new ConstNode(decimal.Parse(_s.Substring(start, _i - start), CultureInfo.InvariantCulture));
            }

            if (char.IsLetter(c) || c == '_')
            {
                int start = _i;
                while (_i < _s.Length && (char.IsLetterOrDigit(_s[_i]) || _s[_i] == '_')) _i++;
                string ident = _s.Substring(start, _i - start);

                if (ident.Equals("null", StringComparison.OrdinalIgnoreCase)) return new ConstNode(null);
                if (ident.Equals("true", StringComparison.OrdinalIgnoreCase)) return new ConstNode(true);
                if (ident.Equals("false", StringComparison.OrdinalIgnoreCase)) return new ConstNode(false);

                if (!Match("(")) throw new FormatException($"Unknown identifier '{ident}'");
                var args = new List<ExprNode>();
                if (!Match(")"))
                {
                    do { args.Add(ParseComparison()); } while (Match(","));
                    if (!Match(")")) throw new FormatException("Missing ')' after arguments");
                }
                return new CallNode(ident.ToUpperInvariant(), args);
            }

            throw new FormatException($"Unexpected '{c}' at {_i}");
        }
    }
}
