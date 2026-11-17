namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Reflection;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.SheetMapper;

/// <summary>Resolved display format of one column (explicit [ColumnFormat] or inferred).</summary>
public readonly record struct EffectiveColumnFormat(
    ColumnFormatKind Kind,
    ZeroBehavior Zero,
    int DecimalPlaces,
    string? CustomUnit,
    string? ExcelNumberFormat);

public static partial class RuleRegistry
{
    private const string Dash = "—";

    /// <summary>Road vehicles: whole kilometres. Zero shows "0 km"; text (the em-dash for null) falls through to the 4th section.</summary>
    public const string OdometerKmFormat = "#,##0 \"km\";-#,##0 \"km\";0 \"km\";@";

    /// <summary>Machinery: engine hours with one decimal.</summary>
    public const string OdometerHoursFormat = "#,##0.0 \"hrs\";-#,##0.0 \"hrs\";0.0 \"hrs\";@";

    // =========================================================================
    // 1. FORMAT RESOLUTION
    // =========================================================================

    /// <summary>
    /// Uses the explicit [ColumnFormat] when present; otherwise infers a consistent default from the
    /// property type and header text so untagged columns still get units and a dash for zero.
    /// </summary>
    internal static EffectiveColumnFormat ResolveEffectiveFormat(
        PropertyInfo prop, string headerName, ColumnFormatAttribute? explicitAttr, ColumnDefaultAttribute? defaultAttr)
    {
        ColumnFormatKind kind;
        ZeroBehavior zero;
        int decimals = -1;
        string? unit = null;

        if (explicitAttr != null)
        {
            kind = explicitAttr.Kind;
            zero = explicitAttr.Zero;
            decimals = explicitAttr.DecimalPlaces;
            unit = explicitAttr.CustomUnit;
        }
        else
        {
            kind = InferKind(prop, headerName, defaultAttr);
            zero = kind is ColumnFormatKind.Text or ColumnFormatKind.General or ColumnFormatKind.None
                ? ZeroBehavior.Blank
                : ZeroBehavior.Dash;
        }

        return new EffectiveColumnFormat(kind, zero, decimals, unit, BuildExcelNumberFormatString(kind, decimals, zero, unit));
    }

    private static ColumnFormatKind InferKind(PropertyInfo prop, string header, ColumnDefaultAttribute? defaultAttr)
    {
        Type t = Nullable.GetUnderlyingType(prop.PropertyType) ?? prop.PropertyType;
        string h = header ?? string.Empty;
        bool Has(string s) => h.Contains(s, StringComparison.OrdinalIgnoreCase);

        if (t == typeof(DateOnly) || t == typeof(DateTime)) return ColumnFormatKind.Date;
        if (t == typeof(TimeOnly)) return ColumnFormatKind.Time;
        if (t == typeof(string) &&
            (prop.Name is "Time" or "TimeStart" or "TimeEnd")) return ColumnFormatKind.Time;

        bool isInt = t == typeof(int) || t == typeof(long);
        bool isDec = t == typeof(decimal) || t == typeof(double) || t == typeof(float);
        if (!isInt && !isDec) return ColumnFormatKind.Text; // string / enum / List<...> / bool

        if (isInt && (defaultAttr?.Kind == DefaultKind.AutoIncrement || h.EndsWith(" ID", StringComparison.OrdinalIgnoreCase)))
            return ColumnFormatKind.Id;

        if (prop.Name is "CurrentOdometer" or "OdometerBefore" or "OdometerIn" or "OdometerOut" or "OdometerDifference")
            return ColumnFormatKind.OdometerDynamic;
        if (Has("Odometer")) return ColumnFormatKind.Kilometers;
        if (Has("Years Old")) return ColumnFormatKind.Years;
        if (h.Equals("Days Worked", StringComparison.OrdinalIgnoreCase)) return ColumnFormatKind.Days;
        if (Has("Hours")) return ColumnFormatKind.Hours;
        if (Has("Shift")) return ColumnFormatKind.Shifts;
        if (Has("Cost") || Has("Price") || Has("Return") || Has("Margin") || Has("Amount") ||
            Has("Claim") || Has("Salary") || Has("Earned"))
            return ColumnFormatKind.Currency;

        return isDec ? ColumnFormatKind.Decimal : ColumnFormatKind.Integer;
    }

    // =========================================================================
    // 2. EXCEL NUMBER FORMAT STRINGS
    // =========================================================================

    /// <summary>
    /// Builds the 4-section "positive;negative;zero;text" Excel format. A numeric 0 always displays as 0 with the
    /// column's decimals and unit; a null is written as the text em-dash and shows through the 4th section (@),
    /// which Excel colour scales / icon sets ignore. ZeroBehavior.Blank leaves the zero section empty.
    /// </summary>
    internal static string? BuildExcelNumberFormatString(ColumnFormatKind kind, int decimalPlaces, ZeroBehavior zero, string? customUnit)
    {
        bool blankZero = zero == ZeroBehavior.Blank;

        static string Digits(int dp, int fallback)
        {
            int n = dp >= 0 ? dp : fallback;
            return n <= 0 ? "#,##0" : "#,##0." + new string('0', n);
        }
        static string ZeroDigits(int dp, int fallback)
        {
            int n = dp >= 0 ? dp : fallback;
            return n <= 0 ? "0" : "0." + new string('0', n);
        }

        string Four(string pos, string zeroPos, string? neg = null) =>
            $"{pos};{neg ?? "-" + pos};{(blankZero ? string.Empty : zeroPos)};@";
        string Unit(string fallbackUnit) => " \"" + (string.IsNullOrWhiteSpace(customUnit) ? fallbackUnit : customUnit) + "\"";

        switch (kind)
        {
            case ColumnFormatKind.Currency:
                return Four("\"RM \"" + Digits(decimalPlaces, 2), "\"RM \"" + ZeroDigits(decimalPlaces, 2));
            case ColumnFormatKind.PricePerLiter:
                return Four("\"RM \"" + Digits(decimalPlaces, 2) + "\"/L\"", "\"RM \"" + ZeroDigits(decimalPlaces, 2) + "\"/L\"");
            case ColumnFormatKind.Liters:
                return Four(Digits(decimalPlaces, 2) + Unit("L"), ZeroDigits(decimalPlaces, 2) + Unit("L"));
            case ColumnFormatKind.Kilometers:
                return Four(Digits(decimalPlaces, 1) + Unit("km"), ZeroDigits(decimalPlaces, 1) + Unit("km"));
            case ColumnFormatKind.Hours:
                return Four(Digits(decimalPlaces, 1) + Unit("hrs"), ZeroDigits(decimalPlaces, 1) + Unit("hrs"));
            case ColumnFormatKind.Shifts:
                return Four((decimalPlaces >= 0 ? Digits(decimalPlaces, 1) : "#,##0.#") + Unit("shifts"), "0" + Unit("shifts"));
            case ColumnFormatKind.Years:
                return Four(Digits(decimalPlaces, 0) + Unit("yrs"), ZeroDigits(decimalPlaces, 0) + Unit("yrs"));
            case ColumnFormatKind.Days:
                return Four(Digits(decimalPlaces, 0) + Unit("days"), ZeroDigits(decimalPlaces, 0) + Unit("days"));
            case ColumnFormatKind.OdometerDynamic:
                return blankZero ? "#,##0 \"km\";-#,##0 \"km\";;@" : OdometerKmFormat;
            case ColumnFormatKind.Integer:
                {
                    string u = string.IsNullOrWhiteSpace(customUnit) ? string.Empty : Unit(customUnit!);
                    return Four(Digits(decimalPlaces, 0) + u, ZeroDigits(decimalPlaces, 0) + u);
                }
            case ColumnFormatKind.Decimal:
                {
                    string u = string.IsNullOrWhiteSpace(customUnit) ? string.Empty : Unit(customUnit!);
                    return Four(Digits(decimalPlaces, 2) + u, ZeroDigits(decimalPlaces, 2) + u);
                }
            case ColumnFormatKind.Id: return blankZero ? "0;-0;;@" : "0;-0;0;@";
            case ColumnFormatKind.Date: return Four("dd/mm/yyyy", "dd/mm/yyyy", "dd/mm/yyyy");
            case ColumnFormatKind.Time: return Four("hh:mm", "hh:mm", "hh:mm");
            case ColumnFormatKind.Text: return "@";
            default: return null;
        }
    }

    /// <summary>Formats worth writing to the sheet; Text / General are left to the template (never force "@" on a column).</summary>
    internal static bool ShouldApplyFormat(PropertyRuleDescriptor? rule) =>
        rule != null && !string.IsNullOrWhiteSpace(rule.ExcelNumberFormat) &&
        rule.EffectiveFormatKind is not (ColumnFormatKind.None or ColumnFormatKind.General or ColumnFormatKind.Text);

    /// <summary>True when an explicit non-text [ColumnFormat] (e.g. Id) should override a stale "@" archetype format.</summary>
    internal static bool OverridesTextFormat(PropertyRuleDescriptor rule) =>
        ShouldApplyFormat(rule) && rule.PropertyType != typeof(string) &&
        !string.Equals(rule.Property?.Name, "NRIC", StringComparison.OrdinalIgnoreCase);

    /// <summary>
    /// Applies the rule's number format to a table column body. Text ("@") template columns are kept as text, and
    /// per-row odometer units are left to populate / row healing (hrs for machinery, km for road vehicles).
    /// </summary>
    internal static void ApplyColumnNumberFormat(dynamic table, int columnIndex, dynamic archCell, PropertyRuleDescriptor? rule)
    {
        if (!ShouldApplyFormat(rule) || rule!.EffectiveFormatKind == ColumnFormatKind.OdometerDynamic) return;
        try
        {
            string archFmt = archCell.NumberFormat?.ToString() ?? string.Empty;
            if (string.Equals(rule!.Property?.Name, "NRIC", StringComparison.OrdinalIgnoreCase)) return;
            // Id columns always get their declared format, overriding stale template "@", "#" or currency formats.
            bool forceId = rule!.EffectiveFormatKind == ColumnFormatKind.Id;
            if (archFmt == "@" && !forceId && !OverridesTextFormat(rule!)) return;

            dynamic? body = table.ListColumns[columnIndex].DataBodyRange;
            if (body != null) body.NumberFormat = rule.ExcelNumberFormat;
        }
        catch { }
    }

    // =========================================================================
    // 3. UNIFIED ZERO / EMPTY DETECTION
    // =========================================================================

    /// <summary>
    /// True for blank cells, COM errors, dashes, zero numbers, "None" / "False", and sentinel dates
    /// (Year &lt;= 1900, including OADate 0 which Excel shows as 00/01/1900).
    /// </summary>
    public static bool IsZeroOrEmptyValue(object? rawVal, PropertyRuleDescriptor? rule = null)
    {
        if (rawVal == null || SheetMapperHelper.IsComError(rawVal)) return true;

        switch (rawVal)
        {
            case DateOnly d: return d.Year <= 1900;
            case DateTime dt: return dt.Year <= 1900;
            case double or decimal or float or int or long or short:
                {
                    double n = Convert.ToDouble(rawVal, CultureInfo.InvariantCulture);
                    if (Math.Abs(n) < 0.00001) return true;
                    // Any small OADate in a date column is a sentinel (0 => 00/01/1900)
                    return rule?.EffectiveFormatKind == ColumnFormatKind.Date && n < 15000;
                }
            case bool b: return !b;
        }

        string s = rawVal.ToString()?.Trim() ?? string.Empty;
        if (s.Length == 0) return true;
        if (s is "-" or Dash or "0" or "0.0" or "0.00" or "00/01/1900" or "01/01/0001" or "1/1/0001" or "0/1/1900") return true;
        return s.Equals("None", StringComparison.OrdinalIgnoreCase) || s.Equals("False", StringComparison.OrdinalIgnoreCase);
    }
}
