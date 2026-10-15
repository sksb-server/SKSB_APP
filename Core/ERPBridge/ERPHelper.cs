namespace SKSB_App.Core.ERPBridge;

using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

public static class ERPHelper
{
    private static readonly string[] DateFormats =
    {
        "yyyy-MM-dd",
        "dd/MM/yyyy",
        "d/M/yyyy",
        "yyyy/MM/dd",
        "dd-MM-yyyy",
        "d-M-yyyy",
        "M/d/yyyy",
        "MM/dd/yyyy"
    };

    public static string NormalizeHeader(string header)
    {
        if (string.IsNullOrWhiteSpace(header)) return string.Empty;
        return Regex.Replace(header.Trim(), @"[\s_\-\.]+", string.Empty).ToUpperInvariant();
    }

    public static string[] SplitDelimitedLine(string line, char delimiter = ',')
    {
        if (string.IsNullOrEmpty(line)) return Array.Empty<string>();

        var pattern = $"{delimiter}(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)";
        return Regex.Split(line, pattern)
                    .Select(col => col.Trim().Trim('"').Replace("\"\"", "\""))
                    .ToArray();
    }

    public static string EscapeDelimitedCell(object? value, char delimiter = ',')
    {
        if (value is null) return string.Empty;

        string stringVal = value switch
        {
            DateOnly d => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            DateTime dt => dt.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            decimal dec => dec.ToString("F2", CultureInfo.InvariantCulture),
            double dbl => dbl.ToString("F2", CultureInfo.InvariantCulture),
            float flt => flt.ToString("F2", CultureInfo.InvariantCulture),
            _ => value.ToString() ?? string.Empty
        };

        if (stringVal.Contains(delimiter) || stringVal.Contains('"') || stringVal.Contains('\n') || stringVal.Contains('\r'))
        {
            return $"\"{stringVal.Replace("\"", "\"\"")}\"";
        }

        return stringVal;
    }

    /// <summary>
    /// Intelligently parses dates from string text, SQL Payroll hash masks, or Excel OADate serial numbers.
    /// </summary>
    public static DateOnly? ParseDateOnly(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        // 1. Strip hash symbols '#' and trim spaces
        string cleaned = value.Replace("#", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleaned)) return null;

        // 2. Check if the value is an Excel numeric OADate serial (e.g. "43160" or "43160.5")
        if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate) && oaDate > 30000 && oaDate < 80000)
        {
            return DateOnly.FromDateTime(DateTime.FromOADate(oaDate));
        }

        // 3. Match explicit standard formats
        if (DateOnly.TryParseExact(cleaned, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            return date;

        // 4. Standard fallback parse
        if (DateOnly.TryParse(cleaned, CultureInfo.InvariantCulture, DateTimeStyles.None, out date))
            return date;

        if (DateTime.TryParse(cleaned, out var dt))
            return DateOnly.FromDateTime(dt);

        return null;
    }

    public static decimal ParseDecimal(string value, decimal defaultValue = 0.00m)
    {
        if (string.IsNullOrWhiteSpace(value)) return defaultValue;
        return decimal.TryParse(value.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal res)
            ? res
            : defaultValue;
    }

    public static long ParseLong(string value, long defaultValue = 0)
    {
        if (string.IsNullOrWhiteSpace(value)) return defaultValue;
        string clean = Regex.Replace(value.Trim(), @"[^\d\-]", string.Empty);
        return long.TryParse(clean, NumberStyles.Any, CultureInfo.InvariantCulture, out long res)
            ? res
            : defaultValue;
    }
}