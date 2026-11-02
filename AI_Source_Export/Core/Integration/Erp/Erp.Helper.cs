namespace SKSB_App.Core.Integration.Erp;

using System;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;

public static class ErpHelper
{
    private static readonly string[] DateFormats =
    {
        "dd/MM/yyyy",
        "d/M/yyyy",
        "dd-MM-yyyy",
        "d-M-yyyy",
        "yyyy-MM-dd",
        "yyyy/MM/dd"
    };

    /// <summary>
    /// Normalizes header strings by stripping whitespace, underscores, dots, and hyphens to allow tolerant lookups.
    /// Internal to prevent Excel-DNA UDF registration collisions.
    /// </summary>
    internal static string NormalizeHeader(string header)
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
            DateOnly d => d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
            DateTime dt => dt.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture),
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

    public static DateOnly? ParseDateOnly(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        string cleaned = value.Replace("#", string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(cleaned)) return null;

        if (double.TryParse(cleaned, NumberStyles.Any, CultureInfo.InvariantCulture, out double oaDate) && oaDate > 30000 && oaDate < 80000)
        {
            return DateOnly.FromDateTime(DateTime.FromOADate(oaDate));
        }

        if (DateOnly.TryParseExact(cleaned, DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            return date;

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