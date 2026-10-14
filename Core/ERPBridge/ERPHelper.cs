namespace SKSB_App.Core.ERPBridge;

using System;
using System.Globalization;
using System.Linq;
using System.Text;
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
        "d-M-yyyy"
    };

    /// <summary>
    /// Normalizes header strings by stripping whitespace, underscores, dots, and hyphens to allow tolerant lookups.
    /// Example: "Emp_Code", "emp code", or "EmpCode" all normalize to "EMPCODE".
    /// </summary>
    public static string NormalizeHeader(string header)
    {
        if (string.IsNullOrWhiteSpace(header)) return string.Empty;
        return Regex.Replace(header.Trim(), @"[\s_\-\.]+", string.Empty).ToUpperInvariant();
    }

    /// <summary>
    /// Splits a delimited string line while respecting quoted values containing delimiters or escaped quotes.
    /// </summary>
    public static string[] SplitDelimitedLine(string line, char delimiter = ',')
    {
        if (string.IsNullOrEmpty(line)) return Array.Empty<string>();

        var pattern = $"{delimiter}(?=(?:[^\"]*\"[^\"]*\")*[^\"]*$)";
        return Regex.Split(line, pattern)
                    .Select(col => col.Trim().Trim('"').Replace("\"\"", "\""))
                    .ToArray();
    }

    /// <summary>
    /// Escapes an individual cell value for delimited outputs (e.g., CSV).
    /// Encloses values in quotes if they contain delimiters, quotes, or newline characters.
    /// </summary>
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

    public static DateOnly? ParseDateOnly(string value)
    {
        if (string.IsNullOrWhiteSpace(value)) return null;

        if (DateOnly.TryParseExact(value.Trim(), DateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out var date))
            return date;

        return DateOnly.TryParse(value.Trim(), out date) ? date : null;
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
        return long.TryParse(value.Trim(), NumberStyles.Any, CultureInfo.InvariantCulture, out long res)
            ? res
            : defaultValue;
    }
}