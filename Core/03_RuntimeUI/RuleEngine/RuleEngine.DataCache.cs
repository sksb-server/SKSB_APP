namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using SKSB_App.Core.WorkbookSave.SheetMapper;

public static class RuleDataCache
{
    private static readonly string[] RecognizedDatePatterns =
{
    "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "MM/dd/yyyy", "M/d/yyyy"
};

    /// <summary>
    /// Extracts clean string values from a source expression (either "TableName[ColName]" or a Defined Name like "VehicleTypeList").
    /// Zero xlApp.Evaluate COM calls.
    /// </summary>
    public static List<string> GetSourceItems(dynamic wb, string sourceExpr, bool preserveDuplicates = false)
    {
        var results = new List<string>();
        if (wb == null || string.IsNullOrWhiteSpace(sourceExpr))
            return results;

        string clean = sourceExpr.Trim().TrimStart('=');

        // 1. COMPOUND / CONCAT EXPRESSION: e.g. "AllVehiclesTable[Brand] + ' ' + AllVehiclesTable[Model]"
        if (clean.Contains("[") && (clean.Contains("+") || clean.Contains("&")))
        {
            var compoundList = EvaluateCompoundExpression(wb, clean);
            return FilterList(compoundList, preserveDuplicates);
        }

        // 2. SINGLE Table[Column] Reference
        if (!clean.Contains("(") && clean.EndsWith("]") && clean.IndexOf('[') > 0)
        {
            int openBracket = clean.IndexOf('[');
            int closeBracket = clean.LastIndexOf(']');
            if (openBracket > 0 && closeBracket > openBracket)
            {
                string tableName = clean.Substring(0, openBracket).Trim();
                string colName = clean.Substring(openBracket + 1, closeBracket - openBracket - 1).Trim('[', ']', ' ');
                var colValues = ReadTableColumnValues(wb, tableName, colName);
                return FilterList(colValues, preserveDuplicates);
            }
        }

        // 3. DIRECT DEFINED NAME LOOKUP (Zero Evaluate / Zero Scratch-Cell Mutation!)
        if (!clean.Contains("(") && !clean.Contains("["))
        {
            try
            {
                dynamic? nameObj = null;
                try { nameObj = wb.Names.Item(clean); } catch { }
                if (nameObj != null)
                {
                    dynamic? targetRange = null;
                    try { targetRange = nameObj.RefersToRange; } catch { }
                    if (targetRange != null)
                    {
                        object? rawVal = targetRange.Value2;
                        var rawList = ExtractRawArrayToList(rawVal);
                        var filtered = FilterList(rawList, preserveDuplicates);
                        if (filtered.Count > 0) return filtered;
                    }
                }
            }
            catch { }
        }

        // 4. Try as C# Enum Name directly in memory (Fastest, zero COM)
        if (clean.EndsWith("List", StringComparison.OrdinalIgnoreCase))
        {
            string enumCandidate = clean.Substring(0, clean.Length - 4);
            var enumVals = TryGetEnumNames(enumCandidate);
            if (enumVals.Count > 0) return enumVals;
        }

        return results;
    }

    /// <summary>
    /// Performs an in-memory XLOOKUP between LookupSource and ResultSource with type-tolerant matching.
    /// </summary>
    public static string? LookupValue(dynamic wb, string searchKey, string lookupSource, string resultSource)
    {
        if (wb == null || string.IsNullOrWhiteSpace(searchKey)) return null;

        var keys = GetSourceItems(wb, lookupSource, preserveDuplicates: true);
        var vals = GetSourceItems(wb, resultSource, preserveDuplicates: true);

        int count = Math.Min(keys.Count, vals.Count);
        if (count == 0) return null;

        string cleanSearch = searchKey.Trim();

        for (int i = 0; i < count; i++)
        {
            if (ValuesMatch(keys[i], cleanSearch))
            {
                return vals[i];
            }
        }

        return null;
    }

    /// <summary>
    /// Reads all raw values in a specific table column directly from its ListObject DataBodyRange.
    /// </summary>
    public static List<string> ReadTableColumnValues(dynamic wb, string tableName, string colHeader)
    {
        var list = new List<string>();
        try
        {
            dynamic? targetTable = FindTableInWorkbook(wb, tableName);
            if (targetTable == null) return list;

            dynamic? bodyRange = targetTable.DataBodyRange;
            if (bodyRange == null) return list;

            int colIndex = FindRelativeColumnIndex(targetTable, colHeader);
            if (colIndex < 1) return list;

            dynamic colRange = targetTable.ListColumns[colIndex].DataBodyRange;
            if (colRange == null) return list;

            object? rawVal = colRange.Value2;
            return ExtractRawArrayToList(rawVal);
        }
        catch
        {
            return list;
        }
    }

    public static dynamic? FindTableInWorkbook(dynamic wb, string tableName)
    {
        if (wb == null || string.IsNullOrWhiteSpace(tableName)) return null;
        string normTarget = RuleRegistry.NormalizeIdentifier(tableName);

        foreach (dynamic ws in wb.Sheets)
        {
            try
            {
                dynamic direct = ws.ListObjects[tableName];
                if (direct != null) return direct;
            }
            catch { }

            try
            {
                foreach (dynamic tbl in ws.ListObjects)
                {
                    string tName = tbl.Name?.ToString() ?? string.Empty;
                    if (RuleRegistry.NormalizeIdentifier(tName).Equals(normTarget, StringComparison.OrdinalIgnoreCase))
                        return tbl;
                }
            }
            catch { }
        }
        return null;
    }

    public static int FindRelativeColumnIndex(dynamic table, string targetHeader)
    {
        if (table == null || string.IsNullOrWhiteSpace(targetHeader)) return -1;
        string normTarget = RuleRegistry.NormalizeIdentifier(targetHeader);

        try
        {
            int colCount = (int)table.ListColumns.Count;
            for (int c = 1; c <= colCount; c++)
            {
                string cName = table.ListColumns[c].Name?.ToString() ?? string.Empty;
                string normCol = RuleRegistry.NormalizeIdentifier(cName);
                if (normCol.Equals(normTarget, StringComparison.OrdinalIgnoreCase) ||
                    normCol.TrimEnd('s').Equals(normTarget.TrimEnd('s'), StringComparison.OrdinalIgnoreCase))
                {
                    return c;
                }
            }
        }
        catch { }

        return -1;
    }

    public static bool ValuesMatch(string? a, string? b)
    {
        if (string.IsNullOrWhiteSpace(a) && string.IsNullOrWhiteSpace(b)) return true;
        if (string.IsNullOrWhiteSpace(a) || string.IsNullOrWhiteSpace(b)) return false;

        string sa = a.Trim();
        string sb = b.Trim();

        if (string.Equals(sa, sb, StringComparison.OrdinalIgnoreCase))
            return true;

        // Enum underscore/space tolerance ("Regular Service" == "Regular_Service")
        if (string.Equals(sa.Replace(" ", "_"), sb.Replace(" ", "_"), StringComparison.OrdinalIgnoreCase))
            return true;

        // Numeric tolerance ("100" == "100.0")
        if (double.TryParse(sa, NumberStyles.Any, CultureInfo.InvariantCulture, out double da) &&
            double.TryParse(sb, NumberStyles.Any, CultureInfo.InvariantCulture, out double db))
        {
            if (Math.Abs(da - db) < 0.0001) return true;
        }

        // Date tolerance
        if (TryParseDate(sa, out DateOnly dta) && TryParseDate(sb, out DateOnly dtb))
        {
            return dta == dtb;
        }

        return false;
    }

    public static bool TryParseDate(object? rawVal, out DateOnly result)
    {
        result = default;
        if (rawVal == null || SheetMapperHelper.IsComError(rawVal)) return false;

        if (rawVal is DateOnly dOnly) { result = dOnly; return true; }
        if (rawVal is DateTime dt) { result = DateOnly.FromDateTime(dt); return true; }
        if (rawVal is double oa && oa >= 15000 && oa <= 80000)
        {
            result = DateOnly.FromDateTime(DateTime.FromOADate(oa));
            return true;
        }

        string s = rawVal.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(s) || s == "—") return false;

        if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out double parsedOa) &&
            parsedOa >= 15000 && parsedOa <= 80000)
        {
            result = DateOnly.FromDateTime(DateTime.FromOADate(parsedOa));
            return true;
        }

        if (DateTime.TryParseExact(s, RecognizedDatePatterns, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime exactDt))
        {
            result = DateOnly.FromDateTime(exactDt);
            return true;
        }

        if (DateTime.TryParse(s, new CultureInfo("en-GB"), DateTimeStyles.None, out DateTime gbDt))
        {
            result = DateOnly.FromDateTime(gbDt);
            return true;
        }

        return false;
    }

    private static List<string> ExtractRawArrayToList(object? raw)
    {
        var list = new List<string>();
        if (raw == null) return list;

        if (raw is object[,] matrix)
        {
            int rStart = matrix.GetLowerBound(0);
            int rEnd = matrix.GetUpperBound(0);
            int cStart = matrix.GetLowerBound(1);
            int cEnd = matrix.GetUpperBound(1);

            for (int r = rStart; r <= rEnd; r++)
            {
                for (int c = cStart; c <= cEnd; c++)
                {
                    list.Add(NormalizeCellToString(matrix[r, c]));
                }
            }
        }
        else if (raw is object[] vec)
        {
            foreach (var item in vec)
                list.Add(NormalizeCellToString(item));
        }
        else
        {
            list.Add(NormalizeCellToString(raw));
        }

        return list;
    }

    private static string NormalizeCellToString(object? val)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        if (val is double d)
        {
            if (Math.Abs(d - Math.Round(d)) < 0.00001 && d < 1e15)
                return ((long)Math.Round(d)).ToString(CultureInfo.InvariantCulture);
            return d.ToString(CultureInfo.InvariantCulture);
        }
        string s = val.ToString()?.Trim() ?? string.Empty;
        if (s.StartsWith("#") || s.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
            return string.Empty;
        return s;
    }

    private static List<string> FilterList(List<string> raw, bool preserveDuplicates)
    {
        if (preserveDuplicates)
            return raw;

        return raw
            .Where(s => !string.IsNullOrWhiteSpace(s) && !s.StartsWith("[") && s != "0" && s != "-" && s != "\u2014")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static List<string> TryGetEnumNames(string enumName)
    {
        var coreAsm = typeof(RuleDataCache).Assembly;
        var enumType = coreAsm.GetTypes().FirstOrDefault(t =>
            t.IsEnum &&
            (t.Name.Equals(enumName, StringComparison.OrdinalIgnoreCase) ||
             t.Name.TrimEnd('s').Equals(enumName.TrimEnd('s'), StringComparison.OrdinalIgnoreCase)));

        if (enumType == null) return new List<string>();

        return Enum.GetNames(enumType)
            .Where(n => !n.Equals("None", StringComparison.OrdinalIgnoreCase))
            .ToList();
    }

    private static List<string> EvaluateCompoundExpression(dynamic wb, string expr)
    {
        // Split on '+' or '&'
        var parts = expr.Split(new[] { '+', '&' }, StringSplitOptions.TrimEntries);
        var tokenColumns = new List<(bool IsColumn, string Table, string Col, string Literal)>();
        int maxRows = 0;

        foreach (var part in parts)
        {
            string token = part.Trim();
            if (token.StartsWith("[") || token.Contains("["))
            {
                int ob = token.IndexOf('[');
                int cb = token.LastIndexOf(']');
                string tName = ob > 0 ? token.Substring(0, ob).Trim() : string.Empty;
                string cName = token.Substring(ob + 1, cb - ob - 1).Trim('[', ']', ' ');

                var vals = ReadTableColumnValues(wb, tName, cName);
                if (vals.Count > maxRows) maxRows = vals.Count;
                tokenColumns.Add((true, tName, cName, string.Empty));
            }
            else
            {
                // Literal string token like ' ' or "-"
                string literal = token.Trim('\'', '\"');
                tokenColumns.Add((false, string.Empty, string.Empty, literal));
            }
        }

        if (maxRows == 0) return new List<string>();

        // Cache each distinct column's rows in memory
        var columnData = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
        foreach (var (isCol, tbl, col, _) in tokenColumns)
        {
            if (isCol)
            {
                string key = $"{tbl}[{col}]";
                if (!columnData.ContainsKey(key))
                {
                    columnData[key] = ReadTableColumnValues(wb, tbl, col);
                }
            }
        }

        var combined = new List<string>(maxRows);
        for (int r = 0; r < maxRows; r++)
        {
            var sb = new System.Text.StringBuilder();
            foreach (var (isCol, tbl, col, literal) in tokenColumns)
            {
                if (isCol)
                {
                    var list = columnData[$"{tbl}[{col}]"];
                    if (r < list.Count && !string.IsNullOrWhiteSpace(list[r]))
                    {
                        sb.Append(list[r]);
                    }
                }
                else
                {
                    sb.Append(literal);
                }
            }

            string resultStr = sb.ToString().Trim();
            if (!string.IsNullOrWhiteSpace(resultStr))
            {
                combined.Add(resultStr);
            }
        }

        return combined;
    }
}