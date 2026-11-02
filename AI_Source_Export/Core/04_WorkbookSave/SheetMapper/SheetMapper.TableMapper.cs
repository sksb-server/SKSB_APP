namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types.Attributes;

public static class TableMapper
{
    private static readonly string[] RecognizedDateFormats =
    {
        "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyy/MM/dd"
    };

    public static dynamic? ResolveTable(dynamic worksheet, params string[] candidateNames)
    {
        if (worksheet == null || candidateNames == null || candidateNames.Length == 0) return null;

        foreach (var name in candidateNames)
        {
            try
            {
                dynamic t = worksheet.ListObjects[name];
                if (t != null) return t;
            }
            catch { }
        }

        try
        {
            foreach (dynamic t in worksheet.ListObjects)
            {
                string tName = NormalizeHeader(t.Name?.ToString() ?? string.Empty);
                foreach (var name in candidateNames)
                {
                    if (tName.Equals(NormalizeHeader(name), StringComparison.OrdinalIgnoreCase))
                        return t;
                }
            }
        }
        catch { }

        return null;
    }

    public static List<T> ExtractFromTable<T>(dynamic worksheet, string tableName) where T : new() =>
        ExtractFromTable<T>(worksheet, new[] { tableName });

    public static List<T> ExtractFromTable<T>(dynamic worksheet, params string[] candidateTableNames) where T : new()
    {
        var records = new List<T>();
        dynamic? table = ResolveTable(worksheet, candidateTableNames);
        if (table == null) return records;

        // Ensure all active rows are healed (auto-increments, partner syncs, defaults, state guards) before reading
        RuleEvaluator.EvaluateAndHealTable(worksheet, table);

        dynamic dataBody = table.DataBodyRange;
        if (dataBody == null) return records;

        int rowCount = 0;
        try { rowCount = (int)dataBody.Rows.Count; } catch { }
        int colCount = (int)table.ListColumns.Count;
        if (rowCount == 0 || colCount == 0) return records;

        object[,] dataGrid = GetRange2DValues(dataBody, rowCount, colCount);
        object[,] headerGrid = GetRange2DValues(table.HeaderRowRange, 1, colCount);

        var headerMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        for (int c = 1; c <= colCount; c++)
        {
            string rawHeader = headerGrid[1, c]?.ToString() ?? string.Empty;
            string normalized = NormalizeHeader(rawHeader);
            if (!string.IsNullOrEmpty(normalized) && !headerMap.ContainsKey(normalized))
            {
                headerMap[normalized] = c;
            }
        }

        if (typeof(T) == typeof(string))
        {
            int targetCol = headerMap.Values.FirstOrDefault(1);
            for (int r = 1; r <= rowCount; r++)
            {
                object? rawVal = dataGrid[r, targetCol];
                if (rawVal == null || SheetMapperHelper.IsComError(rawVal)) continue;

                string s = rawVal.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(s) && !s.Equals("Total", StringComparison.OrdinalIgnoreCase))
                {
                    records.Add((T)(object)s);
                }
            }
            return records;
        }

        var propBindings = ResolvePropertyBindings<T>(headerMap);
        if (propBindings.Count == 0) return records;

        for (int r = 1; r <= rowCount; r++)
        {
            if (IsRowEmptyOrTotal(dataGrid, r, colCount))
                continue;

            T item = new();
            bool hasValidData = false;

            foreach (var binding in propBindings)
            {
                object? rawVal = dataGrid[r, binding.ColumnIndex];
                if (rawVal == null || SheetMapperHelper.IsComError(rawVal)) continue;

                object? converted = ConvertValue(rawVal, binding.Property.PropertyType);
                if (converted != null)
                {
                    object targetObj = binding.TargetInstance(item);
                    if (targetObj != null)
                    {
                        binding.Property.SetValue(targetObj, converted);
                        hasValidData = true;
                    }
                }
            }

            var plateProp = typeof(T).GetProperty("VehiclePlate");
            if (plateProp != null && string.IsNullOrWhiteSpace(plateProp.GetValue(item)?.ToString()))
            {
                continue;
            }

            var staffIdProp = typeof(T).GetProperty("StaffID");
            if (typeof(T).Name.StartsWith("Staff", StringComparison.OrdinalIgnoreCase) &&
                staffIdProp != null &&
                (Convert.ToInt64(staffIdProp.GetValue(item) ?? 0) <= 0))
            {
                continue;
            }

            if (hasValidData)
            {
                records.Add(item);
            }
        }

        return records;
    }

    public static void PopulateTableSafe<T>(dynamic worksheet, string tableName, IList<T> records) =>
        PopulateTableSafe<T>(worksheet, records, tableName);

    public static void PopulateTableSafe<T>(dynamic worksheet, IList<T> records, params string[] candidateTableNames)
    {
        if (worksheet == null || records == null) return;
        dynamic? table = ResolveTable(worksheet, candidateTableNames);
        if (table == null) return;

        string tName = table.Name?.ToString() ?? "Table";
        int rowCount = records.Count;

        using var _ = SKSB_App.Core.Diagnostics.AppLogger.Scope($"PopulateTableSafe<{typeof(T).Name}>({tName}, count={rowCount})");

        dynamic xlApp = worksheet.Application;
        bool origEvents = true;
        bool origUpdating = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

        try
        {
            if (rowCount == 0)
            {
                SheetMapperHelper.ResizeTable(table, 1);
                dynamic? body = table.DataBodyRange;
                if (body != null)
                {
                    int cols = (int)table.ListColumns.Count;
                    int hRow = (int)table.HeaderRowRange.Row;
                    int sCol = (int)table.HeaderRowRange.Column;
                    for (int c = 1; c <= cols; c++)
                    {
                        dynamic cell = worksheet.Cells[hRow + 1, sCol + c - 1];
                        string f = cell.Formula?.ToString() ?? string.Empty;
                        if (!f.StartsWith("="))
                        {
                            try { cell.Value2 = null; } catch { }
                        }
                    }
                }
                return;
            }

            int effectiveRows = Math.Max(1, rowCount);
            SheetMapperHelper.ResizeTable(table, effectiveRows);

            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;

            if (typeof(T) == typeof(string))
            {
                dynamic colRange = worksheet.Range[
                    worksheet.Cells[headerRow + 1, startCol],
                    worksheet.Cells[headerRow + effectiveRows, startCol]
                ];
                object[,] colData = new object[effectiveRows, 1];
                for (int r = 0; r < rowCount; r++)
                {
                    colData[r, 0] = records[r]?.ToString() ?? string.Empty;
                }
                try { colRange.NumberFormat = "@"; } catch { }
                colRange.Value2 = colData;
                return;
            }

            var columnMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int c = 1; c <= colCount; c++)
            {
                string name = table.ListColumns[c].Name?.ToString() ?? string.Empty;
                columnMap[NormalizeHeader(name)] = startCol + c - 1;
            }

            var propBindings = ResolvePropertyBindings<T>(columnMap);

            foreach (var binding in propBindings)
            {
                dynamic archCell = worksheet.Cells[headerRow + 1, binding.ColumnIndex];
                string archFormula = archCell.Formula?.ToString() ?? string.Empty;
                if (archFormula.StartsWith("="))
                {
                    continue;
                }

                object[,] colData = new object[effectiveRows, 1];
                for (int r = 0; r < rowCount; r++)
                {
                    object? target = binding.TargetInstance(records[r]);
                    object? val = target != null ? binding.Property.GetValue(target) : null;
                    colData[r, 0] = FormatForExcel(val);
                }

                dynamic colRange = worksheet.Range[
                    worksheet.Cells[headerRow + 1, binding.ColumnIndex],
                    worksheet.Cells[headerRow + effectiveRows, binding.ColumnIndex]
                ];

                if (binding.Property.PropertyType == typeof(string))
                {
                    try { colRange.NumberFormat = "@"; } catch { }
                }

                colRange.Value2 = colData;
            }

            try
            {
                int finalEndRow = headerRow + effectiveRows;
                int finalEndCol = startCol + colCount - 1;
                dynamic lockRange = worksheet.Range[worksheet.Cells[headerRow, startCol], worksheet.Cells[finalEndRow, finalEndCol]];
                table.Resize(lockRange);
                SheetMapperHelper.ClampAllTablesToActualData(worksheet.Parent);
            }
            catch { }
        }
        finally
        {
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    public static object[,] GetRange2DValues(dynamic range, int expectedRows, int expectedCols)
    {
        object? raw = null;
        try { raw = range.Value2; } catch { }

        if (raw == null) return new object[Math.Max(2, expectedRows + 1), Math.Max(2, expectedCols + 1)];

        if (raw is object[,] matrix)
        {
            return matrix;
        }

        object[,] scalarGrid = new object[2, 2];
        scalarGrid[1, 1] = raw;
        return scalarGrid;
    }

    public static object FormatForExcel(object? val)
    {
        if (val == null) return string.Empty;
        if (val is DateOnly d) return d.ToDateTime(TimeOnly.MinValue).ToOADate();
        if (val is DateTime dt) return dt.ToOADate();
        if (val is decimal dec)
        {
            if (dec == 0.00m) return string.Empty;
            return (double)dec;
        }
        if (val is Enum e)
        {
            if (string.Equals(e.ToString(), "None", StringComparison.OrdinalIgnoreCase))
                return string.Empty;
            return e.ToString();
        }
        if (val is IEnumerable<string> strList && val is not string)
        {
            return string.Join(", ", strList);
        }
        return val;
    }

    public static object? ConvertValue(object? rawValue, Type targetType)
    {
        if (rawValue == null || SheetMapperHelper.IsComError(rawValue)) return null;

        Type underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        string strVal = rawValue.ToString()?.Trim() ?? string.Empty;

        if (string.IsNullOrWhiteSpace(strVal) || strVal == " " || strVal == "-") return null;

        if (underlyingType == typeof(List<string>))
        {
            return strVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(s => s.Trim())
                         .Where(s => !string.IsNullOrEmpty(s))
                         .ToList();
        }

        if (underlyingType == typeof(string))
        {
            if (rawValue is double d && d >= 100000000)
            {
                return d.ToString("F0", CultureInfo.InvariantCulture);
            }
            return strVal;
        }

        if (underlyingType == typeof(DateOnly))
            return ParseToDateOnly(rawValue, strVal);

        if (underlyingType == typeof(DateTime))
        {
            var dOnly = ParseToDateOnly(rawValue, strVal);
            return dOnly.HasValue ? dOnly.Value.ToDateTime(TimeOnly.MinValue) : null;
        }

        if (underlyingType.IsEnum)
        {
            string cleanEnumName = strVal.Replace(" ", "_").Replace("-", "_");
            if (Enum.TryParse(underlyingType, cleanEnumName, true, out var enumResult))
                return enumResult;
            return Enum.ToObject(underlyingType, 0);
        }

        if (underlyingType == typeof(decimal))
            return decimal.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec) ? dec : 0.00m;

        if (underlyingType == typeof(int))
        {
            if (rawValue is double d) return (int)Math.Round(d);
            return int.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out int i) ? i : 0;
        }

        if (underlyingType == typeof(long))
        {
            if (rawValue is double d) return (long)Math.Round(d);
            string digitsOnly = Regex.Replace(strVal, @"[^\d\-]", "");
            return long.TryParse(digitsOnly, NumberStyles.Any, CultureInfo.InvariantCulture, out long l) ? l : 0L;
        }

        if (underlyingType == typeof(double))
            return double.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double dbl) ? dbl : 0.0;

        if (underlyingType == typeof(bool))
        {
            if (bool.TryParse(strVal, out bool b)) return b;
            return strVal.Equals("1") || strVal.Equals("YES", StringComparison.OrdinalIgnoreCase) || strVal.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
        }

        return Convert.ChangeType(rawValue, underlyingType);
    }

    internal static string NormalizeHeader(string header) =>
        RuleRegistry.NormalizeIdentifier(header);

    private static DateOnly? ParseToDateOnly(object rawValue, string strVal)
    {
        if (RuleDataCache.TryParseDate(rawValue, out DateOnly parsed))
            return parsed;

        if (DateOnly.TryParseExact(strVal, RecognizedDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly exactDate))
            return exactDate;

        return null;
    }

    private static bool IsRowEmptyOrTotal(object[,] grid, int row, int colCount)
    {
        bool hasValue = false;
        for (int c = 1; c <= colCount; c++)
        {
            object? val = grid[row, c];
            if (val == null || SheetMapperHelper.IsComError(val)) continue;

            string s = val.ToString()?.Trim() ?? string.Empty;
            if (s.Equals("Total", StringComparison.OrdinalIgnoreCase)) return true;
            if (s == " " || s == "-" || s == "0" || s == "0.00" || s.Equals("None", StringComparison.OrdinalIgnoreCase))
                continue;

            if (!string.IsNullOrEmpty(s)) hasValue = true;
        }
        return !hasValue;
    }

    private record PropertyBinding(
        PropertyInfo Property,
        Func<object, object> TargetInstance,
        int ColumnIndex
    );

    private static List<PropertyBinding> ResolvePropertyBindings<T>(Dictionary<string, int> headerMap)
    {
        var bindings = new List<PropertyBinding>();
        ScanType(typeof(T), root => root);

        void ScanType(Type type, Func<object, object> instanceResolver)
        {
            foreach (var prop in type.GetProperties(BindingFlags.Public | BindingFlags.Instance))
            {
                if (!prop.CanWrite || prop.GetIndexParameters().Length > 0) continue;

                if (typeof(IEnumerable).IsAssignableFrom(prop.PropertyType) &&
                    prop.PropertyType != typeof(string) &&
                    prop.PropertyType != typeof(List<string>))
                {
                    continue;
                }

                if (prop.GetCustomAttribute<DataDetailAttribute>() != null ||
                    (prop.PropertyType.IsClass &&
                     prop.PropertyType != typeof(string) &&
                     prop.PropertyType != typeof(List<string>) &&
                     !prop.PropertyType.IsArray))
                {
                    ScanType(prop.PropertyType, root =>
                    {
                        var parent = instanceResolver(root);
                        if (parent == null) return null!;
                        var child = prop.GetValue(parent);
                        if (child == null)
                        {
                            child = Activator.CreateInstance(prop.PropertyType)!;
                            prop.SetValue(parent, child);
                        }
                        return child;
                    });
                    continue;
                }

                string candidateHeader = prop.Name;
                var colAttr = prop.GetCustomAttribute<ExcelColumnAttribute>();
                if (colAttr != null && !string.IsNullOrWhiteSpace(colAttr.HeaderName))
                {
                    candidateHeader = colAttr.HeaderName;
                }

                string normalized = NormalizeHeader(candidateHeader);
                if (headerMap.TryGetValue(normalized, out int colIndex))
                {
                    bindings.Add(new PropertyBinding(prop, instanceResolver, colIndex));
                }
                else
                {
                    string stripped = normalized.TrimEnd('s');
                    foreach (var (key, val) in headerMap)
                    {
                        if (key.TrimEnd('s').Equals(stripped, StringComparison.OrdinalIgnoreCase))
                        {
                            bindings.Add(new PropertyBinding(prop, instanceResolver, val));
                            break;
                        }
                    }
                }
            }
        }

        if (bindings.Count == 0 && headerMap.Count == 1)
        {
            var singleProp = typeof(T).GetProperties(BindingFlags.Public | BindingFlags.Instance)
                .FirstOrDefault(p => p.CanWrite && p.PropertyType == typeof(string));
            if (singleProp != null)
            {
                bindings.Add(new PropertyBinding(singleProp, root => root, headerMap.Values.First()));
            }
        }

        return bindings;
    }
}