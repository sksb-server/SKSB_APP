namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types.Attributes;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

public static class TableMapper
{
    private static readonly string[] RecognizedDateFormats =
    {
        "yyyy-MM-dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy", "yyyy/MM/dd"
    };

    public static dynamic? ResolveTable(dynamic worksheet, params string[] candidateNames)
    {
        if (worksheet == null || candidateNames == null || candidateNames!.Length == 0) return null;

        foreach (var name in candidateNames!)
        {
            try
            {
                dynamic t = worksheet!.ListObjects[name];
                if (t != null) return t;
            }
            catch { }
        }

        try
        {
            foreach (dynamic t in worksheet!.ListObjects)
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

    public static List<T> ExtractFromTable<T>(dynamic worksheet, params string[] candidateTableNames) where T : new()
    {
        var records = new List<T>();
        dynamic? table = ResolveTable(worksheet, candidateTableNames);
        if (table == null) return records;

        RuleEvaluator.EvaluateAndHealTable(worksheet, table);

        // VALIDATION DEDUPLICATION: Clean and delete duplicate rows before reading body
        var dedupeWarnings = new List<string>();
        int dupes = TableRowManager.CleanTableDuplicateRows(table, dedupeWarnings);

        if (dupes > 0)
        {
            foreach (var w in dedupeWarnings)
            {
                AppLogger.Warn(w);
            }
        }

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

        var propBindings = ResolvePropertyBindings<T>(headerMap);
        if (propBindings.Count == 0) return records;

        for (int r = 1; r <= rowCount; r++)
        {
            if (IsRowEmpty(dataGrid, r, colCount))
                continue;

            T item = new();
            bool hasValidData = false;

            foreach (var binding in propBindings)
            {
                object? rawVal = dataGrid[r, binding.ColumnIndex];
                bool isBlank = rawVal == null || SheetMapperHelper.IsComError(rawVal);

                // Blank / "—" => null; an explicit 0 / "RM 0.00" converts to 0 and is preserved.
                object? converted = isBlank ? null : ConvertValue(rawVal, binding.Property.PropertyType,
                    IsTimeProperty(binding.Property));
                if (converted != null)
                {
                    object targetObj = binding.TargetInstance(item);
                    if (targetObj != null)
                    {
                        binding.Property.SetValue(targetObj, converted);
                        hasValidData = true;
                    }
                }
                else if (Nullable.GetUnderlyingType(binding.Property.PropertyType) != null)
                {
                    // Nullable column: clear any non-null default initializer so "—" round-trips as null.
                    object targetObj = binding.TargetInstance(item);
                    if (targetObj != null) binding.Property.SetValue(targetObj, null);
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
        int recordCount = records!.Count;
        using var _ = AppLogger.Scope($"PopulateTableSafe<{typeof(T).Name}>({tName}, count={recordCount})");

        dynamic xlApp = worksheet!.Application;
        bool origEvents = true;
        bool origUpdating = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

        bool hadTotals = false;
        try { hadTotals = (bool)table.ShowTotals; } catch { }

        try
        {
            // 1. Temporarily disable Total Row (restored in finally)
            if (hadTotals) { try { table.ShowTotals = false; } catch { } }

            // 2. Empty dataset handling
            if (recordCount == 0)
            {
                SheetMapperHelper.ResizeTable(table, 1);
                dynamic? body = table.DataBodyRange;
                if (body != null)
                {
                    int cols = (int)table.ListColumns.Count;
                    int hRow = (int)table.HeaderRowRange.Row;       // <-- Reads Row 3 correctly
                    int sCol = (int)table.HeaderRowRange.Column;
                    for (int c = 1; c <= cols; c++)
                    {
                        dynamic cell = worksheet.Cells[hRow + 1, sCol + c - 1]; // <-- Cleans Row 4 (the archetype row)
                        string f = cell.Formula?.ToString() ?? string.Empty;
                        string hdr = worksheet.Cells[hRow, sCol + c - 1].Value2?.ToString()?.Trim() ?? string.Empty;
                        var cRule = RuleRegistry.GetRule(tName, hdr);
                        if (!f.StartsWith("=") || cRule?.IsStrictlyCalculated == true)
                        {
                            try { cell.Value2 = null; } catch { }
                            RuleRegistry.ApplyCalculatedColumnLock(cell, cRule);
                        }

                        // Templates no longer carry "@" / number formats: seed the empty archetype row so the first row a user
                        // types keeps leading zeros (phone / card / receipt text) and the declared numeric formats.
                        if (cRule != null && !f.StartsWith("="))
                        {
                            try
                            {
                                string? propName = cRule.Property?.Name;
                                bool isIc = string.Equals(propName, "NRIC", StringComparison.OrdinalIgnoreCase) ||
                                            hdr.Equals("IC Number", StringComparison.OrdinalIgnoreCase);
                                bool isTime = cRule.Property != null && IsTimeProperty(cRule.Property);
                                bool isOdo = cRule.EffectiveFormatKind == ColumnFormatKind.OdometerDynamic;
                                if (isTime) { /* hh:mm is applied on populate */ }
                                else if (isIc || (cRule.PropertyType == typeof(string) && !isOdo && !RuleRegistry.ShouldApplyFormat(cRule)))
                                    cell.NumberFormat = "@";
                                else if (!isOdo && RuleRegistry.ShouldApplyFormat(cRule))
                                    cell.NumberFormat = cRule.ExcelNumberFormat;
                            }
                            catch { }
                        }
                    }
                }
                return;
            }

            // 3. Resize table to fit data
            int effectiveRows = recordCount;
            SheetMapperHelper.ResizeTable(table, effectiveRows);
            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;

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
                string colHeader = worksheet.Cells[headerRow, binding.ColumnIndex].Value2?.ToString()?.Trim() ?? string.Empty;
                var rule = RuleRegistry.GetRule(tName, colHeader);
                if (archFormula.StartsWith("=") && rule?.HasCSharpFormula != true)
                {
                    continue;
                }

                // 0. PRE-READ FORMATS (Crucial: Read these BEFORE setting colRange.Style, 
                // because colRange includes archCell and applying a named style overwrites its custom decimals)
                string archFmt = string.Empty;
                dynamic? archFmtLocal = null;
                try
                {
                    archFmt = archCell.NumberFormat?.ToString() ?? string.Empty;
                    archFmtLocal = archCell.NumberFormatLocal;
                }
                catch { }

                // Inside TableMapper.PopulateTableSafe, right after reading archFmt:
                bool isTimeColumn = IsTimeProperty(binding.Property);
                bool isOdometerColumn = binding.Property.Name is "CurrentOdometer" or "OdometerBefore" or "OdometerIn" or "OdometerOut" or "OdometerDifference"
                                         || rule?.EffectiveFormatKind == ColumnFormatKind.OdometerDynamic;
                // IC numbers are always written as zero-padded 12-digit text (0 => blank), never with a numeric format.
                bool isIcColumn = binding.Property.Name.Equals("NRIC", StringComparison.OrdinalIgnoreCase) ||
                                  colHeader.Equals("IC Number", StringComparison.OrdinalIgnoreCase);
                // Explicit non-text [ColumnFormat] (e.g. Id) on a non-string property overrides a stale "@" archetype.
                bool overridesTextFormat = !isIcColumn && !isOdometerColumn && RuleRegistry.ShouldApplyFormat(rule) &&
                                           (binding.Property.PropertyType != typeof(string) ||
                                            rule?.EffectiveFormatKind == ColumnFormatKind.Id);
                bool isTextFormattedCell = !isTimeColumn && !overridesTextFormat &&
                                           (isIcColumn ||
                                            string.Equals(archFmt, "@", StringComparison.OrdinalIgnoreCase) ||
                                            binding.Property.PropertyType == typeof(string));


                object[,] colData = new object[effectiveRows, 1];
                for (int r = 0; r < recordCount; r++)
                {
                    object? target = binding.TargetInstance(records[r]!);
                    object? val = target != null ? binding.Property.GetValue(target) : null;

                    if (isTimeColumn)
                    {
                        if (val is string timeStr && !string.IsNullOrWhiteSpace(timeStr) &&
                            TimeOnly.TryParse(timeStr.Trim(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var parsedTime))
                            colData[r, 0] = parsedTime.ToTimeSpan().TotalDays;
                        else
                            colData[r, 0] = string.Empty;
                    }
                    else if (isTextFormattedCell && isIcColumn && val is long lVal)
                    {
                        // 12-digit IC numbers only: write as zero-padded text (and hide 0)
                        colData[r, 0] = lVal > 0 ? lVal.ToString("D12") : string.Empty;
                    }
                    else
                    {
                        colData[r, 0] = FormatForExcel(val, rule)!;
                    }
                }

                dynamic colRange = worksheet.Range[
                    worksheet.Cells[headerRow + 1, binding.ColumnIndex],
                    worksheet.Cells[headerRow + effectiveRows, binding.ColumnIndex]
                ];

                // 2. Preserve explicit column NumberFormat using the pre-read values
                if (isTextFormattedCell)
                {
                    try { colRange.NumberFormat = "@"; } catch { }
                }
                else
                {
                    try
                    {
                        if (isTimeColumn && (string.IsNullOrWhiteSpace(archFmt) ||
                            archFmt.Equals("General", StringComparison.OrdinalIgnoreCase) || archFmt == "@"))
                        {
                            colRange.NumberFormat = "hh:mm";
                        }
                        else if (!string.IsNullOrWhiteSpace(archFmt) && !archFmt.Equals("General", StringComparison.OrdinalIgnoreCase))
                        {
                            colRange.NumberFormat = archFmt;
                        }
                        if (archFmtLocal != null && !isTimeColumn && rule?.EffectiveFormatKind != ColumnFormatKind.Id)
                        {
                            colRange.NumberFormatLocal = archFmtLocal;
                        }
                    }
                    catch { }
                }

                // 2b. Declarative [ColumnFormat] (units, decimals, dash for zero) wins over the archetype format.
                //     Text template columns keep "@"; per-row odometer units are applied after the write.
                if (!isTextFormattedCell && !isOdometerColumn && RuleRegistry.ShouldApplyFormat(rule))
                {
                    try { colRange.NumberFormat = rule!.ExcelNumberFormat; } catch { }
                }

                // 3. Write data
                colRange.Value2 = colData;
                if (rule?.IsStrictlyCalculated == true)
                    RuleRegistry.ApplyCalculatedColumnLock(colRange, rule);

                // 4. Unit display overrides (applied after the archetype format so they win)
                if (isOdometerColumn)
                {
                    ApplyOdometerUnitFormats(worksheet, records, headerRow, binding.ColumnIndex);
                }
            }

            // 5. Physical size verification
            dynamic? finalBody = table.DataBodyRange;
            int finalPhysicalRows = finalBody != null ? (int)finalBody.Rows.Count : 0;
            string finalRangeAddr = table.Range?.Address?.ToString() ?? "N/A";
            AppLogger.Info($"[{tName}] Physical verify after write: DataRows={finalPhysicalRows}, Range={finalRangeAddr}");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"PopulateTableSafe<{typeof(T).Name}> failed on '{tName}'", ex);
        }
        finally
        {
            if (hadTotals) { try { table.ShowTotals = true; } catch { } }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    /// <summary>String properties holding a time of day, stored in Excel as a day fraction formatted "hh:mm".</summary>
    private static bool IsTimeProperty(PropertyInfo prop) =>
        prop.PropertyType == typeof(string) &&
        (prop.Name.Equals("Time", StringComparison.OrdinalIgnoreCase) ||
         prop.Name.Equals("TimeStart", StringComparison.OrdinalIgnoreCase) ||
         prop.Name.Equals("TimeEnd", StringComparison.OrdinalIgnoreCase));

    private static bool IsHourMeterRecord(object? record) => record switch
    {
        SKSB_App.Core.Types.VehicleRecord v => SKSB_App.Core.Types.VehicleRecord.IsHourMeterType(v.VehicleDetail.Type),
        SKSB_App.Core.Types.OdometerLog o => SKSB_App.Core.Types.VehicleRecord.IsHourMeterType(o.VehicleType),
        SKSB_App.Core.Types.Maintenance m => SKSB_App.Core.Types.VehicleRecord.IsHourMeterType(m.VehicleType),
        _ => false
    };

    /// <summary>
    /// Formats an odometer column as "hrs" (machinery) or "km" (road vehicles). Consecutive rows sharing a unit are
    /// formatted as one contiguous range, so the COM call count scales with unit changes, not with row count.
    /// </summary>
    private static void ApplyOdometerUnitFormats<T>(dynamic worksheet, IList<T> records, int headerRow, int column)
    {
        const string hoursFmt = RuleRegistry.OdometerHoursFormat;
        const string kmFmt = RuleRegistry.OdometerKmFormat;

        int runStart = 0;
        while (runStart < records.Count)
        {
            bool isHours = IsHourMeterRecord(records[runStart]);
            int runEnd = runStart;
            while (runEnd + 1 < records.Count && IsHourMeterRecord(records[runEnd + 1]) == isHours) runEnd++;

            try
            {
                dynamic run = worksheet.Range[
                    worksheet.Cells[headerRow + 1 + runStart, column],
                    worksheet.Cells[headerRow + 1 + runEnd, column]];
                run.NumberFormat = isHours ? hoursFmt : kmFmt;
            }
            catch { }

            runStart = runEnd + 1;
        }
    }

    public static object[,] GetRange2DValues(dynamic range, int expectedRows, int expectedCols)
    {
        object? raw = null;
        try { raw = range.Value2; } catch { }
        if (raw == null) return new object[Math.Max(2, expectedRows + 1), Math.Max(2, expectedCols + 1)];
        if (raw is object[,] matrix) return matrix;

        object[,] scalarGrid = new object[2, 2];
        scalarGrid[1, 1] = raw;
        return scalarGrid;
    }

    /// <summary>
    /// Converts a model value to its Excel cell value. With a <paramref name="rule"/>, zero / empty values follow the
    /// column's <see cref="ZeroBehavior"/>: Dash keeps a true numeric 0 (the column format shows the em-dash, so
    /// calculations and sorting still work; empty dates write the em-dash text itself), Blank writes an empty cell.
    /// Sentinel dates (Year &lt;= 1900, i.e. OADate 0 = 00/01/1900) are never written.
    /// </summary>
    public static object? FormatForExcel(object? val, PropertyRuleDescriptor? rule = null)
    {
        const string dash = "—";

        // null => the em-dash TEXT (shown through the 4th ";@" section; colour scales / icon sets ignore text).
        // A genuine 0 stays the number 0.0 and displays as 0 / RM 0.00 / 0 km through the 3rd section.
        bool dashForNull = rule != null && rule.EffectiveZeroBehavior == ZeroBehavior.Dash &&
                           rule.EffectiveFormatKind is not (ColumnFormatKind.Text or ColumnFormatKind.General or ColumnFormatKind.None);

        // null -> truly empty cell. An empty string would be a text value, and "" >= TODAY()
        // evaluates TRUE in Excel, giving blank date cells a green icon.
        if (val == null) return dashForNull ? dash : null;

        if (val is DateOnly d)
        {
            if (d.Year <= 1900) return dashForNull ? dash : null;
            return d.ToDateTime(TimeOnly.MinValue).ToOADate();
        }
        if (val is DateTime dt)
        {
            if (dt.Year <= 1900) return dashForNull ? dash : null;
            return dt.ToOADate();
        }

        if (val is TimeOnly timeOfDay) return timeOfDay.ToTimeSpan().TotalDays; // day fraction, shown through the hh:mm format

        if (rule != null && val is decimal or double or float or int or long)
        {
            double n = Convert.ToDouble(val, CultureInfo.InvariantCulture);
            if (Math.Abs(n) < 0.00001)
                return rule.EffectiveZeroBehavior == ZeroBehavior.Blank ? null : 0.0;
        }

        if (val is decimal dec)
        {
            if (dec == 0.00m) return rule != null ? 0.0 : string.Empty;
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
            // Multi-select cells round-trip as "Item1, Item2"; an empty list is a truly empty cell.
            string joined = string.Join(", ", strList.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()));
            return joined.Length == 0 ? null : joined;
        }
        if (val is IList enumList && val.GetType().IsGenericType && val.GetType().GetGenericArguments()[0].IsEnum)
        {
            string joinedEnums = string.Join(", ", enumList.Cast<object>()
                .Select(e => e.ToString() ?? string.Empty)
                .Where(n => n.Length > 0 && !n.Equals("None", StringComparison.OrdinalIgnoreCase)));
            return joinedEnums.Length == 0 ? null : joinedEnums;
        }
        return val;
    }

    public static object? ConvertValue(object? rawValue, Type targetType, bool isTimeColumn = false)
    {
        if (rawValue == null || SheetMapperHelper.IsComError(rawValue)) return null;

        if (isTimeColumn && (Nullable.GetUnderlyingType(targetType) ?? targetType) == typeof(string))
        {
            if (rawValue is double oaTime)
            {
                double frac = oaTime - Math.Floor(oaTime);
                var secs = Math.Round(frac * 86400.0);
                if (secs >= 86400) secs = 0;
                return TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(secs)).ToString("HH:mm", CultureInfo.InvariantCulture);
            }
            if (rawValue is DateTime dtTime) return dtTime.ToString("HH:mm", CultureInfo.InvariantCulture);
            string tStr = rawValue.ToString()?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(tStr)) return null;
            if (TimeOnly.TryParse(tStr, CultureInfo.InvariantCulture, DateTimeStyles.None, out var tOnly))
                return tOnly.ToString("HH:mm", CultureInfo.InvariantCulture);
            return tStr;
        }

        Type underlyingType = Nullable.GetUnderlyingType(targetType) ?? targetType;
        string strVal = rawValue.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrWhiteSpace(strVal) || strVal == " " || strVal == "-" || strVal == "—") return null;

        if (underlyingType == typeof(List<string>))
        {
            return strVal.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                         .Select(s => s.Trim())
                         .Where(s => !string.IsNullOrEmpty(s))
                         .Distinct(StringComparer.OrdinalIgnoreCase)
                         .ToList();
        }

        if (underlyingType.IsGenericType && underlyingType.GetGenericTypeDefinition() == typeof(List<>) &&
            underlyingType.GetGenericArguments()[0].IsEnum)
        {
            // List<TEnum> round-trips as "Pipe_Repair, Road_Reinstatement" (spaces tolerated, unknown / None dropped).
            Type enumType = underlyingType.GetGenericArguments()[0];
            var list = (IList)Activator.CreateInstance(underlyingType)!;
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var token in strVal.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries))
            {
                string clean = token.Replace(" ", "_").Replace("-", "_");
                if (Enum.TryParse(enumType, clean, true, out var parsed) &&
                    !string.Equals(parsed!.ToString(), "None", StringComparison.OrdinalIgnoreCase) &&
                    seen.Add(parsed.ToString()!))
                {
                    list.Add(parsed);
                }
            }
            return list;
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

        if (underlyingType == typeof(TimeOnly))
        {
            if (rawValue is double oaTod)
            {
                double secs = Math.Round((oaTod - Math.Floor(oaTod)) * 86400.0);
                return TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(secs >= 86400 ? 0 : secs));
            }
            if (rawValue is DateTime dtTod) return TimeOnly.FromDateTime(dtTod);
            return TimeOnly.TryParse(strVal, CultureInfo.InvariantCulture, DateTimeStyles.None, out var todParsed)
                ? todParsed
                : null;
        }

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
        {
            if (decimal.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out decimal dec)) return dec;
            return decimal.TryParse(StripUnits(strVal), NumberStyles.Any, CultureInfo.InvariantCulture, out dec) ? dec : 0.00m;
        }

        if (underlyingType == typeof(int))
        {
            if (rawValue is double d) return (int)Math.Round(d);
            if (int.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out int i)) return i;
            return int.TryParse(StripUnits(strVal), NumberStyles.Any, CultureInfo.InvariantCulture, out i) ? i : 0;
        }

        if (underlyingType == typeof(long))
        {
            if (rawValue is double d) return (long)Math.Round(d);
            string digitsOnly = Regex.Replace(strVal, @"[^\d\-]", "");
            return long.TryParse(digitsOnly, NumberStyles.Any, CultureInfo.InvariantCulture, out long l) ? l : 0L;
        }

        if (underlyingType == typeof(double))
        {
            if (double.TryParse(strVal, NumberStyles.Any, CultureInfo.InvariantCulture, out double dbl)) return dbl;
            return double.TryParse(StripUnits(strVal), NumberStyles.Any, CultureInfo.InvariantCulture, out dbl) ? dbl : 0.0;
        }

        if (underlyingType == typeof(bool))
        {
            if (bool.TryParse(strVal, out bool b)) return b;
            return strVal.Equals("1") || strVal.Equals("YES", StringComparison.OrdinalIgnoreCase) || strVal.Equals("TRUE", StringComparison.OrdinalIgnoreCase);
        }

        return Convert.ChangeType(rawValue, underlyingType);
    }

    public static string NormalizeHeader(string header) =>
        RuleRegistry.NormalizeIdentifier(
            (header ?? string.Empty).Replace("reciept", "receipt", StringComparison.OrdinalIgnoreCase));

    /// <summary>Removes unit text ("RM", "km", "hrs", "/L", "shifts", "yrs", "days"...) and thousands separators from a typed number.</summary>
    private static string StripUnits(string text) => Regex.Replace(text, @"[^\d\.\-]", string.Empty);

    /// <summary>Parses a date cell. Sentinel dates (Year &lt;= 1900, OADate 0 = 00/01/1900) are treated as empty.</summary>
    private static DateOnly? ParseToDateOnly(object rawValue, string strVal)
    {
        if (RuleDataCache.TryParseDate(rawValue, out DateOnly parsed))
            return parsed.Year <= 1900 ? null : parsed;
        if (DateOnly.TryParseExact(strVal, RecognizedDateFormats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateOnly exactDate))
            return exactDate.Year <= 1900 ? null : exactDate;
        return null;
    }

    private static bool IsRowEmpty(object[,] grid, int row, int colCount)
    {
        bool hasValue = false;
        for (int c = 1; c <= colCount; c++)
        {
            object? val = grid[row, c];
            if (val == null || SheetMapperHelper.IsComError(val)) continue;
            string s = val.ToString()?.Trim() ?? string.Empty;
            if (s == " " || s == "-" || s == "—" || s == "0" || s == "0.00" || s.Equals("None", StringComparison.OrdinalIgnoreCase) ||
                s.Equals("False", StringComparison.OrdinalIgnoreCase))
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
                    !RuleRegistry.IsMultiValueListType(prop.PropertyType))
                {
                    continue;
                }

                if (prop.GetCustomAttribute<DataDetailAttribute>() != null ||
                    (prop.PropertyType.IsClass &&
                     prop.PropertyType != typeof(string) &&
                     !RuleRegistry.IsMultiValueListType(prop.PropertyType) &&
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
                    bool found = false;

                    // 1. Plural match
                    foreach (var (key, val) in headerMap)
                    {
                        if (key.TrimEnd('s').Equals(stripped, StringComparison.OrdinalIgnoreCase))
                        {
                            bindings.Add(new PropertyBinding(prop, instanceResolver, val));
                            found = true;
                            break;
                        }
                    }

                    // 2. Cost suffix tolerance: matches "Tyre" <-> "Tyre Cost", "Brake" <-> "Brake Cost", etc.
                    if (!found)
                    {
                        foreach (var (key, val) in headerMap)
                        {
                            if (key.Equals(normalized + "cost", StringComparison.OrdinalIgnoreCase) ||
                                normalized.Equals(key + "cost", StringComparison.OrdinalIgnoreCase))
                            {
                                bindings.Add(new PropertyBinding(prop, instanceResolver, val));
                                break;
                            }
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
