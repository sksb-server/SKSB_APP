namespace SKSB_App.Core.WorkbookHelper;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using SKSB_App.Core.DropDown;

public record ColumnRule(
    int ColumnIndex,
    string Header,
    bool IsRequired,
    int AlertStyle, // 1 = Stop, 2 = Warning, 3 = Info
    string ErrorMessage,
    Config DropDownConfig
);

public class SheetRulesCache
{
    public string SheetName { get; set; } = string.Empty;
    public List<ColumnRule> ColumnRules { get; set; } = new();
    public DateTime CachedAt { get; set; } = DateTime.UtcNow;
}

public static class CellRulesExtractor
{
    private static readonly Dictionary<string, SheetRulesCache> _cachedRules = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>
    /// Safely extracts and caches validation and dropdown rules.
    /// Completely guarded against COMException 0x800A03EC on cells without data validation.
    /// </summary>
    public static SheetRulesCache InitializeRulesForSheet(dynamic worksheet, int headerRow, int startCol, int columnSpan)
    {
        string sheetName = string.Empty;
        try { sheetName = worksheet.Name?.ToString() ?? "Sheet"; } catch { sheetName = "Sheet"; }

        var cache = new SheetRulesCache { SheetName = sheetName };

        for (int c = 0; c < columnSpan; c++)
        {
            int colIndex = startCol + c;
            string headerText = $"Col_{colIndex}";

            try
            {
                dynamic headerCell = worksheet.Cells[headerRow, colIndex];
                headerText = headerCell.Value2?.ToString()?.Trim() ?? $"Col_{colIndex}";
            }
            catch { }

            bool isRequired = false;
            int alertStyle = 1; // 1 = xlValidAlertStop
            string errorMsg = string.Empty;
            Config config = new();

            try
            {
                dynamic archetypeCell = worksheet.Cells[headerRow + 1, colIndex];

                // SAFE INSPECTION: Excel COM throws 0x800A03EC if .Type is accessed on an unvalidated cell!
                try
                {
                    dynamic val = archetypeCell.Validation;
                    if (val != null)
                    {
                        int vType = (int)val.Type; // If no validation, this throws cleanly to catch
                        isRequired = !(bool)val.IgnoreBlank;
                        alertStyle = (int)val.AlertStyle;
                        errorMsg = val.ErrorMessage?.ToString() ?? string.Empty;
                    }
                }
                catch
                {
                    // Cell has no Data Validation attached; use defaults
                    isRequired = false;
                    alertStyle = 1;
                    errorMsg = string.Empty;
                }

                config = Metadata.ReadConfig(archetypeCell);
            }
            catch { }

            cache.ColumnRules.Add(new ColumnRule(
                ColumnIndex: colIndex,
                Header: headerText,
                IsRequired: isRequired,
                AlertStyle: alertStyle,
                ErrorMessage: errorMsg,
                DropDownConfig: config
            ));
        }

        _cachedRules[sheetName] = cache;
        return cache;
    }

    /// <summary>
    /// Evaluates worksheet data against cached rules on demand (Save/Close/Validation summary).
    /// Does NOT write to the sheet during evaluation to prevent recursive UI events.
    /// </summary>
    public static (List<Dictionary<string, object?>> Rows, List<string> Errors, List<string> Warnings) EvaluateAndExtract(
        dynamic worksheet,
        int startRow,
        int maxRows,
        int startCol,
        int columnSpan,
        int driverColIndex = 1)
    {
        string sheetName = string.Empty;
        try { sheetName = worksheet.Name?.ToString() ?? "Sheet"; } catch { sheetName = "Sheet"; }

        if (!_cachedRules.TryGetValue(sheetName, out var rules))
        {
            rules = InitializeRulesForSheet(worksheet, startRow - 1, startCol, columnSpan);
        }

        var rows = new List<Dictionary<string, object?>>();
        var errors = new List<string>();
        var warnings = new List<string>();

        dynamic xlApp = ExcelDnaUtil.Application;
        int endRow = startRow + maxRows - 1;
        int endCol = startCol + columnSpan - 1;

        object[,] grid;
        try
        {
            grid = worksheet.Range[worksheet.Cells[startRow, startCol], worksheet.Cells[endRow, endCol]].Value2;
        }
        catch
        {
            return (rows, errors, warnings);
        }

        if (grid == null) return (rows, errors, warnings);

        int rowCount = grid.GetLength(0);

        for (int r = 1; r <= rowCount; r++)
        {
            int excelRow = startRow + r - 1;
            string driverVal = DataFormatting.CleanString(grid[r, driverColIndex]);

            // Skip empty rows
            bool hasContent = false;
            for (int c = 1; c <= columnSpan; c++)
            {
                if (!string.IsNullOrWhiteSpace(DataFormatting.CleanString(grid[r, c])))
                {
                    hasContent = true;
                    break;
                }
            }
            if (!hasContent) continue;

            var rowDict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
            bool rowHasError = false;

            for (int c = 0; c < columnSpan; c++)
            {
                var rule = rules.ColumnRules[c];
                object? rawVal = grid[r, c + 1];
                string strVal = DataFormatting.CleanString(rawVal);

                // 1. In-memory Default Fallback Resolution (No COM cell writes during extraction)
                if (string.IsNullOrWhiteSpace(strVal) && !rule.DropDownConfig.IsEmpty && !string.IsNullOrWhiteSpace(rule.DropDownConfig.SyncLeft))
                {
                    try
                    {
                        dynamic cell = worksheet.Cells[excelRow, rule.ColumnIndex];
                        object? autoVal = ResolveDefaultExpression(xlApp, cell, rule.DropDownConfig.SyncLeft);
                        if (autoVal != null && !(autoVal is ExcelError))
                        {
                            strVal = autoVal.ToString()!.Trim();
                            rawVal = autoVal;
                        }
                    }
                    catch { }
                }

                // 2. Validate blanks against the cached rule
                if (string.IsNullOrWhiteSpace(strVal))
                {
                    if (rule.IsRequired && !string.IsNullOrWhiteSpace(driverVal))
                    {
                        string msg = !string.IsNullOrWhiteSpace(rule.ErrorMessage)
                            ? rule.ErrorMessage
                            : $"Field '{rule.Header}' is required and cannot be blank.";

                        string log = $"  Row {excelRow} [{sheetName}]: {msg}";
                        if (rule.AlertStyle == 1) { errors.Add(log); rowHasError = true; }
                        else { warnings.Add(log); }
                    }
                }
                else
                {
                    // 3. Native validation check (Guarded against COMException)
                    try
                    {
                        dynamic cell = worksheet.Cells[excelRow, rule.ColumnIndex];
                        dynamic v = cell.Validation;
                        if (v != null)
                        {
                            int dummy = (int)v.Type; // Verify validation genuinely exists
                            if (!(bool)v.Value)
                            {
                                string msg = !string.IsNullOrWhiteSpace(rule.ErrorMessage)
                                    ? rule.ErrorMessage
                                    : $"Invalid value '{strVal}' for '{rule.Header}'.";

                                string log = $"  Row {excelRow} [{sheetName}]: {msg}";
                                if (rule.AlertStyle == 1) { errors.Add(log); rowHasError = true; }
                                else { warnings.Add(log); }
                            }
                        }
                    }
                    catch { }
                }

                rowDict[rule.Header] = rawVal;
            }

            if (!rowHasError)
            {
                rows.Add(rowDict);
            }
        }

        return (rows, errors, warnings);
    }

    private static object? ResolveDefaultExpression(dynamic xlApp, dynamic cell, string syncExpr)
    {
        try
        {
            int r = (int)cell.Row;
            int c = (int)cell.Column;
            dynamic sheet = cell.Worksheet;

            string leftVal = string.Empty;
            string rightVal = string.Empty;
            try { leftVal = sheet.Cells[r, c - 1].Value2?.ToString()?.Trim() ?? string.Empty; } catch { }
            try { rightVal = sheet.Cells[r, c + 1].Value2?.ToString()?.Trim() ?? string.Empty; } catch { }

            string clean = syncExpr
                .Replace("@to_right", "", StringComparison.OrdinalIgnoreCase)
                .Replace("@to_left", "", StringComparison.OrdinalIgnoreCase)
                .Replace("@left", double.TryParse(leftVal, out _) ? leftVal : $"\"{leftVal}\"")
                .Replace("@right", double.TryParse(rightVal, out _) ? rightVal : $"\"{rightVal}\"")
                .Trim();

            if (clean.StartsWith("=", StringComparison.Ordinal))
                clean = clean.Substring(1).Trim();

            return xlApp.Evaluate(clean);
        }
        catch { return null; }
    }
}