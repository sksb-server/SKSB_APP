namespace SKSB_App.Core.RuntimeUI.DropDown;

using ExcelDna.Integration;
using SKSB_App.Core.WorkbookOpen.Sheet;
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

public record Config
{
    public string Source { get; set; } = string.Empty;
    public string Transform { get; set; } = string.Empty;
    public string Filter { get; set; } = string.Empty;
    public string SyncLeft { get; set; } = string.Empty;
    public bool IsMultiSelect { get; set; } = false;
    public bool IsWarningOnly { get; set; } = false;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Source) &&
        string.IsNullOrWhiteSpace(Transform) &&
        string.IsNullOrWhiteSpace(Filter) &&
        string.IsNullOrWhiteSpace(SyncLeft);
}

public static class Metadata
{
    private static readonly HashSet<string> ModernFunctions = new(StringComparer.OrdinalIgnoreCase)
    {
        "XLOOKUP", "XMATCH", "VSTACK", "HSTACK", "FILTER", "UNIQUE",
        "SORT", "SORTBY", "LET", "LAMBDA", "TEXTSPLIT", "CHOOSEROWS", "CHOOSECOLS"
    };

    public static string GetCellConfigKey(dynamic singleCell)
    {
        try
        {
            string sheetName = singleCell.Worksheet.Name?.ToString() ?? "Sheet";
            string cleanSheet = Regex.Replace(sheetName, @"[^\w]", "_");
            return $"CFG_{cleanSheet}_R{singleCell.Row}C{singleCell.Column}";
        }
        catch { return string.Empty; }
    }

    public static void PurgeConfigIfPresent(dynamic singleCell)
    {
        try
        {
            string key = GetCellConfigKey(singleCell);
            if (string.IsNullOrEmpty(key)) return;
            dynamic wb = singleCell.Worksheet.Parent;
            try { wb.Names.Item(key).Delete(); } catch { }
        }
        catch { }
    }

    public static void WriteConfig(dynamic singleCell, Config config)
    {
        try
        {
            string key = GetCellConfigKey(singleCell);
            if (string.IsNullOrEmpty(key)) return;
            dynamic wb = singleCell.Worksheet.Parent;
            try { wb.Names.Item(key).Delete(); } catch { }

            var parts = new List<string> { config.Source };
            if (!string.IsNullOrWhiteSpace(config.Transform))
                parts.Add($"TRANSFORM: {config.Transform}");
            if (!string.IsNullOrWhiteSpace(config.Filter))
                parts.Add($"FILTER: {config.Filter}");
            if (!string.IsNullOrWhiteSpace(config.SyncLeft))
                parts.Add($"SYNC_LEFT: {config.SyncLeft}");
            parts.Add($"MULTI: {config.IsMultiSelect}");
            parts.Add($"WARN: {config.IsWarningOnly}");

            string composite = string.Join(" | ", parts);
            string safeConfig = composite.Replace("\"", "\"\"");
            wb.Names.Add(Name: key, RefersTo: $"=\"{safeConfig}\"", Visible: false);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Metadata.WriteConfig Error] {ex.Message}");
        }
    }

    public static Config ReadConfig(dynamic singleCell)
    {
        string styleName = singleCell.Style?.Name?.ToString() ?? string.Empty;
        if (!string.Equals(styleName, Commands.DropDownStyleName, StringComparison.OrdinalIgnoreCase))
        {
            return new Config();
        }

        var config = ReadConfigInternal(singleCell);
        if (!config.IsEmpty) return config;

        try
        {
            int row = (int)singleCell.Row;
            int col = (int)singleCell.Column;
            dynamic sheet = singleCell.Worksheet;

            if (row > 2 && Commands.IsInsideSameTableDataBody(singleCell))
            {
                dynamic archetypeCell = sheet.Cells[2, col];
                string archStyle = archetypeCell.Style?.Name?.ToString() ?? string.Empty;
                if (string.Equals(archStyle, Commands.DropDownStyleName, StringComparison.OrdinalIgnoreCase))
                {
                    var inherited = ReadConfigInternal(archetypeCell);
                    if (!inherited.IsEmpty) return inherited;
                }
            }
        }
        catch { }

        return config;
    }

    private static Config ReadConfigInternal(dynamic singleCell)
    {
        var config = new Config();
        string raw = string.Empty;
        try
        {
            string key = GetCellConfigKey(singleCell);
            if (!string.IsNullOrEmpty(key))
            {
                dynamic wb = singleCell.Worksheet.Parent;
                dynamic nameObj = null;
                try { nameObj = wb.Names.Item(key); } catch { }
                if (nameObj != null)
                {
                    string refersTo = nameObj.RefersTo?.ToString() ?? string.Empty;
                    if (refersTo.StartsWith("=\"") && refersTo.EndsWith("\""))
                    {
                        raw = refersTo.Substring(2, refersTo.Length - 3).Replace("\"\"", "\"");
                    }
                    else
                    {
                        raw = refersTo.TrimStart('=');
                    }
                }
            }
        }
        catch { }

        if (string.IsNullOrWhiteSpace(raw)) return config;

        var parts = raw.Split(new[] { " | " }, StringSplitOptions.None);
        config.Source = parts[0].Trim();

        foreach (var p in parts)
        {
            if (p.StartsWith("TRANSFORM:", StringComparison.OrdinalIgnoreCase))
                config.Transform = p.Substring("TRANSFORM:".Length).Trim();
            else if (p.StartsWith("FILTER:", StringComparison.OrdinalIgnoreCase))
                config.Filter = p.Substring("FILTER:".Length).Trim();
            else if (p.StartsWith("SYNC_LEFT:", StringComparison.OrdinalIgnoreCase))
                config.SyncLeft = p.Substring("SYNC_LEFT:".Length).Trim();
            else if (p.StartsWith("MULTI:", StringComparison.OrdinalIgnoreCase))
            {
                if (bool.TryParse(p.Substring("MULTI:".Length).Trim(), out bool m))
                    config.IsMultiSelect = m;
            }
            else if (p.StartsWith("WARN:", StringComparison.OrdinalIgnoreCase))
            {
                if (bool.TryParse(p.Substring("WARN:".Length).Trim(), out bool w))
                    config.IsWarningOnly = w;
            }
        }

        return config;
    }

    public static object? EvaluateFormulaSafely(dynamic xlApp, string formula)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(formula)) return null;
            string clean = formula.Trim();
            if (clean.StartsWith("=", StringComparison.Ordinal))
                clean = clean.Substring(1).Trim();

            clean = ResolveTableReferences(xlApp, clean);

            // Detect if expression relies on modern Excel 365 calculation engine
            bool requiresModernEngine = ContainsModernFunction(clean);

            // Fast-path: xlApp.Evaluate only for simple/legacy formulas under 250 characters
            if (!requiresModernEngine && clean.Length <= 250)
            {
                object? res = xlApp.Evaluate(clean);
                if (res != null && !(res is int iErr && iErr < -2000000000))
                {
                    return res;
                }
            }

            // Reliable Worksheet Evaluation for modern 365 formulas (XLOOKUP, VSTACK, cross-sheet array joins)
            dynamic wb = xlApp.ActiveWorkbook;
            dynamic dataWs = null;
            try { dataWs = wb.Sheets[SheetOperations.DataSheetName]; } catch { }
            if (dataWs == null && wb.Sheets.Count > 0) dataWs = wb.Sheets[1];

            if (dataWs != null)
            {
                dynamic scratchCell = dataWs.Cells[1, 101]; // CW1
                try
                {
                    scratchCell.Formula2 = "=" + clean;
                    object? val = scratchCell.Value2;
                    // CRITICAL FIX: Extract full dynamic array if the formula spills
                    try { val = scratchCell.SpillingToRange.Value2 ?? val; } catch { }
                    return val;
                }
                catch
                {
                    try
                    {
                        scratchCell.Formula = "=" + clean;
                        object? val = scratchCell.Value2;
                        try { val = scratchCell.SpillingToRange.Value2 ?? val; } catch { }
                        return val;
                    }
                    catch { return null; }
                }
                finally
                {
                    try { scratchCell.Value2 = null; } catch { }
                }
            }

            return xlApp.Evaluate(clean);
        }
        catch
        {
            return null;
        }
    }

    private static bool ContainsModernFunction(string formula)
    {
        foreach (var func in ModernFunctions)
        {
            if (formula.IndexOf(func, StringComparison.OrdinalIgnoreCase) >= 0)
                return true;
        }
        return false;
    }

    public static string ResolveTableReferences(dynamic xlApp, string formula)
    {
        if (string.IsNullOrWhiteSpace(formula) || !formula.Contains("["))
            return formula;

        var regex = new Regex(@"\b([A-Za-z0-9_]+)\[([^\]@\r\n]+)\]");
        return regex.Replace(formula, match =>
        {
            string tableName = match.Groups[1].Value;
            string colName = match.Groups[2].Value.Trim('[', ']');

            try
            {
                dynamic wb = xlApp.ActiveWorkbook;
                if (wb == null) return match.Value;

                foreach (dynamic ws in wb.Sheets)
                {
                    dynamic tbl = null;
                    try { tbl = ws.ListObjects[tableName]; } catch { }
                    if (tbl == null)
                    {
                        try
                        {
                            foreach (dynamic t in ws.ListObjects)
                            {
                                if (string.Equals(t.Name?.ToString(), tableName, StringComparison.OrdinalIgnoreCase))
                                {
                                    tbl = t;
                                    break;
                                }
                            }
                        }
                        catch { }
                    }

                    if (tbl != null)
                    {
                        dynamic colObj = null;
                        try { colObj = tbl.ListColumns[colName]; } catch { }
                        if (colObj == null)
                        {
                            try
                            {
                                foreach (dynamic c in tbl.ListColumns)
                                {
                                    if (string.Equals(c.Name?.ToString(), colName, StringComparison.OrdinalIgnoreCase))
                                    {
                                        colObj = c;
                                        break;
                                    }
                                }
                            }
                            catch { }
                        }

                        if (colObj != null)
                        {
                            dynamic colRange = colObj.DataBodyRange ?? colObj.Range;
                            string sName = ws.Name?.ToString() ?? string.Empty;
                            string cleanSheet = sName.Replace("'", "''");

                            string addr = string.Empty;
                            try { addr = colRange.Address?.ToString() ?? string.Empty; } catch { }
                            if (!string.IsNullOrEmpty(addr))
                            {
                                return $"'{cleanSheet}'!{addr}";
                            }
                        }
                    }
                }
            }
            catch { }

            return match.Value;
        });
    }
}