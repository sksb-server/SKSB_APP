namespace SKSB_App.Core.DropDown;

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using ExcelDna.Integration;

public record Config
{
    public string Source { get; set; } = string.Empty;
    public string Transform { get; set; } = string.Empty;
    public string Filter { get; set; } = string.Empty;
    public string SyncLeft { get; set; } = string.Empty;
    public bool IsMultiSelect { get; set; } = false;

    public bool IsEmpty =>
        string.IsNullOrWhiteSpace(Source) &&
        string.IsNullOrWhiteSpace(Transform) &&
        string.IsNullOrWhiteSpace(Filter) &&
        string.IsNullOrWhiteSpace(SyncLeft);
}

public static class Metadata
{
    public static string GetCellConfigKey(dynamic singleCell)
    {
        try
        {
            string sheetName = singleCell.Worksheet.Name?.ToString() ?? "Sheet";
            string cleanSheet = Regex.Replace(sheetName, @"[^\w]", "_");
            return $"CFG_{cleanSheet}_R{singleCell.Row}C{singleCell.Column}";
        }
        catch
        {
            return string.Empty;
        }
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
        var config = ReadConfigInternal(singleCell);
        if (!config.IsEmpty) return config;

        // INHERITANCE: If row > 2 and current cell has no explicit metadata, inherit from Row 2 archetype
        try
        {
            int row = (int)singleCell.Row;
            int col = (int)singleCell.Column;
            dynamic sheet = singleCell.Worksheet;

            if (row > 2)
            {
                dynamic archetypeCell = sheet.Cells[2, col];
                var inherited = ReadConfigInternal(archetypeCell);
                if (!inherited.IsEmpty) return inherited;
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

        // Fallback to Excel Data Validation formula if no custom Name metadata exists
        if (string.IsNullOrWhiteSpace(raw))
        {
            try
            {
                if (singleCell.Validation != null && singleCell.Validation.Type == 3) // 3 = xlValidateList
                {
                    raw = singleCell.Validation.Formula1 ?? string.Empty;
                    config.IsMultiSelect = !singleCell.Validation.ShowError;
                }
            }
            catch { }
        }

        if (string.IsNullOrWhiteSpace(raw)) return config;

        var parts = raw.Split(new[] { " | " }, StringSplitOptions.None);
        config.Source = parts[0].Trim();

        foreach (var p in parts)
        {
            if (p.StartsWith("TRANSFORM:", StringComparison.OrdinalIgnoreCase))
            {
                config.Transform = p.Substring("TRANSFORM:".Length).Trim();
            }
            else if (p.StartsWith("FILTER:", StringComparison.OrdinalIgnoreCase))
            {
                config.Filter = p.Substring("FILTER:".Length).Trim();
            }
            else if (p.StartsWith("SYNC_LEFT:", StringComparison.OrdinalIgnoreCase))
            {
                config.SyncLeft = p.Substring("SYNC_LEFT:".Length).Trim();
            }
            else if (p.StartsWith("MULTI:", StringComparison.OrdinalIgnoreCase))
            {
                if (bool.TryParse(p.Substring("MULTI:".Length).Trim(), out bool m))
                {
                    config.IsMultiSelect = m;
                }
            }
        }

        if (config.Source.StartsWith("=INDIRECT(\"", StringComparison.OrdinalIgnoreCase) && config.Source.EndsWith("\")"))
        {
            config.Source = "=" + config.Source.Substring(11, config.Source.Length - 13);
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

            return xlApp.Evaluate(clean);
        }
        catch
        {
            return null;
        }
    }
}