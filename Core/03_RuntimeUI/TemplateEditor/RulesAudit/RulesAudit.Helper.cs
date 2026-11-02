namespace SKSB_App.Core.RuntimeUI.TemplateEditor.RulesAudit;

using System;
using System.Text.RegularExpressions;

public static class RulesAuditHelper
{
    public class HeaderCellInfo
    {
        public string Address { get; set; } = string.Empty;
        public int Row { get; set; }
        public int Col { get; set; }
        public string Style { get; set; } = "Normal";
        public string NumberFormat { get; set; } = "@";
        public string Formula { get; set; } = string.Empty;
        public string SourceOrFormula { get; set; } = string.Empty;
        public string NativeType { get; set; } = string.Empty;
        public string NativeAlert { get; set; } = "Stop";
        public bool NativeIgnoreBlank { get; set; } = true;
    }

    public static dynamic? ResolveSheetSafely(dynamic wb, string rawName)
    {
        if (wb == null || string.IsNullOrWhiteSpace(rawName)) return null;
        string cleanName = rawName.Trim().Trim('`', '"', '\'');

        try
        {
            dynamic s = wb.Sheets[cleanName];
            if (s != null) return s;
        }
        catch { }

        try
        {
            foreach (dynamic ws in wb.Sheets)
            {
                string sName = ws.Name?.ToString() ?? string.Empty;
                if (string.Equals(sName.Trim(), cleanName, StringComparison.OrdinalIgnoreCase))
                    return ws;
            }
        }
        catch { }

        try
        {
            string cleanTarget = Regex.Replace(cleanName, @"[\s_\-]+", "").ToLowerInvariant();
            foreach (dynamic ws in wb.Sheets)
            {
                string sName = ws.Name?.ToString() ?? string.Empty;
                string cleanSheet = Regex.Replace(sName, @"[\s_\-]+", "").ToLowerInvariant();
                if (cleanSheet.Equals(cleanTarget, StringComparison.OrdinalIgnoreCase))
                    return ws;
                if (cleanSheet.TrimEnd('s').Equals(cleanTarget.TrimEnd('s'), StringComparison.OrdinalIgnoreCase))
                    return ws;
            }
        }
        catch { }

        return null;
    }

    public static string GetStyleName(dynamic cell)
    {
        try { return cell.Style?.Name?.ToString() ?? "Normal"; }
        catch { return "Normal"; }
    }

    public static string GetCellNumberFormat(dynamic cell)
    {
        try
        {
            string fmt = cell.NumberFormat?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(fmt) || fmt.Equals("General", StringComparison.OrdinalIgnoreCase))
                return "General";
            return fmt;
        }
        catch
        {
            return "General";
        }
    }

    public static string ExtractHeaderVal(string line, string key)
    {
        var m = Regex.Match(line, $@"{key}=`?([^`|\r\n]+?)`?(?:\s*\||\s*$)");
        return m.Success ? m.Groups[1].Value.Trim().Trim('`', '\'') : string.Empty;
    }

    public static string EscapePipe(string input) =>
        string.IsNullOrWhiteSpace(input) ? " " : input.Replace("|", "\\|").Replace("\r", "").Replace("\n", " ");

    public static string UnescapePipe(string input) =>
        input.Equals(" ") ? string.Empty : input.Replace("\\|", "|");

    public static string GetExcelColumnLetter(int colNumber)
    {
        string letter = string.Empty;
        while (colNumber > 0)
        {
            int rem = (colNumber - 1) % 26;
            letter = (char)('A' + rem) + letter;
            colNumber = (colNumber - rem) / 26;
        }
        return letter;
    }

    public static int ColumnLetterToIndex(string colLetter)
    {
        int sum = 0;
        foreach (char c in colLetter.ToUpperInvariant())
        {
            sum *= 26;
            sum += c - 'A' + 1;
        }
        return sum;
    }

    public static string GetValidationTypeName(int valType) => valType switch
    {
        1 => "WholeNumber",
        2 => "Decimal",
        3 => "List",
        4 => "Date",
        5 => "Time",
        6 => "TextLength",
        7 => "CustomFormula",
        _ => "Unknown"
    };

    public static int GetValidationTypeCode(string valType) => valType switch
    {
        "WholeNumber" => 1,
        "Decimal" => 2,
        "List" => 3,
        "Date" => 4,
        "Time" => 5,
        "TextLength" => 6,
        "CustomFormula" => 7,
        _ => 3
    };
}