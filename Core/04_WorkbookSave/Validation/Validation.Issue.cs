namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Globalization;
using System.Text;
using System.Text.RegularExpressions;

public enum ValidationSeverity
{
    Error,
    Warning
}

/// <summary>
/// One structured validation finding. Carries the exact sheet / cell location so the Validation Task Pane
/// can navigate straight to the offending cell. For <see cref="IsConcurrencyConflict"/> issues,
/// <see cref="OriginalValue"/> holds the server's value for that cell (written back by "Keep Server Values").
/// </summary>
public sealed record ValidationIssue(
    ValidationSeverity Severity,
    string SheetName,
    string TableName,
    int ExcelRow,
    int ExcelCol,
    string ColumnHeader,
    string Message,
    bool IsStackedHistoryEdit = false,
    object? OriginalValue = null,
    string? RowKey = null,
    bool IsConcurrencyConflict = false)
{
    // =========================================================================
    // 1. DISPLAY HELPERS
    // =========================================================================
    public string CellAddress =>
        ExcelRow > 0 && ExcelCol > 0 ? $"{ColumnIndexToLetters(ExcelCol)}{ExcelRow}"
        : ExcelRow > 0 ? $"Row {ExcelRow}"
        : "—";

    /// <summary>1 -> A, 26 -> Z, 27 -> AA, 703 -> AAA.</summary>
    public static string ColumnIndexToLetters(int colIndex)
    {
        if (colIndex < 1) return string.Empty;
        var sb = new StringBuilder();
        int n = colIndex;
        while (n > 0)
        {
            int rem = (n - 1) % 26;
            sb.Insert(0, (char)('A' + rem));
            n = (n - 1) / 26;
        }
        return sb.ToString();
    }

    public string ToDisplayString()
    {
        var sb = new StringBuilder();
        sb.Append('[').Append(Severity).Append(']');
        if (!string.IsNullOrWhiteSpace(SheetName)) sb.Append(' ').Append(SheetName);
        if (ExcelRow > 0) sb.Append(" (").Append(CellAddress).Append(')');
        if (!string.IsNullOrWhiteSpace(ColumnHeader)) sb.Append(" '").Append(ColumnHeader).Append('\'');
        sb.Append(": ").Append(Message);
        return sb.ToString();
    }

    public override string ToString() => ToDisplayString();

    // =========================================================================
    // 2. LEGACY STRING MESSAGE CONVERSION
    // =========================================================================
    private static readonly Regex RowNumberRegex = new(@"\bRow\s+(\d+)\b", RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // "Row 14 [All Staffs]: msg", "[All Staffs]: msg", "[Daily Schedule - Day Work slot 1]: Row 5: msg"
    private static readonly Regex BracketPrefixRegex = new(
        @"^(?:Row\s+\d+\s*)?\[(?<bracket>[^\]]+)\]\s*:?\s*(?:Row\s+\d+\s*:\s*)?",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    // "Row 5: msg" / "Row 5 - msg"
    private static readonly Regex RowPrefixRegex = new(
        @"^Row\s+\d+\s*[:\-–—]\s*",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Converts a legacy free-text message (e.g. <c>"  Row 14 [All Staffs]: Field 'X' is required."</c>) into a
    /// structured issue: the row number and leading <c>[SheetName]</c> / <c>[SheetName - context]</c> are lifted
    /// into fields and the redundant prefix is stripped from <see cref="Message"/>.
    /// </summary>
    public static ValidationIssue FromLegacyMessage(
        ValidationSeverity severity,
        string rawMessage,
        string defaultSheetName = "",
        string defaultTableName = "")
    {
        string text = (rawMessage ?? string.Empty).Trim();

        int excelRow = 0;
        var rowMatch = RowNumberRegex.Match(text);
        if (rowMatch.Success && int.TryParse(rowMatch.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsedRow))
        {
            excelRow = parsedRow;
        }

        string sheetName = defaultSheetName ?? string.Empty;
        string context = string.Empty;
        string message = text;

        var bracketMatch = BracketPrefixRegex.Match(text);
        if (bracketMatch.Success)
        {
            string bracket = bracketMatch.Groups["bracket"].Value.Trim();
            int sep = bracket.IndexOf(" - ", StringComparison.Ordinal);
            if (sep > 0)
            {
                sheetName = bracket.Substring(0, sep).Trim();
                context = bracket.Substring(sep + 3).Trim();
            }
            else if (bracket.Length > 0)
            {
                sheetName = bracket;
            }
            message = text.Substring(bracketMatch.Length).Trim();
        }
        else
        {
            var rowPrefix = RowPrefixRegex.Match(text);
            if (rowPrefix.Success) message = text.Substring(rowPrefix.Length).Trim();
        }

        if (!string.IsNullOrEmpty(context))
        {
            message = string.IsNullOrEmpty(message) ? context : $"{context}: {message}";
        }
        if (string.IsNullOrEmpty(message)) message = text;

        return new ValidationIssue(
            severity,
            sheetName,
            defaultTableName ?? string.Empty,
            excelRow,
            0,
            string.Empty,
            message);
    }
}
