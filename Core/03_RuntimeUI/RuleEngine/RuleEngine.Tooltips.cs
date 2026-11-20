namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SKSB_App.Core.Types.Attributes;

/// <summary>
/// Auto-generated column tooltips rendered as header-only hover Notes (styled like native Excel
/// tooltips using the Microsoft Office "Paper" colour scheme). No on-cell Data Validation input
/// popups are created; native Excel Data Validation (Type != 0) is always preserved.
/// </summary>
public static partial class RuleRegistry
{
    // =========================================================================
    // 1. CONSTANTS
    // =========================================================================

    private const string NoteFontName = "Segoe UI";
    private const float NoteFontSizePt = 9f;
    private const double NoteMinWidthPt = 140.0;
    private const double NoteMaxWidthPt = 340.0;
    private const double NoteMinHeightPt = 36.0;
    private const double NoteHorizontalPaddingPt = 20.0;
    private const double NoteVerticalPaddingPt = 16.0;

    private const int XlValidateInputOnly = 0;
    private const int MsoTrue = -1;
    private const int MsoFalse = 0;

    // Microsoft Office "Paper" colour scheme
    private static readonly Color NoteFillColor = Color.FromArgb(0xFF, 0xFE, 0xF4);   // Light 2 +80%
    private static readonly Color NoteBorderColor = Color.FromArgb(0xBF, 0xBF, 0xBF); // Light 1 -25%
    private static readonly Color NoteBodyColor = Color.FromArgb(0x22, 0x26, 0x13);   // Dark 2 -50%
    private static readonly Color NoteTitleColor = Color.FromArgb(0x00, 0x00, 0x00);  // Dark 1

    // =========================================================================
    // 2. TOOLTIP TEXT BUILDER
    // =========================================================================

    /// <summary>
    /// Builds a concise (Title, Body) tooltip for a column. Every C#-mapped column yields a tooltip.
    /// </summary>
    internal static (string Title, string Body)? BuildColumnTooltip(PropertyRuleDescriptor rule, string? legacyFormula = null)
    {
        if (rule == null) return null;

        // ---- Output / Calculated / Read-only fields (never mention RoundDownStep) ----
        if (rule.IsStrictlyCalculated)
        {
            return ("Auto-Calculated (Locked)",
                $"{rule.DisplayFormula ?? "Calculated automatically by SKSB."}\nManaged automatically by SKSB (read-only).");
        }

        if (rule.IsReadOnly)
        {
            return ("Read-Only Column", "Managed automatically by SKSB (read-only).");
        }

        if (!string.IsNullOrWhiteSpace(legacyFormula))
        {
            return ("Formula Column", $"Formula: {legacyFormula}");
        }

        // ---- Input fields (editable columns) ----
        var lines = new List<string>();
        bool isNric = string.Equals(rule.Property?.Name, "NRIC", StringComparison.OrdinalIgnoreCase);
        Type baseType = Nullable.GetUnderlyingType(rule.PropertyType) ?? rule.PropertyType;

        if (rule.IsRequired && !rule.IsWarningOnly)
            lines.Add("Required");

        if (isNric)
            lines.Add("Numbers only (no dashes or letters)");
        else if (rule.IsNumericColumn && rule.EffectiveFormatKind != ColumnFormatKind.Id)
            lines.Add(rule.AllowNonNegativeZero ? "Numbers only (>= 0)" : "Numbers only");

        if (rule.IsDateColumn)
            lines.Add("Date (dd/mm/yyyy)");

        if (baseType == typeof(TimeOnly) || rule.EffectiveFormatKind == ColumnFormatKind.Time)
            lines.Add("Time (hh:mm)");

        if (baseType == typeof(bool))
            lines.Add("Yes / No (TRUE / FALSE)");

        if (!string.IsNullOrWhiteSpace(rule.DateMustBeGreaterThanHeader))
        {
            lines.Add(rule.DateMustBeGreaterThanAllowEqual
                ? $"Must be on or after {rule.DateMustBeGreaterThanHeader}"
                : $"Must be after {rule.DateMustBeGreaterThanHeader}");
        }

        if (rule.Sources.Length > 0)
        {
            if (rule.IsMultiSelect)
                lines.Add(rule.IsWarningOnly ? "Multi-select dropdown (comma-separated, custom allowed)" : "Multi-select dropdown (comma-separated)");
            else
                lines.Add(rule.IsWarningOnly ? "Dropdown list (custom allowed)" : "Select from dropdown list");
        }

        if (rule.SyncRules.Count > 0)
        {
            var distinctTargets = rule.SyncRules
                .Select(s => s.TargetHeader)
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (distinctTargets.Count > 0)
                lines.Add($"Auto-fills: {string.Join(", ", distinctTargets)}");
        }

        var syncSources = GetSyncSourcesForColumn(rule.TableName, rule.HeaderName);
        if (syncSources.Count > 0)
            lines.Add($"Auto-filled by: {string.Join(", ", syncSources)}");

        string? defaultLine = DescribeDefault(rule.DefaultRule);
        if (defaultLine != null)
            lines.Add(defaultLine);

        if (!rule.IsStrictlyCalculated && rule.HasCSharpFormula)
            lines.Add($"Auto-default: {rule.DisplayFormula} (editable)");

        if (!rule.IsLocked && rule.RoundDownStep > 0)
            lines.Add($"Rounds down to nearest {rule.RoundDownStep:0.##}");

        string? transformLine = rule.Transform switch
        {
            TransformKind.UpperCase => "Auto-format: UPPERCASE",
            TransformKind.Trim => "Auto-trimmed",
            TransformKind.ShortNameFirstLast => "Auto-format: First + Last name",
            TransformKind.ShortNameThreeWords => "Auto-format: 3-word short name",
            _ => null
        };
        if (transformLine != null)
            lines.Add(transformLine);

        foreach (var guard in rule.StateGuards)
        {
            string? guardLine = DescribeStateGuard(guard);
            if (guardLine != null)
                lines.Add(guardLine);
        }

        if (lines.Count == 0)
            lines.Add("Optional text");

        return (rule.HeaderName, string.Join("\n", lines));
    }

    private static string? DescribeDefault(ColumnDefaultAttribute? def)
    {
        if (def == null || def.Kind == DefaultKind.None) return null;

        return def.Kind switch
        {
            DefaultKind.AutoIncrement => "Auto-default: next ID",
            DefaultKind.Today => "Auto-default: Today",
            DefaultKind.Literal => $"Auto-default: {def.LiteralValue}",
            DefaultKind.DateOffsetDays => $"Auto-default: {def.ReferenceHeader ?? "Today"} + {def.Offset} days",
            DefaultKind.DateOffsetMonths => $"Auto-default: {def.ReferenceHeader ?? "Today"} + {def.Offset} months",
            _ => null
        };
    }

    private static string? DescribeStateGuard(ColumnStateGuardAttribute g)
    {
        if (g == null || g.Condition == StateConditionKind.None) return null;

        string refs = string.Join("/", g.ReferenceHeaders);
        string forced = $"Auto-sets '{g.ForcedValueWhenTrue}'";

        return g.Condition switch
        {
            StateConditionKind.ColumnNotEmpty => $"{forced} when {refs} is filled",
            StateConditionKind.DateExpired => $"{forced} when {refs} is past",
            StateConditionKind.DateReached => $"{forced} when {refs} is reached",
            StateConditionKind.ColumnEquals => $"{forced} when {refs} = '{g.ExpectedValue}'",
            _ => $"{forced} based on {refs}"
        };
    }

    internal static string Truncate(string? value, int max)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Length <= max ? value : value.Substring(0, max);
    }

    // =========================================================================
    // 3. HEADER CELL NOTE (HOVER)
    // =========================================================================

    /// <summary>
    /// Replaces the header cell's Note with the generated tooltip (hidden until hover), styled like a
    /// clean native Excel tooltip using the Office "Paper" colour scheme. Stale notes are always cleared.
    /// </summary>
    internal static void ApplyColumnHeaderNote(dynamic headerCell, PropertyRuleDescriptor rule, string? legacyFormula = null)
    {
        if (headerCell == null || rule == null) return;

        var tooltip = BuildColumnTooltip(rule, legacyFormula);

        try { headerCell.ClearComments(); } catch { }
        if (tooltip == null) return;

        string title = tooltip.Value.Title;
        string body = tooltip.Value.Body;
        string noteText = $"{title}\n{body}";

        try
        {
            dynamic comment = headerCell.AddComment(noteText);
            try { comment.Visible = false; } catch { }

            dynamic? shape = null;
            try { shape = comment.Shape; } catch { }
            if (shape == null) return;

            // ---- Fill ----
            try { shape.Fill.Visible = MsoTrue; } catch { }
            try { shape.Fill.Solid(); } catch { }
            try { shape.Fill.ForeColor.RGB = ColorTranslator.ToOle(NoteFillColor); } catch { }
            try { shape.Fill.Transparency = 0f; } catch { }

            // ---- Border ----
            try { shape.Line.Visible = MsoTrue; } catch { }
            try { shape.Line.Weight = 0.75f; } catch { }
            try { shape.Line.ForeColor.RGB = ColorTranslator.ToOle(NoteBorderColor); } catch { }

            // ---- Shadow ----
            try { shape.Shadow.Visible = MsoFalse; } catch { }

            // ---- Margins & wrap ----
            try { shape.TextFrame.MarginLeft = 7f; } catch { }
            try { shape.TextFrame.MarginRight = 7f; } catch { }
            try { shape.TextFrame.MarginTop = 5f; } catch { }
            try { shape.TextFrame.MarginBottom = 5f; } catch { }
            try { shape.TextFrame2.WordWrap = MsoTrue; } catch { }

            // ---- Typography ----
            dynamic? font = null;
            try { font = shape.TextFrame.Characters().Font; } catch { }
            if (font != null)
            {
                try { font.Name = NoteFontName; } catch { }
                try { font.Size = NoteFontSizePt; } catch { }
                try { font.Bold = false; } catch { }
                try { font.Color = ColorTranslator.ToOle(NoteBodyColor); } catch { }
            }

            dynamic? titleFont = null;
            try { titleFont = shape.TextFrame.Characters(1, title.Length).Font; } catch { }
            if (titleFont != null)
            {
                try { titleFont.Bold = true; } catch { }
                try { titleFont.Color = ColorTranslator.ToOle(NoteTitleColor); } catch { }
            }

            // ---- Sizing ----
            var (widthPt, heightPt) = MeasureNoteSize(noteText);
            try { shape.TextFrame.AutoSize = false; } catch { }
            try { shape.Width = widthPt; } catch { }
            try { shape.Height = heightPt; } catch { }
        }
        catch { }
    }

    /// <summary>
    /// Deterministic GDI measurement (Segoe UI 9pt) so sizing works while ScreenUpdating = false.
    /// Width clamped to 140..340pt, height >= 36pt.
    /// </summary>
    private static (double WidthPt, double HeightPt) MeasureNoteSize(string text)
    {
        double widthPt = NoteMinWidthPt;
        double heightPt = NoteMinHeightPt;

        try
        {
            using var g = Graphics.FromHwnd(IntPtr.Zero);
            using var font = new Font(NoteFontName, NoteFontSizePt, FontStyle.Bold, GraphicsUnit.Point);
            double dpi = g.DpiX > 0 ? g.DpiX : 96.0;
            double pxToPt = 72.0 / dpi;

            const TextFormatFlags flags = TextFormatFlags.WordBreak | TextFormatFlags.NoPadding | TextFormatFlags.TextBoxControl;

            Size natural = TextRenderer.MeasureText(g, text, font, new Size(int.MaxValue, int.MaxValue), TextFormatFlags.NoPadding);
            widthPt = Math.Clamp(natural.Width * pxToPt + NoteHorizontalPaddingPt, NoteMinWidthPt, NoteMaxWidthPt);

            int wrapWidthPx = (int)Math.Ceiling((widthPt - NoteHorizontalPaddingPt) / pxToPt);
            Size wrapped = TextRenderer.MeasureText(g, text, font, new Size(Math.Max(1, wrapWidthPx), int.MaxValue), flags);
            heightPt = Math.Max(NoteMinHeightPt, wrapped.Height * pxToPt + NoteVerticalPaddingPt);
        }
        catch { }

        return (widthPt, heightPt);
    }

    // =========================================================================
    // 4. COLUMN TOOLTIPS (HEADER NOTE ONLY)
    // =========================================================================

    /// <summary>
    /// Applies the header hover Note. No on-cell input popups are created; any stale
    /// xlValidateInputOnly (Type 0) validation is removed while native Data Validation (Type != 0) is untouched.
    /// </summary>
    internal static void ApplyColumnTooltips(dynamic headerCell, dynamic? targetRange, PropertyRuleDescriptor rule)
    {
        if (rule == null) return;

        ApplyColumnHeaderNote(headerCell, rule);

        if (targetRange == null || rule.IsLocked) return;

        try
        {
            // Probing Type throws when no DV exists — nothing to clean up in that case.
            int existingType = (int)targetRange.Cells[1, 1].Validation.Type;
            if (existingType == XlValidateInputOnly)
            {
                try { targetRange.Validation.Delete(); } catch { }
            }
        }
        catch { }
    }
}
