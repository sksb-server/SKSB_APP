namespace SKSB_App.Core.RuntimeUI.RuleEngine;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using SKSB_App.Core.Localization;
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
            return (Loc.Tip("note.calculated.title"),
                $"{rule.DisplayFormula ?? Loc.Tip("note.calculated.defaultBody")}\n{Loc.Tip("note.calculated.managed")}");
        }

        if (rule.IsReadOnly)
        {
            string? source = Loc.ResolveColumnSource(rule.TableName, rule.HeaderName);
            string body = string.IsNullOrWhiteSpace(source)
                ? Loc.Tip("note.readonly.body")
                : $"{source}\n{Loc.Tip("note.readonly.body")}";
            return (Loc.Tip("note.readonly.title"), body);
        }

        if (!string.IsNullOrWhiteSpace(legacyFormula))
        {
            return (Loc.Tip("note.formula.title"), Loc.Tip("note.formula.body", legacyFormula));
        }

        // ---- Input fields (editable columns) ----
        var lines = new List<string>();
        bool isNric = string.Equals(rule.Property?.Name, "NRIC", StringComparison.OrdinalIgnoreCase);
        Type baseType = Nullable.GetUnderlyingType(rule.PropertyType) ?? rule.PropertyType;

        if (rule.IsRequired && !rule.IsWarningOnly)
            lines.Add(Loc.Tip("note.required"));

        if (isNric)
            lines.Add(Loc.Tip("note.nric"));
        else if (rule.IsNumericColumn && rule.EffectiveFormatKind != ColumnFormatKind.Id)
            lines.Add(rule.AllowNonNegativeZero ? Loc.Tip("note.numericNonNeg") : Loc.Tip("note.numeric"));

        if (rule.IsDateColumn)
            lines.Add(Loc.Tip("note.date"));

        if (baseType == typeof(TimeOnly) || rule.EffectiveFormatKind == ColumnFormatKind.Time)
            lines.Add(Loc.Tip("note.time"));

        if (baseType == typeof(bool))
            lines.Add(Loc.Tip("note.yesNo"));

        if (!string.IsNullOrWhiteSpace(rule.DateMustBeGreaterThanHeader))
        {
            lines.Add(rule.DateMustBeGreaterThanAllowEqual
                ? Loc.Tip("note.dateOnOrAfter", rule.DateMustBeGreaterThanHeader)
                : Loc.Tip("note.dateAfter", rule.DateMustBeGreaterThanHeader));
        }

        if (rule.Sources.Length > 0)
        {
            if (rule.IsMultiSelect)
                lines.Add(rule.IsWarningOnly ? Loc.Tip("note.multiDropdownCustom") : Loc.Tip("note.multiDropdown"));
            else
                lines.Add(rule.IsWarningOnly ? Loc.Tip("note.dropdownCustom") : Loc.Tip("note.dropdown"));
        }

        if (rule.SyncRules.Count > 0)
        {
            var distinctTargets = rule.SyncRules
                .Select(s => s.TargetHeader)
                .Where(h => !string.IsNullOrWhiteSpace(h))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            if (distinctTargets.Count > 0)
                lines.Add(Loc.Tip("note.autoFills", string.Join(", ", distinctTargets)));
        }

        var syncSources = GetSyncSourcesForColumn(rule.TableName, rule.HeaderName);
        if (syncSources.Count > 0)
            lines.Add(Loc.Tip("note.autoFilledBy", string.Join(", ", syncSources)));

        string? defaultLine = DescribeDefault(rule.DefaultRule);
        if (defaultLine != null)
            lines.Add(defaultLine);

        if (!rule.IsStrictlyCalculated && rule.HasCSharpFormula)
            lines.Add(Loc.Tip("note.autoDefaultFormula", rule.DisplayFormula ?? string.Empty));

        if (!rule.IsLocked && rule.RoundDownStep > 0)
            lines.Add(Loc.Tip("note.roundDown", $"{rule.RoundDownStep:0.##}"));

        string? transformLine = rule.Transform switch
        {
            TransformKind.UpperCase => Loc.Tip("note.transform.upper"),
            TransformKind.Trim => Loc.Tip("note.transform.trim"),
            TransformKind.ShortNameFirstLast => Loc.Tip("note.transform.shortFirstLast"),
            TransformKind.ShortNameThreeWords => Loc.Tip("note.transform.shortThreeWords"),
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
            lines.Add(Loc.Tip("note.optionalText"));

        string title = rule.HeaderName;
        try
        {
            string? displayName = Loc.ResolveColumnGuide(rule.TableName, rule.HeaderName).DisplayName;
            if (!string.IsNullOrWhiteSpace(displayName) &&
                !string.Equals(displayName, rule.HeaderName, StringComparison.OrdinalIgnoreCase))
            {
                title = $"{displayName} ({rule.HeaderName})";
            }
        }
        catch { }

        return (title, string.Join("\n", lines));
    }

    private static string? DescribeDefault(ColumnDefaultAttribute? def)
    {
        if (def == null || def.Kind == DefaultKind.None) return null;

        return def.Kind switch
        {
            DefaultKind.AutoIncrement => Loc.Tip("note.default.nextId"),
            DefaultKind.Today => Loc.Tip("note.default.today"),
            DefaultKind.Literal => Loc.Tip("note.default.literal", def.LiteralValue ?? string.Empty),
            DefaultKind.DateOffsetDays => Loc.Tip("note.default.dateOffset", def.ReferenceHeader ?? Loc.Tip("note.unit.today"), $"{def.Offset} {Loc.Tip("note.unit.days")}"),
            DefaultKind.DateOffsetMonths => Loc.Tip("note.default.dateOffset", def.ReferenceHeader ?? Loc.Tip("note.unit.today"), $"{def.Offset} {Loc.Tip("note.unit.months")}"),
            _ => Loc.Tip("note.default.generic")
        };
    }

    private static string? DescribeStateGuard(ColumnStateGuardAttribute g)
    {
        if (g == null || g.Condition == StateConditionKind.None) return null;

        string refs = string.Join("/", g.ReferenceHeaders);
        string forced = g.ForcedValueWhenTrue ?? string.Empty;

        return g.Condition switch
        {
            StateConditionKind.ColumnNotEmpty => Loc.Tip("note.guard.notEmpty", forced, refs),
            StateConditionKind.DateExpired => Loc.Tip("note.guard.dateExpired", forced, refs),
            StateConditionKind.DateReached => Loc.Tip("note.guard.dateReached", forced, refs),
            StateConditionKind.ColumnEquals => Loc.Tip("note.guard.equals", forced, refs, g.ExpectedValue ?? string.Empty),
            _ => Loc.Tip("note.guard.other", forced, refs)
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

    /// <summary>
    /// Re-applies localized header hover Notes to every table column across all visible worksheets
    /// in the workbook. Must be called on Excel's macro thread.
    /// </summary>
    public static void RefreshWorkbookHeaderNotes(dynamic? wb)
    {
        if (wb == null) return;

        dynamic? app = null;
        bool origUpdating = true;
        bool origEvents = true;
        bool capturedUpdating = false;
        bool capturedEvents = false;

        try
        {
            app = wb.Application;
            try { origUpdating = (bool)app.ScreenUpdating; capturedUpdating = true; app.ScreenUpdating = false; } catch { }
            try { origEvents = (bool)app.EnableEvents; capturedEvents = true; app.EnableEvents = false; } catch { }

            foreach (dynamic ws in wb.Worksheets)
            {
                try
                {
                    // xlSheetVisible = -1 (skip hidden / very-hidden sheets such as Lookups/Data).
                    if ((int)ws.Visible != -1) continue;

                    dynamic? listObjects = ws.ListObjects;
                    if (listObjects == null) continue;

                    int tableCount = (int)listObjects.Count;
                    for (int t = 1; t <= tableCount; t++)
                    {
                        try
                        {
                            dynamic table = listObjects[t];
                            string? tableName = table.Name?.ToString();
                            if (string.IsNullOrWhiteSpace(tableName)) continue;

                            dynamic? headerRow = table.HeaderRowRange;
                            if (headerRow == null) continue;

                            int colCount = (int)table.ListColumns.Count;
                            for (int c = 1; c <= colCount; c++)
                            {
                                try
                                {
                                    dynamic headerCell = headerRow.Cells[1, c];
                                    string? header = headerCell.Value2?.ToString()?.Trim();
                                    if (string.IsNullOrWhiteSpace(header)) continue;

                                    var rule = GetRule(tableName, header);
                                    if (rule != null)
                                        ApplyColumnHeaderNote(headerCell, rule);
                                }
                                catch { }
                            }
                        }
                        catch { }
                    }
                }
                catch { }
            }
        }
        catch { }
        finally
        {
            if (app != null)
            {
                if (capturedEvents) { try { app.EnableEvents = origEvents; } catch { } }
                if (capturedUpdating) { try { app.ScreenUpdating = origUpdating; } catch { } }
            }
        }
    }
}
