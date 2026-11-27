namespace SKSB_App.Core.RuntimeUI.Guide;

using System;
using System.Linq;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.Validation;
using DropDownCommands = SKSB_App.Core.RuntimeUI.DropDown.Commands;

/// <summary>
/// Reads cell / table / rule / dropdown / formula / validation metadata for the Cell Guide panel.
/// MACRO-THREAD ONLY: every member touches Excel COM. The returned <see cref="CellGuideSnapshot"/> is
/// pure managed data. Every COM read is individually guarded so a protected sheet, a chart sheet or
/// a header-less table degrades gracefully instead of throwing.
/// </summary>
public static class CellContextInspector
{
    // =========================================================================
    // 1. ENTRY POINT
    // =========================================================================
    private const int MaxPreviewLength = 120;

    public static CellGuideSnapshot Inspect(dynamic? sheet, dynamic? targetRange)
    {
        var snap = new CellGuideSnapshot();
        if (sheet == null || targetRange == null) return snap;

        try { snap.SheetName = Convert.ToString(sheet!.Name) ?? string.Empty; } catch { }

        dynamic? cell = null;
        try { cell = targetRange!.Cells[1, 1]; } catch { }
        if (cell == null) return snap;

        try { snap.SelectedCellCount = Convert.ToInt64(targetRange!.CountLarge); } catch { }
        try { snap.CellAddress = (Convert.ToString(cell.Address) ?? string.Empty).Replace("$", string.Empty); } catch { }

        ReadCellContent(cell, snap);

        dynamic? table = null;
        try { table = cell.ListObject; } catch { }
        if (table == null)
        {
            snap.Zone = CellZone.OutsideTable;
            return snap;
        }

        ReadTableContext(sheet, cell, table, snap);
        return snap;
    }

    // =========================================================================
    // 2. CELL CONTENT
    // =========================================================================
    private static void ReadCellContent(dynamic cell, CellGuideSnapshot snap)
    {
        try
        {
            string text = Convert.ToString(cell.Text) ?? string.Empty;
            snap.CurrentValuePreview = Truncate(text.Trim(), MaxPreviewLength);
        }
        catch { }

        try { snap.HasExcelFormula = Convert.ToBoolean(cell.HasFormula); } catch { }
        if (snap.HasExcelFormula)
        {
            try { snap.FormulaHint = Truncate(Convert.ToString(cell.Formula) ?? string.Empty, 220); } catch { }
        }
    }

    // =========================================================================
    // 3. TABLE / ZONE / RULE CONTEXT
    // =========================================================================
    private static void ReadTableContext(dynamic sheet, dynamic cell, dynamic table, CellGuideSnapshot snap)
    {
        string tableName = string.Empty;
        try { tableName = Convert.ToString(table.Name) ?? string.Empty; } catch { }
        snap.TableName = tableName;

        try { snap.TableNature = TableRowManager.GetTableNatureByName(tableName); } catch { }
        try { snap.IsHistoryPreservingTable = StackedEntryGuard.IsStackedTable(tableName); } catch { }

        int row = 0, col = 0, headerRow = -1, totalRow = -1;
        try { row = Convert.ToInt32(cell.Row); } catch { }
        try { col = Convert.ToInt32(cell.Column); } catch { }
        try { headerRow = Convert.ToInt32(table.HeaderRowRange.Row); } catch { }
        try
        {
            if (Convert.ToBoolean(table.ShowTotals))
                totalRow = Convert.ToInt32(table.TotalsRowRange.Row);
        }
        catch { }

        snap.Zone = row == headerRow ? CellZone.Header
                  : row == totalRow ? CellZone.Total
                  : CellZone.Data;

        if (headerRow > 0 && col > 0)
        {
            try { snap.ColumnHeader = (Convert.ToString(sheet.Cells[headerRow, col].Value2) ?? string.Empty).Trim(); } catch { }
        }

        PropertyRuleDescriptor? rule = null;
        if (!string.IsNullOrWhiteSpace(tableName) && !string.IsNullOrWhiteSpace(snap.ColumnHeader))
        {
            try { rule = RuleRegistry.GetRule(tableName, snap.ColumnHeader!); } catch { }
        }

        if (rule != null) ApplyRule(rule, tableName, snap);

        if (snap.Zone == CellZone.Data)
        {
            if (!snap.HasDropdown && !snap.IsLocked)
            {
                try { snap.HasDropdown = DropDownCommands.IsDropdownCell(cell); } catch { }
            }
            ReadExcelValidation(cell, snap);
        }
    }

    private static void ApplyRule(PropertyRuleDescriptor rule, string tableName, CellGuideSnapshot snap)
    {
        snap.HasRule = true;
        snap.IsRequired = rule.IsRequired;
        snap.IsWarningOnly = rule.IsWarningOnly;
        snap.IsStrictlyCalculated = rule.IsStrictlyCalculated;
        snap.IsReadOnly = rule.IsReadOnly;
        snap.HasDropdown = rule.HasDropdown;
        snap.IsMultiSelect = rule.IsMultiSelect;

        if (rule.DefaultRule != null && rule.DefaultRule.Kind != DefaultKind.None)
        {
            snap.HasDefaultValue = true;
            snap.DefaultKind = rule.DefaultRule.Kind;
            snap.DefaultLiteral = rule.DefaultRule.LiteralValue;
        }

        snap.SyncTargets = rule.SyncRules
            .Select(s => s.TargetHeader)
            .Where(h => !string.IsNullOrWhiteSpace(h))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
        try { snap.SyncedFromSources = RuleRegistry.GetSyncSourcesForColumn(tableName, rule.HeaderName); } catch { }
        try { snap.CalculatedFromSources = RuleRegistry.GetCalculationSources(tableName, rule.HeaderName); } catch { }
        try { snap.AffectedColumns = RuleRegistry.GetAffectedColumnsForInput(tableName, rule.HeaderName); } catch { }
        try
        {
            snap.RollupDescriptions = rule.RollupRules
                .Select(r => r.SourceProperties.Length > 0
                    ? $"{r.CollectionOrTable} ({string.Join(", ", r.SourceProperties)})"
                    : r.CollectionOrTable)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch { }

        snap.DateMustBeGreaterThanHeader = rule.DateMustBeGreaterThanHeader;
        snap.DateMustBeGreaterThanAllowEqual = rule.DateMustBeGreaterThanAllowEqual;

        snap.FormatKind = rule.EffectiveFormatKind;
        snap.ExcelNumberFormat = rule.ExcelNumberFormat;
        snap.IsDateColumn = rule.IsDateColumn;
        snap.IsNumericColumn = rule.IsNumericColumn;

        if (rule.HasCSharpFormula && rule.DisplayFormula != null)
            snap.FormulaHint = Truncate(rule.DisplayFormula, 220);
    }

    private static void ReadExcelValidation(dynamic cell, CellGuideSnapshot snap)
    {
        // Throws when the cell carries no Data Validation: that is the normal case.
        try
        {
            dynamic validation = cell.Validation;
            string title = Convert.ToString(validation.InputTitle) ?? string.Empty;
            string message = Convert.ToString(validation.InputMessage) ?? string.Empty;
            if (!string.IsNullOrWhiteSpace(title)) snap.ExcelValidationTitle = title.Trim();
            if (!string.IsNullOrWhiteSpace(message)) snap.ExcelValidationMessage = message.Trim();
        }
        catch { }
    }

    private static string Truncate(string value, int max)
        => value.Length <= max ? value : value.Substring(0, max - 1) + "…";
}
