namespace SKSB_App.Core.RuntimeUI.Guide;

using System;
using System.Collections.Generic;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types.Attributes;

/// <summary>Where the active cell sits relative to an Excel table (ListObject).</summary>
public enum CellZone
{
    /// <summary>No workbook / sheet / selection is available.</summary>
    None,

    /// <summary>The cell is not inside any ListObject.</summary>
    OutsideTable,

    /// <summary>The cell is a table header cell.</summary>
    Header,

    /// <summary>The cell is a table data body cell.</summary>
    Data,

    /// <summary>The cell is in the table Total row.</summary>
    Total
}

/// <summary>
/// Pure managed snapshot of everything the Cell Guide panel shows about one cell.
/// Built on Excel's macro thread by <see cref="CellContextInspector"/>; holds NO COM references, so it is
/// safe to keep inside WinForms controls and to re-render later (e.g. after a language switch).
/// </summary>
public sealed class CellGuideSnapshot
{
    // ---- Location -------------------------------------------------------------------------
    public string SheetName { get; set; } = string.Empty;
    public string CellAddress { get; set; } = string.Empty;
    public long SelectedCellCount { get; set; } = 1;
    public CellZone Zone { get; set; } = CellZone.None;

    // ---- Table ----------------------------------------------------------------------------
    public string? TableName { get; set; }
    public TableNature? TableNature { get; set; }
    public bool IsHistoryPreservingTable { get; set; }
    public string? ColumnHeader { get; set; }

    // ---- Cell content ---------------------------------------------------------------------
    public string? CurrentValuePreview { get; set; }
    public bool HasExcelFormula { get; set; }
    public string? FormulaHint { get; set; }

    // ---- Rule metadata (RuleRegistry) -----------------------------------------------------
    public bool HasRule { get; set; }
    public bool IsRequired { get; set; }
    public bool IsWarningOnly { get; set; }
    public bool IsStrictlyCalculated { get; set; }
    public bool IsReadOnly { get; set; }
    public bool HasDropdown { get; set; }
    public bool IsMultiSelect { get; set; }

    public bool HasDefaultValue { get; set; }
    public DefaultKind DefaultKind { get; set; } = DefaultKind.None;
    public string? DefaultLiteral { get; set; }

    public List<string> SyncTargets { get; set; } = new();
    public List<string> SyncedFromSources { get; set; } = new();

    /// <summary>Sibling columns referenced by this column's C# calculation expression.</summary>
    public List<string> CalculatedFromSources { get; set; } = new();

    /// <summary>Sibling columns that react (sync, count, dropdown, guard, date rule, calculation) when this column changes.</summary>
    public List<string> AffectedColumns { get; set; } = new();

    /// <summary>"Table (Col1, Col2)" descriptions of the child collections this column rolls up from.</summary>
    public List<string> RollupDescriptions { get; set; } = new();

    public string? DateMustBeGreaterThanHeader { get; set; }
    public bool DateMustBeGreaterThanAllowEqual { get; set; }

    public ColumnFormatKind FormatKind { get; set; } = ColumnFormatKind.None;
    public string? ExcelNumberFormat { get; set; }
    public bool IsDateColumn { get; set; }
    public bool IsNumericColumn { get; set; }

    // ---- Native Excel Data Validation input message ---------------------------------------
    public string? ExcelValidationTitle { get; set; }
    public string? ExcelValidationMessage { get; set; }

    /// <summary>Plain-language description of the default rule in the ACTIVE language (computed on demand).</summary>
    public string? DefaultKindDescription => !HasDefaultValue ? null : DefaultKind switch
    {
        DefaultKind.Literal => Loc.Tip("guide.default.literal", DefaultLiteral ?? string.Empty),
        DefaultKind.Today => Loc.Tip("guide.default.today"),
        DefaultKind.DateOffsetDays or DefaultKind.DateOffsetMonths => Loc.Tip("guide.default.dateOffset"),
        DefaultKind.AutoIncrement => Loc.Tip("guide.default.autoIncrement"),
        _ => Loc.Tip("guide.default.generic")
    };

    public bool IsLocked => IsStrictlyCalculated || IsReadOnly;

    public static CellGuideSnapshot Empty() => new() { Zone = CellZone.None };
}
