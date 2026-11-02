namespace SKSB_App.Core.Types.Attributes;

using System;

/// <summary>
/// Binds a C# record/class to one or more Excel ListObject table names.
/// </summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = true)]
public sealed class ExcelTableAttribute : Attribute
{
    public string TableName { get; }

    public ExcelTableAttribute(string tableName)
    {
        TableName = tableName;
    }
}

public enum DefaultKind
{
    None,
    Literal,            // Uses LiteralValue (e.g. "Active", "0")
    Today,              // DateOnly.FromDateTime(DateTime.Today)
    DateOffsetDays,     // ReferenceHeader date + Offset days (Fallback: Today + Offset)
    DateOffsetMonths,   // ReferenceHeader date + Offset months (Fallback: Today + Offset)
    AutoIncrement       // Max(existing column integer values) + 1
}

public enum StateConditionKind
{
    None,
    DateExpired,        // True if ANY ReferenceHeaders date < Today
    ColumnNotEmpty,     // True if ANY ReferenceHeaders cell is not blank (e.g. Sold Date != "")
    ColumnEquals        // True if ReferenceHeaders cell equals ExpectedValue
}

public enum TransformKind
{
    None,
    Trim,
    UpperCase,
    ShortNameFirstLast, // First word + Last word
    ShortNameThreeWords // 3-word name formatting rule
}

/// <summary>
/// Primary column rule governing dropdown sources, required status, and text transformation.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnRuleAttribute : Attribute
{
    /// <summary>
    /// One or more Defined Names (e.g. "VehicleTypeList") or "TableName[ColumnHeader]" sources.
    /// Multiple sources are stacked in sections (Section 0 at top, Section 1 below, etc.).
    /// </summary>
    public string[] Sources { get; }
    public bool IsRequired { get; set; } = true;
    public bool IsWarningOnly { get; set; } = false;
    public bool IsMultiSelect { get; set; } = false;
    public bool AllowNonNegativeZero { get; set; } = true;
    public TransformKind Transform { get; set; } = TransformKind.None;

    public ColumnRuleAttribute(params string[] sources)
    {
        Sources = sources ?? Array.Empty<string>();
    }
}

/// <summary>
/// Defines automatic adjacent/target column lookup when this column is edited or selected.
/// AllowMultiple = true supports fallback chains (e.g. check Vehicles first, then Staff)
/// and multi-column fills (e.g. fill both Workshop Address and Workshop Contact).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
public sealed class ColumnSyncAttribute : Attribute
{
    public string TargetHeader { get; }
    public string LookupSource { get; }
    public string ResultSource { get; }

    public ColumnSyncAttribute(string targetHeader, string lookupSource, string resultSource)
    {
        TargetHeader = targetHeader;
        LookupSource = lookupSource;
        ResultSource = resultSource;
    }
}

/// <summary>
/// Defines how an empty cell in an active row is automatically defaulted.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnDefaultAttribute : Attribute
{
    public DefaultKind Kind { get; }
    public string? LiteralValue { get; set; }
    public string? ReferenceHeader { get; set; }
    public int Offset { get; set; }

    public ColumnDefaultAttribute(
        DefaultKind kind,
        string? literalValue = null,
        string? referenceHeader = null,
        int offset = 0)
    {
        Kind = kind;
        LiteralValue = literalValue;
        ReferenceHeader = referenceHeader;
        Offset = offset;
    }
}

/// <summary>
/// Self-healing state guard that forces or forbids specific dropdown values based on other columns in the same row.
/// Evaluated in order of Priority (lowest number first).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
public sealed class ColumnStateGuardAttribute : Attribute
{
    public int Priority { get; }
    public StateConditionKind Condition { get; }
    public string[] ReferenceHeaders { get; }
    public string ForcedValueWhenTrue { get; }
    public bool DisallowForcedValueWhenFalse { get; set; } = true;
    public string FallbackValueWhenFalse { get; set; } = "Active";
    public string? ExpectedValue { get; set; }

    public ColumnStateGuardAttribute(
        int priority,
        StateConditionKind condition,
        string forcedValueWhenTrue,
        string fallbackValueWhenFalse = "Active",
        params string[] referenceHeaders)
    {
        Priority = priority;
        Condition = condition;
        ForcedValueWhenTrue = forcedValueWhenTrue;
        FallbackValueWhenFalse = fallbackValueWhenFalse;
        ReferenceHeaders = referenceHeaders ?? Array.Empty<string>();
    }
}