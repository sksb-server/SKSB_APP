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
    ColumnEquals,       // True if ReferenceHeaders cell equals ExpectedValue
    DateGreaterThan,
    DateReached,        // True if ANY ReferenceHeaders date <= Today
    DateExpiredOrNumberReached // 3 ReferenceHeaders: [0]=date column, [1]=current numeric column, [2]=target numeric column; true if date < Today OR (target > 0 && current >= target)
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
/// AllowMultiple = true supports per-table rules via <see cref="TableName"/> (a null/blank TableName is the fallback).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
public sealed class ColumnRuleAttribute : Attribute
{
    /// <summary>
    /// One or more Defined Names (e.g. "VehicleTypeList") or "TableName[ColumnHeader]" sources.
    /// Multiple sources are stacked in sections (Section 0 at top, Section 1 below, etc.).
    /// </summary>
    public string[] Sources { get; }

    /// <summary>When set, this rule only applies to this Excel table (case-insensitive); null/blank is the fallback.</summary>
    public string? TableName { get; set; }

    public bool IsRequired { get; set; } = true;
    public bool IsWarningOnly { get; set; } = false;
    public bool IsMultiSelect { get; set; } = false;

    /// <summary>
    /// When true, users cannot modify this column in Excel (blocked via '=FALSE' validation and auto-undone on SheetChange),
    /// while C# COM writes continue to work normally.
    /// </summary>
    public bool IsReadOnly { get; set; } = false;

    /// <summary>
    /// Multi-select only: header of an integer column that receives the count of distinct selected tokens
    /// whenever this column is edited (e.g. Workers -> Worker Presence).
    /// </summary>
    public string? CountTargetHeader { get; set; }
    public bool AllowNonNegativeZero { get; set; } = true;
    public TransformKind Transform { get; set; } = TransformKind.None;

    /// <summary>
    /// When > 0 (e.g. 0.5), numeric values entered into this column are rounded down to the nearest multiple
    /// of this step: Math.Floor(val / step) * step (so 1.8 -> 1.5, 1.4 -> 1.0).
    /// </summary>
    public double RoundDownStep { get; set; } = 0.0;

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

    /// <summary>When set, the sync only applies to this Excel table (case-insensitive); null/blank applies to every table.</summary>
    public string? TableName { get; set; }

    /// <summary>When true, the target cell is only filled while it is still blank (never overwrites a user value).</summary>
    public bool OnlyWhenTargetBlank { get; set; } = false;

    public ColumnSyncAttribute(string targetHeader, string lookupSource, string resultSource)
    {
        TargetHeader = targetHeader;
        LookupSource = lookupSource;
        ResultSource = resultSource;
    }
}

/// <summary>
/// Defines how an empty cell in an active row is automatically defaulted.
/// AllowMultiple = true supports per-table defaults via <see cref="TableName"/> (a null/blank TableName is the fallback).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
public sealed class ColumnDefaultAttribute : Attribute
{
    public DefaultKind Kind { get; }
    public string? LiteralValue { get; set; }

    /// <summary>When set, this default only applies to this Excel table (case-insensitive).</summary>
    public string? TableName { get; set; }
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
    public string? SkipIfCurrentEquals { get; set; }

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

/// <summary>
/// Enforces that a date column must be strictly greater (later) than a referenced date column.
/// Set <see cref="AllowEqual"/> = true to enforce <c>&gt;=</c> (on or after) instead of strictly <c>&gt;</c>.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnDateGreaterThanAttribute : Attribute
{
    public string ReferenceHeader { get; }

    /// <summary>When true, the date may equal the reference date (on or after, <c>&gt;=</c>).</summary>
    public bool AllowEqual { get; set; } = false;

    public ColumnDateGreaterThanAttribute(string referenceHeader)
    {
        ReferenceHeader = referenceHeader;
    }
}
// =========================================================================
// CALCULATED COLUMNS, ROLLUPS & SHEET PARAMETERS
// =========================================================================

public enum CalculationMode
{
    Always,         // Strict C#-evaluated column: always recomputes and overwrites user edits
    DefaultOnly     // Soft default: computes only when cell is blank or 0; user can override
}

/// <summary>
/// Declares a column whose value is computed in C# from other columns of the same row
/// (replaces an Excel cell formula).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnCalculateAttribute : Attribute
{
    public string Expression { get; }
    public CalculationMode Mode { get; set; } = CalculationMode.Always;
    public int DecimalPlaces { get; set; } = 2;

    public ColumnCalculateAttribute(string expression, CalculationMode mode = CalculationMode.Always)
    {
        Expression = expression;
        Mode = mode;
    }
}

public enum RollupKind
{
    Sum,
    Count,
    DistinctCount,
    Min,
    Max,
    Average,
    DistinctJoin,   // Joins distinct non-empty string values with ", " (replaces TEXTJOIN(UNIQUE(FILTER(...))))
    PriorMax        // Maximum value from prior rows matching the same key (replaces Odometer Before LET/FILTER)
}

/// <summary>
/// Declares either:
/// (A) an in-memory child collection rollup on a master record (e.g. Maintenances, Insurances), OR
/// (B) a cross-table workbook rollup matching ForeignKeyHeader == LocalKeyHeader (e.g. WorkshopsTable &lt;- MaintenancesTable).
/// AllowMultiple = true sums multiple rollups on the same property (e.g. RoadTaxes + Inspections).
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = true, Inherited = true)]
public sealed class ColumnRollupAttribute : Attribute
{
    /// <summary>Either a child collection property name on the master record OR a source Excel table name.</summary>
    public string CollectionOrTable { get; }
    public RollupKind Kind { get; }

    /// <summary>Source properties (A) or source column headers (B). Several are summed per item.</summary>
    public string[] SourceProperties { get; }

    /// <summary>Cross-table: column header in the source table matched against <see cref="LocalKeyHeader"/>.</summary>
    public string? ForeignKeyHeader { get; set; }

    /// <summary>Cross-table: column header on this table whose value is matched.</summary>
    public string? LocalKeyHeader { get; set; }

    /// <summary>Optional date property/header on the child/source table for date-scoped filtering (B2 anchor + C2 scope).</summary>
    public string? DateProperty { get; set; }

    /// <summary>Optional fallback key used when the computed value is 0 (e.g. "Odometer", "Fuel").</summary>
    public string? FallbackKey { get; set; }

    public CalculationMode Mode { get; set; } = CalculationMode.Always;

    public bool IsCrossTable => !string.IsNullOrWhiteSpace(ForeignKeyHeader);

    public ColumnRollupAttribute(string collectionOrTable, RollupKind kind, params string[] sourceProperties)
    {
        CollectionOrTable = collectionOrTable;
        Kind = kind;
        SourceProperties = sourceProperties ?? Array.Empty<string>();
    }
}

/// <summary>Declares a named sheet parameter cell (e.g. @WorkerHourlyRate) usable inside [ColumnCalculate] expressions.</summary>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = true)]
public sealed class SheetParameterAttribute : Attribute
{
    public string Name { get; }
    public string CellAddress { get; }
    public double DefaultValue { get; }
    public string? FallbackSheet { get; set; }
    public string? FallbackAddress { get; set; }

    public SheetParameterAttribute(string name, string cellAddress, double defaultValue = 0.0)
    {
        Name = name;
        CellAddress = cellAddress;
        DefaultValue = defaultValue;
    }
}


// =========================================================================
// COLUMN DISPLAY FORMAT (unit suffixes, decimals, zero / empty display)
// =========================================================================

public enum ColumnFormatKind
{
    None,
    General,
    Text,               // "@"
    Integer,            // #,##0
    Id,                 // 0 (blank when <= 0)
    Decimal,            // #,##0.00
    Currency,           // "RM "#,##0.00
    PricePerLiter,      // "RM "#,##0.00"/L"
    Liters,             // #,##0.00 "L"
    Kilometers,         // #,##0.0 "km"
    Hours,              // #,##0.0 "hrs"
    Shifts,             // #,##0.# "shifts"
    Years,              // #,##0 "yrs"
    Days,               // #,##0 "days"
    OdometerDynamic,    // Per row: hours (machinery) or km (road vehicles)
    CurrencyPerHour,    // "RM "#,##0.00"/hr"
    WageDynamic,        // Per row: "RM "#,##0.00"/mth" (monthly) or "/day" (daily)
    Date,              // dd/mm/yyyy
    Time                // hh:mm
}

public enum ZeroBehavior
{
    Dash,               // Displays an em-dash when 0 / empty (3rd section of the Excel NumberFormat; dates write the dash itself)
    Blank               // Truly blank cell when 0 / empty (IDs, optional text)
}

/// <summary>
/// Declares the Excel display format, unit suffix and zero / empty behaviour of a column.
/// Columns without this attribute get a sensible default inferred from property type and header name.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnFormatAttribute : Attribute
{
    public ColumnFormatKind Kind { get; }
    public ZeroBehavior Zero { get; set; } = ZeroBehavior.Dash;

    public ColumnFormatAttribute(ColumnFormatKind kind, ZeroBehavior zero = ZeroBehavior.Dash)
    {
        Kind = kind;
        Zero = zero;
    }
}
