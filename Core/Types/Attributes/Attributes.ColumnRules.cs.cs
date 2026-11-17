
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
    Date,               // dd/mm/yyyy
    Time                // hh:mm
}

public enum ZeroBehavior
{
    Dash,               // Displays an em-dash when 0 / empty (3rd section of the Excel NumberFormat; dates write the dash itself)
    Blank,              // Truly blank cell when 0 / empty (IDs, optional text)
    LiteralZero         // Displays a literal 0
}

/// <summary>
/// Declares the Excel display format, unit suffix, decimal places and zero / empty behaviour of a column.
/// Columns without this attribute get a sensible default inferred from property type and header name.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ColumnFormatAttribute : Attribute
{
    public ColumnFormatKind Kind { get; }
    public string? CustomUnit { get; set; }
    public int DecimalPlaces { get; set; } = -1;
    public ZeroBehavior Zero { get; set; } = ZeroBehavior.Dash;

    public ColumnFormatAttribute(
        ColumnFormatKind kind,
        int decimalPlaces = -1,
        ZeroBehavior zero = ZeroBehavior.Dash,
        string? customUnit = null)
    {
        Kind = kind;
        DecimalPlaces = decimalPlaces;
        Zero = zero;
        CustomUnit = customUnit;
    }
}
