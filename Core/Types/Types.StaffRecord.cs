namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. STAFF DOMAIN ENUMS (Generates Lookups Named Ranges)
// =========================================================================

public enum Department
{
    Office,
    Worker,
    Admin,
    None
}

public enum ContractStatus
{
    Active,
    Suspended,
    Resigned,
    Terminated,
    On_Leave,
    None
}

public enum VehicleLicenseClass
{
    None,
    B2,
    B,
    D,
    DA,
    E,
    E1,
    E2,
    GDL_D,
    GDL_E,
    GDL_E_Full,
    GDL,
    PSV,
    Forklift,
    H,
    I
}

public enum VehicleLicenseStatus
{
    Valid,
    Invalid,
    Suspended
}

public enum SkillLicenseStatus
{
    Valid,
    Invalid,
    Suspended
}

public enum SalaryBasis
{
    Monthly,
    Daily,
    None
}

// =========================================================================
// 2. ROOT STAFF RECORD & MASTER TABLE ("AllStaffsTable")
// =========================================================================

[ExcelTable("AllStaffsTable")]
public record StaffRecord
{
    [DataList("StaffIDList")]
    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    public long StaffID { get; set; } = 0;

    [DataDetail]
    public StaffDetails StaffDetails { get; set; } = new();

    // Child Collections
    public List<VehicleLicense> VehicleLicenses { get; set; } = new();
    public List<SkillLicense> SkillLicenses { get; set; } = new();
    public List<StaffSalary> Salaries { get; set; } = new();
    public List<StaffLeave> Leaves { get; set; } = new();

    /// <summary>staff_master.updated_at as last read from the server (optimistic concurrency token; never in Excel/JSON).</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }

    /// <summary>
    /// Salary effective on <paramref name="asOf"/>: the latest positive-rate row on/before that date; when every row
    /// starts after it, the earliest positive-rate row (so early attendance still resolves a rate).
    /// </summary>
    public StaffSalary? GetEffectiveSalary(DateOnly asOf)
    {
        if (Salaries == null || Salaries.Count == 0) return null;
        var positive = Salaries.Where(s => s != null && s.SalaryRate > 0m).ToList();
        if (positive.Count == 0) return null;

        return positive
                   .Where(s => s.EffectiveDate <= asOf)
                   .OrderByDescending(s => s.EffectiveDate)
                   .ThenByDescending(s => s.SalaryID)
                   .FirstOrDefault()
               ?? positive
                   .OrderBy(s => s.EffectiveDate)
                   .ThenBy(s => s.SalaryID)
                   .First();
    }

    /// <summary>Copies today's effective salary into the read-only "Current Salary" column.</summary>
    public void RefreshCurrentSalary()
    {
        StaffDetails ??= new StaffDetails();
        var salary = GetEffectiveSalary(DateOnly.FromDateTime(DateTime.Today));
        StaffDetails.CurrentSalary = salary?.SalaryRate;
        StaffDetails.CurrentSalaryBasis = salary?.SalaryBasis ?? SalaryBasis.None;
        ResolveSalaryBeforeBases();
    }

    /// <summary>
    /// Re-derives the display-only <see cref="StaffSalary.SalaryBeforeBasis"/> (not an Excel column and not in JSON)
    /// from the chronologically preceding row, so rows extracted from the sheet / offline snapshot keep it.
    /// </summary>
    public void ResolveSalaryBeforeBases()
    {
        if (Salaries == null || Salaries.Count == 0) return;
        StaffSalary? previous = null;
        foreach (var s in Salaries.Where(s => s != null).OrderBy(s => s.EffectiveDate).ThenBy(s => s.SalaryID))
        {
            if (!s.SalaryBefore.HasValue) s.SalaryBeforeBasis = SalaryBasis.None;
            else if (s.SalaryBeforeBasis == SalaryBasis.None) s.SalaryBeforeBasis = previous?.SalaryBasis ?? s.SalaryBasis;
            previous = s;
        }
    }
}

public record StaffDetails
{
    [DataList("StaffNameList")]
    [ExcelColumn("Staff Name", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true, Transform = TransformKind.Trim)]
    public string StaffName { get; set; } = string.Empty;

    [DataList("StaffShortNameList")]
    [ExcelColumn("Short Name")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true, IsReadOnly = true, Transform = TransformKind.ShortNameFirstLast)]
    public string StaffShortName
    {
        get
        {
            if (!string.IsNullOrWhiteSpace(_staffShortName)) return _staffShortName;

            if (string.IsNullOrWhiteSpace(StaffName)) return string.Empty;

            return SKSB_App.Core.RuntimeUI.RuleEngine.RuleEvaluator.ApplyTextTransform(StaffName, TransformKind.ShortNameFirstLast);
        }
        set => _staffShortName = value;
    }
    private string _staffShortName = string.Empty;

    [ExcelColumn("Department", isRequired: true)]
    [ColumnRule("DepartmentList", IsRequired = true, IsReadOnly = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Worker")]
    public string Department { get; set; } = "Worker";

    // Restored as long for backward compatibility with legacy importer/database mapping
    [ExcelColumn("IC Number")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public long NRIC { get; set; } = 0;

    [ExcelColumn("Passport Number")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string Passport { get; set; } = string.Empty;

    [ExcelColumn("Phone Number")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string Phone { get; set; } = string.Empty;

    [ExcelColumn("Join Date", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly JoinDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Resign Date")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnDateGreaterThan("Join Date", AllowEqual = true)]
    public DateOnly? ResignDate { get; set; } = null;

    [ExcelColumn("Contract Status", isRequired: true)]
    [ColumnRule("ContractStatusList", IsRequired = true, IsReadOnly = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Resigned", "Active", "Resign Date")]
    public ContractStatus ContractStatus
    {
        get => GetContractStatusOn(DateOnly.FromDateTime(DateTime.Today));
        set => _contractStatus = value;
    }
    private ContractStatus _contractStatus = ContractStatus.Active;

    public ContractStatus GetContractStatusOn(DateOnly targetDate)
    {
        if (ResignDate.HasValue && targetDate >= ResignDate.Value)
            return ContractStatus.Resigned;

        if (targetDate < JoinDate)
            return ContractStatus.Resigned;

        if (_contractStatus == ContractStatus.Suspended)
            return ContractStatus.Suspended;

        return ContractStatus.Active;
    }

    /// <summary>Today's effective salary (read-only; resolved from <see cref="StaffRecord.Salaries"/>).</summary>
    [ExcelColumn("Current Salary")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.WageDynamic)]
    public decimal? CurrentSalary { get; set; } = null;

    /// <summary>Basis of <see cref="CurrentSalary"/>; drives the per-row "/mth" vs "/day" number format.</summary>
    [JsonIgnore]
    public SalaryBasis CurrentSalaryBasis { get; set; } = SalaryBasis.None;
}

// =========================================================================
// 2b. STAFF SALARY HISTORY ("StaffSalaryTable"; synced from SQL Payroll, read-only)
// =========================================================================

[ExcelTable("StaffSalaryTable")]
public record StaffSalary
{
    [ExcelColumn("Date", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    public DateOnly EffectiveDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Salary ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public long SalaryID { get; set; } = 0;

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllStaffsTable[Staff ID]", IsRequired = true, IsReadOnly = true)]
    public long StaffID { get; set; } = 0;

    [ExcelColumn("Staff Name")]
    [JsonIgnore]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string StaffName { get; set; } = string.Empty;

    /// <summary>Salary in force immediately before this revision (null for the initial row).</summary>
    [ExcelColumn("Salary Before")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.WageDynamic)]
    public decimal? SalaryBefore { get; set; } = null;

    /// <summary>Basis of <see cref="SalaryBefore"/>; drives its per-row "/mth" vs "/day" number format.</summary>
    [JsonIgnore]
    public SalaryBasis SalaryBeforeBasis { get; set; } = SalaryBasis.None;

    [ExcelColumn("New Salary Basis", isRequired: true)]
    [ColumnRule("SalaryBasisList", IsRequired = true, IsReadOnly = true)]
    public SalaryBasis SalaryBasis { get; set; } = SalaryBasis.Monthly;

    [ExcelColumn("Salary Rate", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.WageDynamic)]
    public decimal SalaryRate { get; set; } = 0m;

    /// <summary>Display-only hourly rate: daily / 8, or monthly / 26 / 8 (rounded to sen).</summary>
    [ExcelColumn("Hourly Rate")]
    [JsonIgnore]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.CurrencyPerHour)]
    public decimal? HourlyRate
    {
        get => SalaryRate <= 0m
            ? null
            : (SalaryBasis == SalaryBasis.Daily
                ? decimal.Round(SalaryRate / 8m, 2, MidpointRounding.AwayFromZero)
                : decimal.Round(SalaryRate / 26m / 8m, 2, MidpointRounding.AwayFromZero));
        set { /* computed: the setter only lets the sheet mapper bind the column on extraction */ }
    }

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string Remarks { get; set; } = string.Empty;

    /// <summary>staff_salaries.updated_at as last read from the server (optimistic concurrency token).</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }

    /// <summary>Unrounded hourly rate used for salary calculations.</summary>
    public decimal GetHourlyRateExact() =>
        SalaryRate <= 0m ? 0m : (SalaryBasis == SalaryBasis.Daily ? SalaryRate / 8m : SalaryRate / 26m / 8m);
}

// =========================================================================
// 2c. STAFF LEAVES ("StaffLeavesTable"; synced from SQL Payroll LV_APP, read-only)
// =========================================================================

[ExcelTable("StaffLeavesTable")]
public record StaffLeave
{
    [ExcelColumn("Date", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    public DateOnly LeaveDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    /// <summary>SQL Payroll <c>LV_APP.DOCKEY</c> (stored as <c>staff_leaves.payroll_dockey</c>).</summary>
    [ExcelColumn("Leave ID", isRequired: true)]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public long LeaveID { get; set; }

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllStaffsTable[Staff ID]", IsRequired = true, IsReadOnly = true)]
    public long StaffID { get; set; }

    [ExcelColumn("Staff Name")]
    [JsonIgnore]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string StaffName { get; set; } = string.Empty;

    /// <summary><see cref="AttendanceStatus"/> name (e.g. <c>Annual_Leave</c>) mapped from the raw LV_APP.LEAVETYPE.</summary>
    [ExcelColumn("Leave Type", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true)]
    public string LeaveType { get; set; } = string.Empty;

    [ExcelColumn("Days", isRequired: true)]
    [ColumnRule(IsRequired = true, IsReadOnly = true, AllowNonNegativeZero = false)]
    [ColumnFormat(ColumnFormatKind.Decimal)]
    public decimal LeaveDay { get; set; } = 1.0m;

    [ExcelColumn("Description")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string Description { get; set; } = string.Empty;

    /// <summary>staff_leaves.updated_at as last read from the server.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

/// <summary>Shared staff-license columns: owning Staff ID / Name (two-way lookup), renew / expiry dates and cost.</summary>
public abstract record StaffLicenseBase
{
    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllStaffsTable[Staff ID]", IsRequired = true)]
    [ColumnSync("Staff Name", "AllStaffsTable[Staff ID]", "AllStaffsTable[Staff Name]")]
    public long StaffID { get; set; } = 0;

    [ExcelColumn("Staff Name")]
    [ColumnRule("AllStaffsTable[Staff Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "AllStaffsTable[Staff Name]", "AllStaffsTable[Staff ID]")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Renew Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly RenewDate
    {
        get => _renewDate;
        set => _renewDate = value;
    }
    private DateOnly _renewDate = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Renew Date", offset: 365)]
    [ColumnDateGreaterThan("Renew Date")]
    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? _renewDate.AddYears(1);
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

    [ExcelColumn("Cost")]
    public decimal Cost { get; set; } = 0.00m;
}

// =========================================================================
// 3. VEHICLE LICENSES TABLE ("VehicleLicensesTable")
// =========================================================================

[ExcelTable("VehicleLicensesTable")]
public record VehicleLicense : StaffLicenseBase
{
    [ExcelColumn("Vehicle License ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long VehicleLicenseID { get; set; } = 0;

    /// <summary>staff_vehicle_licenses.updated_at as last read from the server (optimistic concurrency token).</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }

    [ExcelColumn("Vehicle License Class", isRequired: true)]
    [ColumnRule("VehicleLicenseClassList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "D")]
    public VehicleLicenseClass VehicleLicenseClass { get; set; } = VehicleLicenseClass.D;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("VehicleLicenseStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Valid")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Invalid", "Valid", "Expiry Date")]
    public VehicleLicenseStatus VehicleLicenseStatus
    {
        get
        {
            if (_status == VehicleLicenseStatus.Suspended)
                return VehicleLicenseStatus.Suspended;

            return DateOnly.FromDateTime(DateTime.Today) > ExpiryDate
                ? VehicleLicenseStatus.Invalid
                : _status;
        }
        set => _status = value;
    }
    private VehicleLicenseStatus _status = VehicleLicenseStatus.Valid;
}

// =========================================================================
// 4. SKILL LICENSES & TYPES ("SkillLicensesTable", "SkillLicenseTypesTable")
// =========================================================================

[ExcelTable("SkillLicensesTable")]
public record SkillLicense : StaffLicenseBase
{
    [ExcelColumn("Skill License ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long SkillLicenseID { get; set; } = 0;

    /// <summary>staff_skill_licenses.updated_at as last read from the server (optimistic concurrency token).</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }

    [ExcelColumn("Skill License Type", isRequired: true)]
    [ColumnRule("SkillLicenseTypesTable[Skill License Types]", IsRequired = true)]
    public string SkillLicenseType { get; set; } = string.Empty;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("SkillLicenseStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Valid")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Invalid", "Valid", "Expiry Date")]
    public SkillLicenseStatus SkillLicenseStatus
    {
        get
        {
            if (_status == SkillLicenseStatus.Suspended)
                return SkillLicenseStatus.Suspended;

            return DateOnly.FromDateTime(DateTime.Today) > ExpiryDate
                ? SkillLicenseStatus.Invalid
                : _status;
        }
        set => _status = value;
    }
    private SkillLicenseStatus _status = SkillLicenseStatus.Valid;
}

[ExcelTable("SkillLicenseTypesTable")]
public record SkillLicenseType
{
    public long SkillLicenseTypeID { get; set; } = 0;

    [DataList("SkillLicenseTypeList")]
    [ExcelColumn("Skill License Type", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string SkillLicenseName { get; set; } = string.Empty;

    public string Description { get; set; } = string.Empty;

    [ExcelColumn("Display Icon")]
    [ColumnRule("DisplayIconList", IsRequired = false, IsWarningOnly = true)]
    public string DisplayIcon { get; set; } = string.Empty;
}

// =========================================================================
// 5. SKILL CATALOG (shared display icons + vehicle-license role skills)
// =========================================================================

/// <summary>
/// Shared skill/icon catalog: the emoji palette behind the <c>DisplayIconList</c> dropdown and the
/// built-in mapping from <see cref="VehicleLicenseClass"/> to a schedulable role skill.
/// </summary>
public static class SkillCatalog
{
    public const string GeneralWorker = "General Worker";
    public const string GeneralWorkerIcon = "🧰";

    public static readonly string[] DisplayIcons =
    {
        // Existing / core
        "🚗", "🚜", "🔧", "👷", "🧰", "🚚", "🚛", "🏗", "🏍", "🚐", "🚙", "⛏", "🛠",
        "⚡", "🔥", "🦺", "🚧", "🔨", "⚙", "🚰", "💧", "🚒", "🔗", "🛣", "📋",
        // Welding & metalwork
        "🧑‍🏭", "🥽", "🎇", "🔩", "🧲", "⛓",
        // Tools & hardware
        "🪛", "🪚", "🪜", "🪝", "🪣", "🔦", "🔌", "🔋", "📐", "📏",
        // Vehicles & traffic
        "🛻", "🚌", "⛽", "🚦", "🛑",
        // Safety & inspection
        "🧯", "🩹", "⚠", "🛡", "🔍", "✅", "🪪",
        // Site, utilities & materials
        "🧱", "🪨", "🪵", "🌲", "🌊", "📦", "🗑", "🏭", "📡",
        // Admin & general
        "📝", "📅", "🗺", "📞", "⭐"
    };

    public const string Driver = "Driver";
    public const string ExcavatorOperator = "Excavator Operator";
    public const string ForkliftOperator = "Forklift Operator";
    public const string MotorcycleRider = "Motorcycle Rider";

    public static string? RoleSkillFor(VehicleLicenseClass licenseClass) => licenseClass switch
    {
        VehicleLicenseClass.D or VehicleLicenseClass.DA or VehicleLicenseClass.E or VehicleLicenseClass.E1
            or VehicleLicenseClass.E2 or VehicleLicenseClass.GDL_D or VehicleLicenseClass.GDL_E
            or VehicleLicenseClass.GDL_E_Full or VehicleLicenseClass.GDL or VehicleLicenseClass.PSV => Driver,
        VehicleLicenseClass.H or VehicleLicenseClass.I => ExcavatorOperator,
        VehicleLicenseClass.Forklift => ForkliftOperator,
        VehicleLicenseClass.B or VehicleLicenseClass.B2 => MotorcycleRider,
        _ => null
    };

    public static IReadOnlyDictionary<string, string> VehicleLicenseRoleIcons { get; } =
        new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase)
        {
            [Driver] = "🚗",
            [ExcavatorOperator] = "🚜",
            [ForkliftOperator] = "🏗",
            [MotorcycleRider] = "🏍"
        };

    public static IReadOnlyList<string> VehicleLicenseRoleSkills { get; } =
        new[] { Driver, ExcavatorOperator, ForkliftOperator, MotorcycleRider };
}

// =========================================================================
// STAFF SUMMARY (read-only "StaffSummaryTable"; never saved)
// =========================================================================

[ExcelTable("StaffSummaryTable")]
public record StaffSummaryRow
{
    [ExcelColumn("Date")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public DateOnly Date { get; set; }

    [ExcelColumn("Time")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Text)]
    public string Time { get; set; } = string.Empty;

    [ExcelColumn("Salary Rate")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.CurrencyPerHour)]
    public decimal? SalaryRate { get; set; }

    [ExcelColumn("Status")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public AttendanceStatus Status { get; set; } = AttendanceStatus.Absence;

    [ExcelColumn("Projects")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string Projects { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Driven")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string VehicleDriven { get; set; } = string.Empty;

    [ExcelColumn("Regular Shift")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public decimal RegularShift { get; set; }

    [ExcelColumn("Additional Shifts")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public decimal AdditionalShifts { get; set; }

    [ExcelColumn("OT Hours")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal OTHours { get; set; }

    [ExcelColumn("Additional OT Hours")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal AdditionalOTHours { get; set; }

    [ExcelColumn("Flatten Hours")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal FlattenHours { get; set; }

    [ExcelColumn("Total Salary")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? TotalSalary { get; set; }

    [ExcelColumn("Claim Made")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? ClaimMade { get; set; }
}