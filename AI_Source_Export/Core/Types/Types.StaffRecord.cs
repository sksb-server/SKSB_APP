namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. STAFF DOMAIN ENUMS (Generates Lookups Named Ranges)
// =========================================================================

public enum Department
{
    Worker,
    Driver,
    Supervisor,
    Foreman,
    Operator,
    Mechanic,
    Admin,
    Management,
    Office,
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

// =========================================================================
// 2. ROOT STAFF RECORD & MASTER TABLE ("AllStaffsTable")
// =========================================================================

[ExcelTable("AllStaffsTable")]
public record StaffRecord
{
    [DataList("StaffIDList")]
    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long StaffID { get; set; } = 0;

    [DataDetail]
    public StaffDetails StaffDetails { get; set; } = new();

    // Child Collections
    public List<VehicleLicense> VehicleLicenses { get; set; } = new();
    public List<SkillLicense> SkillLicenses { get; set; } = new();

    public ContractStatus OverallStatus => StaffDetails.ContractStatus;

    public ContractStatus EvaluateContractStatus(DateOnly asOfDate)
    {
        if (StaffDetails.ResignDate.HasValue && StaffDetails.ResignDate.Value <= asOfDate)
            return ContractStatus.Resigned;

        return StaffDetails.ContractStatus == ContractStatus.None
            ? ContractStatus.Active
            : StaffDetails.ContractStatus;
    }
}

public record StaffDetails
{
    // Restored on StaffDetails for backward compatibility with Importer and Sync commands
    public long StaffID { get; set; } = 0;

    [DataList("StaffNameList")]
    [ExcelColumn("Staff Name", isRequired: true)]
    [ColumnRule(IsRequired = true, Transform = TransformKind.Trim)]
    public string StaffName { get; set; } = string.Empty;

    [DataList("StaffShortNameList")]
    [ExcelColumn("Short Name")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true, Transform = TransformKind.ShortNameFirstLast)]
    public string StaffShortName { get; set; } = string.Empty;

    [ExcelColumn("Department", isRequired: true)]
    [ColumnRule("DepartmentList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Worker")]
    public string Department { get; set; } = "Worker";

    // Restored as long for backward compatibility with legacy importer/database mapping
    [ExcelColumn("IC Number")]
    [ColumnRule(IsRequired = false)]
    public long NRIC { get; set; } = 0;

    [ExcelColumn("Passport Number")]
    [ColumnRule(IsRequired = false)]
    public string Passport { get; set; } = string.Empty;

    [ExcelColumn("Phone Number")]
    [ColumnRule(IsRequired = false)]
    public string Phone { get; set; } = string.Empty;

    [ExcelColumn("Join Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly JoinDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Resign Date")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? ResignDate { get; set; } = null;

    [ExcelColumn("Contract Status", isRequired: true)]
    [ColumnRule("ContractStatusList", IsRequired = true)]
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

    public bool IsActiveOn(DateOnly targetDate) => GetContractStatusOn(targetDate) == ContractStatus.Active;
}

// =========================================================================
// 3. VEHICLE LICENSES TABLE ("VehicleLicensesTable")
// =========================================================================

[ExcelTable("VehicleLicensesTable")]
public record VehicleLicense
{
    [ExcelColumn("Vehicle License ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long VehicleLicenseID { get; set; } = 0;

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllStaffsTable[Staff ID]", IsRequired = true)]
    [ColumnSync("Staff Name", "AllStaffsTable[Staff ID]", "AllStaffsTable[Staff Name]")]
    public long StaffID { get; set; } = 0;

    [ExcelColumn("Staff Name")]
    [ColumnRule("AllStaffsTable[Staff Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "AllStaffsTable[Staff Name]", "AllStaffsTable[Staff ID]")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Vehicle License Class", isRequired: true)]
    [ColumnRule("VehicleLicenseClassList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "D")]
    public VehicleLicenseClass VehicleLicenseClass { get; set; } = VehicleLicenseClass.D;

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
    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? _renewDate.AddYears(1);
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

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

    [ExcelColumn("Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Cost { get; set; } = 0.00m;
}

// =========================================================================
// 4. SKILL LICENSES & TYPES ("SkillLicensesTable", "SkillLicenseTypesTable")
// =========================================================================

[ExcelTable("SkillLicensesTable")]
public record SkillLicense
{
    [ExcelColumn("Skill License ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long SkillLicenseID { get; set; } = 0;

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllStaffsTable[Staff ID]", IsRequired = true)]
    [ColumnSync("Staff Name", "AllStaffsTable[Staff ID]", "AllStaffsTable[Staff Name]")]
    public long StaffID { get; set; } = 0;

    [ExcelColumn("Staff Name")]
    [ColumnRule("AllStaffsTable[Staff Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "AllStaffsTable[Staff Name]", "AllStaffsTable[Staff ID]")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Skill License Type", isRequired: true)]
    [ColumnRule("SkillLicenseTypesTable[Skill License Type]", IsRequired = true)]
    public string SkillLicenseType { get; set; } = string.Empty;

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
    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? _renewDate.AddYears(1);
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

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

    [ExcelColumn("Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Cost { get; set; } = 0.00m;
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
}