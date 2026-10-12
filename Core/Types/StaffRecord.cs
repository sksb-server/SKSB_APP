namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;

public record StaffRecord
{
    public long StaffID { get; set; } = 0;
    public StaffDetails StaffDetails { get; set; } = new();
    public List<VehicleLicense> VehicleLicenses { get; set; } = new();
    public List<SkillLicense> SkillLicenses { get; set; } = new();

    /// <summary>
    /// Evaluates aggregate staff readiness. If the employee is resigned,
    /// or if their contract is suspended, they cannot be deployed.
    /// </summary>
    public ContractStatus OverallStatus => StaffDetails.ContractStatus;
}

public enum Department
{
    Admin,
    Worker,
    Management,
    Operations,
    None
}

public enum PayType
{
    Hourly,
    Monthly,
    Daily,
    None
}

public enum ContractStatus
{
    Active,
    Suspended,
    Resigned
}

public record StaffDetails
{
    public long StaffID { get; set; } = 0;
    public string FamilyName { get; set; } = string.Empty;
    public string MiddleName { get; set; } = string.Empty;
    public string GivenName { get; set; } = string.Empty;

    /// <summary>
    /// Evaluates to "FamilyName GivenName". Can be set explicitly or generated automatically.
    /// </summary>
    public string Name
    {
        get => string.IsNullOrWhiteSpace(_name)
            ? $"{FamilyName} {GivenName}".Trim()
            : _name;
        set => _name = value;
    }
    private string _name = string.Empty;

    public string NRIC { get; set; } = string.Empty;
    public string Address { get; set; } = string.Empty;
    public string Email { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public Department Department { get; set; } = Department.Worker;
    public PayType PayType { get; set; } = PayType.Hourly;
    public decimal BaseRate { get; set; } = 11.00m;

    public ContractStatus ContractStatus
    {
        get
        {
            if (_contractStatus == ContractStatus.Suspended)
                return ContractStatus.Suspended;

            DateOnly today = DateOnly.FromDateTime(DateTime.Today);
            if (ResignDate.HasValue && today >= ResignDate.Value)
                return ContractStatus.Resigned;

            return _contractStatus;
        }
        set => _contractStatus = value;
    }
    private ContractStatus _contractStatus = ContractStatus.Active;

    // Financial & Statutory Identity
    public string BankName { get; set; } = string.Empty;
    public string BankNumber { get; set; } = string.Empty;
    public string EPFNumber { get; set; } = string.Empty;
    public string SocsoNumber { get; set; } = string.Empty;
    public string TINNumber { get; set; } = string.Empty;

    // Employment Dates
    public DateOnly JoinDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? ConfirmDate { get; set; } = null;
    public DateOnly? ResignDate { get; set; } = null;
}

public enum LicenseClass
{
    B2,         // Motorcycle <= 250cc
    B,          // Motorcycle > 500cc (Full B)
    D,          // Car / Light Vehicle (BTM <= 3500kg)
    DA,         // Automatic Car
    E,          // Heavy Motor Car (Rigid BTM > 7500kg)
    GDL_D,      // Goods Vehicle Driver's Licence (Van / 4WD / Light Pickup)
    GDL_E,      // Goods Vehicle Driver's Licence (Rigid Lorry > 7.5T)
    GDL_E_Full, // Articulated / Trailer / Prime Mover
    Forklift,   // Internal / Industrial Certificate
    None
}

public enum LicenseStatus
{
    Valid,
    Invalid,
    Suspended
}

public record VehicleLicense
{
    public long LicenseID { get; set; } = 0;
    public long StaffID { get; set; } = 0;
    public LicenseClass LicenseClass { get; set; } = LicenseClass.None;

    public DateOnly RenewDate
    {
        get => _renewDate;
        set => _renewDate = value;
    }
    private DateOnly _renewDate = DateOnly.FromDateTime(DateTime.Today);

    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? _renewDate.AddYears(1);
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

    public LicenseStatus LicenseStatus
    {
        get
        {
            if (_licenseStatus == LicenseStatus.Suspended)
                return LicenseStatus.Suspended;

            return DateOnly.FromDateTime(DateTime.Today) > ExpiryDate
                ? LicenseStatus.Invalid
                : _licenseStatus;
        }
        set => _licenseStatus = value;
    }
    private LicenseStatus _licenseStatus = LicenseStatus.Valid;
}

public record SkillLicense
{
    public long LicenseID { get; set; } = 0;
    public long StaffID { get; set; } = 0;
    public string LicenseType { get; set; } = string.Empty; // e.g., "CIDB Green Card", "SPAN Welder", "Confined Space AESP"

    public DateOnly RenewDate
    {
        get => _renewDate;
        set => _renewDate = value;
    }
    private DateOnly _renewDate = DateOnly.FromDateTime(DateTime.Today);

    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? _renewDate.AddYears(1);
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

    public LicenseStatus LicenseStatus
    {
        get
        {
            if (_licenseStatus == LicenseStatus.Suspended)
                return LicenseStatus.Suspended;

            return DateOnly.FromDateTime(DateTime.Today) > ExpiryDate
                ? LicenseStatus.Invalid
                : _licenseStatus;
        }
        set => _licenseStatus = value;
    }
    private LicenseStatus _licenseStatus = LicenseStatus.Valid;
}