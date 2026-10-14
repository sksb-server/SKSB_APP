namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;


public record WorkerRecord
{
    public long WorkerID { get; set; } = 0;
    public WorkerDetails WorkerDetails { get; set; } = new();
    public List<VehicleLicense> VehicleLicenses { get; set; } = new();
    public List<SkillLicense> SkillLicenses { get; set; } = new();

    public ContractStatus OverallStatus => WorkerDetails.ContractStatus;
}


public enum ContractStatus
{
    Active,
    Suspended
}

public record WorkerDetails
{
    public long WorkerID { get; set; } = 0;
    public string WorkerName { get; set; } = string.Empty;
    public long NRIC { get; set; } = 0;
    public string Passport { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;
    public ContractStatus ContractStatus { get; set; } = ContractStatus.Active;
}

public enum LicenseClass
{
    B2,
    B,
    D,
    DA,
    E,
    GDL_D,
    GDL_E,
    GDL_E_Full,
    Forklift,
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
    public long WorkerID { get; set; } = 0;
    public LicenseClass LicenseClass { get; set; } = LicenseClass.None;

    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? RenewDate.AddYears(1);
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
    public long WorkerID { get; set; } = 0;
    public string LicenseType { get; set; } = string.Empty;

    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? RenewDate.AddYears(1);
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