namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;

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
    Suspended,
    Resigned
}

public record WorkerDetails
{
    public long WorkerID { get; set; } = 0;
    public string WorkerName { get; set; } = string.Empty;

    /// <summary>
    /// Derives the short name using the 3-word preservation rule:
    /// - 1 to 3 words: Keeps the entire name intact (preserves Chinese and 3-part names like "Tan Ah Kow").
    /// - 4 or more words: Takes the first 2 words plus the last word ("Muhammad Haziq Bin Rosli" -> "Muhammad Haziq Rosli").
    /// </summary>
    public string WorkerShortName
    {
        get
        {
            if (string.IsNullOrWhiteSpace(WorkerName)) return string.Empty;

            var tokens = WorkerName.Trim()
                                   .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);

            if (tokens.Length <= 3)
            {
                return string.Join(" ", tokens);
            }

            return $"{tokens[0]} {tokens[1]} {tokens[^1]}";
        }
    }

    public long NRIC { get; set; } = 0;
    public string Passport { get; set; } = string.Empty;
    public string Phone { get; set; } = string.Empty;

    public DateOnly JoinDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? ResignDate { get; set; } = null;

    /// <summary>
    /// Dynamic contract status evaluated against today's date.
    /// If today is on or past ResignDate, automatically evaluates to Resigned.
    /// </summary>
    public ContractStatus ContractStatus
    {
        get => GetContractStatusOn(DateOnly.FromDateTime(DateTime.Today));
        set => _contractStatus = value;
    }
    private ContractStatus _contractStatus = ContractStatus.Active;

    /// <summary>
    /// Evaluates what this worker's contract status was (or will be) on a specific target date.
    /// </summary>
    public ContractStatus GetContractStatusOn(DateOnly targetDate)
    {
        // 1. Resignation check takes highest priority
        if (ResignDate.HasValue && targetDate >= ResignDate.Value)
            return ContractStatus.Resigned;

        // 2. Pre-employment check
        if (targetDate < JoinDate)
            return ContractStatus.Resigned;

        // 3. Manual administrative suspension override
        if (_contractStatus == ContractStatus.Suspended)
            return ContractStatus.Suspended;

        return ContractStatus.Active;
    }

    public bool IsActiveOn(DateOnly targetDate) => GetContractStatusOn(targetDate) == ContractStatus.Active;
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
    public long WorkerID { get; set; } = 0;
    public string LicenseType { get; set; } = string.Empty;

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