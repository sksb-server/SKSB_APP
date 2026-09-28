namespace SKSB_App.Core.ServerManager;

using System;

public record StaffMasterDto
{
    public long StaffId { get; init; }
    public string FamilyName { get; init; } = string.Empty;
    public string MiddleName { get; init; } = string.Empty;
    public string GivenName { get; init; } = string.Empty;
    public string FullName { get; init; } = string.Empty;
    public string Nric { get; init; } = string.Empty;
    public string Department { get; init; } = string.Empty;
    public string PayType { get; init; } = string.Empty;
    public decimal BaseRate { get; init; }
    public decimal OvertimeRate { get; init; }
    public string ActiveStatus { get; init; } = string.Empty;
}

public record StaffPreferenceDto
{
    public long StaffId { get; init; }
    public string PrimaryTeam { get; init; } = string.Empty;
    public string SecondaryTeam { get; init; } = string.Empty;
    public string PrimarySkill { get; init; } = string.Empty;
    public string SecondarySkill { get; init; } = string.Empty;
    public string BlacklistTask { get; init; } = string.Empty;
    public string PreferredShift { get; init; } = string.Empty;
    public int MaxDaysPerWeek { get; init; }
    public string PreferredVehicleType { get; init; } = string.Empty;
    public long? DedicatedVehicleId { get; init; }
}

public record StaffLicenseDto
{
    public long StaffId { get; init; }
    public string LicenseType { get; init; } = string.Empty;
    public DateTime ExpiryDate { get; init; }
    public string Status { get; init; } = string.Empty;
}

public record StaffLeaveDto
{
    public long LeaveId { get; init; }
    public long StaffId { get; init; }
    public string LeaveType { get; init; } = string.Empty;
    public DateTime DateStart { get; init; }
    public DateTime DateEnd { get; init; }
    public string SessionType { get; init; } = string.Empty;
    public double TotalDays { get; init; }
    public string ApprovalStatus { get; init; } = string.Empty;
}