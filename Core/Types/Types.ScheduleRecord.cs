namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. SCHEDULE & ATTENDANCE ENUMS (Auto-discovered into 'Lookups' sheet)
// =========================================================================

public enum Shifts
{
    Morning,
    Afternoon,
    Night,
    None
}

public enum AttendanceStatus
{
    Presence,
    Late,
    NightShift,
    Unavailable,
    Absence,
    None
}

public enum BiometricStatus
{
    Thumbprint,
    Face,
    None
}

// =========================================================================
// 2. ROOT SCHEDULE AGGREGATE
// =========================================================================

public record ScheduleRecord
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public List<DailyProjectSchedule> Projects { get; set; } = new();
    public List<DailyWorkerAttendance> WorkerAttendance { get; set; } = new();
    public List<WorkerPreference> Preferences { get; set; } = new();
}

// =========================================================================
// 3. DAILY PROJECT SCHEDULE TABLE ("DailyScheduleTable", "ProjectScheduleTable")
// =========================================================================

[ExcelTable("DailyScheduleTable")]
[ExcelTable("ProjectScheduleTable")]
public record DailyProjectSchedule
{
    [ExcelColumn("Project ID", isRequired: true)]
    [ColumnRule("ProjectIDList", IsRequired = true)]
    [ColumnSync("Project Name", "ProjectIDList", "ProjectNameList")]
    public string ProjectID { get; set; } = string.Empty;

    [ExcelColumn("Project Name")]
    [ColumnRule("ProjectNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Project ID", "ProjectNameList", "ProjectIDList")]
    public string ProjectName { get; set; } = string.Empty;

    [ExcelColumn("Shift", isRequired: true)]
    [ColumnRule("ShiftsList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Morning")]
    public Shifts Shift { get; set; } = Shifts.Morning;

    [ExcelColumn("Workers")]
    [ColumnRule("StaffNameList", IsRequired = false, IsMultiSelect = true)]
    public List<string> Workers { get; set; } = new();

    [ExcelColumn("Vehicles")]
    [ColumnRule("VehiclePlateList", IsRequired = false, IsMultiSelect = true)]
    public List<string> Vehicles { get; set; } = new();

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 4. DAILY WORKER ATTENDANCE TABLE ("DailyAttendanceTable", "TimeAttendanceTable")
// =========================================================================

[ExcelTable("DailyAttendanceTable")]
[ExcelTable("TimeAttendanceTable")]
public record DailyWorkerAttendance
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("StaffIDList", IsRequired = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffNameList")]
    public long StaffID { get; set; } = 0;

    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Project Done")]
    [ColumnRule("ProjectIDList", IsRequired = false, IsWarningOnly = true)]
    public string ProjectDone { get; set; } = string.Empty;

    [ExcelColumn("Checked In Time")]
    [ColumnRule(IsRequired = false)]
    public TimeOnly ClockInTime { get; set; } = TimeOnly.MinValue;

    [ExcelColumn("Checked Out Time")]
    [ColumnRule(IsRequired = false)]
    public TimeOnly ClockOutTime { get; set; } = TimeOnly.MaxValue;

    [ExcelColumn("Base Hours")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal BaseHours { get; set; } = 0.0m;

    [ExcelColumn("OT Hours")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal OTHours { get; set; } = 0.0m;

    [ExcelColumn("Additional OT Hours")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal AdditionalOTHours { get; set; } = 0.0m;

    [ExcelColumn("Biometric Status")]
    [ColumnRule("BiometricStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "None")]
    public BiometricStatus BiometricStatus { get; set; } = BiometricStatus.None;

    [ExcelColumn("Manual Override")]
    [ColumnRule("AttendanceStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Absence")]
    public AttendanceStatus ManualOverride { get; set; } = AttendanceStatus.Absence;

    [ExcelColumn("Override Reason")]
    [ColumnRule(IsRequired = false)]
    public string OverrideReason { get; set; } = string.Empty;

    [ExcelColumn("Attendance Status")]
    [ColumnRule("AttendanceStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Absence")]
    public AttendanceStatus AttendanceStatus { get; set; } = AttendanceStatus.Absence;

    public bool IsAllocated { get; set; } = false;
}

// =========================================================================
// 5. WORKER PREFERENCE TABLE ("WorkerPreferencesTable")
// =========================================================================

[ExcelTable("WorkerPreferencesTable")]
public record WorkerPreference
{
    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("StaffIDList", IsRequired = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffNameList")]
    public long StaffID { get; set; }

    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Preferred Vehicle")]
    [ColumnRule("VehiclePlateList", IsRequired = false, IsWarningOnly = true)]
    public string PreferredVehicle { get; set; } = string.Empty;

    [ExcelColumn("Preferred Project")]
    [ColumnRule("ProjectIDList", IsRequired = false, IsWarningOnly = true)]
    public string PreferredProject { get; set; } = string.Empty;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}