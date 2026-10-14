namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;

public record ScheduleRecord
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public List<DailyProjectSchedule> Projects { get; set; } = new();
    public List<DailyWorkerAttendance> WorkerAttendance { get; set; } = new();
    public List<WorkerPreference> Preferences { get; set; } = new();

    // Aggregated salary snapshot for the daily batch
    public DailySalary? DailySalary { get; set; }
}

public enum Shifts
{
    Morning,
    Afternoon,
    Night,
    None
}

public record DailyProjectSchedule
{
    public string ProjectID { get; set; } = string.Empty;
    public Shifts Shift { get; set; } = Shifts.Morning;
    public List<string> Workers { get; set; } = new();   // Worker IDs
    public List<string> Vehicles { get; set; } = new();  // Vehicle Plates
    public string Remarks { get; set; } = string.Empty;
}

public enum AttendanceStatus
{
    Presence,
    Late,
    NightShift,
    Unavailable,
    Absence
}

public enum BiometricStatus
{
    Thumbprint,
    Face,
    None
}

public record DailyWorkerAttendance
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public long WorkerID { get; set; } = 0;
    public TimeOnly ClockInTime { get; set; } = TimeOnly.MinValue;
    public TimeOnly ClockOutTime { get; set; } = TimeOnly.MaxValue;
    public BiometricStatus BiometricStatus { get; set; } = BiometricStatus.None;
    public AttendanceStatus ManualOverride { get; set; } = AttendanceStatus.Absence;
    public string OverrideReason { get; set; } = string.Empty;
    public AttendanceStatus AttendanceStatus { get; set; } = AttendanceStatus.Absence;
    public bool IsAllocated { get; set; } = false;

    // Snapshot base rate locked on shift date for salary calculations
    public decimal AppliedBaseRate { get; set; } = 0.00m;
}

public enum Days
{
    Monday,
    Tuesday,
    Wednesday,
    Thursday, // Corrected spelling
    Friday,
    Saturday,
    Sunday,
    None
}

public record WorkerPreference
{
    public long WorkerID { get; set; } = 0;
    public string Team { get; set; } = string.Empty;
    public string PrimarySkill { get; set; } = string.Empty;
    public List<string> OtherSkills { get; set; } = new();
    public List<string> BlackListedSkills { get; set; } = new();
    public VehicleType PreferredVehicleType { get; set; } = VehicleType.None;
    public string DedicatedVehiclePlate { get; set; } = string.Empty;
    public Shifts PreferredShift { get; set; } = Shifts.None;
    public int MaxDaysAWeek { get; set; } = 6;
    public List<Days> RestDayPreference { get; set; } = new();
}

public record DailySalary
{
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public decimal TotalBaseSalary { get; set; } = 0.00m;
    public decimal TotalOTSalary { get; set; } = 0.00m;
    public decimal TotalEarned { get; set; } = 0.00m;
}