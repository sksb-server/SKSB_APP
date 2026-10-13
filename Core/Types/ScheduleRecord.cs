namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;


public record ScheduleRecord
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Now);
    public List<DailyProjectSchedule> Projects { get; set; } = new List<DailyProjectSchedule>();
    public List<DailyWorkerAttendance> WorkerAttendance { get; set; } = new List<DailyWorkerAttendance>();

    /// <summary>
    /// Stored or evaluated aggregate summary of the project.
    /// </summary>
    public DailySalary DailySalary
    {
        get => _dailySalary ?? DailySalary.Create(ProjectDetails, DailyProjectSummary);
        set => _dailySalary = value;
    }
    private DailySalary? _dailySalary;
}

public record DailyProjectSchedule
{
    public List<string> Workers { get; set; } = new List<string>();
    public List<string> Vehicles { get; set; } = new List<string>();
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
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Now);
    public long StaffID { get; set; } = 0;
    public TimeOnly ClockInTime { get; set; } = TimeOnly.MinValue;
    public TimeOnly ClockOutTime { get; set; } = TimeOnly.MaxValue;
    public BiometricStatus BiometricStatus { get; set; } = BiometricStatus.None;
    public AttendanceStatus ManualOverride { get; set; } = AttendanceStatus.Absence;
    public string OverrideReason { get; set; } = string.Empty;
    public AttendanceStatus AttendanceStatus { get; set; } = AttendanceStatus.Absence;
}

public enum Shifts
{
    Morning,
    Afternoon,
    Night,
    None
}

public enum Days
{
    Monday,
    Tuesday,
    Wednesday,
    Thrusday,
    Friday,
    Saturday,
    Sunday,
    None
}

public record WorkerPreference
{
    public long StaffID { get; set; } = 0;
    public string Team { get; set; } = string.Empty;
    public string PrimarySkill { get; set; } = string.Empty;
    public List<string> OtherSkills { get; set; } = new List<string>();
    public List<string> BlackListedSkills { get; set; } = new List<string>();
    public VehicleType PreferredVehicleType { get; set; } = VehicleType.None;
    public string DedicatedVehiclePlate { get; set; } = string.Empty;
    public Shifts PreferredShift { get; set; } = Shifts.None;
    public int MaxDaysAWeek { get; set; } = 6;
    public List<Days> RestDayPreference = new List<Days>();
}