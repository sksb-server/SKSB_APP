namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
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
    Annual_Leave,
    Medical_Leave,
    Unpaid_Leave,
    Hospitalization_Leave,
    Maternity_Leave,
    Time_Off,
    On_Leave,
    Unavailable,
    Absence,
    None
}

/// <summary>Maps SQL Payroll <c>LV_APP.LEAVETYPE</c> codes onto the leave values of <see cref="AttendanceStatus"/>.</summary>
public static class PayrollLeaveMapper
{
    /// <summary>Alias of <see cref="MapLeaveType"/>.</summary>
    public static AttendanceStatus MapStatus(string? leaveType) => MapLeaveType(leaveType);

    /// <summary>
    /// Accepts both raw SQL Payroll codes ("AL", "MC", "UL", ...) and <see cref="AttendanceStatus"/> leave names
    /// ("Annual_Leave", "Medical_Leave", ...) as stored by Staff Manager in <c>staff_leaves</c>.
    /// </summary>
    public static AttendanceStatus MapLeaveType(string? leaveType)
    {
        string raw = (leaveType ?? string.Empty).Trim();
        if (raw.Length > 0 && !char.IsDigit(raw[0]) && raw[0] != '-' &&
            Enum.TryParse<AttendanceStatus>(raw, ignoreCase: true, out var named) && IsLeave(named))
            return named;

        string t = raw.ToUpperInvariant();
        return t switch
        {
            "" => AttendanceStatus.None,
            "AL" => AttendanceStatus.Annual_Leave,
            "MC" => AttendanceStatus.Medical_Leave,
            "UL" or "UL (X)" => AttendanceStatus.Unpaid_Leave,
            "HOSPITALIZATION LEAV" or "HL" => AttendanceStatus.Hospitalization_Leave,
            "BERSALIN" => AttendanceStatus.Maternity_Leave,
            "T/O" => AttendanceStatus.Time_Off,
            _ => AttendanceStatus.On_Leave
        };
    }

    /// <summary>True for every leave status (un-punched staff on approved leave, never "Absence").</summary>
    public static bool IsLeave(AttendanceStatus status) => status is
        AttendanceStatus.Annual_Leave or AttendanceStatus.Medical_Leave or AttendanceStatus.Unpaid_Leave or
        AttendanceStatus.Hospitalization_Leave or AttendanceStatus.Maternity_Leave or AttendanceStatus.Time_Off or
        AttendanceStatus.On_Leave;
}

/// <summary>One staff leave day mirrored from SQL Payroll <c>LV_APP</c> (cached in PostgreSQL <c>staff_leaves</c>).</summary>
public record StaffLeaveEntry
{
    public long DocKey { get; set; }
    public long StaffID { get; set; }
    public DateOnly LeaveDate { get; set; }
    public string LeaveType { get; set; } = string.Empty;
    public decimal LeaveDay { get; set; } = 1m;
    public string Description { get; set; } = string.Empty;
    public AttendanceStatus MappedStatus { get; set; } = AttendanceStatus.On_Leave;
}

public enum AttendanceEntryType
{
    CheckIn,
    CheckOut
}

public enum AttendanceSource
{
    Thumbprint,
    Manual,
    SKSB_App,
    WhatsApp,
    Admin
}

public enum Days
{
    None,
    Monday,
    Tuesday,
    Wednesday,
    Thursday,
    Friday,
    Saturday,
    Sunday
}

// =========================================================================
// 2. ROOT SCHEDULE AGGREGATE (one working date + the global reference tables)
// =========================================================================

public record ScheduleRecord
{
    /// <summary>The working date the date-scoped collections below belong to.</summary>
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    /// <summary>
    /// Day Start Time for <see cref="DailyScheduleDate"/> (Daily Attendances C2). Default 06:00. Working-day boundary:
    /// punches before this time belong to the previous working day (night shifts).
    /// </summary>
    public TimeOnly DayStartTime { get; set; } = new(6, 0);

    /// <summary>Day Start Time for <see cref="DailyScheduleDate"/> + 1 (end boundary of this working day). Default 06:00.</summary>
    public TimeOnly NextDayStartTime { get; set; } = new(6, 0);

    /// <summary>Late Checked In Cut Off Time for <see cref="DailyScheduleDate"/> (Daily Attendances D2). Default 10:30.</summary>
    public TimeOnly LateCutOffTime { get; set; } = new(10, 30);

    /// <summary>OT Start Time for <see cref="DailyScheduleDate"/> (Daily Attendances E2). Default 18:00.</summary>
    public TimeOnly OTStartTime { get; set; } = new(18, 0);

    /// <summary>Day / Night dispatch board slots (8 per band) for <see cref="DailyScheduleDate"/>.</summary>
    public List<DailyProjectSchedule> Projects { get; set; } = new();

    /// <summary>
    /// Auto-scheduled project IDs the user removed from the board for <see cref="DailyScheduleDate"/> ("✕ Remove Slot");
    /// they are not auto-populated again for this date until re-added.
    /// </summary>
    public List<string> DismissedProjects { get; set; } = new();

    /// <summary>Per-worker rollup for <see cref="DailyScheduleDate"/>.</summary>
    public List<DailyWorkerAttendance> WorkerAttendance { get; set; } = new();

    /// <summary>Append-only punch log for <see cref="DailyScheduleDate"/>.</summary>
    public List<AttendanceEntry> AttendanceEntries { get; set; } = new();

    /// <summary>
    /// Save-time delta of <see cref="AttendanceEntries"/>. When non-null, <c>ServerDispatcher.PersistSchedule</c> only
    /// upserts these new/modified entries; when null (e.g. an offline pending upload), it upserts all
    /// <see cref="AttendanceEntries"/>.
    /// </summary>
    [JsonIgnore]
    public List<AttendanceEntry>? ModifiedAttendanceEntries { get; set; }

    /// <summary>
    /// Concurrency token for the date-scoped board: MAX(<c>updated_at</c>) across <c>daily_schedule_settings</c>,
    /// <c>daily_project_schedules</c> and <c>daily_attendances</c> for <see cref="DailyScheduleDate"/> (null = no rows).
    /// </summary>
    [JsonIgnore]
    public DateTime? BoardServerUpdatedAt { get; set; }

    /// <summary>Public holiday flag for <see cref="DailyScheduleDate"/> (SQL Payroll LV_CALENDAR_DAYS; read-only).</summary>
    public bool IsHoliday { get; set; }

    /// <summary>SQL Payroll description of the public holiday (empty when not a holiday).</summary>
    public string HolidayDescription { get; set; } = string.Empty;

    /// <summary>Approved leave days on <see cref="DailyScheduleDate"/> (SQL Payroll LV_APP; read-only).</summary>
    public List<StaffLeaveEntry> Leaves { get; set; } = new();

    public List<WorkerPreference> Preferences { get; set; } = new();
    public List<TeamRecord> Teams { get; set; } = new();
    public List<ScheduleSkillRecord> Skills { get; set; } = new();
    public List<ProjectSkillRequirement> ProjectSkillRequirements { get; set; } = new();
}

// =========================================================================
// 3. ATTENDANCE ENTRIES ("AttendanceEntriesTable") - APPEND-ONLY EVENT LOG
// =========================================================================

[ExcelTable("AttendanceEntriesTable")]
public record AttendanceEntry
{
    /// <summary>Database surrogate key; never shown in Excel.</summary>
    public long EntryId { get; set; }

    [ExcelColumn("Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Time", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnFormat(ColumnFormatKind.Time, ZeroBehavior.Blank)]
    public TimeOnly Time { get; set; } = TimeOnly.MinValue;

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("StaffIDList", IsRequired = true, IsWarningOnly = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffNameList")]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long StaffID { get; set; }

    [JsonIgnore]
    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Project ID")]
    [ColumnRule("ProjectIDList", IsRequired = false, IsWarningOnly = true)]
    public string ProjectID { get; set; } = string.Empty;

    [ExcelColumn("Entry Type", isRequired: true)]
    [ColumnRule("AttendanceEntryTypeList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "CheckIn")]
    public AttendanceEntryType EntryType { get; set; } = AttendanceEntryType.CheckIn;

    [ExcelColumn("Source", isRequired: true)]
    [ColumnRule("AttendanceSourceList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Manual")]
    public AttendanceSource Source { get; set; } = AttendanceSource.Manual;

    [ExcelColumn("Override Reason")]
    [ColumnRule(IsRequired = false)]
    public string OverrideReason { get; set; } = string.Empty;

    /// <summary>Thumbprint / Manual / SKSB_App / Admin rows default to approved; WhatsApp rows wait for an admin.</summary>
    [ExcelColumn("Is Approved")]
    [ColumnRule(IsRequired = false)]
    public bool IsApproved { get; set; } = true;

    [ExcelColumn("WhatsApp Picture")]
    [ColumnRule(IsRequired = false)]
    public string WhatsAppPicture { get; set; } = string.Empty;
}

// =========================================================================
// 4. DAILY WORKER ATTENDANCE ("DailyAttendanceTable") - ONE ROW PER WORKER PER DAY
// =========================================================================

[ExcelTable("DailyAttendanceTable")]
public record DailyWorkerAttendance
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("StaffIDList", IsRequired = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffNameList")]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long StaffID { get; set; }

    [JsonIgnore]
    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Status")]
    [ColumnRule("AttendanceStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Absence")]
    public AttendanceStatus AttendanceStatus { get; set; } = AttendanceStatus.Absence;

    [ExcelColumn("Clock-In Time")]
    [ColumnRule(IsRequired = false)]
    [ColumnFormat(ColumnFormatKind.Time, ZeroBehavior.Blank)]
    public TimeOnly? ClockInTime { get; set; }

    [ExcelColumn("Clock-Out Time")]
    [ColumnRule(IsRequired = false)]
    [ColumnFormat(ColumnFormatKind.Time, ZeroBehavior.Blank)]
    public TimeOnly? ClockOutTime { get; set; }

    /// <summary>Distinct projects checked into that day, comma-separated (e.g. <c>PROJ-A, PROJ-B</c>), derived strictly from Attendance Entries.</summary>
    [ExcelColumn("Project Check-Ins")]
    [ColumnRule(IsRequired = false)]
    public string ProjectCheckIns { get; set; } = string.Empty;

    /// <summary>
    /// Projects the worker is planned on, comma-separated: on-board slots listing the worker, plus projects whose slot was
    /// removed from the board (kept so re-adding the project restores its workers; shown in grey).
    /// </summary>
    [ExcelColumn("Planned Projects")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string PlannedProjects { get; set; } = string.Empty;

    /// <summary>Distinct projects worked that day (derived from <see cref="ProjectCheckIns"/>; not an Excel column).</summary>
    public List<string> ProjectsDone { get; set; } = new();

    [ExcelColumn("Regular Shift")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public decimal RegularShift { get; set; }

    /// <summary>+0.5 per distinct project beyond two: <c>Math.Max(0, distinctProjects - 2) * 0.5</c>.</summary>
    [ExcelColumn("Additional Shifts")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public decimal AdditionalShifts { get; set; }

    [ExcelColumn("OT Hours")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal OTHours { get; set; }

    /// <summary>+1.0 per premix hole fixed by the worker's slots that day.</summary>
    [ExcelColumn("Additional OT Hours")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal AdditionalOTHours { get; set; }

    [ColumnCalculate("ROUND((([Regular Shift] + [Additional Shifts]) * 8) + (([OT Hours] + [Additional OT Hours]) * 1.5), 2)")]
    [ExcelColumn("Total Flattened Hours")]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal TotalFlattenedHours { get; set; }

    /// <summary>Persisted for DB / JSON compatibility; no longer an Excel column (replaced by <see cref="PlannedProjects"/>).</summary>
    public bool IsAllocated { get; set; }
}

// =========================================================================
// 5. DAILY PROJECT SCHEDULE (board slot; not a ListObject - see ScheduleBoardMapper)
// =========================================================================

public record DailyProjectSchedule
{
    public DateOnly DailyScheduleDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    /// <summary><see cref="Shifts.Morning"/> = "Day Work" band, <see cref="Shifts.Night"/> = "Night Work" band.</summary>
    public Shifts Shift { get; set; } = Shifts.Morning;

    /// <summary>Board slot 0..7 (left to right).</summary>
    public int SlotIndex { get; set; }

    public string ProjectID { get; set; } = string.Empty;

    [JsonIgnore]
    public string ProjectName { get; set; } = string.Empty;

    /// <summary>Premix hole fixes for this project on this day (later also fed by claims data).</summary>
    public int PremixHoles { get; set; }

    public List<string> Workers { get; set; } = new();
    public List<string> Vehicles { get; set; } = new();
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 6. WORKER PREFERENCE TABLE ("WorkerPreferencesTable")
// =========================================================================

[ExcelTable("WorkerPreferencesTable")]
public record WorkerPreference
{
    [ExcelColumn("Staff ID", isRequired: true)]
    [ColumnRule("AllFieldStaffIDList", IsRequired = true)]
    [ColumnSync("Staff Name", "AllFieldStaffIDList", "AllFieldStaffNameList")]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long StaffID { get; set; }

    [JsonIgnore]
    [ExcelColumn("Staff Name")]
    [ColumnRule("AllFieldStaffNameList", IsRequired = false)]
    [ColumnSync("Staff ID", "AllFieldStaffNameList", "AllFieldStaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Team")]
    [ColumnRule("TeamTable[Team Name]", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> Team { get; set; } = new();

    /// <summary>Single dedicated role, e.g. Driver, Excavator Operator, Pipe Fitter, Foreman, General Worker.</summary>
    [ExcelColumn("Primary Skill")]
    [ColumnRule(IsRequired = false)]
    public string PrimarySkill { get; set; } = string.Empty;

    /// <summary>Read-only: derived from StaffManager's valid vehicle / skill licenses on every load.</summary>
    [ExcelColumn("Licensed Skills")]
    [ColumnRule(IsRequired = false)]
    public List<string> LicensedSkills { get; set; } = new();

    [ExcelColumn("Other Skills")]
    [ColumnRule("OtherSkillList", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> OtherSkills { get; set; } = new();

    [ExcelColumn("Blacklisted Skills")]
    [ColumnRule("StaffSkillList", "SkillsTable[Skills]", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> BlacklistedSkills { get; set; } = new();

    [ExcelColumn("Preferred Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false)]
    public VehicleType PreferredVehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Dedicated Vehicle Plate")]
    [ColumnRule("VehiclePlateList", IsRequired = false, IsWarningOnly = true)]
    public string DedicatedVehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Preferred Shift")]
    [ColumnRule("ShiftsList", IsRequired = false)]
    public Shifts PreferredShift { get; set; } = Shifts.None;

    [ExcelColumn("Max days a week")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int MaxDaysPerWeek { get; set; }

    [ExcelColumn("Rest Day Preference")]
    [ColumnRule("DaysList", IsRequired = false)]
    public Days RestDayPreference { get; set; } = Days.None;

    /// <summary>Server <c>updated_at</c> (optimistic-concurrency token); never written to Excel or JSON.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

// =========================================================================
// 7. TEAM TABLE ("TeamTable" on the Preference sheet)
// =========================================================================

[ExcelTable("TeamTable")]
public record TeamRecord
{
    [ExcelColumn("Team Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string TeamName { get; set; } = string.Empty;

    /// <summary>Leader's short name (Excel); resolved to a Staff ID on save.</summary>
    [ExcelColumn("Team Leader")]
    [ColumnRule("AllFieldStaffNameList", IsRequired = false)]
    public string TeamLeader { get; set; } = string.Empty;

    [JsonIgnore]
    public long TeamLeaderStaffID { get; set; }

    [ExcelColumn("Default Vehicles")]
    [ColumnRule("VehiclePlateList", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> DefaultVehicles { get; set; } = new();

    /// <summary>Server <c>updated_at</c> (optimistic-concurrency token); never written to Excel or JSON.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

// =========================================================================
// 7b. SKILLS TABLE ("SkillsTable" on the Preference sheet)
// =========================================================================

[ExcelTable("SkillsTable")]
public record ScheduleSkillRecord
{
    [DataList("StaffSkillList")]
    [ExcelColumn("Skills", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string SkillName { get; set; } = string.Empty;

    [ExcelColumn("Display Icon")]
    [ColumnRule("DisplayIconList", IsRequired = false, IsWarningOnly = true)]
    public string DisplayIcon { get; set; } = string.Empty;

    /// <summary>Server <c>updated_at</c> (optimistic-concurrency token); never written to Excel or JSON.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

// =========================================================================
// 7c. PROJECT SKILL REQUIREMENTS TABLE ("ProjectRequirementsTable" on Preference)
// =========================================================================

[ExcelTable("ProjectRequirementsTable")]
public record ProjectSkillRequirement
{
    [ExcelColumn("Project Type", isRequired: true)]
    [ColumnRule("ProjectTypeList", IsRequired = true)]
    public ProjectType ProjectType { get; set; } = ProjectType.None;

    [ExcelColumn("Required Skills")]
    [ColumnRule("StaffSkillList", "SkillsTable[Skills]", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> Skills { get; set; } = new();

    [ExcelColumn("Required Vehicles")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<VehicleType> Vehicles { get; set; } = new();

    /// <summary>Server <c>updated_at</c> (optimistic-concurrency token); never written to Excel or JSON.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

// =========================================================================
// 8. PROJECT-TYPE REQUIREMENTS (editable mapping on 'Lookups')
// =========================================================================

[ExcelTable("ProjectTypeRequirementsTable")]
public record ProjectTypeRequirement
{
    [ExcelColumn("Project Type", isRequired: true)]
    [ColumnRule("ProjectTypeList", IsRequired = true)]
    public ProjectType ProjectType { get; set; } = ProjectType.None;

    [ExcelColumn("Required Skills")]
    [ColumnRule("StaffSkillList", "SkillsTable[Skills]", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<string> RequiredSkills { get; set; } = new();

    [ExcelColumn("Required Vehicle Types")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public List<VehicleType> RequiredVehicleTypes { get; set; } = new();
}

// =========================================================================
// 9. ICON LEGEND (editable mapping on 'Lookups')
// =========================================================================

[ExcelTable("SkillIconsTable")]
public record SkillIcon
{
    [ExcelColumn("Skill / Vehicle Type", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string Key { get; set; } = string.Empty;

    [ExcelColumn("Icon")]
    [ColumnRule(IsRequired = false)]
    public string Icon { get; set; } = string.Empty;
}
