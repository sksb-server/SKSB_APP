namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using Dapper;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.Types;

public static partial class ServerDispatcher
{
    private static bool _scheduleSchemaEnsured;

    // =========================================================================
    // 1. SCHEMA BOOTSTRAP (embedded schedule_schema.sql, idempotent)
    // =========================================================================
    public static void EnsureScheduleSchema()
    {
        if (_scheduleSchemaEnsured) return;

        using Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("schedule_schema.sql");
        if (stream == null) throw new InvalidOperationException("Embedded resource 'schedule_schema.sql' is missing from Core.");
        using var reader = new StreamReader(stream);
        string ddl = reader.ReadToEnd();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute(ddl, transaction: trans);
            conn.Execute(SqlPayrollCalendarSchema, transaction: trans);
        });
        _scheduleSchemaEnsured = true;
    }

    /// <summary>SQL Payroll calendar mirror: per-date holiday flag, public holidays and date-scoped leave days.</summary>
    private const string SqlPayrollCalendarSchema = @"
        ALTER TABLE daily_schedule_settings ADD COLUMN IF NOT EXISTS is_holiday BOOLEAN NOT NULL DEFAULT FALSE;

        CREATE TABLE IF NOT EXISTS public_holidays (
            holiday_date DATE PRIMARY KEY,
            description  TEXT DEFAULT '',
            updated_at   TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );

        CREATE TABLE IF NOT EXISTS staff_leaves (
            leave_id       BIGSERIAL PRIMARY KEY,
            payroll_dockey BIGINT UNIQUE,
            staff_id       BIGINT NOT NULL,
            leave_date     DATE NOT NULL,
            leave_type     VARCHAR(64) NOT NULL,
            leave_day      NUMERIC(5,2) NOT NULL DEFAULT 1.0,
            description    TEXT DEFAULT '',
            updated_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        );
        CREATE INDEX IF NOT EXISTS ix_staff_leaves_leave_date ON staff_leaves (leave_date);";

    private static string[] CleanList(IEnumerable<string>? items) =>
        (items ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToArray();

    // =========================================================================
    // 2. SCHEDULE ROW PROJECTIONS & SQL
    // =========================================================================
    /// <summary>
    /// Concurrency token for one working date's board: MAX(updated_at) across the three date-scoped tables
    /// (NULL when the date has no rows yet). Parameter: <c>@ScheduleDate</c>.
    /// </summary>
    public const string SqlScheduleBoardMaxUpdatedAt = @"
        SELECT MAX(u) FROM (
            SELECT updated_at AS u FROM daily_schedule_settings WHERE schedule_date = @ScheduleDate
            UNION ALL SELECT updated_at AS u FROM daily_project_schedules WHERE schedule_date = @ScheduleDate
            UNION ALL SELECT updated_at AS u FROM daily_attendances WHERE schedule_date = @ScheduleDate) t;";

    private static bool IsValidPreference(WorkerPreference p) => p.StaffID > 0;
    private static bool IsValidTeam(TeamRecord t) => !string.IsNullOrWhiteSpace(t.TeamName);
    private static bool IsValidScheduleSkill(ScheduleSkillRecord s) => !string.IsNullOrWhiteSpace(s.SkillName);
    private static bool IsValidProjectSkillRequirement(ProjectSkillRequirement r) => r.ProjectType != ProjectType.None;

    private static object ToPreferenceRow(WorkerPreference p, DateTime? expectedUpdatedAt = null) => new
    {
        StaffId = p.StaffID,
        Team = CleanList(p.Team),
        PrimarySkill = (p.PrimarySkill ?? string.Empty).Trim(),
        OtherSkills = CleanList(p.OtherSkills),
        BlacklistedSkills = CleanList(p.BlacklistedSkills),
        PreferredVehicleType = p.PreferredVehicleType.ToString(),
        DedicatedVehiclePlate = (p.DedicatedVehiclePlate ?? string.Empty).Trim().ToUpperInvariant(),
        PreferredShift = p.PreferredShift.ToString(),
        MaxDaysPerWeek = Math.Max(0, p.MaxDaysPerWeek),
        RestDayPreference = p.RestDayPreference.ToString(),
        ExpectedUpdatedAt = expectedUpdatedAt
    };

    private static object ToTeamRow(TeamRecord t, DateTime? expectedUpdatedAt = null) => new
    {
        TeamName = (t.TeamName ?? string.Empty).Trim(),
        TeamLeaderStaffId = t.TeamLeaderStaffID > 0 ? (long?)t.TeamLeaderStaffID : null,
        DefaultVehicles = CleanList(t.DefaultVehicles).Select(v => v.ToUpperInvariant()).ToArray(),
        ExpectedUpdatedAt = expectedUpdatedAt
    };

    private static object ToScheduleSkillRow(ScheduleSkillRecord s, DateTime? expectedUpdatedAt = null) => new
    {
        SkillName = (s.SkillName ?? string.Empty).Trim(),
        DisplayIcon = (s.DisplayIcon ?? string.Empty).Trim(),
        ExpectedUpdatedAt = expectedUpdatedAt
    };

    private static object ToProjectSkillRequirementRow(ProjectSkillRequirement r, DateTime? expectedUpdatedAt = null) => new
    {
        ProjectType = r.ProjectType.ToString(),
        RequiredSkills = CleanList(r.Skills),
        RequiredVehicles = (r.Vehicles ?? new List<VehicleType>())
            .Where(v => v != VehicleType.None)
            .Distinct()
            .Select(v => v.ToString())
            .ToArray(),
        ExpectedUpdatedAt = expectedUpdatedAt
    };

    // ---- worker_preferences ----
    private const string SqlPrefInsert = @"
        INSERT INTO worker_preferences (
            staff_id, team, primary_skill, other_skills, blacklisted_skills, preferred_vehicle_type,
            dedicated_vehicle_plate, preferred_shift, max_days_per_week, rest_day_preference, updated_at)
        VALUES (
            @StaffId, @Team, @PrimarySkill, @OtherSkills, @BlacklistedSkills, @PreferredVehicleType,
            @DedicatedVehiclePlate, @PreferredShift, @MaxDaysPerWeek, @RestDayPreference, clock_timestamp())";

    private const string SqlPrefUpsert = SqlPrefInsert + @"
        ON CONFLICT (staff_id) DO UPDATE SET
            team                    = EXCLUDED.team,
            primary_skill           = EXCLUDED.primary_skill,
            other_skills            = EXCLUDED.other_skills,
            blacklisted_skills      = EXCLUDED.blacklisted_skills,
            preferred_vehicle_type  = EXCLUDED.preferred_vehicle_type,
            dedicated_vehicle_plate = EXCLUDED.dedicated_vehicle_plate,
            preferred_shift         = EXCLUDED.preferred_shift,
            max_days_per_week       = EXCLUDED.max_days_per_week,
            rest_day_preference     = EXCLUDED.rest_day_preference,
            updated_at              = GREATEST(clock_timestamp(), worker_preferences.updated_at + INTERVAL '1 millisecond');";

    private const string SqlPrefInsertIfAbsent = SqlPrefInsert + @"
        ON CONFLICT (staff_id) DO NOTHING;";

    private const string SqlPrefConditionalUpdate = @"
        UPDATE worker_preferences SET
            team                    = @Team,
            primary_skill           = @PrimarySkill,
            other_skills            = @OtherSkills,
            blacklisted_skills      = @BlacklistedSkills,
            preferred_vehicle_type  = @PreferredVehicleType,
            dedicated_vehicle_plate = @DedicatedVehiclePlate,
            preferred_shift         = @PreferredShift,
            max_days_per_week       = @MaxDaysPerWeek,
            rest_day_preference     = @RestDayPreference,
            updated_at              = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE staff_id = @StaffId
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    // ---- teams_master ----
    private const string SqlTeamInsert = @"
        INSERT INTO teams_master (team_name, team_leader_staff_id, default_vehicles, updated_at)
        VALUES (@TeamName, @TeamLeaderStaffId, @DefaultVehicles, clock_timestamp())";

    private const string SqlTeamUpsert = SqlTeamInsert + @"
        ON CONFLICT (team_name) DO UPDATE SET
            team_leader_staff_id = EXCLUDED.team_leader_staff_id,
            default_vehicles     = EXCLUDED.default_vehicles,
            updated_at           = GREATEST(clock_timestamp(), teams_master.updated_at + INTERVAL '1 millisecond');";

    private const string SqlTeamInsertIfAbsent = SqlTeamInsert + @"
        ON CONFLICT (team_name) DO NOTHING;";

    private const string SqlTeamConditionalUpdate = @"
        UPDATE teams_master SET
            team_leader_staff_id = @TeamLeaderStaffId,
            default_vehicles     = @DefaultVehicles,
            updated_at           = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE LOWER(team_name) = LOWER(@TeamName)
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    // ---- lookup_schedule_skills ----
    private const string SqlSkillInsert = @"
        INSERT INTO lookup_schedule_skills (skill_name, display_icon, updated_at)
        VALUES (@SkillName, @DisplayIcon, clock_timestamp())";

    private const string SqlSkillUpsert = SqlSkillInsert + @"
        ON CONFLICT (skill_name) DO UPDATE SET
            display_icon = EXCLUDED.display_icon,
            updated_at   = GREATEST(clock_timestamp(), lookup_schedule_skills.updated_at + INTERVAL '1 millisecond');";

    private const string SqlSkillInsertIfAbsent = SqlSkillInsert + @"
        ON CONFLICT (skill_name) DO NOTHING;";

    private const string SqlSkillConditionalUpdate = @"
        UPDATE lookup_schedule_skills SET
            display_icon = @DisplayIcon,
            updated_at   = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE LOWER(skill_name) = LOWER(@SkillName)
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    // ---- project_skill_requirements ----
    private const string SqlProjSkillReqInsert = @"
        INSERT INTO project_skill_requirements (project_type, required_skills, required_vehicles, updated_at)
        VALUES (@ProjectType, @RequiredSkills, @RequiredVehicles, clock_timestamp())";

    private const string SqlProjSkillReqUpsert = SqlProjSkillReqInsert + @"
        ON CONFLICT (project_type) DO UPDATE SET
            required_skills   = EXCLUDED.required_skills,
            required_vehicles = EXCLUDED.required_vehicles,
            updated_at        = GREATEST(clock_timestamp(), project_skill_requirements.updated_at + INTERVAL '1 millisecond');";

    private const string SqlProjSkillReqInsertIfAbsent = SqlProjSkillReqInsert + @"
        ON CONFLICT (project_type) DO NOTHING;";

    private const string SqlProjSkillReqConditionalUpdate = @"
        UPDATE project_skill_requirements SET
            required_skills   = @RequiredSkills,
            required_vehicles = @RequiredVehicles,
            updated_at        = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE project_type = @ProjectType
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    // ---- date-scoped board / settings / daily attendance ----
    private const string SqlSlotUpsert = @"
        INSERT INTO daily_project_schedules (schedule_date, shift, slot_index, project_id, premix_holes, workers, vehicles, remarks, updated_at)
        VALUES (@ScheduleDate, @Shift, @SlotIndex, @ProjectId, @PremixHoles, @Workers, @Vehicles, @Remarks, clock_timestamp())
        ON CONFLICT (schedule_date, shift, slot_index) DO UPDATE SET
            project_id   = EXCLUDED.project_id,
            premix_holes = EXCLUDED.premix_holes,
            workers      = EXCLUDED.workers,
            vehicles     = EXCLUDED.vehicles,
            remarks      = EXCLUDED.remarks,
            updated_at   = GREATEST(clock_timestamp(), daily_project_schedules.updated_at + INTERVAL '1 millisecond');";

    private const string SqlDailyAttendanceUpsert = @"
        INSERT INTO daily_attendances (
            schedule_date, staff_id, attendance_status, clock_in_time, clock_out_time, project_check_ins, projects_done,
            regular_shift, additional_shifts, ot_hours, additional_ot_hours, total_flattened_hours,
            is_allocated, updated_at)
        VALUES (
            @ScheduleDate, @StaffId, @AttendanceStatus, @ClockInTime, @ClockOutTime, @ProjectCheckIns, @ProjectsDone,
            @RegularShift, @AdditionalShifts, @OtHours, @AdditionalOtHours, @TotalFlattenedHours,
            @IsAllocated, clock_timestamp())
        ON CONFLICT (schedule_date, staff_id) DO UPDATE SET
            attendance_status     = EXCLUDED.attendance_status,
            clock_in_time         = EXCLUDED.clock_in_time,
            clock_out_time        = EXCLUDED.clock_out_time,
            project_check_ins     = EXCLUDED.project_check_ins,
            projects_done         = EXCLUDED.projects_done,
            regular_shift         = EXCLUDED.regular_shift,
            additional_shifts     = EXCLUDED.additional_shifts,
            ot_hours              = EXCLUDED.ot_hours,
            additional_ot_hours   = EXCLUDED.additional_ot_hours,
            total_flattened_hours = EXCLUDED.total_flattened_hours,
            is_allocated          = EXCLUDED.is_allocated,
            updated_at            = GREATEST(clock_timestamp(), daily_attendances.updated_at + INTERVAL '1 millisecond');";

    private const string SqlScheduleSettingsUpsert = @"
        INSERT INTO daily_schedule_settings (schedule_date, day_start_time, late_cut_off_time, ot_start_time, is_holiday, updated_at)
        VALUES (@ScheduleDate, @DayStartTime, @LateCutOffTime, @OtStartTime, @IsHoliday, clock_timestamp())
        ON CONFLICT (schedule_date) DO UPDATE SET
            day_start_time    = EXCLUDED.day_start_time,
            late_cut_off_time = EXCLUDED.late_cut_off_time,
            ot_start_time     = EXCLUDED.ot_start_time,
            is_holiday        = EXCLUDED.is_holiday,
            updated_at        = GREATEST(clock_timestamp(), daily_schedule_settings.updated_at + INTERVAL '1 millisecond');";

    /// <summary>
    /// Upserts the settings row and reconciles the board slots / daily attendances of ONE working date
    /// (rows removed from the board disappear for that date only). Returns (slots, attendances) written.
    /// </summary>
    private static (int Slots, int Attendances) WriteScheduleDateScoped(IDbConnection conn, IDbTransaction trans, ScheduleRecord schedule)
    {
        DateTime date = schedule.DailyScheduleDate.ToDateTime(TimeOnly.MinValue);

        var slots = schedule.Projects
            .Where(p => p.SlotIndex >= 0 &&
                        (p.Workers.Count > 0 || p.Vehicles.Count > 0 || p.PremixHoles > 0))
            .Select(p => new
            {
                ScheduleDate = date,
                Shift = p.Shift.ToString(),
                p.SlotIndex,
                ProjectId = (p.ProjectID ?? string.Empty).Trim().ToUpperInvariant(),
                PremixHoles = Math.Max(0, p.PremixHoles),
                Workers = CleanList(p.Workers),
                Vehicles = CleanList(p.Vehicles).Select(v => v.ToUpperInvariant()).ToArray(),
                Remarks = p.Remarks ?? string.Empty
            }).ToList();
        string[] activeSlotKeys = slots.Select(s => $"{s.Shift}:{s.SlotIndex}").ToArray();

        var attendances = schedule.WorkerAttendance.Where(a => a.StaffID > 0).Select(a => new
        {
            ScheduleDate = date,
            StaffId = a.StaffID,
            AttendanceStatus = a.AttendanceStatus.ToString(),
            ClockInTime = a.ClockInTime.HasValue ? (TimeSpan?)a.ClockInTime.Value.ToTimeSpan() : null,
            ClockOutTime = a.ClockOutTime.HasValue ? (TimeSpan?)a.ClockOutTime.Value.ToTimeSpan() : null,
            ProjectCheckIns = a.ProjectCheckIns ?? string.Empty,
            ProjectsDone = CleanList(a.ProjectsDone),
            a.RegularShift,
            a.AdditionalShifts,
            OtHours = a.OTHours,
            AdditionalOtHours = a.AdditionalOTHours,
            a.TotalFlattenedHours,
            a.IsAllocated
        }).ToList();
        long[] activeAttendanceIds = attendances.Select(a => a.StaffId).ToArray();

        conn.Execute(SqlScheduleSettingsUpsert, new
        {
            ScheduleDate = date,
            DayStartTime = schedule.DayStartTime.ToTimeSpan(),
            LateCutOffTime = schedule.LateCutOffTime.ToTimeSpan(),
            OtStartTime = schedule.OTStartTime.ToTimeSpan(),
            schedule.IsHoliday
        }, trans);

        // Date-scoped reconciliation: removed slots / workers disappear for THIS date only.
        conn.Execute(
            "DELETE FROM daily_project_schedules WHERE schedule_date = @ScheduleDate AND (shift || ':' || slot_index::text) != ALL(@ActiveKeys);",
            new { ScheduleDate = date, ActiveKeys = activeSlotKeys }, trans);
        conn.Execute(
            "DELETE FROM daily_attendances WHERE schedule_date = @ScheduleDate AND staff_id != ALL(@ActiveIds);",
            new { ScheduleDate = date, ActiveIds = activeAttendanceIds }, trans);

        if (slots.Count > 0) conn.Execute(SqlSlotUpsert, slots, trans);
        if (attendances.Count > 0) conn.Execute(SqlDailyAttendanceUpsert, attendances, trans);
        return (slots.Count, attendances.Count);
    }

    // =========================================================================
    // 2b. SQL PAYROLL CALENDAR MIRROR (public holidays + date-scoped leaves)
    // =========================================================================
    private const string SqlPublicHolidayUpsert = @"
        INSERT INTO public_holidays (holiday_date, description, updated_at)
        VALUES (@HolidayDate, @Description, CURRENT_TIMESTAMP)
        ON CONFLICT (holiday_date) DO UPDATE SET
            description = EXCLUDED.description,
            updated_at  = CURRENT_TIMESTAMP
        WHERE public_holidays.description IS DISTINCT FROM EXCLUDED.description;";

    private const string SqlStaffLeaveUpsert = @"
        INSERT INTO staff_leaves (payroll_dockey, staff_id, leave_date, leave_type, leave_day, description, updated_at)
        VALUES (@DocKey, @StaffId, @LeaveDate, @LeaveType, @LeaveDay, @Description, CURRENT_TIMESTAMP)
        ON CONFLICT (payroll_dockey) DO UPDATE SET
            staff_id    = EXCLUDED.staff_id,
            leave_date  = EXCLUDED.leave_date,
            leave_type  = EXCLUDED.leave_type,
            leave_day   = EXCLUDED.leave_day,
            description = EXCLUDED.description,
            updated_at  = CURRENT_TIMESTAMP
        WHERE staff_leaves.staff_id    IS DISTINCT FROM EXCLUDED.staff_id
           OR staff_leaves.leave_date  IS DISTINCT FROM EXCLUDED.leave_date
           OR staff_leaves.leave_type  IS DISTINCT FROM EXCLUDED.leave_type
           OR staff_leaves.leave_day   IS DISTINCT FROM EXCLUDED.leave_day
           OR staff_leaves.description IS DISTINCT FROM EXCLUDED.description;";

    /// <summary>
    /// Mirrors SQL Payroll public holidays into <c>public_holidays</c> (upsert by date; SQL Payroll is authoritative
    /// but rows are never hard-deleted, so history survives a calendar clean-up there). Returns rows written.
    /// </summary>
    public static int UpsertPublicHolidays(IEnumerable<PayrollHolidayRow> holidays)
    {
        EnsureScheduleSchema();
        var rows = (holidays ?? Enumerable.Empty<PayrollHolidayRow>())
            .GroupBy(h => h.HolidayDate)
            .Select(g => new
            {
                HolidayDate = g.Key.ToDateTime(TimeOnly.MinValue),
                Description = (g.OrderByDescending(h => h.AutoKey).First().Description ?? string.Empty).Trim()
            })
            .ToList();
        if (rows.Count == 0) return 0;

        int written = 0;
        Db.ExecuteInTransaction((conn, trans) => { written = conn.Execute(SqlPublicHolidayUpsert, rows, trans); });
        AppLogger.Info($"[ServerDispatcher] public_holidays synced: {written}/{rows.Count} row(s) written.");
        return written;
    }

    /// <summary>
    /// Makes <c>staff_leaves</c> for <paramref name="date"/> match SQL Payroll: upserts by <c>payroll_dockey</c> and
    /// deletes rows of that date no longer present in SQL Payroll (a cache, not user data). Returns rows upserted.
    /// </summary>
    public static int ReplaceLeavesForDate(DateOnly date, IEnumerable<StaffLeaveEntry> leaves)
    {
        EnsureScheduleSchema();
        DateTime d = date.ToDateTime(TimeOnly.MinValue);
        var rows = (leaves ?? Enumerable.Empty<StaffLeaveEntry>())
            .Where(l => l != null && l.DocKey > 0 && l.StaffID > 0 && l.LeaveDate == date)
            .GroupBy(l => l.DocKey)
            .Select(g => g.Last())
            .Select(l => new
            {
                l.DocKey,
                StaffId = l.StaffID,
                LeaveDate = d,
                LeaveType = (l.LeaveType ?? string.Empty).Trim(),
                l.LeaveDay,
                Description = (l.Description ?? string.Empty).Trim()
            })
            .ToList();
        long[] activeKeys = rows.Select(r => r.DocKey).ToArray();

        int written = 0;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute("DELETE FROM staff_leaves WHERE leave_date = @D AND payroll_dockey != ALL(@ActiveKeys);",
                new { D = d, ActiveKeys = activeKeys }, trans);
            if (rows.Count > 0) written = conn.Execute(SqlStaffLeaveUpsert, rows, trans);
        });
        return written;
    }

    /// <summary>
    /// Records the SQL Payroll holiday flag on an EXISTING <c>daily_schedule_settings</c> row without touching
    /// <c>updated_at</c>, so the date's board concurrency token is unaffected. New dates get the flag on save.
    /// </summary>
    public static void UpdateScheduleHolidayFlag(DateOnly date, bool isHoliday)
    {
        EnsureScheduleSchema();
        Db.ExecuteInTransaction((conn, trans) => conn.Execute(
            "UPDATE daily_schedule_settings SET is_holiday = @IsHoliday WHERE schedule_date = @D AND is_holiday IS DISTINCT FROM @IsHoliday;",
            new { D = date.ToDateTime(TimeOnly.MinValue), IsHoliday = isHoliday }, trans));
    }

    // =========================================================================
    // 3. LEGACY / OFFLINE SCHEDULE PERSISTENCE (upsert-only, single transaction)
    // =========================================================================
    /// <summary>
    /// Persists one working date plus the global reference tables (offline pending upload / generic dispatch):
    /// date-scoped board slots and daily attendances are reconciled for <see cref="ScheduleRecord.DailyScheduleDate"/>;
    /// worker preferences, teams, skills and project skill requirements are upsert-only (never hard-deleted);
    /// attendance entries are append-only (a wrong punch is voided by clearing "Is Approved").
    /// </summary>
    public static void PersistSchedule(ScheduleRecord schedule)
    {
        EnsureScheduleSchema();

        // ModifiedAttendanceEntries == null (offline pending upload / no baseline) => safe full upsert.
        var entrySource = schedule.ModifiedAttendanceEntries ?? schedule.AttendanceEntries;
        int entriesUpserted = 0;
        (int Slots, int Attendances) written = (0, 0);

        var prefs = DistinctBy(schedule.Preferences.Where(IsValidPreference), p => p.StaffID.ToString(CultureInfo.InvariantCulture))
            .Select(p => ToPreferenceRow(p)).ToList();
        var teams = DistinctBy(schedule.Teams.Where(IsValidTeam), t => t.TeamName.Trim().ToUpperInvariant())
            .Select(t => ToTeamRow(t)).ToList();
        var skills = DistinctBy((schedule.Skills ?? new List<ScheduleSkillRecord>()).Where(IsValidScheduleSkill), s => s.SkillName.Trim().ToUpperInvariant())
            .Select(s => ToScheduleSkillRow(s)).ToList();
        var projSkillReqs = DistinctBy((schedule.ProjectSkillRequirements ?? new List<ProjectSkillRequirement>()).Where(IsValidProjectSkillRequirement), r => r.ProjectType.ToString())
            .Select(r => ToProjectSkillRequirementRow(r)).ToList();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            written = WriteScheduleDateScoped(conn, trans, schedule);

            // Reference tables: upsert-only (no hard deletes; rows missing from the workbook are kept).
            if (teams.Count > 0) conn.Execute(SqlTeamUpsert, teams, trans);
            if (prefs.Count > 0) conn.Execute(SqlPrefUpsert, prefs, trans);
            if (skills.Count > 0) conn.Execute(SqlSkillUpsert, skills, trans);
            if (projSkillReqs.Count > 0) conn.Execute(SqlProjSkillReqUpsert, projSkillReqs, trans);

            entriesUpserted = UpsertAttendanceEntriesBulk(conn, trans, entrySource);
        });

        AppLogger.Info(
            $"[PostgreSQL Upload - Schedule {schedule.DailyScheduleDate:yyyy-MM-dd}] " +
            $"dayStart={schedule.DayStartTime:HH:mm}, lateCutOff={schedule.LateCutOffTime:HH:mm}, otStart={schedule.OTStartTime:HH:mm}, " +
            $"slots={written.Slots}, attendances={written.Attendances}, entries={entriesUpserted}/{schedule.AttendanceEntries.Count}, prefs={prefs.Count}, teams={teams.Count}, skills={skills.Count}, projSkillReqs={projSkillReqs.Count}");
    }

    private static IEnumerable<T> DistinctBy<T>(IEnumerable<T> rows, Func<T, string> key)
    {
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (var r in rows)
            if (seen.Add(key(r))) yield return r;
    }

    // =========================================================================
    // 4. DELTA SAVE WITH OPTIMISTIC CONCURRENCY
    // =========================================================================
    private sealed class ScheduleRefSpec<T>
    {
        public string DbTable { get; init; } = string.Empty;
        public string ExcelTable { get; init; } = string.Empty;
        public string KeyHeader { get; init; } = string.Empty;
        public string Label { get; init; } = string.Empty;
        public Func<T, bool> IsValid { get; init; } = _ => true;
        /// <summary>Raw key as typed in the Excel key column (also the dedupe key after upper-casing).</summary>
        public Func<T, string> RawKey { get; init; } = _ => string.Empty;
        public Func<T, DateTime?, object> ToParams { get; init; } = (_, _) => new { };
        public string Upsert { get; init; } = string.Empty;
        public string InsertIfAbsent { get; init; } = string.Empty;
        public string ConditionalUpdate { get; init; } = string.Empty;
        /// <summary>Returns the server <c>updated_at</c> for the row; parameter <c>@Key</c>.</summary>
        public string ServerUpdatedAtSql { get; init; } = string.Empty;
        public Func<T, object> KeyParam { get; init; } = _ => new { };
        public Func<IDbConnection, IDbTransaction, T, List<FieldConflict>> Diff { get; init; } = (_, _, _) => new List<FieldConflict>();
    }

    private static readonly ScheduleRefSpec<TeamRecord> TeamRefSpec = new()
    {
        DbTable = "teams_master", ExcelTable = "TeamTable", KeyHeader = "Team Name", Label = "Team",
        IsValid = IsValidTeam, RawKey = t => (t.TeamName ?? string.Empty).Trim(), ToParams = ToTeamRow,
        Upsert = SqlTeamUpsert, InsertIfAbsent = SqlTeamInsertIfAbsent, ConditionalUpdate = SqlTeamConditionalUpdate,
        ServerUpdatedAtSql = "SELECT updated_at FROM teams_master WHERE LOWER(team_name) = LOWER(@Key) LIMIT 1;",
        KeyParam = t => new { Key = (t.TeamName ?? string.Empty).Trim() },
        Diff = ScheduleConflictDiff.ForTeam
    };

    private static readonly ScheduleRefSpec<WorkerPreference> PreferenceRefSpec = new()
    {
        DbTable = "worker_preferences", ExcelTable = "WorkerPreferencesTable", KeyHeader = "Staff ID", Label = "Worker Preference",
        IsValid = IsValidPreference, RawKey = p => p.StaffID.ToString(CultureInfo.InvariantCulture), ToParams = ToPreferenceRow,
        Upsert = SqlPrefUpsert, InsertIfAbsent = SqlPrefInsertIfAbsent, ConditionalUpdate = SqlPrefConditionalUpdate,
        ServerUpdatedAtSql = "SELECT updated_at FROM worker_preferences WHERE staff_id = @Key;",
        KeyParam = p => new { Key = p.StaffID },
        Diff = ScheduleConflictDiff.ForPreference
    };

    private static readonly ScheduleRefSpec<ScheduleSkillRecord> SkillRefSpec = new()
    {
        DbTable = "lookup_schedule_skills", ExcelTable = "SkillsTable", KeyHeader = "Skills", Label = "Skill",
        IsValid = IsValidScheduleSkill, RawKey = s => (s.SkillName ?? string.Empty).Trim(), ToParams = ToScheduleSkillRow,
        Upsert = SqlSkillUpsert, InsertIfAbsent = SqlSkillInsertIfAbsent, ConditionalUpdate = SqlSkillConditionalUpdate,
        ServerUpdatedAtSql = "SELECT updated_at FROM lookup_schedule_skills WHERE LOWER(skill_name) = LOWER(@Key) LIMIT 1;",
        KeyParam = s => new { Key = (s.SkillName ?? string.Empty).Trim() },
        Diff = ScheduleConflictDiff.ForSkill
    };

    private static readonly ScheduleRefSpec<ProjectSkillRequirement> ProjSkillReqRefSpec = new()
    {
        DbTable = "project_skill_requirements", ExcelTable = "ProjectRequirementsTable", KeyHeader = "Project Type", Label = "Project Requirement",
        IsValid = IsValidProjectSkillRequirement, RawKey = r => r.ProjectType.ToString(), ToParams = ToProjectSkillRequirementRow,
        Upsert = SqlProjSkillReqUpsert, InsertIfAbsent = SqlProjSkillReqInsertIfAbsent, ConditionalUpdate = SqlProjSkillReqConditionalUpdate,
        ServerUpdatedAtSql = "SELECT updated_at FROM project_skill_requirements WHERE project_type = @Key;",
        KeyParam = r => new { Key = r.ProjectType.ToString() },
        Diff = ScheduleConflictDiff.ForProjectSkillRequirement
    };

    /// <summary>
    /// Writes one reference table's delta: fallback rows are upserted; Added rows are inserted only when absent;
    /// Modified rows are updated only when the server <c>updated_at</c> still matches the baseline. Never deletes.
    /// </summary>
    private static void ApplyScheduleRefDelta<T>(
        IDbConnection conn, IDbTransaction trans, ScheduleRefSpec<T> spec, TableDelta<T> delta,
        bool forceOverwrite, List<ConcurrencyConflict> conflicts)
    {
        string DedupeKey(T r) => spec.RawKey(r).ToUpperInvariant();

        var fallback = DistinctBy(delta.FallbackUpsert.Where(spec.IsValid), DedupeKey).Select(r => spec.ToParams(r, null)).ToList();
        if (fallback.Count > 0) conn.Execute(spec.Upsert, fallback, trans);

        void AddConflict(T r, string what, DateTime? expected)
        {
            string key = spec.RawKey(r);
            DateTime? serverTs = null;
            try { serverTs = conn.ExecuteScalar<DateTime?>(spec.ServerUpdatedAtSql, spec.KeyParam(r), trans); }
            catch (Exception ex) { AppLogger.Warn($"[ServerDispatcher] {spec.DbTable} updated_at lookup failed: {ex.Message}"); }

            conflicts.Add(new ConcurrencyConflict(spec.DbTable, $"{spec.Label} {key}",
                $"{spec.Label} '{key}' {what}",
                expected, serverTs,
                ExcelTableName: spec.ExcelTable,
                KeyColumnHeader: spec.KeyHeader,
                RawKeyValue: key,
                FieldDiffs: spec.Diff(conn, trans, r)));
        }

        foreach (var r in DistinctBy(delta.Added.Where(spec.IsValid), DedupeKey))
        {
            if (forceOverwrite)
            {
                conn.Execute(spec.Upsert, spec.ToParams(r, null), trans);
                continue;
            }
            if (conn.Execute(spec.InsertIfAbsent, spec.ToParams(r, null), trans) == 0)
                AddConflict(r, "already exists on the server (added by another user).", null);
        }

        var seenModified = new HashSet<string>(StringComparer.Ordinal);
        foreach (var item in delta.Modified.Where(m => spec.IsValid(m.Record)))
        {
            var r = item.Record;
            if (!seenModified.Add(DedupeKey(r))) continue;

            if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
            {
                if (conn.Execute(spec.ConditionalUpdate, spec.ToParams(r, item.ExpectedUpdatedAt), trans) == 0)
                    AddConflict(r, "was modified on the server by another user.", item.ExpectedUpdatedAt);
            }
            else
            {
                conn.Execute(spec.Upsert, spec.ToParams(r, null), trans);
            }
        }
    }

    /// <summary>
    /// Delta save for one working date: the four global reference tables write only Added / Modified rows (never
    /// hard-deleted) under row-level <c>updated_at</c> checks; attendance entries stay append-only; the date-scoped
    /// board / settings / daily attendance are rewritten only when they changed, guarded by the date's
    /// MAX(updated_at). Any conflict throws <see cref="ConcurrencyConflictException"/> so the whole batch rolls back.
    /// <paramref name="forceOverwrite"/> skips every check.
    /// </summary>
    public static void PersistScheduleDelta(ScheduleDeltaPayload delta, bool forceOverwrite = false)
    {
        EnsureScheduleSchema();

        var schedule = delta.Schedule;
        DateTime date = schedule.DailyScheduleDate.ToDateTime(TimeOnly.MinValue);
        var entrySource = schedule.ModifiedAttendanceEntries ?? schedule.AttendanceEntries;
        bool writeBoard = !delta.HasDateBaseline || delta.BoardOrAttendanceChanged || forceOverwrite;
        int entriesUpserted = 0;
        (int Slots, int Attendances) written = (0, 0);

        Db.ExecuteInTransaction((conn, trans) =>
        {
            var conflicts = new List<ConcurrencyConflict>();

            // a. Global reference tables (teams first: preferences reference them by name).
            ApplyScheduleRefDelta(conn, trans, TeamRefSpec, delta.Teams, forceOverwrite, conflicts);
            ApplyScheduleRefDelta(conn, trans, PreferenceRefSpec, delta.Preferences, forceOverwrite, conflicts);
            ApplyScheduleRefDelta(conn, trans, SkillRefSpec, delta.Skills, forceOverwrite, conflicts);
            ApplyScheduleRefDelta(conn, trans, ProjSkillReqRefSpec, delta.ProjectSkillRequirements, forceOverwrite, conflicts);

            // b. Attendance entries: append-only delta upsert.
            entriesUpserted = UpsertAttendanceEntriesBulk(conn, trans, entrySource);

            // c. Date-scoped board + settings + daily attendance (untouched when only reference tables changed).
            if (writeBoard)
            {
                bool boardConflict = false;
                if (delta.HasDateBaseline && delta.BoardOrAttendanceChanged && !forceOverwrite)
                {
                    DateTime? serverMax = conn.ExecuteScalar<DateTime?>(SqlScheduleBoardMaxUpdatedAt, new { ScheduleDate = date }, trans);
                    if (serverMax.HasValue &&
                        (!delta.ExpectedBoardMaxUpdatedAt.HasValue ||
                         Math.Abs((serverMax.Value - delta.ExpectedBoardMaxUpdatedAt.Value).TotalSeconds) >= 0.001))
                    {
                        boardConflict = true;
                        conflicts.Add(new ConcurrencyConflict(
                            "daily_project_schedules",
                            $"Schedule {schedule.DailyScheduleDate:yyyy-MM-dd}",
                            $"Daily Schedule / Attendance for {schedule.DailyScheduleDate:dd/MM/yyyy} was modified on the server by another user.",
                            delta.ExpectedBoardMaxUpdatedAt,
                            serverMax,
                            ExcelTableName: "DailyAttendanceTable",
                            KeyColumnHeader: "",
                            RawKeyValue: schedule.DailyScheduleDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)));
                    }
                }

                if (!boardConflict) written = WriteScheduleDateScoped(conn, trans, schedule);
            }

            // d. Any conflict rolls back the whole batch (ExecuteInTransaction catches, rolls back, rethrows).
            if (conflicts.Count > 0) throw new ConcurrencyConflictException(conflicts);
        });

        static string Stats<T>(string name, TableDelta<T> d) =>
            $"{name} +{d.Added.Count}/~{d.Modified.Count}/fallback {d.FallbackUpsert.Count}/={d.UnchangedCount}";

        AppLogger.Info(
            $"[ServerDispatcher] Schedule delta saved {schedule.DailyScheduleDate:yyyy-MM-dd} (force={forceOverwrite}): " +
            $"{Stats("teams", delta.Teams)}; {Stats("prefs", delta.Preferences)}; {Stats("skills", delta.Skills)}; " +
            $"{Stats("projSkillReqs", delta.ProjectSkillRequirements)}; entries={entriesUpserted}; " +
            (writeBoard ? $"board slots={written.Slots}, attendances={written.Attendances}." : "board unchanged (skipped)."));
    }

    // =========================================================================
    // 5. SCHEDULE CONFLICT FIELD DIFFS ("Keep Server Values" cell write-back)
    // =========================================================================
    private static class ScheduleConflictDiff
    {
        private const string Empty = "(empty)";

        public static List<FieldConflict> ForPreference(IDbConnection conn, IDbTransaction trans, WorkerPreference local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM worker_preferences WHERE staff_id = @Key;", new { Key = local.StaffID });
            if (row == null) return diffs;

            AddList(diffs, "Team", Arr(row, "team"), local.Team);
            AddText(diffs, "Primary Skill", Text(row, "primary_skill"), local.PrimarySkill);
            AddList(diffs, "Other Skills", Arr(row, "other_skills"), local.OtherSkills);
            AddList(diffs, "Blacklisted Skills", Arr(row, "blacklisted_skills"), local.BlacklistedSkills);
            AddEnum(diffs, "Preferred Vehicle Type", Text(row, "preferred_vehicle_type"), local.PreferredVehicleType);
            AddText(diffs, "Dedicated Vehicle Plate", Text(row, "dedicated_vehicle_plate").ToUpperInvariant(),
                (local.DedicatedVehiclePlate ?? string.Empty).Trim().ToUpperInvariant());
            AddEnum(diffs, "Preferred Shift", Text(row, "preferred_shift"), local.PreferredShift);

            int serverMax = int.TryParse(Text(row, "max_days_per_week"), NumberStyles.Integer, CultureInfo.InvariantCulture, out int m) ? m : 0;
            Add(diffs, "Max days a week", serverMax.ToString(CultureInfo.InvariantCulture),
                Math.Max(0, local.MaxDaysPerWeek).ToString(CultureInfo.InvariantCulture), serverMax);

            AddEnum(diffs, "Rest Day Preference", Text(row, "rest_day_preference"), local.RestDayPreference);
            return diffs;
        }

        public static List<FieldConflict> ForTeam(IDbConnection conn, IDbTransaction trans, TeamRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, @"
                SELECT t.*, COALESCE(NULLIF(TRIM(s.short_name), ''), s.staff_name, '') AS leader_name
                FROM teams_master t LEFT JOIN staff_master s ON s.staff_id = t.team_leader_staff_id
                WHERE LOWER(t.team_name) = LOWER(@Key) LIMIT 1;",
                new { Key = (local.TeamName ?? string.Empty).Trim() });
            if (row == null) return diffs;

            long serverLeader = long.TryParse(Text(row, "team_leader_staff_id"), NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) ? id : 0;
            if (serverLeader != local.TeamLeaderStaffID)
            {
                string serverName = Text(row, "leader_name");
                if (serverName.Length == 0 && serverLeader > 0) serverName = serverLeader.ToString(CultureInfo.InvariantCulture);
                Add(diffs, "Team Leader", ShowText(serverName), ShowText(local.TeamLeader), serverName.Length == 0 ? null : serverName);
            }

            AddList(diffs, "Default Vehicles",
                Arr(row, "default_vehicles").Select(v => v.ToUpperInvariant()).ToList(),
                (local.DefaultVehicles ?? new List<string>()).Select(v => (v ?? string.Empty).ToUpperInvariant()).ToList());
            return diffs;
        }

        public static List<FieldConflict> ForSkill(IDbConnection conn, IDbTransaction trans, ScheduleSkillRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM lookup_schedule_skills WHERE LOWER(skill_name) = LOWER(@Key) LIMIT 1;",
                new { Key = (local.SkillName ?? string.Empty).Trim() });
            if (row == null) return diffs;

            AddText(diffs, "Display Icon", Text(row, "display_icon"), local.DisplayIcon);
            return diffs;
        }

        public static List<FieldConflict> ForProjectSkillRequirement(IDbConnection conn, IDbTransaction trans, ProjectSkillRequirement local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM project_skill_requirements WHERE project_type = @Key;",
                new { Key = local.ProjectType.ToString() });
            if (row == null) return diffs;

            AddList(diffs, "Required Skills", Arr(row, "required_skills"), local.Skills);
            static List<string> Vehicles(IEnumerable<string> raw) => raw
                .Select(v => Enum.TryParse<VehicleType>(v, ignoreCase: true, out var vt) ? vt : VehicleType.None)
                .Where(v => v != VehicleType.None)
                .Distinct()
                .Select(v => v.ToString())
                .ToList();
            AddList(diffs, "Required Vehicles", Vehicles(Arr(row, "required_vehicles")),
                Vehicles((local.Vehicles ?? new List<VehicleType>()).Select(v => v.ToString())));
            return diffs;
        }

        // --- Diff collection ---
        private static void Add(List<FieldConflict> diffs, string header, string serverDisplay, string localDisplay, object? serverExcelValue)
        {
            if (!string.Equals(serverDisplay, localDisplay, StringComparison.Ordinal))
                diffs.Add(new FieldConflict(header, serverDisplay, localDisplay, serverExcelValue));
        }

        private static void AddText(List<FieldConflict> diffs, string header, string? server, string? local)
        {
            string s = (server ?? string.Empty).Trim();
            Add(diffs, header, ShowText(s), ShowText(local), s.Length == 0 ? null : s);
        }

        /// <summary>Multi-select cells: comma-separated, blanks dropped.</summary>
        private static void AddList(List<FieldConflict> diffs, string header, IEnumerable<string>? server, IEnumerable<string>? local)
        {
            string s = string.Join(", ", CleanList(server));
            string l = string.Join(", ", CleanList(local));
            Add(diffs, header, ShowText(s), ShowText(l), s.Length == 0 ? null : s);
        }

        /// <summary>Unparseable / NULL server text falls back to <c>None</c> (shown as empty).</summary>
        private static void AddEnum<TEnum>(List<FieldConflict> diffs, string header, string serverRaw, TEnum local) where TEnum : struct, Enum
        {
            TEnum server = Enum.TryParse<TEnum>(serverRaw, ignoreCase: true, out var parsed) ? parsed
                : Enum.TryParse<TEnum>("None", ignoreCase: true, out var none) ? none : default;
            string s = server.ToString();
            bool isNone = string.Equals(s, "None", StringComparison.OrdinalIgnoreCase);
            Add(diffs, header, isNone ? Empty : s, ShowEnum(local), isNone ? null : s);
        }

        private static string ShowText(string? s) => string.IsNullOrWhiteSpace(s) ? Empty : s.Trim();

        private static string ShowEnum<TEnum>(TEnum v) where TEnum : struct, Enum
        {
            string s = v.ToString();
            return string.Equals(s, "None", StringComparison.OrdinalIgnoreCase) ? Empty : s;
        }

        // --- Server row access ---
        private static IDictionary<string, object>? QueryRow(IDbConnection conn, IDbTransaction trans, string sql, object param)
        {
            try
            {
                return (object?)conn.QueryFirstOrDefault(sql, param, trans) as IDictionary<string, object>;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[ServerDispatcher] Schedule conflict diff query failed: {ex.Message}");
                return null;
            }
        }

        private static object? Raw(IDictionary<string, object> row, string col) =>
            row.TryGetValue(col, out var v) && v is not null && v is not DBNull ? v : null;

        private static string Text(IDictionary<string, object> row, string col) =>
            (Convert.ToString(Raw(row, col), CultureInfo.InvariantCulture) ?? string.Empty).Trim();

        private static List<string> Arr(IDictionary<string, object> row, string col) => Raw(row, col) switch
        {
            string[] a => a.ToList(),
            IEnumerable<string> e => e.ToList(),
            _ => new List<string>()
        };
    }

    // =========================================================================
    // 6. ATTENDANCE ENTRY ROWS (bulk UNNEST, shared by the save path and the terminal sync)
    // =========================================================================
    private const string SqlEntryBulkInsertSelect = @"
        INSERT INTO attendance_entries (entry_date, entry_time, staff_id, project_id, entry_type, source, override_reason, is_approved, whatsapp_picture)
        SELECT t.entry_date::date, t.entry_time::time, t.staff_id, t.project_id, t.entry_type, t.source, t.override_reason, t.is_approved, t.whatsapp_picture
        FROM UNNEST(
            @EntryDates::text[], @EntryTimes::text[], @StaffIds::bigint[], @ProjectIds::text[], @EntryTypes::text[],
            @Sources::text[], @OverrideReasons::text[], @IsApproveds::boolean[], @WhatsappPictures::text[])
            AS t(entry_date, entry_time, staff_id, project_id, entry_type, source, override_reason, is_approved, whatsapp_picture)";

    private const string SqlEntryBulkUpsert = SqlEntryBulkInsertSelect + @"
        ON CONFLICT (entry_date, entry_time, staff_id, entry_type, source) DO UPDATE SET
            project_id       = EXCLUDED.project_id,
            override_reason  = EXCLUDED.override_reason,
            is_approved      = EXCLUDED.is_approved,
            whatsapp_picture = EXCLUDED.whatsapp_picture,
            updated_at       = CURRENT_TIMESTAMP
        WHERE attendance_entries.project_id       IS DISTINCT FROM EXCLUDED.project_id
           OR attendance_entries.override_reason  IS DISTINCT FROM EXCLUDED.override_reason
           OR attendance_entries.is_approved      IS DISTINCT FROM EXCLUDED.is_approved
           OR attendance_entries.whatsapp_picture IS DISTINCT FROM EXCLUDED.whatsapp_picture;";

    // DO NOTHING: terminal sync never overwrites manual is_approved / override_reason / project_id edits on stored punches.
    private const string SqlEntryBulkInsertIgnore = SqlEntryBulkInsertSelect + @"
        ON CONFLICT (entry_date, entry_time, staff_id, entry_type, source) DO NOTHING;";

    /// <summary>Inserts new attendance entries across any dates, skipping rows that already exist. Returns rows actually inserted.</summary>
    public static int InsertAttendanceEntries(IEnumerable<AttendanceEntry> entries)
    {
        EnsureScheduleSchema();
        var arrays = BuildEntryArrays(entries);
        if (arrays == null) return 0;

        int inserted = 0;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            inserted = conn.Execute(SqlEntryBulkInsertIgnore, arrays, trans);
        });
        AppLogger.Info($"[PostgreSQL Upload - Terminal Sync] inserted={inserted}, attempted={arrays.EntryDates.Length}");
        return inserted;
    }

    /// <summary>
    /// Upserts <paramref name="entries"/> in a single <c>INSERT ... SELECT FROM UNNEST(...)</c> statement.
    /// Unchanged rows are skipped by the <c>DO UPDATE ... WHERE IS DISTINCT FROM</c> guard. Returns rows inserted/updated.
    /// </summary>
    private static int UpsertAttendanceEntriesBulk(IDbConnection conn, IDbTransaction trans, IEnumerable<AttendanceEntry> entries)
    {
        var arrays = BuildEntryArrays(entries);
        return arrays == null ? 0 : conn.Execute(SqlEntryBulkUpsert, arrays, trans);
    }

    private sealed class EntryArrays
    {
        public string[] EntryDates { get; init; } = Array.Empty<string>();
        public string[] EntryTimes { get; init; } = Array.Empty<string>();
        public long[] StaffIds { get; init; } = Array.Empty<long>();
        public string[] ProjectIds { get; init; } = Array.Empty<string>();
        public string[] EntryTypes { get; init; } = Array.Empty<string>();
        public string[] Sources { get; init; } = Array.Empty<string>();
        public string[] OverrideReasons { get; init; } = Array.Empty<string>();
        public bool[] IsApproveds { get; init; } = Array.Empty<bool>();
        public string[] WhatsappPictures { get; init; } = Array.Empty<string>();
    }

    /// <summary>
    /// Projects entries into parallel UNNEST arrays, deduplicated on the conflict key
    /// (entry_date, entry_time, staff_id, entry_type, source) keeping the LAST occurrence, so a single statement never
    /// hits SQLSTATE 21000 ("ON CONFLICT DO UPDATE command cannot affect row a second time"). Null when empty.
    /// </summary>
    private static EntryArrays? BuildEntryArrays(IEnumerable<AttendanceEntry> entries)
    {
        var byKey = new Dictionary<(string Date, string Time, long StaffId, string EntryType, string Source), int>();
        var rows = new List<(string Date, string Time, long StaffId, string ProjectId, string EntryType, string Source, string OverrideReason, bool IsApproved, string WhatsappPicture)>();

        foreach (var e in entries)
        {
            if (e == null || e.StaffID <= 0) continue;

            var row = (
                Date: e.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                // Sub-second digits are kept (trimmed when zero) so the conflict key matches what the previous TimeSpan parameter stored.
                Time: e.Time.ToString("HH:mm:ss.FFFFFF", CultureInfo.InvariantCulture),
                StaffId: e.StaffID,
                ProjectId: (e.ProjectID ?? string.Empty).Trim().ToUpperInvariant(),
                EntryType: e.EntryType.ToString(),
                Source: e.Source.ToString(),
                OverrideReason: e.OverrideReason ?? string.Empty,
                IsApproved: e.IsApproved,
                WhatsappPicture: e.WhatsAppPicture ?? string.Empty);

            var key = (row.Date, row.Time, row.StaffId, row.EntryType, row.Source);
            if (byKey.TryGetValue(key, out int idx)) rows[idx] = row; // keep last occurrence
            else
            {
                byKey[key] = rows.Count;
                rows.Add(row);
            }
        }

        if (rows.Count == 0) return null;

        return new EntryArrays
        {
            EntryDates = rows.Select(r => r.Date).ToArray(),
            EntryTimes = rows.Select(r => r.Time).ToArray(),
            StaffIds = rows.Select(r => r.StaffId).ToArray(),
            ProjectIds = rows.Select(r => r.ProjectId).ToArray(),
            EntryTypes = rows.Select(r => r.EntryType).ToArray(),
            Sources = rows.Select(r => r.Source).ToArray(),
            OverrideReasons = rows.Select(r => r.OverrideReason).ToArray(),
            IsApproveds = rows.Select(r => r.IsApproved).ToArray(),
            WhatsappPictures = rows.Select(r => r.WhatsappPicture).ToArray()
        };
    }
}
