namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using Dapper;
using SKSB_App.Core.Diagnostics;
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

        Db.ExecuteInTransaction((conn, trans) => conn.Execute(ddl, transaction: trans));
        _scheduleSchemaEnsured = true;
    }

    private static string[] CleanList(IEnumerable<string>? items) =>
        (items ?? Enumerable.Empty<string>()).Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).ToArray();

    // =========================================================================
    // 2. SCHEDULE PERSISTENCE (single transaction)
    // =========================================================================
    /// <summary>
    /// Persists one working date plus the global reference tables:
    /// date-scoped board slots and daily attendances are reconciled for <see cref="ScheduleRecord.DailyScheduleDate"/>;
    /// worker preferences and teams are full-batch; attendance entries are append-only (upsert, never deleted -
    /// a wrong punch is voided by clearing "Is Approved").
    /// </summary>
    public static void PersistSchedule(ScheduleRecord schedule)
    {
        EnsureScheduleSchema();

        DateTime date = schedule.DailyScheduleDate.ToDateTime(TimeOnly.MinValue);

        // ---- board slots ----------------------------------------------------
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

        // ---- daily attendances ---------------------------------------------
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

        // ---- attendance entries (append-only upsert) ------------------------
        var entries = schedule.AttendanceEntries.Where(e => e.StaffID > 0).Select(ToEntryRow).ToList();

        // ---- preferences & teams (full batch) ------------------------------
        var prefs = schedule.Preferences
            .Where(p => p.StaffID > 0)
            .GroupBy(p => p.StaffID)
            .Select(g => g.First())
            .Select(p => new
            {
                StaffId = p.StaffID,
                Team = CleanList(p.Team),
                PrimarySkill = p.PrimarySkill ?? string.Empty,
                OtherSkills = CleanList(p.OtherSkills),
                BlacklistedSkills = CleanList(p.BlacklistedSkills),
                PreferredVehicleType = p.PreferredVehicleType.ToString(),
                DedicatedVehiclePlate = (p.DedicatedVehiclePlate ?? string.Empty).Trim().ToUpperInvariant(),
                PreferredShift = p.PreferredShift.ToString(),
                MaxDaysPerWeek = Math.Max(0, p.MaxDaysPerWeek),
                RestDayPreference = p.RestDayPreference.ToString()
            }).ToList();
        long[] activePrefIds = prefs.Select(p => p.StaffId).ToArray();

        var teams = schedule.Teams
            .Where(t => !string.IsNullOrWhiteSpace(t.TeamName))
            .GroupBy(t => t.TeamName.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .Select(t => new
            {
                TeamName = t.TeamName.Trim(),
                TeamLeaderStaffId = t.TeamLeaderStaffID > 0 ? (long?)t.TeamLeaderStaffID : null,
                DefaultVehicles = CleanList(t.DefaultVehicles).Select(v => v.ToUpperInvariant()).ToArray()
            }).ToList();
        string[] activeTeamNames = teams.Select(t => t.TeamName).ToArray();

        var skills = (schedule.Skills ?? new List<ScheduleSkillRecord>())
            .Where(s => !string.IsNullOrWhiteSpace(s.SkillName))
            .GroupBy(s => s.SkillName.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .Select(s => new
            {
                SkillName = s.SkillName.Trim(),
                DisplayIcon = (s.DisplayIcon ?? string.Empty).Trim()
            }).ToList();
        string[] activeSkillNames = skills.Select(s => s.SkillName).ToArray();

        var projSkillReqs = (schedule.ProjectSkillRequirements ?? new List<ProjectSkillRequirement>())
            .Where(r => r.ProjectType != ProjectType.None)
            .GroupBy(r => r.ProjectType)
            .Select(g => g.First())
            .Select(r => new
            {
                ProjectType = r.ProjectType.ToString(),
                RequiredSkills = CleanList(r.Skills),
                RequiredVehicles = (r.Vehicles ?? new List<VehicleType>())
                    .Where(v => v != VehicleType.None)
                    .Distinct()
                    .Select(v => v.ToString())
                    .ToArray()
            }).ToList();
        string[] activeProjectTypes = projSkillReqs.Select(r => r.ProjectType).ToArray();

        const string sqlSlotUpsert = @"
        INSERT INTO daily_project_schedules (schedule_date, shift, slot_index, project_id, premix_holes, workers, vehicles, remarks)
        VALUES (@ScheduleDate, @Shift, @SlotIndex, @ProjectId, @PremixHoles, @Workers, @Vehicles, @Remarks)
        ON CONFLICT (schedule_date, shift, slot_index) DO UPDATE SET
            project_id   = EXCLUDED.project_id,
            premix_holes = EXCLUDED.premix_holes,
            workers      = EXCLUDED.workers,
            vehicles     = EXCLUDED.vehicles,
            remarks      = EXCLUDED.remarks,
            updated_at   = CURRENT_TIMESTAMP;";

        const string sqlAttendanceUpsert = @"
        INSERT INTO daily_attendances (
            schedule_date, staff_id, attendance_status, clock_in_time, clock_out_time, project_check_ins, projects_done,
            regular_shift, additional_shifts, ot_hours, additional_ot_hours, total_flattened_hours,
            is_allocated)
        VALUES (
            @ScheduleDate, @StaffId, @AttendanceStatus, @ClockInTime, @ClockOutTime, @ProjectCheckIns, @ProjectsDone,
            @RegularShift, @AdditionalShifts, @OtHours, @AdditionalOtHours, @TotalFlattenedHours,
            @IsAllocated)
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
            updated_at            = CURRENT_TIMESTAMP;";

        const string sqlPrefUpsert = @"
        INSERT INTO worker_preferences (
            staff_id, team, primary_skill, other_skills, blacklisted_skills, preferred_vehicle_type,
            dedicated_vehicle_plate, preferred_shift, max_days_per_week, rest_day_preference)
        VALUES (
            @StaffId, @Team, @PrimarySkill, @OtherSkills, @BlacklistedSkills, @PreferredVehicleType,
            @DedicatedVehiclePlate, @PreferredShift, @MaxDaysPerWeek, @RestDayPreference)
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
            updated_at              = CURRENT_TIMESTAMP;";

        const string sqlTeamUpsert = @"
        INSERT INTO teams_master (team_name, team_leader_staff_id, default_vehicles)
        VALUES (@TeamName, @TeamLeaderStaffId, @DefaultVehicles)
        ON CONFLICT (team_name) DO UPDATE SET
            team_leader_staff_id = EXCLUDED.team_leader_staff_id,
            default_vehicles     = EXCLUDED.default_vehicles,
            updated_at           = CURRENT_TIMESTAMP;";

        const string sqlSkillUpsert = @"
        INSERT INTO lookup_schedule_skills (skill_name, display_icon)
        VALUES (@SkillName, @DisplayIcon)
        ON CONFLICT (skill_name) DO UPDATE SET
            display_icon = EXCLUDED.display_icon,
            updated_at   = CURRENT_TIMESTAMP;";

        const string sqlProjectSkillReqUpsert = @"
        INSERT INTO project_skill_requirements (project_type, required_skills, required_vehicles)
        VALUES (@ProjectType, @RequiredSkills, @RequiredVehicles)
        ON CONFLICT (project_type) DO UPDATE SET
            required_skills   = EXCLUDED.required_skills,
            required_vehicles = EXCLUDED.required_vehicles,
            updated_at        = CURRENT_TIMESTAMP;";

        const string sqlSettingsUpsert = @"
        INSERT INTO daily_schedule_settings (schedule_date, late_cut_off_time, ot_start_time)
        VALUES (@ScheduleDate, @LateCutOffTime, @OtStartTime)
        ON CONFLICT (schedule_date) DO UPDATE SET
            late_cut_off_time = EXCLUDED.late_cut_off_time,
            ot_start_time     = EXCLUDED.ot_start_time,
            updated_at        = CURRENT_TIMESTAMP;";

        var settings = new
        {
            ScheduleDate = date,
            LateCutOffTime = schedule.LateCutOffTime.ToTimeSpan(),
            OtStartTime = schedule.OTStartTime.ToTimeSpan()
        };

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute(sqlSettingsUpsert, settings, trans);

            // Date-scoped reconciliation: removed slots / workers disappear for THIS date only.
            conn.Execute(
                "DELETE FROM daily_project_schedules WHERE schedule_date = @ScheduleDate AND (shift || ':' || slot_index::text) != ALL(@ActiveKeys);",
                new { ScheduleDate = date, ActiveKeys = activeSlotKeys }, trans);
            conn.Execute(
                "DELETE FROM daily_attendances WHERE schedule_date = @ScheduleDate AND staff_id != ALL(@ActiveIds);",
                new { ScheduleDate = date, ActiveIds = activeAttendanceIds }, trans);

            // Reference tables are full-batch, but an empty sheet never wipes the server copy.
            if (activePrefIds.Length > 0)
                conn.Execute("DELETE FROM worker_preferences WHERE staff_id != ALL(@ActiveIds);", new { ActiveIds = activePrefIds }, trans);
            if (activeTeamNames.Length > 0)
                conn.Execute("DELETE FROM teams_master WHERE team_name != ALL(@ActiveNames);", new { ActiveNames = activeTeamNames }, trans);
            if (activeSkillNames.Length > 0)
            {
                conn.Execute("DELETE FROM lookup_schedule_skills WHERE skill_name != ALL(@ActiveNames);", new { ActiveNames = activeSkillNames }, trans);
                conn.Execute(sqlSkillUpsert, skills, trans);
            }
            if (activeProjectTypes.Length > 0)
            {
                conn.Execute("DELETE FROM project_skill_requirements WHERE project_type != ALL(@ActiveTypes);", new { ActiveTypes = activeProjectTypes }, trans);
                conn.Execute(sqlProjectSkillReqUpsert, projSkillReqs, trans);
            }

            if (teams.Count > 0) conn.Execute(sqlTeamUpsert, teams, trans);
            if (prefs.Count > 0) conn.Execute(sqlPrefUpsert, prefs, trans);
            if (slots.Count > 0) conn.Execute(sqlSlotUpsert, slots, trans);
            if (entries.Count > 0) conn.Execute(SqlEntryUpsert, entries, trans);
            if (attendances.Count > 0) conn.Execute(sqlAttendanceUpsert, attendances, trans);
        });

        AppLogger.Info(
            $"[PostgreSQL Upload - Schedule {schedule.DailyScheduleDate:yyyy-MM-dd}] " +
            $"lateCutOff={schedule.LateCutOffTime:HH:mm}, otStart={schedule.OTStartTime:HH:mm}, " +
            $"slots={slots.Count}, attendances={attendances.Count}, entries={entries.Count}, prefs={prefs.Count}, teams={teams.Count}, skills={skills.Count}, projSkillReqs={projSkillReqs.Count}");
    }

    // =========================================================================
    // 3. ATTENDANCE ENTRY ROWS (shared by the save path and the terminal sync)
    // =========================================================================
    private const string SqlEntryUpsert = @"
        INSERT INTO attendance_entries (entry_date, entry_time, staff_id, project_id, entry_type, source, override_reason, is_approved, whatsapp_picture)
        VALUES (@EntryDate, @EntryTime, @StaffId, @ProjectId, @EntryType, @Source, @OverrideReason, @IsApproved, @WhatsappPicture)
        ON CONFLICT (entry_date, entry_time, staff_id, entry_type, source) DO UPDATE SET
            project_id       = EXCLUDED.project_id,
            override_reason  = EXCLUDED.override_reason,
            is_approved      = EXCLUDED.is_approved,
            whatsapp_picture = EXCLUDED.whatsapp_picture,
            updated_at       = CURRENT_TIMESTAMP;";

    private static object ToEntryRow(AttendanceEntry e) => new
    {
        EntryDate = e.Date.ToDateTime(TimeOnly.MinValue),
        EntryTime = e.Time.ToTimeSpan(),
        StaffId = e.StaffID,
        ProjectId = (e.ProjectID ?? string.Empty).Trim().ToUpperInvariant(),
        EntryType = e.EntryType.ToString(),
        Source = e.Source.ToString(),
        OverrideReason = e.OverrideReason ?? string.Empty,
        e.IsApproved,
        WhatsappPicture = e.WhatsAppPicture ?? string.Empty
    };
}
