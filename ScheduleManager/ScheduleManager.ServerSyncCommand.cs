namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Storage;

internal static class ServerSyncCommands
{
    private const string Module = "ScheduleManager";

    // =========================================================================
    // 1. SAVE (local snapshot + transactional PostgreSQL dispatch)
    // =========================================================================
    internal static bool SaveScheduleSilent(ScheduleRecord schedule, out string errorMessage)
    {
        AppLogger.Info(
            $"[{Module}.SaveScheduleSilent] {schedule.DailyScheduleDate:yyyy-MM-dd}: slots={schedule.Projects.Count}, " +
            $"attendances={schedule.WorkerAttendance.Count}, entries={schedule.AttendanceEntries.Count}, " +
            $"prefs={schedule.Preferences.Count}, teams={schedule.Teams.Count}");

        // OfflineSnapshot.Save(record, writeToDatabase: true) -> ServerDispatcher.Dispatch -> PersistSchedule.
        return ServerSyncHelper.SaveSilent(schedule, out errorMessage);
    }

    // =========================================================================
    // 2. FETCH ONE WORKING DATE WITH OFFLINE FALLBACK
    // =========================================================================
    internal static ScheduleRecord FetchSchedule(DateOnly date)
    {
        var record = ServerSyncHelper.FetchWithOfflineFallback<ScheduleRecord>(
            $"{Module}.Fetch schedule {date:yyyy-MM-dd}",
            db => BuildFromServer(db, date),
            debugFormatter: r =>
                $"  date={r.DailyScheduleDate:yyyy-MM-dd} slots={r.Projects.Count} attendances={r.WorkerAttendance.Count} " +
                $"entries={r.AttendanceEntries.Count} prefs={r.Preferences.Count} teams={r.Teams.Count}");

        if (record == null) return new ScheduleRecord { DailyScheduleDate = date };

        // The local snapshot holds whichever date was fetched last: keep the reference tables, drop foreign-date rows.
        if (record.DailyScheduleDate != date)
        {
            AppLogger.Warn($"[{Module}] Offline snapshot is for {record.DailyScheduleDate:yyyy-MM-dd}; showing {date:yyyy-MM-dd} without date-scoped rows.");
            record = new ScheduleRecord { DailyScheduleDate = date, Preferences = record.Preferences, Teams = record.Teams };
        }

        DeriveLicensedSkills(record.Preferences);
        return record;
    }

    private static ScheduleRecord BuildFromServer(PostgreSqlService db, DateOnly date)
    {
        try { ServerDispatcher.EnsureScheduleSchema(); }
        catch (Exception ex) { AppLogger.Warn($"[{Module}] Schedule schema bootstrap skipped: {ex.Message}"); }

        DateTime d = date.ToDateTime(TimeOnly.MinValue);
        var record = new ScheduleRecord { DailyScheduleDate = date };

        foreach (var r in db.Query<dynamic>("SELECT * FROM daily_project_schedules WHERE schedule_date = @D ORDER BY shift, slot_index;", new { D = d }))
        {
            record.Projects.Add(new DailyProjectSchedule
            {
                DailyScheduleDate = date,
                Shift = Parse(r.shift, Shifts.Morning),
                SlotIndex = Convert.ToInt32(r.slot_index),
                ProjectID = Text(r.project_id),
                PremixHoles = Convert.ToInt32(r.premix_holes),
                Workers = Arr(r.workers),
                Vehicles = Arr(r.vehicles),
                Remarks = Text(r.remarks)
            });
        }

        foreach (var r in db.Query<dynamic>("SELECT * FROM attendance_entries WHERE entry_date = @D ORDER BY entry_time, entry_id;", new { D = d }))
        {
            record.AttendanceEntries.Add(new AttendanceEntry
            {
                EntryId = Convert.ToInt64(r.entry_id),
                Date = date,
                Time = Time(r.entry_time) ?? TimeOnly.MinValue,
                StaffID = Convert.ToInt64(r.staff_id),
                ProjectID = Text(r.project_id),
                EntryType = Parse(r.entry_type, AttendanceEntryType.CheckIn),
                Source = Parse(r.source, AttendanceSource.Manual),
                OverrideReason = Text(r.override_reason),
                IsApproved = (bool)r.is_approved,
                WhatsAppPicture = Text(r.whatsapp_picture)
            });
        }

        foreach (var r in db.Query<dynamic>("SELECT * FROM daily_attendances WHERE schedule_date = @D ORDER BY staff_id;", new { D = d }))
        {
            record.WorkerAttendance.Add(new DailyWorkerAttendance
            {
                DailyScheduleDate = date,
                StaffID = Convert.ToInt64(r.staff_id),
                AttendanceStatus = Parse(r.attendance_status, AttendanceStatus.Absence),
                ClockInTime = Time(r.clock_in_time),
                ClockOutTime = Time(r.clock_out_time),
                ProjectCheckIns = Text(r.project_check_ins),
                ProjectsDone = Arr(r.projects_done),
                RegularShift = ServerSyncHelper.Dec(r.regular_shift),
                AdditionalShifts = ServerSyncHelper.Dec(r.additional_shifts),
                OTHours = ServerSyncHelper.Dec(r.ot_hours),
                AdditionalOTHours = ServerSyncHelper.Dec(r.additional_ot_hours),
                TotalFlattenedHours = ServerSyncHelper.Dec(r.total_flattened_hours),
                ManualOverride = Parse(r.manual_override, AttendanceStatus.None),
                OverrideReason = Text(r.override_reason),
                IsAllocated = (bool)r.is_allocated
            });
        }

        foreach (var r in db.Query<dynamic>("SELECT * FROM worker_preferences ORDER BY staff_id;"))
        {
            record.Preferences.Add(new WorkerPreference
            {
                StaffID = Convert.ToInt64(r.staff_id),
                Team = Arr(r.team),
                PrimarySkill = Text(r.primary_skill),
                OtherSkills = Arr(r.other_skills),
                BlacklistedSkills = Arr(r.blacklisted_skills),
                PreferredVehicleType = Parse(r.preferred_vehicle_type, VehicleType.None),
                DedicatedVehiclePlate = Text(r.dedicated_vehicle_plate),
                PreferredShift = Parse(r.preferred_shift, Shifts.None),
                MaxDaysPerWeek = Convert.ToInt32(r.max_days_per_week),
                RestDayPreference = Parse(r.rest_day_preference, Days.None)
            });
        }

        foreach (var r in db.Query<dynamic>("SELECT * FROM teams_master ORDER BY team_name;"))
        {
            object? leader = r.team_leader_staff_id;
            record.Teams.Add(new TeamRecord
            {
                TeamName = Text(r.team_name),
                TeamLeaderStaffID = leader != null ? Convert.ToInt64(leader) : 0,
                DefaultVehicles = Arr(r.default_vehicles)
            });
        }

        // Licensed skills are derived from StaffManager data, never stored with the preference.
        var vehicleLicenses = db.Query<dynamic>("SELECT staff_id, vehicle_license_class, expiry_date, status FROM staff_vehicle_licenses;")
            .Select(r => new VehicleLicense
            {
                StaffID = Convert.ToInt64(r.staff_id),
                VehicleLicenseClass = Parse(r.vehicle_license_class, VehicleLicenseClass.None),
                ExpiryDate = DateOnly.FromDateTime(Convert.ToDateTime(r.expiry_date)),
                VehicleLicenseStatus = Parse(r.status, VehicleLicenseStatus.Invalid)
            }).ToList();
        var skillLicenses = db.Query<dynamic>("SELECT staff_id, skill_license_type, expiry_date, status FROM staff_skill_licenses;")
            .Select(r => new SkillLicense
            {
                StaffID = Convert.ToInt64(r.staff_id),
                SkillLicenseType = Text(r.skill_license_type),
                ExpiryDate = DateOnly.FromDateTime(Convert.ToDateTime(r.expiry_date)),
                SkillLicenseStatus = Parse(r.status, SkillLicenseStatus.Invalid)
            }).ToList();
        ApplyLicensedSkills(record.Preferences, vehicleLicenses, skillLicenses);

        return record;
    }

    // =========================================================================
    // 3. LICENSED SKILL DERIVATION (StaffManager licences -> WorkerPreference.LicensedSkills)
    // =========================================================================
    /// <summary>Offline path: re-derives LicensedSkills from the local StaffRecord snapshot.</summary>
    private static void DeriveLicensedSkills(List<WorkerPreference> prefs)
    {
        if (prefs.Count == 0 || prefs.All(p => p.LicensedSkills.Count > 0)) return;
        var staff = OfflineSnapshot.Load<List<StaffRecord>>();
        if (staff == null || staff.Count == 0) return;

        ApplyLicensedSkills(
            prefs,
            staff.SelectMany(s => s.VehicleLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })),
            staff.SelectMany(s => s.SkillLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })));
    }

    /// <summary>
    /// A licence counts when its status is Valid and it has not expired. Any valid vehicle licence class grants
    /// "Driver" plus "License &lt;class&gt;"; a valid skill licence grants its type name.
    /// </summary>
    internal static void ApplyLicensedSkills(
        IEnumerable<WorkerPreference> prefs,
        IEnumerable<VehicleLicense> vehicleLicenses,
        IEnumerable<SkillLicense> skillLicenses)
    {
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);
        var skillsByStaff = new Dictionary<long, List<string>>();

        void Add(long staffId, string skill)
        {
            if (!skillsByStaff.TryGetValue(staffId, out var list)) skillsByStaff[staffId] = list = new List<string>();
            if (!list.Contains(skill, StringComparer.OrdinalIgnoreCase)) list.Add(skill);
        }

        foreach (var l in vehicleLicenses.Where(l => l.VehicleLicenseStatus == VehicleLicenseStatus.Valid && l.ExpiryDate >= today && l.VehicleLicenseClass != VehicleLicenseClass.None))
        {
            Add(l.StaffID, "Driver");
            Add(l.StaffID, $"License {l.VehicleLicenseClass.ToString().Replace('_', ' ')}");
        }
        foreach (var l in skillLicenses.Where(l => l.SkillLicenseStatus == SkillLicenseStatus.Valid && l.ExpiryDate >= today && !string.IsNullOrWhiteSpace(l.SkillLicenseType)))
            Add(l.StaffID, l.SkillLicenseType.Trim());

        foreach (var p in prefs)
            p.LicensedSkills = skillsByStaff.TryGetValue(p.StaffID, out var s) ? s : new List<string>();
    }

    // =========================================================================
    // 4. LOOKUP LISTS FOR THE 'Data' SHEET
    // =========================================================================
    internal static void SyncLookupLists()
    {
        Roster.Invalidate();

        // StaffIDList and StaffNameList stay index-aligned (Staff ID <-> Staff Name sync).
        var staff = Roster.Staff.Where(s => s.Id > 0 && s.ShortName.Length > 0).OrderBy(s => s.Id).ToList();
        if (staff.Count > 0)
        {
            DataList.SyncValues("StaffIDList", staff.Select(s => s.Id.ToString()).ToList());
            DataList.SyncValues("StaffNameList", staff.Select(s => s.ShortName).ToList());
        }

        var plates = Roster.Vehicles.Select(v => v.Plate).Where(p => p.Length > 0).Distinct().OrderBy(p => p).ToList();
        if (plates.Count > 0) DataList.SyncValues("VehiclePlateList", plates);

        var projects = Roster.Projects.Select(p => p.Id).Where(p => p.Length > 0).Distinct().ToList();
        if (projects.Count > 0) DataList.SyncValues("ProjectIDList", projects);
    }

    // =========================================================================
    // 5. DYNAMIC ROW HELPERS
    // =========================================================================
    private static string Text(object? v) => v?.ToString() ?? string.Empty;

    private static List<string> Arr(object? v) => v switch
    {
        string[] a => a.ToList(),
        IEnumerable<string> e => e.ToList(),
        _ => new List<string>()
    };

    private static TimeOnly? Time(object? v) => v switch
    {
        TimeSpan ts => TimeOnly.FromTimeSpan(ts),
        TimeOnly t => t,
        DateTime dt => TimeOnly.FromDateTime(dt),
        _ => null
    };

    private static TEnum Parse<TEnum>(object? v, TEnum fallback) where TEnum : struct, Enum =>
        Enum.TryParse<TEnum>(v?.ToString(), true, out var parsed) ? parsed : fallback;
}
