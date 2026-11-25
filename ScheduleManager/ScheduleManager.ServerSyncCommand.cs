namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Storage;

internal static class ServerSyncCommands
{
    private const string Module = "ScheduleManager";

    // =========================================================================
    // 1. SAVE (delta-only, optimistic concurrency, single PostgreSQL transaction)
    // =========================================================================
    /// <summary>
    /// Non-interactive save (startup pending upload / contracts without conflict checks). Without baselines every
    /// reference row is upserted and the board is written; with baselines only changed rows are written, overwriting the server.
    /// </summary>
    internal static bool SaveScheduleSilent(ScheduleRecord schedule, out string errorMessage) =>
        SaveScheduleWithDelta(schedule, forceOverwrite: true, out errorMessage, out _);

    /// <summary>
    /// Delta save: only reference rows that differ from the last server load are written (never hard-deleted); the
    /// date-scoped board / settings / daily attendance are written only when they changed. Modified rows and the
    /// board are checked against their server <c>updated_at</c>; on conflict nothing is saved and
    /// <paramref name="conflicts"/> lists the affected rows.
    /// </summary>
    internal static bool SaveScheduleWithDelta(ScheduleRecord schedule, bool forceOverwrite, out string errorMessage, out IReadOnlyList<ConcurrencyConflict>? conflicts)
    {
        errorMessage = string.Empty;
        conflicts = null;
        try
        {
            DateOnly date = schedule.DailyScheduleDate;
            bool hasDateBaseline = _baselineDate == date && _baselineBoardFingerprint != null;
            bool boardChanged = !hasDateBaseline ||
                                !string.Equals(_baselineBoardFingerprint, BoardFingerprint(schedule), StringComparison.Ordinal);

            var payload = new ScheduleDeltaPayload
            {
                Schedule = schedule,
                HasDateBaseline = hasDateBaseline,
                BoardOrAttendanceChanged = boardChanged,
                ExpectedBoardMaxUpdatedAt = hasDateBaseline ? _baselineBoardMaxUpdatedAt : null,
                Preferences = PrefBaseline.Diff(schedule.Preferences.Where(p => p.StaffID > 0), p => p.StaffID, PrefFingerprint),
                Teams = TeamBaseline.Diff(schedule.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)), TeamKey, TeamFingerprint),
                Skills = SkillBaseline.Diff((schedule.Skills ?? new List<ScheduleSkillRecord>()).Where(s => !string.IsNullOrWhiteSpace(s.SkillName)), SkillKey, SkillFingerprint),
                ProjectSkillRequirements = ProjSkillReqBaseline.Diff(
                    (schedule.ProjectSkillRequirements ?? new List<ProjectSkillRequirement>()).Where(r => r.ProjectType != ProjectType.None),
                    r => r.ProjectType, ProjSkillReqFingerprint)
            };
            bool boardWritten = !payload.HasDateBaseline || payload.BoardOrAttendanceChanged || forceOverwrite;

            AppLogger.Info(
                $"[{Module}.SaveScheduleWithDelta] {date:yyyy-MM-dd} force={forceOverwrite}: dateBaseline={hasDateBaseline}, boardChanged={boardChanged}, " +
                $"slots={schedule.Projects.Count}, attendances={schedule.WorkerAttendance.Count}, entries={schedule.AttendanceEntries.Count}, " +
                $"prefs={schedule.Preferences.Count}, teams={schedule.Teams.Count}, skills={schedule.Skills?.Count ?? 0}, projSkillReqs={schedule.ProjectSkillRequirements?.Count ?? 0}");

            // 1. PostgreSQL: single transaction, rolled back on any conflict.
            ServerDispatcher.PersistScheduleDelta(payload, forceOverwrite);

            // 2. Local offline snapshot (only after the server commit succeeded).
            OfflineSnapshot.Save(schedule, writeToDatabase: false);

            // 3. Move the baselines to what was just written (only the written rows / board, so concurrent server
            //    edits to rows this workbook never touched are still detected on the next save).
            RefreshBaselinesAfterSave(payload, boardWritten);
            return true;
        }
        catch (ConcurrencyConflictException cex)
        {
            conflicts = cex.Conflicts;
            errorMessage = cex.Message;
            AppLogger.Warn($"[{Module}] Delta save rolled back: {cex.Conflicts.Count} concurrency conflict(s).");
            return false;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error($"[{Module}] SaveScheduleWithDelta failed", ex);
            return false;
        }
    }

    private static void RefreshBaselinesAfterSave(ScheduleDeltaPayload payload, bool boardWritten)
    {
        try
        {
            var db = new PostgreSqlService();

            static List<T> Written<T>(TableDelta<T> d) =>
                d.Added.Concat(d.Modified.Select(m => m.Record)).Concat(d.FallbackUpsert).ToList();

            // key (normalised exactly like the baseline key) -> server updated_at
            Dictionary<string, DateTime?> UpdatedAtByKey(string sql, object param, string keyColumn)
            {
                var map = new Dictionary<string, DateTime?>(StringComparer.Ordinal);
                foreach (var row in db.Query<dynamic>(sql, param))
                {
                    var r = (IDictionary<string, object>)row;
                    string key = Norm(Convert.ToString(r[keyColumn], CultureInfo.InvariantCulture)).ToUpperInvariant();
                    map[key] = ToUtcDateTime(r["updated_at"]);
                }
                return map;
            }

            var prefs = Written(payload.Preferences);
            if (prefs.Count > 0)
            {
                var ts = UpdatedAtByKey("SELECT staff_id, updated_at FROM worker_preferences WHERE staff_id = ANY(@Ids);",
                    new { Ids = prefs.Select(p => p.StaffID).Distinct().ToArray() }, "staff_id");
                foreach (var p in prefs)
                    if (ts.TryGetValue(p.StaffID.ToString(CultureInfo.InvariantCulture), out DateTime? u)) PrefBaseline.UpsertEntry(p.StaffID, PrefFingerprint(p), u);
            }

            var teams = Written(payload.Teams);
            if (teams.Count > 0)
            {
                var ts = UpdatedAtByKey("SELECT team_name, updated_at FROM teams_master WHERE UPPER(TRIM(team_name)) = ANY(@Keys);",
                    new { Keys = teams.Select(TeamKey).Distinct().ToArray() }, "team_name");
                foreach (var t in teams)
                    if (ts.TryGetValue(TeamKey(t), out DateTime? u)) TeamBaseline.UpsertEntry(TeamKey(t), TeamFingerprint(t), u);
            }

            var skills = Written(payload.Skills);
            if (skills.Count > 0)
            {
                var ts = UpdatedAtByKey("SELECT skill_name, updated_at FROM lookup_schedule_skills WHERE UPPER(TRIM(skill_name)) = ANY(@Keys);",
                    new { Keys = skills.Select(SkillKey).Distinct().ToArray() }, "skill_name");
                foreach (var s in skills)
                    if (ts.TryGetValue(SkillKey(s), out DateTime? u)) SkillBaseline.UpsertEntry(SkillKey(s), SkillFingerprint(s), u);
            }

            var reqs = Written(payload.ProjectSkillRequirements);
            if (reqs.Count > 0)
            {
                var ts = UpdatedAtByKey("SELECT project_type, updated_at FROM project_skill_requirements WHERE project_type = ANY(@Keys);",
                    new { Keys = reqs.Select(r => r.ProjectType.ToString()).Distinct().ToArray() }, "project_type");
                foreach (var r in reqs)
                    if (ts.TryGetValue(r.ProjectType.ToString().ToUpperInvariant(), out DateTime? u)) ProjSkillReqBaseline.UpsertEntry(r.ProjectType, ProjSkillReqFingerprint(r), u);
            }

            if (boardWritten)
            {
                var schedule = payload.Schedule;
                _baselineDate = schedule.DailyScheduleDate;
                _baselineBoardFingerprint = BoardFingerprint(schedule);
                _baselineBoardMaxUpdatedAt = QueryBoardMaxUpdatedAt(db, schedule.DailyScheduleDate);
                schedule.BoardServerUpdatedAt = _baselineBoardMaxUpdatedAt;
            }
        }
        catch (Exception ex)
        {
            // The save itself committed: a stale baseline only costs an extra (harmless) concurrency check next time.
            AppLogger.Warn($"[{Module}] Post-save baseline refresh failed: {ex.Message}");
        }
    }

    // =========================================================================
    // 1b. DELTA BASELINES & FINGERPRINTS
    // =========================================================================
    // Rows as last loaded from (or saved to) the server; diffing the sheet against them yields the delta to save.
    private static readonly TableBaseline<long> PrefBaseline = new();
    private static readonly TableBaseline<string> TeamBaseline = new();
    private static readonly TableBaseline<string> SkillBaseline = new();
    private static readonly TableBaseline<ProjectType> ProjSkillReqBaseline = new();

    // Date-scoped board / settings / daily attendance baseline (one working date at a time).
    private static DateOnly? _baselineDate;
    private static string? _baselineBoardFingerprint;
    private static DateTime? _baselineBoardMaxUpdatedAt;

    private const string Sep = "\u001F";

    private static string Norm(string? s) => (s ?? string.Empty).Trim();
    private static string NormList(IEnumerable<string>? items, bool upper = false) =>
        string.Join(",", (items ?? Enumerable.Empty<string>())
            .Where(s => !string.IsNullOrWhiteSpace(s))
            .Select(s => upper ? s.Trim().ToUpperInvariant() : s.Trim()));
    private static string NormDec(decimal d) => decimal.Round(d, 2).ToString("F2", CultureInfo.InvariantCulture);

    /// <summary>Whole seconds (Excel serial times carry floating-point noise).</summary>
    private static string NormTime(TimeOnly? t) =>
        t.HasValue ? ((long)Math.Round(t.Value.ToTimeSpan().TotalSeconds)).ToString(CultureInfo.InvariantCulture) : string.Empty;

    private static string TeamKey(TeamRecord t) => Norm(t.TeamName).ToUpperInvariant();
    private static string SkillKey(ScheduleSkillRecord s) => Norm(s.SkillName).ToUpperInvariant();

    private static string PrefFingerprint(WorkerPreference p) => string.Join(Sep,
        p.StaffID.ToString(CultureInfo.InvariantCulture),
        NormList(p.Team),
        Norm(p.PrimarySkill),
        NormList(p.OtherSkills),
        NormList(p.BlacklistedSkills),
        p.PreferredVehicleType.ToString(),
        Norm(p.DedicatedVehiclePlate).ToUpperInvariant(),
        p.PreferredShift.ToString(),
        Math.Max(0, p.MaxDaysPerWeek).ToString(CultureInfo.InvariantCulture),
        p.RestDayPreference.ToString());

    private static string TeamFingerprint(TeamRecord t) => string.Join(Sep,
        Norm(t.TeamName),
        t.TeamLeaderStaffID.ToString(CultureInfo.InvariantCulture),
        NormList(t.DefaultVehicles, upper: true));

    private static string SkillFingerprint(ScheduleSkillRecord s) => string.Join(Sep,
        Norm(s.SkillName),
        Norm(s.DisplayIcon));

    private static string ProjSkillReqFingerprint(ProjectSkillRequirement r) => string.Join(Sep,
        r.ProjectType.ToString(),
        NormList(r.Skills),
        string.Join(",", (r.Vehicles ?? new List<VehicleType>()).Where(v => v != VehicleType.None).Distinct().Select(v => v.ToString())));

    /// <summary>
    /// Fingerprint of everything written for one working date: settings, persisted board slots (same filter as the
    /// dispatcher) and daily attendance rows. Derived / display-only fields (project names, staff names, ProjectsDone,
    /// Total Flattened Hours, Next Day Start Time) are excluded.
    /// </summary>
    private static string BoardFingerprint(ScheduleRecord s)
    {
        var parts = new List<string>
        {
            s.DailyScheduleDate.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            NormTime(s.DayStartTime), NormTime(s.LateCutOffTime), NormTime(s.OTStartTime)
        };

        foreach (var p in s.Projects
                     .Where(p => p.SlotIndex >= 0 && (p.Workers.Count > 0 || p.Vehicles.Count > 0 || p.PremixHoles > 0))
                     .OrderBy(p => p.Shift.ToString(), StringComparer.Ordinal).ThenBy(p => p.SlotIndex))
        {
            parts.Add(string.Join("|",
                p.Shift.ToString(), p.SlotIndex.ToString(CultureInfo.InvariantCulture),
                Norm(p.ProjectID).ToUpperInvariant(), Math.Max(0, p.PremixHoles).ToString(CultureInfo.InvariantCulture),
                NormList(p.Workers), NormList(p.Vehicles, upper: true), Norm(p.Remarks)));
        }

        foreach (var a in s.WorkerAttendance.Where(a => a.StaffID > 0).OrderBy(a => a.StaffID))
        {
            parts.Add(string.Join("|",
                a.StaffID.ToString(CultureInfo.InvariantCulture), a.AttendanceStatus.ToString(),
                NormTime(a.ClockInTime), NormTime(a.ClockOutTime), Norm(a.ProjectCheckIns),
                NormDec(a.RegularShift), NormDec(a.AdditionalShifts), NormDec(a.OTHours), NormDec(a.AdditionalOTHours),
                a.IsAllocated ? "1" : "0"));
        }

        return string.Join(Sep, parts);
    }

    private static DateTime? QueryBoardMaxUpdatedAt(PostgreSqlService db, DateOnly date) =>
        ToUtcDateTime(db.Query<DateTime?>(ServerDispatcher.SqlScheduleBoardMaxUpdatedAt, new { ScheduleDate = date.ToDateTime(TimeOnly.MinValue) })
            .FirstOrDefault());

    /// <summary>
    /// Server <c>updated_at</c> as a concurrency token. Kept exactly as Npgsql returned it (UTC for <c>timestamptz</c>,
    /// Unspecified for <c>timestamp</c>) so it round-trips against the column's own type.
    /// </summary>
    private static DateTime? ToUtcDateTime(object? raw) => raw switch
    {
        DateTime dt => dt.Kind == DateTimeKind.Local ? DateTime.SpecifyKind(dt.ToUniversalTime(), DateTimeKind.Utc) : dt,
        DateTimeOffset dto => DateTime.SpecifyKind(dto.UtcDateTime, DateTimeKind.Utc),
        _ => null
    };

    // =========================================================================
    // 1c. CONCURRENCY CONFLICT RESOLUTION ("Keep Server Values")
    // =========================================================================
    /// <summary>
    /// After the pane wrote the server's values into the conflicting cells, re-reads ONLY the conflicting reference
    /// rows and moves their baselines to the server's fingerprint + <c>updated_at</c>. A board conflict re-fetches the
    /// date and re-renders the board + Daily Attendances (never the AttendanceEntriesTable) from the server copy.
    /// Runs on the Excel macro thread (the pane wraps the action).
    /// </summary>
    internal static void AcceptServerConflictBaselines(IReadOnlyList<ConcurrencyConflict> conflicts)
    {
        if (conflicts == null || conflicts.Count == 0) return;
        if (!ServerEnvironment.IsOnline)
        {
            AppLogger.Warn($"[{Module}] Keep Server Values: server offline, conflict baselines not refreshed.");
            return;
        }

        List<string> KeysFor(string table) => conflicts
            .Where(c => string.Equals(c.TableName, table, StringComparison.OrdinalIgnoreCase))
            .Select(c => (c.RawKeyValue ?? string.Empty).Trim())
            .Where(k => k.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        long[] prefIds = KeysFor("worker_preferences")
            .Select(k => long.TryParse(k, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) ? id : 0)
            .Where(id => id > 0).Distinct().ToArray();
        string[] teamKeys = KeysFor("teams_master").Select(k => k.ToUpperInvariant()).ToArray();
        string[] skillKeys = KeysFor("lookup_schedule_skills").Select(k => k.ToUpperInvariant()).ToArray();
        string[] reqKeys = KeysFor("project_skill_requirements").ToArray();
        var boardDates = KeysFor("daily_project_schedules")
            .Select(k => DateOnly.TryParseExact(k, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out var d) ? (DateOnly?)d : null)
            .Where(d => d.HasValue).Select(d => d!.Value).Distinct().ToList();
        int refreshed = 0;

        try
        {
            var db = new PostgreSqlService();

            if (prefIds.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT * FROM worker_preferences WHERE staff_id = ANY(@Ids);", new { Ids = prefIds }))
                {
                    WorkerPreference p = MapPreference(r);
                    PrefBaseline.UpsertEntry(p.StaffID, PrefFingerprint(p), p.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (teamKeys.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT * FROM teams_master WHERE UPPER(TRIM(team_name)) = ANY(@Keys);", new { Keys = teamKeys }))
                {
                    TeamRecord t = MapTeam(r);
                    TeamBaseline.UpsertEntry(TeamKey(t), TeamFingerprint(t), t.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (skillKeys.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT skill_name, display_icon, updated_at FROM lookup_schedule_skills WHERE UPPER(TRIM(skill_name)) = ANY(@Keys);", new { Keys = skillKeys }))
                {
                    ScheduleSkillRecord s = MapSkill(r);
                    SkillBaseline.UpsertEntry(SkillKey(s), SkillFingerprint(s), s.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (reqKeys.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT project_type, required_skills, required_vehicles, updated_at FROM project_skill_requirements WHERE project_type = ANY(@Keys);", new { Keys = reqKeys }))
                {
                    ProjectSkillRequirement? req = MapProjectSkillRequirement(r);
                    if (req == null) continue;
                    ProjSkillReqBaseline.UpsertEntry(req.ProjectType, ProjSkillReqFingerprint(req), req.ServerUpdatedAt);
                    refreshed++;
                }
            }

            foreach (var date in boardDates)
            {
                // Reference baselines are NOT re-captured here: the Preference sheet is not repopulated, so its
                // untouched rows must keep their original baselines.
                var fresh = BuildFromServer(db, date, captureReferenceBaselines: false);
                Hydrate(fresh);
                SheetMapper.RefreshAfterSave(fresh);
                refreshed++;
            }

            AppLogger.Info($"[{Module}] Keep Server Values: refreshed {refreshed} conflict baseline(s) " +
                $"(prefs {prefIds.Length}, teams {teamKeys.Length}, skills {skillKeys.Length}, projSkillReqs {reqKeys.Length}, board dates {boardDates.Count}).");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{Module}] AcceptServerConflictBaselines failed", ex);
        }
    }

    // =========================================================================
    // 2. FETCH ONE WORKING DATE WITH OFFLINE FALLBACK
    // =========================================================================
    /// <param name="refreshPayrollCalendar">
    /// True on startup / "Load from System": re-reads the SQL Payroll public holiday calendar (otherwise the session
    /// cache is reused). Leaves are always read for <paramref name="date"/> only.
    /// </param>
    internal static ScheduleRecord FetchSchedule(DateOnly date, bool refreshPayrollCalendar = false)
    {
        var record = FetchScheduleCore(date);
        ApplyPayrollCalendar(record, refreshPayrollCalendar);
        return record;
    }

    private static ScheduleRecord FetchScheduleCore(DateOnly date)
    {
        var record = ServerSyncHelper.FetchWithOfflineFallback<ScheduleRecord>(
            $"{Module}.Fetch schedule {date:yyyy-MM-dd}",
            db => BuildFromServer(db, date),
            debugFormatter: r =>
                $"  date={r.DailyScheduleDate:yyyy-MM-dd} slots={r.Projects.Count} attendances={r.WorkerAttendance.Count} " +
                $"entries={r.AttendanceEntries.Count} prefs={r.Preferences.Count} teams={r.Teams.Count}");

        if (record == null) return new ScheduleRecord { DailyScheduleDate = date };

        // The local snapshot holds whichever date was fetched last: keep the reference tables and all-time attendance entries, drop foreign-date rows.
        if (record.DailyScheduleDate != date)
        {
            AppLogger.Warn($"[{Module}] Offline snapshot is for {record.DailyScheduleDate:yyyy-MM-dd}; showing {date:yyyy-MM-dd} without date-scoped rows.");
            record = new ScheduleRecord { DailyScheduleDate = date, AttendanceEntries = record.AttendanceEntries, Preferences = record.Preferences, Teams = record.Teams, Skills = record.Skills, ProjectSkillRequirements = record.ProjectSkillRequirements };
        }

        // Offline Mode only: Online Mode already got LicensedSkills from staff_vehicle_licenses / staff_skill_licenses
        // in BuildFromServer. Licence validity is date-effective: re-derived for every date the board is moved to.
        if (ServerEnvironment.IsOffline)
        {
            DeriveLicensedSkills(record.Preferences, date);
            BoardContext.SetSkillLicenseTypes(LoadOfflineSkillLicenseTypes());
        }
        BoardContext.SetSkills(record.Skills);
        BoardContext.SetProjectSkillRequirements(record.ProjectSkillRequirements);
        Hydrate(record);
        return record;
    }

    // =========================================================================
    // 2b. SQL PAYROLL CALENDAR (public holidays + date-scoped leaves)
    // =========================================================================
    // Public holidays: SQL Payroll LV_CALENDAR_DAYS (DAYTYPE = 2), cached once per session (refreshed on Load from
    // System) and mirrored into PostgreSQL public_holidays. Leaves: SQL Payroll LV_APP for the loaded date only,
    // mirrored into staff_leaves. When SQL Payroll is unreachable both fall back to the PostgreSQL mirror.
    private static readonly object _calendarSync = new();
    private static Dictionary<DateOnly, string>? _holidaysByDate;
    /// <summary>Set when the last calendar refresh could not reach SQL Payroll: leaves then go straight to PostgreSQL.</summary>
    private static bool _payrollUnreachable;
    private static bool _holidayAdvisoryShown;
    private static string? _pendingHolidayAdvisory;

    /// <summary>Sets <see cref="ScheduleRecord.IsHoliday"/> / <see cref="ScheduleRecord.HolidayDescription"/> / <see cref="ScheduleRecord.Leaves"/>.</summary>
    private static void ApplyPayrollCalendar(ScheduleRecord record, bool refreshHolidays)
    {
        DateOnly date = record.DailyScheduleDate;
        try
        {
            EnsurePublicHolidays(force: refreshHolidays);

            Dictionary<DateOnly, string>? holidays;
            lock (_calendarSync) holidays = _holidaysByDate;
            if (holidays != null)
            {
                record.IsHoliday = holidays.TryGetValue(date, out string? description);
                record.HolidayDescription = record.IsHoliday ? (description ?? string.Empty) : string.Empty;

                if (ServerEnvironment.IsOnline)
                {
                    try { ServerDispatcher.UpdateScheduleHolidayFlag(date, record.IsHoliday); }
                    catch (Exception ex) { AppLogger.Warn($"[{Module}] daily_schedule_settings.is_holiday update failed: {ex.Message}"); }
                }
            }
            // holidays == null: neither SQL Payroll nor PostgreSQL answered -> keep the offline snapshot's values.

            var leaves = FetchLeavesForDate(date);
            if (leaves != null) record.Leaves = leaves;
            foreach (var l in record.Leaves) l.MappedStatus = PayrollLeaveMapper.MapLeaveType(l.LeaveType);

            AppLogger.Info($"[{Module}] Payroll calendar {date:yyyy-MM-dd}: holiday={record.IsHoliday}" +
                $"{(record.IsHoliday ? $" ({record.HolidayDescription})" : string.Empty)}, leaves={record.Leaves.Count}.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] Applying SQL Payroll calendar for {date:yyyy-MM-dd} failed: {ex.Message}");
        }
    }

    /// <summary>Loads the public holiday cache (SQL Payroll first, PostgreSQL mirror second) when empty or forced.</summary>
    private static void EnsurePublicHolidays(bool force)
    {
        lock (_calendarSync)
        {
            if (!force && _holidaysByDate != null) return;
        }

        // Online Mode only, like the Staff Manager payroll sync (respects ServerEnvironment.SimulateOffline).
        if (!ServerEnvironment.IsOnline) return;

        try
        {
            var rows = new FirebirdPayrollService().FetchPublicHolidays();
            var map = rows
                .GroupBy(h => h.HolidayDate)
                .ToDictionary(g => g.Key, g => (g.OrderByDescending(h => h.AutoKey).First().Description ?? string.Empty).Trim());
            lock (_calendarSync)
            {
                _holidaysByDate = map;
                _payrollUnreachable = false;
            }

            try { ServerDispatcher.UpsertPublicHolidays(rows); }
            catch (Exception ex) { AppLogger.Warn($"[{Module}] public_holidays mirror failed: {ex.Message}"); }

            int year = DateTime.Today.Year;
            if (!map.Keys.Any(d => d.Year == year))
            {
                string advisory =
                    $"SQL Payroll has no public holidays for {year}.\n\n" +
                    "'Is Holiday' on the Daily Schedule is evaluated from the SQL Payroll calendar, so no day this year will be " +
                    "marked as a holiday until it is maintained.\n\n" +
                    $"HR: in SQL Payroll, open Leave -> Maintain Calendar and enter the {year} public holidays, then click 'Load from System'.";
                AppLogger.Warn($"[{Module}] SQL Payroll LV_CALENDAR_DAYS has 0 public holidays for {year}.");
                lock (_calendarSync)
                {
                    if (!_holidayAdvisoryShown) _pendingHolidayAdvisory = advisory;
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] SQL Payroll public holidays unavailable, using PostgreSQL public_holidays: {ex.Message}");
            var fallback = QueryHolidaysFromDb();
            lock (_calendarSync)
            {
                _payrollUnreachable = true;
                if (fallback != null) _holidaysByDate = fallback;
            }
        }
    }

    private static Dictionary<DateOnly, string>? QueryHolidaysFromDb()
    {
        try
        {
            ServerDispatcher.EnsureScheduleSchema();
            var map = new Dictionary<DateOnly, string>();
            foreach (var r in new PostgreSqlService().Query<dynamic>("SELECT holiday_date, COALESCE(description, '') AS description FROM public_holidays;"))
            {
                DateOnly? d = Date(r.holiday_date);
                if (d.HasValue) map[d.Value] = Text(r.description).Trim();
            }
            return map;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] public_holidays query failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// Leave days on <paramref name="date"/>: SQL Payroll LV_APP (mirrored into staff_leaves), else the PostgreSQL
    /// staff_leaves cache. Null when neither answered (the caller keeps the offline snapshot's leaves).
    /// </summary>
    private static List<StaffLeaveEntry>? FetchLeavesForDate(DateOnly date)
    {
        if (!ServerEnvironment.IsOnline) return null;

        bool skipPayroll;
        lock (_calendarSync) skipPayroll = _payrollUnreachable;

        if (!skipPayroll)
        {
            try
            {
                var leaves = new List<StaffLeaveEntry>();
                foreach (var row in new FirebirdPayrollService().FetchLeavesForDate(date))
                {
                    if (!long.TryParse((row.Code ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long staffId) || staffId <= 0)
                    {
                        AppLogger.Warn($"[{Module}] SQL Payroll leave DOCKEY {row.DocKey}: employee code '{row.Code}' is not a Staff ID (skipped).");
                        continue;
                    }
                    leaves.Add(new StaffLeaveEntry
                    {
                        DocKey = row.DocKey,
                        StaffID = staffId,
                        LeaveDate = row.LeaveDate,
                        LeaveType = row.LeaveType,
                        LeaveDay = row.LeaveDay,
                        Description = row.Description,
                        MappedStatus = PayrollLeaveMapper.MapLeaveType(row.LeaveType)
                    });
                }

                try { ServerDispatcher.ReplaceLeavesForDate(date, leaves); }
                catch (Exception ex) { AppLogger.Warn($"[{Module}] staff_leaves mirror for {date:yyyy-MM-dd} failed: {ex.Message}"); }
                return leaves;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[{Module}] SQL Payroll leaves for {date:yyyy-MM-dd} unavailable, using PostgreSQL staff_leaves: {ex.Message}");
                lock (_calendarSync) _payrollUnreachable = true;
            }
        }

        try
        {
            ServerDispatcher.EnsureScheduleSchema();
            var cached = new List<StaffLeaveEntry>();
            foreach (var r in new PostgreSqlService().Query<dynamic>(
                "SELECT payroll_dockey, staff_id, leave_date, leave_type, leave_day, COALESCE(description, '') AS description FROM staff_leaves WHERE leave_date = @D ORDER BY staff_id, payroll_dockey;",
                new { D = date.ToDateTime(TimeOnly.MinValue) }))
            {
                string type = Text(r.leave_type).Trim();
                cached.Add(new StaffLeaveEntry
                {
                    DocKey = r.payroll_dockey != null ? Convert.ToInt64(r.payroll_dockey) : 0,
                    StaffID = Convert.ToInt64(r.staff_id),
                    LeaveDate = date,
                    LeaveType = type,
                    LeaveDay = ServerSyncHelper.Dec(r.leave_day),
                    Description = Text(r.description).Trim(),
                    MappedStatus = PayrollLeaveMapper.MapLeaveType(type)
                });
            }
            return cached;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] staff_leaves query for {date:yyyy-MM-dd} failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>
    /// After-load hook (macro thread, Excel state already restored): shows the "no public holidays this year"
    /// advisory once per session. Informational only; the loaded schedule is unaffected.
    /// </summary>
    internal static void ShowPendingHolidayAdvisory()
    {
        string? advisory;
        lock (_calendarSync)
        {
            advisory = _pendingHolidayAdvisory;
            _pendingHolidayAdvisory = null;
            if (advisory != null) _holidayAdvisoryShown = true;
        }
        if (advisory == null) return;

        try
        {
            // Modal dialog exception (CLAUDE.md Rule 5.2): Excel must repaint & respond while the dialog is open.
            dynamic xlApp = ExcelDnaUtil.Application;
            try { xlApp.ScreenUpdating = true; } catch { }
            try { xlApp.EnableEvents = true; } catch { }
            try { Arrow.HideAll(); } catch { }

            MessageBox.Show(advisory, "Schedule Manager - SQL Payroll Calendar", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] Holiday advisory failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Fills the [JsonIgnore] display columns (staff names, team leader, project label) so a record is complete the
    /// moment it is loaded - from the server or from an offline snapshot - without waiting for the rule engine.
    /// Names resolve for former field workers too, so past dates keep their attendance names.
    /// </summary>
    internal static void Hydrate(ScheduleRecord record)
    {
        foreach (var e in record.AttendanceEntries) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
        foreach (var a in record.WorkerAttendance) a.StaffName = Roster.NameOf(a.StaffID, a.StaffName);
        foreach (var p in record.Preferences) p.StaffName = Roster.NameOf(p.StaffID, p.StaffName);
        foreach (var t in record.Teams)
        {
            if (t.TeamLeaderStaffID > 0) t.TeamLeader = Roster.NameOf(t.TeamLeaderStaffID, t.TeamLeader);
        }
        // project_master carries no project name, so the label is the project code itself.
        foreach (var slot in record.Projects)
            slot.ProjectName = Roster.FindProject(slot.ProjectID)?.Id ?? slot.ProjectName;
    }

    /// <param name="captureReferenceBaselines">
    /// False when only the board is re-rendered (Keep Server Values): the reference-table baselines must then keep
    /// matching the Preference sheet, which is not repopulated.
    /// </param>
    private static ScheduleRecord BuildFromServer(PostgreSqlService db, DateOnly date, bool captureReferenceBaselines = true)
    {
        try { ServerDispatcher.EnsureScheduleSchema(); }
        catch (Exception ex) { AppLogger.Warn($"[{Module}] Schedule schema bootstrap skipped: {ex.Message}"); }

        DateTime d = date.ToDateTime(TimeOnly.MinValue);
        var record = new ScheduleRecord { DailyScheduleDate = date };

        // Per-date attendance parameters; no row => ScheduleRecord defaults (06:00 / 10:30 / 18:00).
        // The next date's Day Start Time is the exclusive end boundary of this working day.
        DateTime dNext = date.AddDays(1).ToDateTime(TimeOnly.MinValue);
        foreach (var r in db.Query<dynamic>(
            "SELECT schedule_date, day_start_time, late_cut_off_time, ot_start_time FROM daily_schedule_settings WHERE schedule_date IN (@D, @DNext);",
            new { D = d, DNext = dNext }))
        {
            DateOnly rowDate = DateOnly.FromDateTime(Convert.ToDateTime(r.schedule_date));
            if (rowDate == date)
            {
                record.DayStartTime = Time(r.day_start_time) ?? record.DayStartTime;
                record.LateCutOffTime = Time(r.late_cut_off_time) ?? record.LateCutOffTime;
                record.OTStartTime = Time(r.ot_start_time) ?? record.OTStartTime;
            }
            else if (rowDate == date.AddDays(1))
            {
                record.NextDayStartTime = Time(r.day_start_time) ?? record.NextDayStartTime;
            }
        }

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

        // All-time attendance entries are loaded; the Attendance Entries sheet filters its Date column to the selected date.
        const string sqlEntries = @"
        SELECT * FROM attendance_entries
        ORDER BY entry_date, entry_time, entry_id;";
        foreach (var r in db.Query<dynamic>(sqlEntries))
        {
            record.AttendanceEntries.Add(new AttendanceEntry
            {
                EntryId = Convert.ToInt64(r.entry_id),
                Date = DateOnly.FromDateTime(Convert.ToDateTime(r.entry_date)),
                Time = Time(r.entry_time) ?? TimeOnly.MinValue,
                StaffID = Convert.ToInt64(r.staff_id),
                ProjectID = AttendanceRollupService.NormalizeProjectId(Text(r.project_id)),
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
                IsAllocated = (bool)r.is_allocated
            });
        }

        // Board concurrency token + baseline for this date (fingerprint taken from the rows just loaded).
        record.BoardServerUpdatedAt = QueryBoardMaxUpdatedAt(db, date);
        _baselineDate = date;
        _baselineBoardFingerprint = BoardFingerprint(record);
        _baselineBoardMaxUpdatedAt = record.BoardServerUpdatedAt;

        foreach (var r in db.Query<dynamic>("SELECT * FROM worker_preferences ORDER BY staff_id;"))
            record.Preferences.Add(MapPreference(r));

        foreach (var r in db.Query<dynamic>("SELECT * FROM teams_master ORDER BY team_name;"))
            record.Teams.Add(MapTeam(r));

        // Preference[SkillsTable] (source 3 of the unified skill list).
        foreach (var r in db.Query<dynamic>("SELECT skill_name, display_icon, updated_at FROM lookup_schedule_skills ORDER BY skill_name;"))
            record.Skills.Add(MapSkill(r));
        BoardContext.SetSkillLicenseTypes(QuerySkillLicenseTypes(db));
        BoardContext.SetSkills(record.Skills);

        foreach (var r in db.Query<dynamic>("SELECT project_type, required_skills, required_vehicles, updated_at FROM project_skill_requirements ORDER BY project_type;"))
        {
            ProjectSkillRequirement? req = MapProjectSkillRequirement(r);
            if (req != null) record.ProjectSkillRequirements.Add(req);
        }
        BoardContext.SetProjectSkillRequirements(record.ProjectSkillRequirements);

        // Only reached when the server answered: an empty result is an authoritative (empty) baseline.
        if (captureReferenceBaselines)
        {
            PrefBaseline.Capture(record.Preferences.Where(p => p.StaffID > 0), p => p.StaffID, PrefFingerprint, p => p.ServerUpdatedAt);
            TeamBaseline.Capture(record.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)), TeamKey, TeamFingerprint, t => t.ServerUpdatedAt);
            SkillBaseline.Capture(record.Skills.Where(s => !string.IsNullOrWhiteSpace(s.SkillName)), SkillKey, SkillFingerprint, s => s.ServerUpdatedAt);
            ProjSkillReqBaseline.Capture(record.ProjectSkillRequirements, r => r.ProjectType, ProjSkillReqFingerprint, r => r.ServerUpdatedAt);
        }

        // Licensed skills are derived from StaffManager data, never stored with the preference.
        var vehicleLicenses = QueryVehicleLicenses(db);
        var skillLicenses = QuerySkillLicenses(db);
        BoardContext.SetLicenses(vehicleLicenses, skillLicenses);
        ApplyLicensedSkills(record.Preferences, vehicleLicenses, skillLicenses, date);

        return record;
    }

    private static List<VehicleLicense> QueryVehicleLicenses(PostgreSqlService db) =>
        db.Query<dynamic>("SELECT staff_id, vehicle_license_class, renew_date, expiry_date, status FROM staff_vehicle_licenses;")
            .Select(r => new VehicleLicense
            {
                StaffID = Convert.ToInt64(r.staff_id),
                VehicleLicenseClass = Parse(r.vehicle_license_class, VehicleLicenseClass.None),
                RenewDate = DateOnly.FromDateTime(Convert.ToDateTime(r.renew_date)),
                ExpiryDate = DateOnly.FromDateTime(Convert.ToDateTime(r.expiry_date)),
                VehicleLicenseStatus = Parse(r.status, VehicleLicenseStatus.Invalid)
            }).ToList();

    private static List<SkillLicense> QuerySkillLicenses(PostgreSqlService db) =>
        db.Query<dynamic>("SELECT staff_id, skill_license_type, renew_date, expiry_date, status FROM staff_skill_licenses;")
            .Select(r => new SkillLicense
            {
                StaffID = Convert.ToInt64(r.staff_id),
                SkillLicenseType = Text(r.skill_license_type),
                RenewDate = DateOnly.FromDateTime(Convert.ToDateTime(r.renew_date)),
                ExpiryDate = DateOnly.FromDateTime(Convert.ToDateTime(r.expiry_date)),
                SkillLicenseStatus = Parse(r.status, SkillLicenseStatus.Invalid)
            }).ToList();

    /// <summary>StaffManager's skill licence types (source 1), tolerating a database without the display_icon column.</summary>
    private static List<SkillLicenseType> QuerySkillLicenseTypes(PostgreSqlService db)
    {
        try
        {
            bool hasIcon = db.Query<string>(
                "SELECT column_name FROM information_schema.columns WHERE table_name = 'lookup_skill_license_types' AND column_name = 'display_icon';").Any();
            string iconExpr = hasIcon ? "COALESCE(display_icon, '')" : "''";
            return db.Query<dynamic>($"SELECT skill_license_type, {iconExpr} AS display_icon FROM lookup_skill_license_types ORDER BY skill_license_type;")
                .Select(r => new SkillLicenseType
                {
                    SkillLicenseName = Text(r.skill_license_type),
                    DisplayIcon = Text(r.display_icon)
                })
                .Where(t => t.SkillLicenseName.Length > 0)
                .ToList();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] Skill licence types unavailable: {ex.Message}");
            return new List<SkillLicenseType>();
        }
    }

    private static List<SkillLicenseType> LoadOfflineSkillLicenseTypes()
    {
        try { return OfflineSnapshot.Load<List<SkillLicenseType>>() ?? new List<SkillLicenseType>(); }
        catch { return new List<SkillLicenseType>(); }
    }

    // =========================================================================
    // 3. LICENSED SKILL DERIVATION (StaffManager licences -> WorkerPreference.LicensedSkills)
    // =========================================================================
    /// <summary>Offline path: re-derives LicensedSkills from the local StaffRecord snapshot.</summary>
    private static void DeriveLicensedSkills(List<WorkerPreference> prefs, DateOnly scheduleDate)
    {
        var (veh, skl) = LoadOfflineLicenses();
        BoardContext.SetLicenses(veh, skl);
        if (prefs.Count == 0) return;

        ApplyLicensedSkills(prefs, veh, skl, scheduleDate);
    }

    private static (List<VehicleLicense> Vehicle, List<SkillLicense> Skill) LoadOfflineLicenses()
    {
        var staff = OfflineSnapshot.Load<List<StaffRecord>>();
        if (staff == null || staff.Count == 0) return (new List<VehicleLicense>(), new List<SkillLicense>());

        return (
            staff.SelectMany(s => s.VehicleLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList(),
            staff.SelectMany(s => s.SkillLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList());
    }

    /// <summary>Makes sure the licence cache is populated (server when online, local snapshot otherwise).</summary>
    internal static void EnsureLicensesLoaded()
    {
        if (BoardContext.HasLicensesLoaded) return;
        if (ServerEnvironment.IsOnline)
        {
            try
            {
                var db = new PostgreSqlService();
                BoardContext.SetLicenses(QueryVehicleLicenses(db), QuerySkillLicenses(db));
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[{Module}] Licence cache load failed: {ex.Message}");
            }
        }
        else
        {
            var (veh, skl) = LoadOfflineLicenses();
            BoardContext.SetLicenses(veh, skl);
        }
    }

    /// <summary>
    /// A licence counts for <c>scheduleDate</c> when it is not Suspended and RenewDate &lt;= scheduleDate &lt;= ExpiryDate
    /// (so a licence that has since expired still counts for past dates, and a future renewal does not apply early).
    /// Any qualifying vehicle licence class grants its role skill (<see cref="SkillCatalog.RoleSkillFor"/>, falling back
    /// to "Driver"); a valid skill licence grants its type name.
    /// </summary>
    internal static void ApplyLicensedSkills(
        IEnumerable<WorkerPreference> prefs,
        IEnumerable<VehicleLicense> vehicleLicenses,
        IEnumerable<SkillLicense> skillLicenses,
        DateOnly scheduleDate)
    {
        var skillsByStaff = BuildLicensedSkillsMap(vehicleLicenses, skillLicenses, scheduleDate);
        foreach (var p in prefs)
            p.LicensedSkills = skillsByStaff.TryGetValue(p.StaffID, out var s) ? s : new List<string>();
    }

    /// <summary>Single source of truth: staffId -> licensed skills valid on <paramref name="scheduleDate"/>.</summary>
    internal static Dictionary<long, List<string>> BuildLicensedSkillsMap(
        IEnumerable<VehicleLicense> vehicleLicenses,
        IEnumerable<SkillLicense> skillLicenses,
        DateOnly scheduleDate)
    {
        var skillsByStaff = new Dictionary<long, List<string>>();

        void Add(long staffId, string skill)
        {
            if (!skillsByStaff.TryGetValue(staffId, out var list)) skillsByStaff[staffId] = list = new List<string>();
            if (!list.Contains(skill, StringComparer.OrdinalIgnoreCase)) list.Add(skill);
        }

        foreach (var l in vehicleLicenses.Where(l => l.VehicleLicenseStatus != VehicleLicenseStatus.Suspended && l.RenewDate <= scheduleDate && scheduleDate <= l.ExpiryDate && l.VehicleLicenseClass != VehicleLicenseClass.None))
            Add(l.StaffID, SkillCatalog.RoleSkillFor(l.VehicleLicenseClass) ?? SkillCatalog.Driver);
        foreach (var l in skillLicenses.Where(l => l.SkillLicenseStatus != SkillLicenseStatus.Suspended && l.RenewDate <= scheduleDate && scheduleDate <= l.ExpiryDate && !string.IsNullOrWhiteSpace(l.SkillLicenseType)))
            Add(l.StaffID, l.SkillLicenseType.Trim());

        return skillsByStaff;
    }

    // =========================================================================
    // 4. LOOKUP LISTS FOR THE 'Data' SHEET
    // =========================================================================
    /// <summary>
    /// StaffIDList/StaffNameList offer strictly the Worker, Office and Admin staff employed on <paramref name="date"/>
    /// (resigned / not-yet-joined excluded; Daily Attendances / Attendance Entries);
    /// AllFieldStaffIDList/AllFieldStaffNameList include every field worker (Office / Admin excluded), resigned included (Preference / Teams).
    /// </summary>
    internal static void SyncLookupLists(DateOnly date)
    {
        Roster.Invalidate();
        SyncStaffLookupLists(date, synchronous: false);
        SyncOtherLookupLists();
    }

    internal static void SyncStaffLookupLists(DateOnly date, bool synchronous = false)
    {
        // StaffIDList and StaffNameList stay index-aligned (Staff ID <-> Staff Name sync).
        var staff = Roster.ActiveAttendanceOn(date)
            .Where(s => s.Id > 0 && s.ShortName.Length > 0)
            .OrderBy(s => s.Id).ToList();
        // AllFieldStaff* lists include resigned workers but never Office / Admin (used by the Preference sheet).
        var allStaff = Roster.Staff
            .Where(s => s.Id > 0 && s.ShortName.Length > 0)
            .OrderBy(s => s.Id).ToList();
        if (staff.Count > 0 || allStaff.Count > 0)
        {
            DataList.SyncAlignedLists(new Dictionary<string, IEnumerable<string>>
            {
                ["StaffIDList"] = staff.Select(s => s.Id.ToString()).ToList(),
                ["StaffNameList"] = staff.Select(s => s.ShortName).ToList(),
                ["AllFieldStaffIDList"] = allStaff.Select(s => s.Id.ToString()).ToList(),
                ["AllFieldStaffNameList"] = allStaff.Select(s => s.ShortName).ToList(),
            }, synchronous);
        }
    }

    private static void SyncOtherLookupLists()
    {
        // VehiclePlateList and VehiclePlateTypeList stay index-aligned ("PLATE — Type" dropdown labels).
        var vehicles = Roster.Vehicles.Where(v => v.Plate.Length > 0)
            .GroupBy(v => v.Plate, StringComparer.OrdinalIgnoreCase).Select(g => g.First())
            .OrderBy(v => v.Plate, StringComparer.OrdinalIgnoreCase).ToList();
        if (vehicles.Count > 0)
        {
            DataList.SyncAlignedLists(new Dictionary<string, IEnumerable<string>>
            {
                ["VehiclePlateList"] = vehicles.Select(v => v.Plate).ToList(),
                ["VehiclePlateTypeList"] = vehicles.Select(v => v.Type == VehicleType.None ? string.Empty : v.Type.ToString()).ToList(),
            });
        }

        var projects = Roster.Projects.Select(p => p.Id).Where(p => p.Length > 0).Distinct().ToList();
        if (!projects.Contains(AttendanceRollupService.OfficeProjectName, StringComparer.OrdinalIgnoreCase))
            projects.Add(AttendanceRollupService.OfficeProjectName);
        DataList.SyncValues("ProjectIDList", projects);

        // Emoji palette behind SkillsTable[Display Icon].
        DataList.SyncValues("DisplayIconList", SkillCatalog.DisplayIcons);

        // Unified skill list: SkillsTable + SkillLicenseTypesTable + VehicleLicenseClass role skills.
        if (BoardContext.SkillLicenseTypes.Count == 0)
        {
            if (ServerEnvironment.IsOnline)
            {
                try { BoardContext.SetSkillLicenseTypes(QuerySkillLicenseTypes(new PostgreSqlService())); }
                catch (Exception ex) { AppLogger.Warn($"[{Module}] Skill licence types fetch skipped: {ex.Message}"); }
            }
            else
            {
                BoardContext.SetSkillLicenseTypes(LoadOfflineSkillLicenseTypes());
            }
        }
        SyncSkillList();
    }

    /// <summary>
    /// Pushes <see cref="BoardContext.AllConfiguredSkills"/> into the 'Data' sheet's StaffSkillList and
    /// <see cref="BoardContext.UnlicensedSkills"/> (no licensed skills) into OtherSkillList.
    /// </summary>
    internal static void SyncSkillList()
    {
        var allSkills = BoardContext.AllConfiguredSkills();
        if (allSkills.Count > 0) DataList.SyncValues("StaffSkillList", allSkills);
        var unlicensedSkills = BoardContext.UnlicensedSkills();
        if (unlicensedSkills.Count > 0) DataList.SyncValues("OtherSkillList", unlicensedSkills);
    }

    /// <summary>All stored Thumbprint entries from <paramref name="minDate"/> onwards (used to dedupe terminal punches).</summary>
    internal static List<AttendanceEntry> FetchThumbprintEntriesSince(DateOnly minDate)
    {
        ServerDispatcher.EnsureScheduleSchema();
        var db = new PostgreSqlService();
        const string sql = @"
        SELECT * FROM attendance_entries
        WHERE source = 'Thumbprint' AND entry_date >= @MinDate
        ORDER BY entry_date, entry_time, entry_id;";
        var list = new List<AttendanceEntry>();
        foreach (var r in db.Query<dynamic>(sql, new { MinDate = minDate.ToDateTime(TimeOnly.MinValue) }))
        {
            list.Add(new AttendanceEntry
            {
                EntryId = Convert.ToInt64(r.entry_id),
                Date = DateOnly.FromDateTime(Convert.ToDateTime(r.entry_date)),
                Time = Time(r.entry_time) ?? TimeOnly.MinValue,
                StaffID = Convert.ToInt64(r.staff_id),
                ProjectID = AttendanceRollupService.NormalizeProjectId(Text(r.project_id)),
                EntryType = Parse(r.entry_type, AttendanceEntryType.CheckIn),
                Source = Parse(r.source, AttendanceSource.Manual),
                OverrideReason = Text(r.override_reason),
                IsApproved = (bool)r.is_approved,
                WhatsAppPicture = Text(r.whatsapp_picture)
            });
        }
        return list;
    }

    /// <summary>
    /// Day Start Time stored for <paramref name="date"/> (online only); no row, offline or error =>
    /// <see cref="BoardContext.DefaultDayStartTime"/>.
    /// </summary>
    internal static TimeOnly FetchDayStartTime(DateOnly date)
    {
        if (!ServerEnvironment.IsOnline) return BoardContext.DefaultDayStartTime;
        try
        {
            ServerDispatcher.EnsureScheduleSchema();
            var db = new PostgreSqlService();
            foreach (var r in db.Query<dynamic>(
                "SELECT day_start_time FROM daily_schedule_settings WHERE schedule_date = @D;",
                new { D = date.ToDateTime(TimeOnly.MinValue) }))
            {
                TimeOnly? t = Time(r.day_start_time);
                if (t.HasValue) return t.Value;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] Day Start Time for {date:yyyy-MM-dd} unavailable: {ex.Message}");
        }
        return BoardContext.DefaultDayStartTime;
    }

    /// <summary>Every stored Day Start Time from <paramref name="minDate"/> onwards (dates without a row are absent).</summary>
    internal static Dictionary<DateOnly, TimeOnly> FetchDayStartTimesSince(DateOnly minDate)
    {
        var result = new Dictionary<DateOnly, TimeOnly>();
        try
        {
            ServerDispatcher.EnsureScheduleSchema();
            var db = new PostgreSqlService();
            foreach (var r in db.Query<dynamic>(
                "SELECT schedule_date, day_start_time FROM daily_schedule_settings WHERE schedule_date >= @MinDate;",
                new { MinDate = minDate.ToDateTime(TimeOnly.MinValue) }))
            {
                TimeOnly? t = Time(r.day_start_time);
                DateOnly rowDate = DateOnly.FromDateTime(Convert.ToDateTime(r.schedule_date));
                if (t.HasValue) result[rowDate] = t.Value;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] Day Start Times since {minDate:yyyy-MM-dd} unavailable: {ex.Message}");
        }
        return result;
    }

    // =========================================================================
    // 5. REFERENCE ROW MAPPERS (shared by the load and the conflict re-query)
    // =========================================================================
    private static WorkerPreference MapPreference(dynamic r) => new()
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
        RestDayPreference = Parse(r.rest_day_preference, Days.None),
        ServerUpdatedAt = ToUtcDateTime((object?)r.updated_at)
    };

    private static TeamRecord MapTeam(dynamic r)
    {
        object? leader = r.team_leader_staff_id;
        return new TeamRecord
        {
            TeamName = Text(r.team_name),
            TeamLeaderStaffID = leader != null ? Convert.ToInt64(leader) : 0,
            DefaultVehicles = Arr(r.default_vehicles),
            ServerUpdatedAt = ToUtcDateTime((object?)r.updated_at)
        };
    }

    private static ScheduleSkillRecord MapSkill(dynamic r) => new()
    {
        SkillName = Text(r.skill_name),
        DisplayIcon = Text(r.display_icon),
        ServerUpdatedAt = ToUtcDateTime((object?)r.updated_at)
    };

    /// <summary>Null for an unknown / <c>None</c> project type.</summary>
    private static ProjectSkillRequirement? MapProjectSkillRequirement(dynamic r)
    {
        ProjectType pt = Parse((object?)r.project_type, ProjectType.None);
        if (pt == ProjectType.None) return null;
        return new ProjectSkillRequirement
        {
            ProjectType = pt,
            Skills = Arr((object?)r.required_skills),
            Vehicles = Arr((object?)r.required_vehicles)
                .Select(v => Parse(v, VehicleType.None))
                .Where(v => v != VehicleType.None)
                .Distinct()
                .ToList(),
            ServerUpdatedAt = ToUtcDateTime((object?)r.updated_at)
        };
    }

    // =========================================================================
    // 6. DYNAMIC ROW HELPERS
    // =========================================================================
    private static string Text(object? v) => v?.ToString() ?? string.Empty;

    private static List<string> Arr(object? v) => v switch
    {
        string[] a => a.ToList(),
        IEnumerable<string> e => e.ToList(),
        _ => new List<string>()
    };

    private static DateOnly? Date(object? v) => v switch
    {
        DateOnly d => d,
        DateTime dt => DateOnly.FromDateTime(dt),
        _ => DateOnly.TryParse(v?.ToString(), out var parsed) ? parsed : null
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
