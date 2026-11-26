namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;

/// <summary>Start / end of one project's activity on the board header; <see cref="End"/> is null while someone is still checked in.</summary>
internal sealed record ProjectWindow(
    TimeOnly? Start,
    TimeOnly? End,
    IReadOnlyList<string> StartWorkers,
    IReadOnlyList<WindowContributor> EndContributors,
    IReadOnlyList<string> OpenWorkers);

/// <summary>A worker whose punch ended a span that set the window's end (check-out, or check-in to another project).</summary>
internal readonly record struct WindowContributor(string WorkerName, AttendanceEntryType? OutEntryType = null, string? OutProject = null);

/// <summary>
/// One worker's span on one project, used for the "Check in and out Time" cell. <see cref="InOffset"/> /
/// <see cref="OutOffset"/> are exact offsets from the working-day start (they may exceed 24 h on long working days).
/// </summary>
internal readonly record struct WorkerProjectSpan(TimeOnly In, TimeOnly? Out, TimeSpan InOffset, TimeSpan? OutOffset, AttendanceEntryType? OutEntryType = null, string? OutProject = null);

/// <summary>
/// Rolls approved <see cref="AttendanceEntry"/> punches and the dispatch board up into one
/// <see cref="DailyWorkerAttendance"/> row per Worker, Office and Admin staff member. Pure logic: no COM, no database.
/// </summary>
internal static class AttendanceRollupService
{
    private const decimal MinutesPerHour = 60m;

    /// <summary>Project name auto-derived for Office staff.</summary>
    internal const string OfficeProjectName = "Office";

    /// <summary>Trims a project id; real codes become upper-case, "Office" keeps its canonical casing.</summary>
    internal static string NormalizeProjectId(string? projectId)
    {
        string trimmed = (projectId ?? string.Empty).Trim();
        if (trimmed.Length == 0) return string.Empty;
        return trimmed.Equals(OfficeProjectName, StringComparison.OrdinalIgnoreCase)
            ? OfficeProjectName
            : trimmed.ToUpperInvariant();
    }

    /// <summary>A night-shift worker who stays under this many hours earns a half shift instead of a full one.</summary>
    private const decimal HalfShiftBelowHours = 5m;

    // =========================================================================
    // 1. PUNCH ANALYSIS (shared with the board header / worker cells)
    // =========================================================================

    /// <summary>
    /// The project the Daily Schedule board assigns to a worker for a punch at <paramref name="entryDate"/> @
    /// <paramref name="entryTime"/> within <paramref name="workingDate"/>'s working day.
    /// With several bands, night is preferred once the punch is at/after the OT start, morning otherwise.
    /// Office staff always resolve to <see cref="OfficeProjectName"/>; Admin staff are never on the board and resolve to "".
    /// </summary>
    internal static string ScheduledProjectFor(
        long staffId, DateOnly entryDate, TimeOnly entryTime, DateOnly workingDate,
        IReadOnlyList<DailyProjectSchedule> slots, TimeOnly dayStartTime, TimeOnly otStartTime)
    {
        if (Roster.IsOfficeStaff(staffId)) return OfficeProjectName;

        string name = Roster.NameOf(staffId);
        if (name.Length == 0) return string.Empty;

        var matching = slots
            .Where(s => !string.IsNullOrWhiteSpace(s.ProjectID) && s.Workers.Contains(name, StringComparer.OrdinalIgnoreCase))
            .ToList();
        if (matching.Count == 0) return string.Empty;

        DailyProjectSchedule chosen = matching[0];
        if (matching.Count > 1)
        {
            var preferred = ShiftOffset(entryDate, entryTime, workingDate, dayStartTime) >= ShiftOffset(otStartTime, dayStartTime)
                ? Shifts.Night
                : Shifts.Morning;
            chosen = matching.FirstOrDefault(s => s.Shift == preferred) ?? matching[0];
        }
        return NormalizeProjectId(chosen.ProjectID);
    }

    /// <summary>The project recorded on an Attendance Entry (already auto-derived or manually overridden).</summary>
    internal static string EntryProject(AttendanceEntry entry) => NormalizeProjectId(entry.ProjectID);

    /// <summary>
    /// Offset of <paramref name="time"/> from the working-day start (0..24 h), so times after midnight sort after the
    /// evening of the same working day.
    /// </summary>
    internal static TimeSpan ShiftOffset(TimeOnly time, TimeOnly dayStart)
    {
        TimeSpan diff = time.ToTimeSpan() - dayStart.ToTimeSpan();
        return diff < TimeSpan.Zero ? diff + TimeSpan.FromHours(24) : diff;
    }

    /// <summary>
    /// Exact offset of a punch (real calendar date + time) from <paramref name="workingDate"/> @ <paramref name="dayStart"/>.
    /// Unlike the <see cref="TimeOnly"/> overload it never wraps at 24 h, so long working days stay ordered.
    /// </summary>
    internal static TimeSpan ShiftOffset(DateOnly punchDate, TimeOnly punchTime, DateOnly workingDate, TimeOnly dayStart) =>
        punchDate.ToDateTime(punchTime) - workingDate.ToDateTime(dayStart);

    private static List<AttendanceEntry> ApprovedFor(IEnumerable<AttendanceEntry> entries, DateOnly date, TimeOnly dayStartTime, TimeOnly? nextDayStartTime = null) =>
        entries.Where(e => e.IsApproved && e.StaffID > 0 && DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date, dayStartTime, nextDayStartTime))
               .OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();

    /// <summary>An approved Manual entry with a non-empty Override Reason: an intentional, audited override.</summary>
    internal static bool IsManualOverride(AttendanceEntry entry) =>
        entry.Source == AttendanceSource.Manual && !string.IsNullOrWhiteSpace(entry.OverrideReason);

    /// <summary>
    /// One worker's chronologically ordered approved punches with manual-override precedence applied:
    /// an override CheckOut drops all non-override CheckOuts; an override CheckIn drops non-override CheckIns
    /// that are earlier than the earliest override CheckIn or belong to a project the overrides cover.
    /// </summary>
    internal static List<AttendanceEntry> EffectiveTimeline(IReadOnlyList<AttendanceEntry> orderedWorkerPunches)
    {
        var overrideOuts = orderedWorkerPunches.Any(p => p.EntryType == AttendanceEntryType.CheckOut && IsManualOverride(p));
        var overrideIns = orderedWorkerPunches.Where(p => p.EntryType == AttendanceEntryType.CheckIn && IsManualOverride(p)).ToList();

        AttendanceEntry? earliestOverrideIn = overrideIns.Count > 0 ? overrideIns[0] : null;
        var overrideInProjects = overrideIns.Select(EntryProject).Where(p => p.Length > 0).ToHashSet(StringComparer.OrdinalIgnoreCase);

        var result = new List<AttendanceEntry>(orderedWorkerPunches.Count);
        foreach (var p in orderedWorkerPunches)
        {
            if (IsManualOverride(p)) { result.Add(p); continue; }
            if (p.EntryType == AttendanceEntryType.CheckOut && overrideOuts) continue;
            if (p.EntryType == AttendanceEntryType.CheckIn && earliestOverrideIn != null)
            {
                bool earlier = (p.Date, p.Time).CompareTo((earliestOverrideIn.Date, earliestOverrideIn.Time)) < 0;
                if (earlier || overrideInProjects.Contains(EntryProject(p))) continue;
            }
            result.Add(p);
        }
        return result;
    }

    /// <summary>
    /// Per (staff, project): every in -> out span. A span ends at the worker's next punch of any kind
    /// (a CheckOut, or a CheckIn to another project), and stays open when nothing follows.
    /// Manual overrides (Manual source with an Override Reason) take precedence via <see cref="EffectiveTimeline"/>.
    /// With <paramref name="slots"/> and <paramref name="otStartTime"/>, a CheckIn with a blank Project ID (thumbprint-only
    /// punch) is attributed in memory to the worker's on-board planned project (<see cref="ScheduledProjectFor"/>); the
    /// <see cref="AttendanceEntry"/> itself is never modified.
    /// </summary>
    internal static Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> BuildSpans(
        IEnumerable<AttendanceEntry> entries, DateOnly date, TimeOnly dayStartTime, TimeOnly? nextDayStartTime = null,
        IReadOnlyList<DailyProjectSchedule>? slots = null, TimeOnly? otStartTime = null)
    {
        var result = new Dictionary<(long, string), List<WorkerProjectSpan>>();
        foreach (var worker in ApprovedFor(entries, date, dayStartTime, nextDayStartTime).GroupBy(e => e.StaffID))
        {
            string ProjectOf(AttendanceEntry e)
            {
                string p = EntryProject(e);
                if (p.Length == 0 && slots != null && otStartTime.HasValue)
                    p = ScheduledProjectFor(worker.Key, e.Date, e.Time, date, slots, dayStartTime, otStartTime.Value);
                return p;
            }

            var timeline = EffectiveTimeline(worker.ToList());
            for (int i = 0; i < timeline.Count; i++)
            {
                if (timeline[i].EntryType != AttendanceEntryType.CheckIn) continue;

                string project = ProjectOf(timeline[i]);
                if (project.Length == 0) continue;

                bool hasNext = i + 1 < timeline.Count;
                TimeOnly? end = hasNext ? timeline[i + 1].Time : null;
                TimeSpan inOffset = ShiftOffset(timeline[i].Date, timeline[i].Time, date, dayStartTime);
                TimeSpan? outOffset = hasNext ? ShiftOffset(timeline[i + 1].Date, timeline[i + 1].Time, date, dayStartTime) : null;
                var key = (worker.Key, project);
                if (!result.TryGetValue(key, out var spans)) result[key] = spans = new List<WorkerProjectSpan>();
                AttendanceEntryType? outType = hasNext ? timeline[i + 1].EntryType : null;
                string? outProject = hasNext && timeline[i + 1].EntryType == AttendanceEntryType.CheckIn ? ProjectOf(timeline[i + 1]) : null;
                spans.Add(new WorkerProjectSpan(timeline[i].Time, end, inOffset, outOffset, outType, outProject));
            }
        }
        return result;
    }

    /// <summary>
    /// Header window of a project: earliest approved CheckIn of any worker, and the latest moment a worker left it
    /// (CheckOut or check-in elsewhere). End is null while any worker is still checked in.
    /// </summary>
    internal static ProjectWindow WindowFor(
        string projectId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans)
    {
        var mine = spans.Where(kv => kv.Key.ProjectId.Equals(projectId, StringComparison.OrdinalIgnoreCase))
                        .SelectMany(kv => kv.Value.Select(sp => (kv.Key.StaffId, Span: sp))).ToList();
        if (mine.Count == 0)
            return new ProjectWindow(null, null, Array.Empty<string>(), Array.Empty<WindowContributor>(), Array.Empty<string>());

        string NameFor(long id) => Roster.NameOf(id, $"#{id}");

        // Ordered by exact offset from the working-day start so a midnight-crossing window reads e.g. 21:00 - 03:00.
        TimeSpan minIn = mine.Min(x => x.Span.InOffset);
        TimeOnly start = mine.First(x => x.Span.InOffset == minIn).Span.In;
        var startWorkers = mine.Where(x => x.Span.InOffset == minIn)
            .Select(x => NameFor(x.StaffId)).Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();

        if (mine.Any(x => x.Span.Out == null || x.Span.OutOffset == null))
        {
            var openWorkers = mine.Where(x => x.Span.Out == null || x.Span.OutOffset == null)
                .Select(x => NameFor(x.StaffId)).Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
            return new ProjectWindow(start, null, startWorkers, Array.Empty<WindowContributor>(), openWorkers);
        }

        TimeSpan maxOut = mine.Max(x => x.Span.OutOffset!.Value);
        TimeOnly end = mine.First(x => x.Span.OutOffset!.Value == maxOut).Span.Out!.Value;
        var endContributors = mine.Where(x => x.Span.OutOffset!.Value == maxOut)
            .Select(x => new WindowContributor(NameFor(x.StaffId), x.Span.OutEntryType, x.Span.OutProject))
            .Distinct()
            .OrderBy(c => c.WorkerName, StringComparer.OrdinalIgnoreCase).ToList();
        return new ProjectWindow(start, end, startWorkers, endContributors, Array.Empty<string>());
    }

    // =========================================================================
    // 2. ADMIN AUTO-ATTENDANCE
    // =========================================================================

    /// <summary>
    /// Approved CheckIn (day start) and CheckOut (OT start) entries for every active Admin staff member who has no
    /// punch of any kind in <paramref name="date"/>'s working day. Weekdays only, never for future dates, never on a
    /// public holiday, and never for Admin staff on SQL Payroll leave that day.
    /// </summary>
    internal static List<AttendanceEntry> BuildAdminAutoEntries(
        DateOnly date, IReadOnlyList<AttendanceEntry> existingEntries, TimeOnly dayStartTime, TimeOnly otStartTime, TimeOnly? nextDayStartTime = null,
        IReadOnlyDictionary<long, StaffLeaveEntry>? leavesByStaffId = null, bool isHoliday = false)
    {
        var result = new List<AttendanceEntry>();
        if (date > DateOnly.FromDateTime(DateTime.Today)) return result;
        if (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) return result;
        if (isHoliday) return result;

        var admins = Roster.ActiveAdminOn(date);
        if (admins.Count == 0) return result;

        var hasEntry = existingEntries
            .Where(e => e.StaffID > 0 && DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date, dayStartTime, nextDayStartTime))
            .Select(e => e.StaffID)
            .ToHashSet();

        DateOnly checkOutDate = otStartTime < dayStartTime ? date.AddDays(1) : date;
        foreach (var admin in admins.OrderBy(a => a.Id))
        {
            if (hasEntry.Contains(admin.Id)) continue;
            if (leavesByStaffId != null && leavesByStaffId.ContainsKey(admin.Id)) continue;
            result.Add(NewAdminEntry(date, dayStartTime, admin, AttendanceEntryType.CheckIn));
            result.Add(NewAdminEntry(checkOutDate, otStartTime, admin, AttendanceEntryType.CheckOut));
        }
        return result;
    }

    private static AttendanceEntry NewAdminEntry(DateOnly date, TimeOnly time, StaffInfo admin, AttendanceEntryType type) => new()
    {
        Date = date,
        Time = time,
        StaffID = admin.Id,
        StaffName = admin.ShortName,
        ProjectID = string.Empty,
        EntryType = type,
        Source = AttendanceSource.Admin,
        OverrideReason = string.Empty,
        IsApproved = true,
        WhatsAppPicture = string.Empty
    };

    // =========================================================================
    // 3. ROLLUP
    // =========================================================================

    /// <summary>
    /// Builds the daily attendance rows. Unapproved entries (pending WhatsApp confirmations) are ignored.
    /// Worker, Office and Admin staff employed on <paramref name="date"/> (<see cref="Roster.ActiveAttendanceOn(DateOnly)"/>)
    /// appear, and each always gets a row, even when unscheduled or unpunched. Office staff are auto-allocated to
    /// "Office" unless absent; Admin staff are covered by their auto check-in / check-out entries.
    /// Punches/slots for staff resigned before <paramref name="date"/> (or not yet joined) are excluded.
    /// Day-band overtime is the time clocked after <paramref name="otStartTime"/> (floored to 0.5 h).
    /// Un-punched staff with a SQL Payroll leave in <paramref name="leavesByStaffId"/> get that leave status instead of
    /// <see cref="AttendanceStatus.Absence"/>.
    /// Manual overrides (Manual source with an Override Reason) take precedence: the earliest override CheckIn and
    /// latest override CheckOut define clock-in / clock-out, and conflicting non-manual punches are ignored.
    /// </summary>
    internal static List<DailyWorkerAttendance> Run(
        DateOnly date,
        IReadOnlyList<AttendanceEntry> entries,
        IReadOnlyList<DailyProjectSchedule> slots,
        IReadOnlyList<DailyWorkerAttendance> existing,
        TimeOnly dayStartTime,
        TimeOnly lateCutOff,
        TimeOnly otStartTime,
        TimeOnly? nextDayStartTime = null,
        IReadOnlyDictionary<long, StaffLeaveEntry>? leavesByStaffId = null)
    {
        var approved = ApprovedFor(entries, date, dayStartTime, nextDayStartTime);
        TimeSpan lateCutOffOffset = ShiftOffset(lateCutOff, dayStartTime);
        var byStaff = approved.GroupBy(e => e.StaffID).ToDictionary(g => g.Key, g => g.ToList());
        var existingById = existing.Where(a => a.StaffID > 0).GroupBy(a => a.StaffID).ToDictionary(g => g.Key, g => g.First());

        // Board assignments: staff id -> slots they are listed in (field workers only).
        var scheduled = new Dictionary<long, List<DailyProjectSchedule>>();
        foreach (var slot in slots)
        {
            foreach (var workerName in slot.Workers)
            {
                var staff = Roster.FindStaff(workerName);
                if (staff == null) continue;
                if (!scheduled.TryGetValue(staff.Id, out var list)) scheduled[staff.Id] = list = new List<DailyProjectSchedule>();
                list.Add(slot);
            }
        }

        // Projects currently drawn on the board: a planned project missing from here had its slot removed.
        var boardProjects = new HashSet<string>(
            slots.Select(s => NormalizeProjectId(s.ProjectID)).Where(id => id.Length > 0),
            StringComparer.OrdinalIgnoreCase);

        // Active Worker, Office and Admin staff are rolled up.
        // Everyone employed on this date gets a row, even without a punch, slot, or override.
        var ids = new SortedSet<long>(Roster.ActiveAttendanceOn(date).Select(s => s.Id));

        var result = new List<DailyWorkerAttendance>();
        foreach (long id in ids)
        {
            bool isOffice = Roster.IsOfficeStaff(id);
            byStaff.TryGetValue(id, out var punches);
            punches ??= new List<AttendanceEntry>();
            scheduled.TryGetValue(id, out var mySlots);
            mySlots ??= new List<DailyProjectSchedule>();
            existingById.TryGetValue(id, out var prior);

            var effectivePunches = EffectiveTimeline(punches);
            var checkIns = effectivePunches.Where(p => p.EntryType == AttendanceEntryType.CheckIn).ToList();
            var checkOuts = effectivePunches.Where(p => p.EntryType == AttendanceEntryType.CheckOut).ToList();
            var manualCheckIns = checkIns.Where(IsManualOverride).ToList();
            var manualCheckOuts = checkOuts.Where(IsManualOverride).ToList();
            AttendanceEntry? clockInEntry = manualCheckIns.Count > 0 ? manualCheckIns[0] : checkIns.Count > 0 ? checkIns[0] : null;
            AttendanceEntry? clockOutEntry = manualCheckOuts.Count > 0 ? manualCheckOuts[^1] : checkOuts.Count > 0 ? checkOuts[^1] : null;
            if (clockOutEntry != null && clockInEntry != null && clockOutEntry.Date == clockInEntry.Date && clockOutEntry.Time == clockInEntry.Time)
                clockOutEntry = null;
            TimeOnly? clockIn = clockInEntry?.Time;
            TimeOnly? clockOut = clockOutEntry?.Time;
            // Exact offsets from the working-day start (never wrap at 24 h, so >24 h working days stay correct).
            TimeSpan inOffset = clockInEntry != null ? ShiftOffset(clockInEntry.Date, clockInEntry.Time, date, dayStartTime) : TimeSpan.Zero;
            TimeSpan outOffset = clockOutEntry != null ? ShiftOffset(clockOutEntry.Date, clockOutEntry.Time, date, dayStartTime) : inOffset;

            // Distinct checked-in projects in chronological order, strictly from Attendance Entries.
            // Collected from all punches so an override CheckIn with a blank Project ID never erases ProjectCheckIns.
            var projectsDone = new List<string>();
            var seenProjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ci in punches.Where(p => p.EntryType == AttendanceEntryType.CheckIn))
            {
                string project = EntryProject(ci);
                if (project.Length > 0 && seenProjects.Add(project))
                    projectsDone.Add(project);
            }

            // Planned Projects: on-board slots listing the worker, then previously planned projects whose slot was
            // removed from the board (unassigning the worker from an on-board slot drops that project instead).
            var activePlanned = mySlots
                .Select(s => NormalizeProjectId(s.ProjectID))
                .Where(p => p.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
            var preservedOffBoard = SplitProjects(prior?.PlannedProjects)
                .Where(p => !boardProjects.Contains(p) && !activePlanned.Contains(p, StringComparer.OrdinalIgnoreCase))
                .ToList();
            string plannedProjects = string.Join(", ", activePlanned.Concat(preservedOffBoard));

            bool present = clockIn.HasValue;
            if (isOffice && present && projectsDone.Count == 0) projectsDone.Add(OfficeProjectName);
            bool nightBand = mySlots.Any(s => s.Shift == Shifts.Night);
            bool isLate = present && !nightBand && inOffset > lateCutOffOffset;
            // Early checkout before lateCutOff (outOffset > inOffset avoids penalizing midnight-crossing OT).
            bool earlyOut = present && !nightBand && clockOut.HasValue
                && outOffset > inOffset && outOffset < lateCutOffOffset;

            decimal regular = !present ? 0m
                : nightBand ? (clockOut.HasValue && WorkedHours(inOffset, outOffset) < HalfShiftBelowHours ? 0.5m : 1m)
                : (isLate || earlyOut ? 0.5m : 1m);
            decimal ot = present && clockOut.HasValue && !nightBand
                ? OvertimeHours(inOffset, outOffset, otStartTime, dayStartTime)
                : 0m;
            decimal additionalShifts = Math.Max(0, projectsDone.Count - 2) * 0.5m;
            // +1.0 OT hour per premix hole on Road_Reinstatement projects, no thresholds - but only for a worker who actually clocked in.
            decimal additionalOt = present
                ? mySlots.Where(s => Roster.SupportsPremixHoles(s.ProjectID)).Sum(s => Math.Max(0, s.PremixHoles)) * 1.0m
                : 0m;

            AttendanceStatus status =
                !present ? AbsentStatusFor(id, leavesByStaffId)
                : nightBand ? AttendanceStatus.NightShift
                : isLate ? AttendanceStatus.Late
                : AttendanceStatus.Presence;

            result.Add(new DailyWorkerAttendance
            {
                DailyScheduleDate = date,
                StaffID = id,
                StaffName = Roster.NameOf(id, prior?.StaffName ?? string.Empty),
                AttendanceStatus = status,
                ClockInTime = clockIn,
                ClockOutTime = clockOut,
                ProjectCheckIns = string.Join(", ", projectsDone),
                PlannedProjects = plannedProjects,
                ProjectsDone = projectsDone,
                RegularShift = regular,
                AdditionalShifts = additionalShifts,
                OTHours = ot,
                AdditionalOTHours = additionalOt,
                TotalFlattenedHours = Flatten(regular, additionalShifts, ot, additionalOt),
                IsAllocated = isOffice
                    ? status is AttendanceStatus.Presence or AttendanceStatus.Late or AttendanceStatus.NightShift
                    : mySlots.Count > 0
            });
        }

        return result;
    }

    /// <summary>Splits a comma-separated project list into distinct normalized project IDs (blanks dropped, order kept).</summary>
    internal static List<string> SplitProjects(string? csv) =>
        (csv ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(NormalizeProjectId)
            .Where(p => p.Length > 0)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

    /// <summary>Status of an un-punched staff member: their SQL Payroll leave status, otherwise Absence.</summary>
    private static AttendanceStatus AbsentStatusFor(long staffId, IReadOnlyDictionary<long, StaffLeaveEntry>? leavesByStaffId) =>
        leavesByStaffId != null && leavesByStaffId.TryGetValue(staffId, out var leave) && PayrollLeaveMapper.IsLeave(leave.MappedStatus)
            ? leave.MappedStatus
            : AttendanceStatus.Absence;

    /// <summary>The [ColumnCalculate] expression on TotalFlattenedHours, in C# so persisted values never depend on Excel.</summary>
    internal static decimal Flatten(decimal regular, decimal additionalShifts, decimal ot, decimal additionalOt) =>
        Math.Round(((regular + additionalShifts) * 8m) + ((ot + additionalOt) * 1.5m), 2);

    /// <summary>Hours between two exact working-day offsets (floored at 0).</summary>
    private static decimal WorkedHours(TimeSpan inOffset, TimeSpan outOffset)
    {
        double minutes = (outOffset - inOffset).TotalMinutes;
        return minutes > 0 ? (decimal)minutes / MinutesPerHour : 0m;
    }

    /// <summary>Hours clocked after the OT start time; offsets are exact (date-aware) from the working-day start.</summary>
    private static decimal OvertimeHours(TimeSpan start, TimeSpan end, TimeOnly otStartTime, TimeOnly dayStartTime)
    {
        TimeSpan otStart = ShiftOffset(otStartTime, dayStartTime);
        TimeSpan otBegin = start > otStart ? start : otStart;
        return end > otBegin ? RuleFloor((decimal)(end - otBegin).TotalMinutes / MinutesPerHour, 0.5m) : 0m;
    }

    private static decimal RuleFloor(decimal value, decimal step) => Math.Floor(value / step) * step;
}
