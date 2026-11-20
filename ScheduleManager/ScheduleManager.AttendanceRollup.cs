namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;

/// <summary>Start / end of one project's activity on the board header; <see cref="End"/> is null while someone is still checked in.</summary>
internal readonly record struct ProjectWindow(TimeOnly? Start, TimeOnly? End);

/// <summary>
/// One worker's span on one project, used for the "Check in and out Time" cell. <see cref="InOffset"/> /
/// <see cref="OutOffset"/> are exact offsets from the working-day start (they may exceed 24 h on long working days).
/// </summary>
internal readonly record struct WorkerProjectSpan(TimeOnly In, TimeOnly? Out, TimeSpan InOffset, TimeSpan? OutOffset);

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

    /// <summary>
    /// Per (staff, project): every in -> out span. A span ends at the worker's next punch of any kind
    /// (a CheckOut, or a CheckIn to another project), and stays open when nothing follows.
    /// </summary>
    internal static Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> BuildSpans(
        IEnumerable<AttendanceEntry> entries, DateOnly date, TimeOnly dayStartTime, TimeOnly? nextDayStartTime = null)
    {
        var result = new Dictionary<(long, string), List<WorkerProjectSpan>>();
        foreach (var worker in ApprovedFor(entries, date, dayStartTime, nextDayStartTime).GroupBy(e => e.StaffID))
        {
            var timeline = worker.ToList();
            for (int i = 0; i < timeline.Count; i++)
            {
                if (timeline[i].EntryType != AttendanceEntryType.CheckIn) continue;

                string project = EntryProject(timeline[i]);
                if (project.Length == 0) continue;

                bool hasNext = i + 1 < timeline.Count;
                TimeOnly? end = hasNext ? timeline[i + 1].Time : null;
                TimeSpan inOffset = ShiftOffset(timeline[i].Date, timeline[i].Time, date, dayStartTime);
                TimeSpan? outOffset = hasNext ? ShiftOffset(timeline[i + 1].Date, timeline[i + 1].Time, date, dayStartTime) : null;
                var key = (worker.Key, project);
                if (!result.TryGetValue(key, out var spans)) result[key] = spans = new List<WorkerProjectSpan>();
                spans.Add(new WorkerProjectSpan(timeline[i].Time, end, inOffset, outOffset));
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
                        .SelectMany(kv => kv.Value).ToList();
        if (mine.Count == 0) return new ProjectWindow(null, null);

        // Ordered by exact offset from the working-day start so a midnight-crossing window reads e.g. 21:00 - 03:00.
        TimeOnly start = mine.OrderBy(s => s.InOffset).First().In;
        TimeOnly? end = mine.Any(s => s.Out == null || s.OutOffset == null)
            ? null
            : mine.OrderBy(s => s.OutOffset!.Value).Last().Out;
        return new ProjectWindow(start, end);
    }

    // =========================================================================
    // 2. ADMIN AUTO-ATTENDANCE
    // =========================================================================

    /// <summary>
    /// Approved CheckIn (day start) and CheckOut (OT start) entries for every active Admin staff member who has no
    /// punch of any kind in <paramref name="date"/>'s working day. Weekdays only, never for future dates.
    /// </summary>
    internal static List<AttendanceEntry> BuildAdminAutoEntries(
        DateOnly date, IReadOnlyList<AttendanceEntry> existingEntries, TimeOnly dayStartTime, TimeOnly otStartTime, TimeOnly? nextDayStartTime = null)
    {
        var result = new List<AttendanceEntry>();
        if (date > DateOnly.FromDateTime(DateTime.Today)) return result;
        if (date.DayOfWeek is DayOfWeek.Saturday or DayOfWeek.Sunday) return result;

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
    /// </summary>
    internal static List<DailyWorkerAttendance> Run(
        DateOnly date,
        IReadOnlyList<AttendanceEntry> entries,
        IReadOnlyList<DailyProjectSchedule> slots,
        IReadOnlyList<DailyWorkerAttendance> existing,
        TimeOnly dayStartTime,
        TimeOnly lateCutOff,
        TimeOnly otStartTime,
        TimeOnly? nextDayStartTime = null)
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

            var checkIns = punches.Where(p => p.EntryType == AttendanceEntryType.CheckIn).ToList();
            var checkOuts = punches.Where(p => p.EntryType == AttendanceEntryType.CheckOut).ToList();
            TimeOnly? clockIn = checkIns.Count > 0 ? checkIns[0].Time : null;
            TimeOnly? clockOut = checkOuts.Count > 0 && (clockIn == null || checkOuts[^1].Date != checkIns[0].Date || checkOuts[^1].Time != clockIn)
                ? checkOuts[^1].Time
                : null;
            // Exact offsets from the working-day start (never wrap at 24 h, so >24 h working days stay correct).
            TimeSpan inOffset = checkIns.Count > 0 ? ShiftOffset(checkIns[0].Date, checkIns[0].Time, date, dayStartTime) : TimeSpan.Zero;
            TimeSpan outOffset = clockOut.HasValue ? ShiftOffset(checkOuts[^1].Date, checkOuts[^1].Time, date, dayStartTime) : inOffset;

            // Distinct checked-in projects in chronological order, strictly from Attendance Entries.
            var projectsDone = new List<string>();
            var seenProjects = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var ci in checkIns)
            {
                string project = EntryProject(ci);
                if (project.Length > 0 && seenProjects.Add(project))
                    projectsDone.Add(project);
            }

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
            // +1.0 OT hour per premix hole, no thresholds - but only for a worker who actually clocked in.
            decimal additionalOt = present ? mySlots.Sum(s => Math.Max(0, s.PremixHoles)) * 1.0m : 0m;

            AttendanceStatus status =
                !present ? AttendanceStatus.Absence
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
                ProjectsDone = projectsDone,
                RegularShift = regular,
                AdditionalShifts = additionalShifts,
                OTHours = ot,
                AdditionalOTHours = additionalOt,
                TotalFlattenedHours = Flatten(regular, additionalShifts, ot, additionalOt),
                IsAllocated = isOffice ? status != AttendanceStatus.Absence : mySlots.Count > 0
            });
        }

        return result;
    }

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
