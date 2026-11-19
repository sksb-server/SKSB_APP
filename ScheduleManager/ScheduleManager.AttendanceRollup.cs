namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;

/// <summary>Start / end of one project's activity on the board header; <see cref="End"/> is null while someone is still checked in.</summary>
internal readonly record struct ProjectWindow(TimeOnly? Start, TimeOnly? End);

/// <summary>One worker's span on one project, used for the "Check in and out Time" cell.</summary>
internal readonly record struct WorkerProjectSpan(TimeOnly In, TimeOnly? Out);

/// <summary>
/// Rolls approved <see cref="AttendanceEntry"/> punches and the dispatch board up into one
/// <see cref="DailyWorkerAttendance"/> row per worker. Pure logic: no COM, no database.
/// </summary>
internal static class AttendanceRollupService
{
    private const decimal MinutesPerHour = 60m;

    /// <summary>A night-shift worker who stays under this many hours earns a half shift instead of a full one.</summary>
    private const decimal HalfShiftBelowHours = 5m;

    // =========================================================================
    // 1. PUNCH ANALYSIS (shared with the board header / worker cells)
    // =========================================================================

    /// <summary>
    /// The project the Daily Schedule board assigns to a worker for a punch at <paramref name="entryTime"/>.
    /// With several bands, night is preferred once the punch is at/after the OT start, morning otherwise.
    /// </summary>
    internal static string ScheduledProjectFor(
        long staffId, TimeOnly entryTime, IReadOnlyList<DailyProjectSchedule> slots, TimeOnly dayStartTime, TimeOnly otStartTime)
    {
        string name = Roster.NameOf(staffId);
        if (name.Length == 0) return string.Empty;

        var matching = slots
            .Where(s => !string.IsNullOrWhiteSpace(s.ProjectID) && s.Workers.Contains(name, StringComparer.OrdinalIgnoreCase))
            .ToList();
        if (matching.Count == 0) return string.Empty;

        DailyProjectSchedule chosen = matching[0];
        if (matching.Count > 1)
        {
            var preferred =ShiftOffset(entryTime, dayStartTime) >= ShiftOffset(otStartTime, dayStartTime)
                ? Shifts.Night
                : Shifts.Morning;
            chosen = matching.FirstOrDefault(s => s.Shift == preferred) ?? matching[0];
        }
        return chosen.ProjectID.Trim().ToUpperInvariant();
    }

    /// <summary>The project recorded on an Attendance Entry (already auto-derived or manually overridden).</summary>
    internal static string EntryProject(AttendanceEntry entry) => (entry.ProjectID ?? string.Empty).Trim().ToUpperInvariant();

    /// <summary>
    /// Offset of <paramref name="time"/> from the working-day start (0..24 h), so times after midnight sort after the
    /// evening of the same working day.
    /// </summary>
    internal static TimeSpan ShiftOffset(TimeOnly time, TimeOnly dayStart)
    {
        TimeSpan diff = time.ToTimeSpan() - dayStart.ToTimeSpan();
        return diff < TimeSpan.Zero ? diff + TimeSpan.FromHours(24) : diff;
    }

    private static List<AttendanceEntry> ApprovedFor(IEnumerable<AttendanceEntry> entries, DateOnly date, TimeOnly dayStartTime) =>
        entries.Where(e => e.IsApproved && e.StaffID > 0 && DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date, dayStartTime))
               .OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();

    /// <summary>
    /// Per (staff, project): every in -> out span. A span ends at the worker's next punch of any kind
    /// (a CheckOut, or a CheckIn to another project), and stays open when nothing follows.
    /// </summary>
    internal static Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> BuildSpans(
        IEnumerable<AttendanceEntry> entries, DateOnly date, TimeOnly dayStartTime)
    {
        var result = new Dictionary<(long, string), List<WorkerProjectSpan>>();
        foreach (var worker in ApprovedFor(entries, date, dayStartTime).GroupBy(e => e.StaffID))
        {
            var timeline = worker.ToList();
            for (int i = 0; i < timeline.Count; i++)
            {
                if (timeline[i].EntryType != AttendanceEntryType.CheckIn) continue;

                string project = EntryProject(timeline[i]);
                if (project.Length == 0) continue;

                TimeOnly? end = i + 1 < timeline.Count ? timeline[i + 1].Time : null;
                var key = (worker.Key, project);
                if (!result.TryGetValue(key, out var spans)) result[key] = spans = new List<WorkerProjectSpan>();
                spans.Add(new WorkerProjectSpan(timeline[i].Time, end));
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
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        TimeOnly dayStartTime)
    {
        var mine = spans.Where(kv => kv.Key.ProjectId.Equals(projectId, StringComparison.OrdinalIgnoreCase))
                        .SelectMany(kv => kv.Value).ToList();
        if (mine.Count == 0) return new ProjectWindow(null, null);

        // Ordered by offset from the working-day start so a midnight-crossing window reads e.g. 21:00 - 03:00.
        TimeOnly start = mine.OrderBy(s => ShiftOffset(s.In, dayStartTime)).First().In;
        TimeOnly? end = mine.Any(s => s.Out == null)
            ? null
            : mine.OrderBy(s => ShiftOffset(s.Out!.Value, dayStartTime)).Last().Out;
        return new ProjectWindow(start, end);
    }

    // =========================================================================
    // 2. ROLLUP
    // =========================================================================

    /// <summary>
    /// Builds the daily attendance rows. Unapproved entries (pending WhatsApp confirmations) are ignored.
    /// Only field workers employed on <paramref name="date"/> (<see cref="Roster.ActiveOn(DateOnly)"/>) appear, and each
    /// always gets a row, even when unscheduled or unpunched. Punches/slots for workers resigned before
    /// <paramref name="date"/> (or not yet joined) are excluded.
    /// Day-band overtime is the time clocked after <paramref name="otStartTime"/> (floored to 0.5 h).
    /// </summary>
    internal static List<DailyWorkerAttendance> Run(
        DateOnly date,
        IReadOnlyList<AttendanceEntry> entries,
        IReadOnlyList<DailyProjectSchedule> slots,
        IReadOnlyList<DailyWorkerAttendance> existing,
        TimeOnly dayStartTime,
        TimeOnly lateCutOff,
        TimeOnly otStartTime)
    {
        var approved = ApprovedFor(entries, date, dayStartTime);
        var byStaff = approved.GroupBy(e => e.StaffID).ToDictionary(g => g.Key, g => g.ToList());
        var existingById = existing.Where(a => a.StaffID > 0).GroupBy(a => a.StaffID).ToDictionary(g => g.Key, g => g.First());

        // Board assignments: staff id -> slots they are listed in.
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

        // Only active field workers are rolled up (office / admin staff may still punch the terminal).
        // Every field worker employed on this date gets a row, even without a punch, slot, or override.
        var ids = new SortedSet<long>(Roster.ActiveOn(date).Select(s => s.Id));

        var result = new List<DailyWorkerAttendance>();
        foreach (long id in ids)
        {
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
            bool nightBand = mySlots.Any(s => s.Shift == Shifts.Night);
            bool isLate = present && !nightBand && ShiftOffset(clockIn!.Value, dayStartTime) > ShiftOffset(lateCutOff, dayStartTime);
            // Same-day early checkout before lateCutOff (clockOut > clockIn avoids penalizing midnight-crossing OT).
            bool earlyOut = present && !nightBand && clockOut.HasValue
                && clockOut.Value > clockIn!.Value && clockOut.Value < lateCutOff;

            decimal regular = !present ? 0m
                : nightBand ? (clockOut.HasValue && WorkedHours(clockIn, clockOut) < HalfShiftBelowHours ? 0.5m : 1m)
                : (isLate || earlyOut ? 0.5m : 1m);
            decimal ot = present && clockOut.HasValue && !nightBand
                ? OvertimeHours(clockIn!.Value, clockOut.Value, otStartTime, dayStartTime)
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
                IsAllocated = mySlots.Count > 0
            });
        }

        return result;
    }

    /// <summary>The [ColumnCalculate] expression on TotalFlattenedHours, in C# so persisted values never depend on Excel.</summary>
    internal static decimal Flatten(decimal regular, decimal additionalShifts, decimal ot, decimal additionalOt) =>
        Math.Round(((regular + additionalShifts) * 8m) + ((ot + additionalOt) * 1.5m), 2);

    private static decimal WorkedHours(TimeOnly? clockIn, TimeOnly? clockOut)
    {
        if (!clockIn.HasValue || !clockOut.HasValue) return 0m;
        double minutes = (clockOut.Value.ToTimeSpan() - clockIn.Value.ToTimeSpan()).TotalMinutes;
        if (minutes < 0) minutes += 24 * 60; // night shift crossing midnight
        return (decimal)minutes / MinutesPerHour;
    }

    /// <summary>Hours clocked after the OT start time (a clock-out past midnight runs on into the next day).</summary>
    private static decimal OvertimeHours(TimeOnly clockIn, TimeOnly clockOut, TimeOnly otStartTime, TimeOnly dayStartTime)
    {
        // All offsets are measured from the working-day start, so a clock-out after midnight lands after the evening.
        TimeSpan start = ShiftOffset(clockIn, dayStartTime);
        TimeSpan end = ShiftOffset(clockOut, dayStartTime);
        if (end < start) end += TimeSpan.FromHours(24);

        TimeSpan otStart = ShiftOffset(otStartTime, dayStartTime);
        TimeSpan otBegin = start > otStart ? start : otStart;
        return end > otBegin ? RuleFloor((decimal)(end - otBegin).TotalMinutes / MinutesPerHour, 0.5m) : 0m;
    }

    private static decimal RuleFloor(decimal value, decimal step) => Math.Floor(value / step) * step;
}
