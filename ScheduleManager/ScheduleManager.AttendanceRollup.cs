namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
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

    /// <summary>A worker who stays under this many hours earns a half shift instead of a full one.</summary>
    private const decimal HalfShiftBelowHours = 5m;

    /// <summary>Hours beyond this count as overtime.</summary>
    private const decimal RegularShiftHours = 8m;

    // =========================================================================
    // 1. PUNCH ANALYSIS (shared with the board header / worker cells)
    // =========================================================================

    /// <summary>
    /// The project an entry belongs to. Terminal punches carry no project, so they are attributed to the worker's
    /// scheduled project when that is unambiguous (otherwise the first scheduled slot, otherwise empty).
    /// </summary>
    internal static string EffectiveProject(AttendanceEntry entry, IReadOnlyList<DailyProjectSchedule> slots)
    {
        if (!string.IsNullOrWhiteSpace(entry.ProjectID)) return entry.ProjectID.Trim().ToUpperInvariant();

        string name = Roster.FindStaff(entry.StaffID)?.ShortName ?? string.Empty;
        if (name.Length == 0) return string.Empty;

        return slots
            .Where(s => !string.IsNullOrWhiteSpace(s.ProjectID) && s.Workers.Contains(name, StringComparer.OrdinalIgnoreCase))
            .Select(s => s.ProjectID.Trim().ToUpperInvariant())
            .FirstOrDefault() ?? string.Empty;
    }

    private static List<AttendanceEntry> ApprovedFor(IEnumerable<AttendanceEntry> entries, DateOnly date) =>
        entries.Where(e => e.IsApproved && e.Date == date && e.StaffID > 0).OrderBy(e => e.Time).ToList();

    /// <summary>
    /// Per (staff, project): every in -> out span. A span ends at the worker's next punch of any kind
    /// (a CheckOut, or a CheckIn to another project), and stays open when nothing follows.
    /// </summary>
    internal static Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> BuildSpans(
        IEnumerable<AttendanceEntry> entries, DateOnly date, IReadOnlyList<DailyProjectSchedule> slots)
    {
        var result = new Dictionary<(long, string), List<WorkerProjectSpan>>();
        foreach (var worker in ApprovedFor(entries, date).GroupBy(e => e.StaffID))
        {
            var timeline = worker.ToList();
            for (int i = 0; i < timeline.Count; i++)
            {
                if (timeline[i].EntryType != AttendanceEntryType.CheckIn) continue;

                string project = EffectiveProject(timeline[i], slots);
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
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans)
    {
        var mine = spans.Where(kv => kv.Key.ProjectId.Equals(projectId, StringComparison.OrdinalIgnoreCase))
                        .SelectMany(kv => kv.Value).ToList();
        if (mine.Count == 0) return new ProjectWindow(null, null);

        TimeOnly start = mine.Min(s => s.In);
        TimeOnly? end = mine.Any(s => s.Out == null) ? null : mine.Max(s => s.Out);
        return new ProjectWindow(start, end);
    }

    // =========================================================================
    // 2. ROLLUP
    // =========================================================================

    /// <summary>
    /// Builds the daily attendance rows. Unapproved entries (pending WhatsApp confirmations) are ignored.
    /// Workers appear when they have an approved punch, are scheduled on the board, or already carry a manual
    /// override in <paramref name="existing"/> (which is preserved).
    /// </summary>
    internal static List<DailyWorkerAttendance> Run(
        DateOnly date,
        IReadOnlyList<AttendanceEntry> entries,
        IReadOnlyList<DailyProjectSchedule> slots,
        IReadOnlyList<DailyWorkerAttendance> existing,
        TimeOnly lateCutOff)
    {
        var approved = ApprovedFor(entries, date);
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

        var ids = new SortedSet<long>(byStaff.Keys);
        ids.UnionWith(scheduled.Keys);
        foreach (var kv in existingById)
        {
            if (kv.Value.ManualOverride != AttendanceStatus.None || !string.IsNullOrWhiteSpace(kv.Value.OverrideReason)) ids.Add(kv.Key);
        }

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
            TimeOnly? clockOut = checkOuts.Count > 0 && (clockIn == null || checkOuts[^1].Time != clockIn) ? checkOuts[^1].Time : null;

            // Project timeline: "PROJ-A @ 08:05; PROJ-B @ 13:15" (consecutive repeats of the same project collapse).
            var timeline = new List<(string Project, TimeOnly At)>();
            foreach (var ci in checkIns)
            {
                string project = EffectiveProject(ci, slots);
                if (project.Length == 0) continue;
                if (timeline.Count > 0 && timeline[^1].Project.Equals(project, StringComparison.OrdinalIgnoreCase)) continue;
                timeline.Add((project, ci.Time));
            }
            var projectsDone = timeline.Select(t => t.Project).Distinct(StringComparer.OrdinalIgnoreCase).ToList();

            bool present = clockIn.HasValue;
            bool nightBand = mySlots.Any(s => s.Shift == Shifts.Night);
            decimal worked = WorkedHours(clockIn, clockOut);

            decimal regular = !present ? 0m : (clockOut.HasValue && worked < HalfShiftBelowHours ? 0.5m : 1m);
            decimal ot = present && clockOut.HasValue
                ? RuleFloor(Math.Max(0m, worked - RegularShiftHours), 0.5m)
                : 0m;
            decimal additionalShifts = Math.Max(0, projectsDone.Count - 2) * 0.5m;
            // +1.0 OT hour per premix hole, no thresholds - but only for a worker who actually clocked in.
            decimal additionalOt = present ? mySlots.Sum(s => Math.Max(0, s.PremixHoles)) * 1.0m : 0m;

            AttendanceStatus status =
                !present ? AttendanceStatus.Absence
                : nightBand ? AttendanceStatus.NightShift
                : clockIn!.Value > lateCutOff ? AttendanceStatus.Late
                : AttendanceStatus.Presence;

            var manual = prior?.ManualOverride ?? AttendanceStatus.None;
            if (manual != AttendanceStatus.None) status = manual;

            result.Add(new DailyWorkerAttendance
            {
                DailyScheduleDate = date,
                StaffID = id,
                StaffName = Roster.FindStaff(id)?.ShortName ?? prior?.StaffName ?? string.Empty,
                AttendanceStatus = status,
                ClockInTime = clockIn,
                ClockOutTime = clockOut,
                ProjectCheckIns = string.Join("; ", timeline.Select(t => $"{t.Project} @ {t.At:HH:mm}")),
                ProjectsDone = projectsDone,
                RegularShift = regular,
                AdditionalShifts = additionalShifts,
                OTHours = ot,
                AdditionalOTHours = additionalOt,
                TotalFlattenedHours = Flatten(regular, additionalShifts, ot, additionalOt),
                ManualOverride = manual,
                OverrideReason = prior?.OverrideReason ?? string.Empty,
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

    private static decimal RuleFloor(decimal value, decimal step) => Math.Floor(value / step) * step;
}
