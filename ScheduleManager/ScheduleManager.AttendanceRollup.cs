namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;

/// <summary>
/// Start / end of one project's activity on the board header. <see cref="Start"/> = manual Start override, else the
/// LATEST worker check-in; <see cref="End"/> = manual End override, else the EARLIEST worker exit among closed spans
/// (even while other workers are still checked in - they are listed in <see cref="OpenWorkers"/>).
/// </summary>
internal sealed record ProjectWindow(
    TimeOnly? Start,
    TimeOnly? End,
    IReadOnlyList<string> StartWorkers,
    IReadOnlyList<WindowContributor> EndContributors,
    IReadOnlyList<string> OpenWorkers)
{
    /// <summary>True when <see cref="Start"/> was auto-filled from the latest check-in (no manual Start).</summary>
    public bool StartIsFallback { get; init; }

    /// <summary>True when <see cref="End"/> was auto-filled from the earliest exit (no manual End).</summary>
    public bool EndIsFallback { get; init; }

    /// <summary>True when the effective End is before the effective Start.</summary>
    public bool IsInverted { get; init; }

    /// <summary>Auto fallback Start (latest check-in), whether or not a manual Start overrides it.</summary>
    public TimeOnly? FallbackStart { get; init; }

    /// <summary>Auto fallback End (earliest exit among closed spans), whether or not a manual End overrides it.</summary>
    public TimeOnly? FallbackEnd { get; init; }

    /// <summary>Exact offset of <see cref="Start"/> from the working-day start.</summary>
    public TimeSpan? StartOffset { get; init; }

    /// <summary>Exact offset of <see cref="End"/> from the working-day start.</summary>
    public TimeSpan? EndOffset { get; init; }

    /// <summary>Provenance note for the Start time (manual vs auto fallback, contributing check-ins).</summary>
    public string StartNote { get; init; } = string.Empty;

    /// <summary>Provenance note for the End time (manual vs auto fallback, contributing exits, open workers).</summary>
    public string EndNote { get; init; } = string.Empty;
}

/// <summary>A worker whose punch ended a span that set the window's end (check-out, or check-in to another project).</summary>
internal readonly record struct WindowContributor(string WorkerName, AttendanceEntryType? OutEntryType = null, string? OutProject = null);

/// <summary>
/// One worker's span on one project, used for the "Check in and out Time" cell. <see cref="InOffset"/> /
/// <see cref="OutOffset"/> are exact offsets from the working-day start (they may exceed 24 h on long working days).
/// </summary>
internal readonly record struct WorkerProjectSpan(
    TimeOnly In, TimeOnly? Out, TimeSpan InOffset, TimeSpan? OutOffset,
    AttendanceEntryType? OutEntryType = null, string? OutProject = null,
    AttendanceSource InSource = AttendanceSource.Thumbprint, AttendanceSource? OutSource = null,
    bool InIsManualOverride = false, bool OutIsManualOverride = false);

/// <summary>
/// A worker's span on a project after clamping to the project window: effective in = max(punch, project Start),
/// effective out = min(punch, project End) - or the project End when the worker never checked out.
/// </summary>
internal sealed record ResolvedWorkerSpan(
    string ProjectId,
    WorkerProjectSpan Raw,
    TimeOnly EffectiveIn,
    TimeSpan InOffset,
    bool InIsFallback,
    string InFormat,
    string InNote,
    TimeOnly? EffectiveOut,
    TimeSpan? OutOffset,
    bool OutIsFallback,
    string OutFormat,
    string OutNote);

/// <summary>
/// Project windows of one working day, keyed by board slot (Shift, Project) with a per-project fallback for projects
/// checked into without a board slot. Manual Start / End overrides come from <see cref="DailyProjectSchedule.ActualStartTime"/> /
/// <see cref="DailyProjectSchedule.ActualEndTime"/>.
/// </summary>
internal sealed class ProjectWindowIndex
{
    private readonly Dictionary<(Shifts Shift, string ProjectId), ProjectWindow> _bySlot = new();
    private readonly Dictionary<string, ProjectWindow> _byProject = new(StringComparer.OrdinalIgnoreCase);
    private readonly IReadOnlyList<DailyProjectSchedule> _slots;
    private readonly Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> _spans;
    private readonly TimeOnly _dayStartTime;

    internal ProjectWindowIndex(
        IReadOnlyList<DailyProjectSchedule> slots,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        TimeOnly dayStartTime)
    {
        _slots = slots;
        _spans = spans;
        _dayStartTime = dayStartTime;

        var groups = slots
            .Select(s => (Slot: s, ProjectId: AttendanceRollupService.NormalizeProjectId(s.ProjectID)))
            .Where(x => x.ProjectId.Length > 0)
            .GroupBy(x => (x.Slot.Shift, x.ProjectId))
            .OrderBy(g => g.Key.Shift == Shifts.Morning ? 0 : 1);

        foreach (var g in groups)
        {
            TimeOnly? manualStart = g.Select(x => x.Slot.ActualStartTime).FirstOrDefault(t => t.HasValue);
            TimeOnly? manualEnd = g.Select(x => x.Slot.ActualEndTime).FirstOrDefault(t => t.HasValue);
            var window = AttendanceRollupService.WindowFor(g.Key.ProjectId, spans, manualStart, manualEnd, dayStartTime);
            _bySlot[g.Key] = window;
            _byProject.TryAdd(g.Key.ProjectId, window);
        }
    }

    /// <summary>On-board slot windows, keyed by (Shift, Project).</summary>
    internal IReadOnlyDictionary<(Shifts Shift, string ProjectId), ProjectWindow> BySlot => _bySlot;

    /// <summary>
    /// Window that applies to <paramref name="staffId"/>'s span on <paramref name="projectId"/>: the slot listing the
    /// worker, else the project's first on-board slot, else an auto-only window from the punches.
    /// </summary>
    internal ProjectWindow For(long staffId, string projectId)
    {
        projectId = AttendanceRollupService.NormalizeProjectId(projectId);
        string name = Roster.NameOf(staffId);
        if (name.Length > 0)
        {
            var slot = _slots.FirstOrDefault(s =>
                AttendanceRollupService.NormalizeProjectId(s.ProjectID).Equals(projectId, StringComparison.OrdinalIgnoreCase)
                && s.Workers.Contains(name, StringComparer.OrdinalIgnoreCase));
            if (slot != null && _bySlot.TryGetValue((slot.Shift, projectId), out var slotWindow)) return slotWindow;
        }

        if (_byProject.TryGetValue(projectId, out var projectWindow)) return projectWindow;

        var auto = AttendanceRollupService.WindowFor(projectId, _spans, null, null, _dayStartTime);
        _byProject[projectId] = auto;
        return auto;
    }
}

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
    // 0. SOURCE ICONS & DISPLAY FORMATS
    // =========================================================================

    /// <summary>Icon for a time taken from a manual project Start / End.</summary>
    internal const string ProjectIcon = "📋";

    /// <summary>Icon for a time auto-filled from a fallback (latest check-in / earliest exit / missing punch).</summary>
    internal const string FallbackIcon = "⚠";

    /// <summary>Icon describing where a punch came from.</summary>
    internal static string SourceIcon(AttendanceSource s) => s switch
    {
        AttendanceSource.Thumbprint => "👆",
        AttendanceSource.Manual => "📝",
        AttendanceSource.SKSB_App => "📱",
        AttendanceSource.WhatsApp => "💬",
        AttendanceSource.Admin => "⚙",
        _ => "👆"
    };

    /// <summary>Excel custom NumberFormat showing <paramref name="icon"/> before the time and an optional "(tag)" after it.</summary>
    internal static string BuildTimeFormat(string icon, string? tag = null) =>
        string.IsNullOrWhiteSpace(tag)
            ? $"\"{icon} \"hh:mm"
            : $"\"{icon} \"hh:mm\" ({tag})\"";

    private static string Hm(TimeOnly t) => t.ToString("HH:mm", CultureInfo.InvariantCulture);

    private static string Hm(TimeOnly? t) => t.HasValue ? Hm(t.Value) : "--:--";

    /// <summary>Clock time of an exact working-day offset (wraps at 24 h; display only).</summary>
    private static string HmOffset(TimeSpan offset, TimeOnly dayStart) => Hm(dayStart.Add(offset));

    private static string Hrs(decimal h) => h.ToString("0.0", CultureInfo.InvariantCulture);

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
                spans.Add(new WorkerProjectSpan(
                    timeline[i].Time, end, inOffset, outOffset, outType, outProject,
                    InSource: timeline[i].Source,
                    OutSource: hasNext ? timeline[i + 1].Source : null,
                    InIsManualOverride: IsManualOverride(timeline[i]),
                    OutIsManualOverride: hasNext && IsManualOverride(timeline[i + 1])));
            }
        }
        return result;
    }

    /// <summary>Auto-only header window (no manual Start / End; working day assumed to start at 06:00).</summary>
    internal static ProjectWindow WindowFor(
        string projectId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans) =>
        WindowFor(projectId, spans, manualStart: null, manualEnd: null, dayStartTime: new TimeOnly(6, 0));

    /// <summary>
    /// Header window of a project. Start = <paramref name="manualStart"/>, else the LATEST approved CheckIn of any worker.
    /// End = <paramref name="manualEnd"/>, else the EARLIEST moment a worker left it (CheckOut or check-in elsewhere)
    /// among closed spans - even while other workers are still checked in (listed in <see cref="ProjectWindow.OpenWorkers"/>).
    /// </summary>
    internal static ProjectWindow WindowFor(
        string projectId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        TimeOnly? manualStart,
        TimeOnly? manualEnd,
        TimeOnly dayStartTime)
    {
        projectId = NormalizeProjectId(projectId);
        var mine = spans.Where(kv => kv.Key.ProjectId.Equals(projectId, StringComparison.OrdinalIgnoreCase))
                        .SelectMany(kv => kv.Value.Select(sp => (kv.Key.StaffId, Span: sp))).ToList();

        string NameFor(long id) => Roster.NameOf(id, $"#{id}");

        // ---- auto fallback Start: LATEST check-in (exact offsets, so midnight-crossing windows stay ordered) ----
        TimeOnly? fallbackStart = null;
        TimeSpan? fallbackStartOffset = null;
        List<string> startWorkers = new();
        if (mine.Count > 0)
        {
            TimeSpan maxIn = mine.Max(x => x.Span.InOffset);
            fallbackStart = mine.First(x => x.Span.InOffset == maxIn).Span.In;
            fallbackStartOffset = maxIn;
            startWorkers = mine.Where(x => x.Span.InOffset == maxIn)
                .Select(x => NameFor(x.StaffId)).Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();
        }

        // ---- auto fallback End: EARLIEST exit among closed spans ----
        var closed = mine.Where(x => x.Span.Out.HasValue && x.Span.OutOffset.HasValue).ToList();
        TimeOnly? fallbackEnd = null;
        TimeSpan? fallbackEndOffset = null;
        List<WindowContributor> endContributors = new();
        if (closed.Count > 0)
        {
            TimeSpan minOut = closed.Min(x => x.Span.OutOffset!.Value);
            fallbackEnd = closed.First(x => x.Span.OutOffset!.Value == minOut).Span.Out!.Value;
            fallbackEndOffset = minOut;
            endContributors = closed.Where(x => x.Span.OutOffset!.Value == minOut)
                .Select(x => new WindowContributor(NameFor(x.StaffId), x.Span.OutEntryType, x.Span.OutProject))
                .Distinct()
                .OrderBy(c => c.WorkerName, StringComparer.OrdinalIgnoreCase).ToList();
        }

        var openWorkers = mine.Where(x => x.Span.Out == null || x.Span.OutOffset == null)
            .Select(x => NameFor(x.StaffId)).Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList();

        // ---- effective window ----
        TimeOnly? start = manualStart ?? fallbackStart;
        TimeOnly? end = manualEnd ?? fallbackEnd;
        bool startIsFallback = !manualStart.HasValue && fallbackStart.HasValue;
        bool endIsFallback = !manualEnd.HasValue && fallbackEnd.HasValue;
        TimeSpan? startOffset = manualStart.HasValue ? ShiftOffset(manualStart.Value, dayStartTime) : fallbackStartOffset;
        TimeSpan? endOffset = manualEnd.HasValue ? ShiftOffset(manualEnd.Value, dayStartTime) : fallbackEndOffset;
        bool inverted = startOffset.HasValue && endOffset.HasValue && endOffset.Value < startOffset.Value;

        // ---- Start note ----
        var sn = new StringBuilder();
        if (manualStart.HasValue)
        {
            sn.Append(projectId).Append(" Start: ").Append(Hm(manualStart.Value)).AppendLine(" [Manual]");
            if (fallbackStart.HasValue)
                sn.Append("Auto fallback (latest check-in) would be ").AppendLine(Hm(fallbackStart.Value));
        }
        else if (fallbackStart.HasValue)
            sn.Append(projectId).Append(" Start: ").Append(Hm(fallbackStart.Value)).AppendLine(" [Auto Fallback = latest check-in]");
        else
            sn.Append(projectId).AppendLine(" Start: -- (no manual Start and no check-in yet)");

        if (mine.Count > 0)
        {
            sn.AppendLine("Check-ins:");
            foreach (var x in mine.OrderBy(x => x.Span.InOffset).ThenBy(x => NameFor(x.StaffId), StringComparer.OrdinalIgnoreCase))
            {
                sn.Append("  ").Append(SourceIcon(x.Span.InSource)).Append(' ').Append(NameFor(x.StaffId))
                  .Append(' ').Append(Hm(x.Span.In))
                  .Append(x.Span.InIsManualOverride ? " (manual override)" : string.Empty)
                  .Append(startIsFallback && x.Span.InOffset == fallbackStartOffset ? " ← latest" : string.Empty)
                  .AppendLine();
            }
        }

        // ---- End note ----
        var en = new StringBuilder();
        if (manualEnd.HasValue)
        {
            en.Append(projectId).Append(" End: ").Append(Hm(manualEnd.Value)).AppendLine(" [Manual]");
            if (fallbackEnd.HasValue)
                en.Append("Auto fallback (earliest exit) would be ").AppendLine(Hm(fallbackEnd.Value));
        }
        else if (fallbackEnd.HasValue)
            en.Append(projectId).Append(" End: ").Append(Hm(fallbackEnd.Value)).AppendLine(" [Auto Fallback = earliest exit]");
        else
            en.Append(projectId).AppendLine(" End: -- (no manual End and no worker has left yet)");

        if (closed.Count > 0)
        {
            en.AppendLine("Exits:");
            foreach (var x in closed.OrderBy(x => x.Span.OutOffset!.Value).ThenBy(x => NameFor(x.StaffId), StringComparer.OrdinalIgnoreCase))
            {
                string how = x.Span.OutEntryType == AttendanceEntryType.CheckIn && !string.IsNullOrEmpty(x.Span.OutProject)
                    ? $"→ {x.Span.OutProject}"
                    : "CheckOut";
                en.Append("  ").Append(SourceIcon(x.Span.OutSource ?? AttendanceSource.Thumbprint)).Append(' ')
                  .Append(NameFor(x.StaffId)).Append(' ').Append(Hm(x.Span.Out!.Value)).Append(" (").Append(how).Append(')')
                  .Append(x.Span.OutIsManualOverride ? " (manual override)" : string.Empty)
                  .Append(endIsFallback && x.Span.OutOffset == fallbackEndOffset ? " ← earliest" : string.Empty)
                  .AppendLine();
            }
        }
        if (openWorkers.Count > 0)
            en.Append("Still checked in (no CheckOut punch): ").AppendLine(string.Join(", ", openWorkers));
        if (inverted)
            en.Append(FallbackIcon).Append(" Inverted window: End ").Append(Hm(end)).Append(" is before Start ").Append(Hm(start)).AppendLine("!");

        return new ProjectWindow(start, end, startWorkers, endContributors, openWorkers)
        {
            StartIsFallback = startIsFallback,
            EndIsFallback = endIsFallback,
            IsInverted = inverted,
            FallbackStart = fallbackStart,
            FallbackEnd = fallbackEnd,
            StartOffset = startOffset,
            EndOffset = endOffset,
            StartNote = sn.ToString().TrimEnd(),
            EndNote = en.ToString().TrimEnd()
        };
    }

    /// <summary>
    /// Clamps one worker span to its project window: a late worker keeps their own check-in, otherwise the project
    /// Start wins; a worker leaving early keeps their own exit, otherwise the project End caps it; a worker with no
    /// CheckOut punch falls back to the project End.
    /// </summary>
    internal static ResolvedWorkerSpan ResolveWorkerSpan(WorkerProjectSpan span, ProjectWindow window, string projectId)
    {
        projectId = NormalizeProjectId(projectId);
        string inIcon = SourceIcon(span.InSource);
        string inSrc = $"{inIcon} {span.InSource} {Hm(span.In)}";

        // ---- Effective In ----
        TimeOnly effIn;
        TimeSpan inOffset;
        bool inFallback;
        string inFormat, inNote;
        if (window.Start.HasValue && window.StartOffset.HasValue)
        {
            if (span.InOffset > window.StartOffset.Value)
            {
                effIn = span.In;
                inOffset = span.InOffset;
                inFallback = false;
                inFormat = BuildTimeFormat(inIcon, "late");
                inNote = $"Effective In: {Hm(span.In)} = max({inSrc}, {projectId} Start {Hm(window.Start.Value)}) [Late]";
            }
            else
            {
                effIn = window.Start.Value;
                inOffset = window.StartOffset.Value;
                inFallback = window.StartIsFallback;
                inFormat = BuildTimeFormat(window.StartIsFallback ? FallbackIcon : ProjectIcon);
                inNote = $"Effective In: {Hm(window.Start.Value)} = max({inSrc}, {projectId} Start {Hm(window.Start.Value)}{(window.StartIsFallback ? " [Auto Fallback]" : string.Empty)})";
            }
        }
        else
        {
            effIn = span.In;
            inOffset = span.InOffset;
            inFallback = false;
            inFormat = BuildTimeFormat(inIcon);
            inNote = $"Effective In: {Hm(span.In)} = {inSrc} (no {projectId} Start time)";
        }

        // ---- Effective Out ----
        TimeOnly? effOut;
        TimeSpan? outOffset;
        bool outFallback;
        string outFormat, outNote;
        if (span.Out.HasValue && span.OutOffset.HasValue)
        {
            AttendanceSource outSource = span.OutSource ?? AttendanceSource.Thumbprint;
            string outIcon = SourceIcon(outSource);
            string outSrc = $"{outIcon} {outSource} {Hm(span.Out.Value)}";
            if (window.End.HasValue && window.EndOffset.HasValue)
            {
                if (span.OutOffset.Value < window.EndOffset.Value)
                {
                    effOut = span.Out.Value;
                    outOffset = span.OutOffset.Value;
                    outFallback = false;
                    outFormat = BuildTimeFormat(outIcon, "leave early");
                    outNote = $"Effective Out: {Hm(span.Out.Value)} = min({outSrc}, {projectId} End {Hm(window.End.Value)}) [Leave Early]";
                }
                else
                {
                    effOut = window.End.Value;
                    outOffset = window.EndOffset.Value;
                    outFallback = window.EndIsFallback;
                    outFormat = BuildTimeFormat(window.EndIsFallback ? FallbackIcon : ProjectIcon);
                    outNote = $"Effective Out: {Hm(window.End.Value)} = min({outSrc}, {projectId} End {Hm(window.End.Value)}{(window.EndIsFallback ? " [Auto Fallback]" : string.Empty)})";
                }
            }
            else
            {
                effOut = span.Out.Value;
                outOffset = span.OutOffset.Value;
                outFallback = false;
                outFormat = BuildTimeFormat(outIcon);
                outNote = $"Effective Out: {Hm(span.Out.Value)} = {outSrc} (no {projectId} End time)";
            }
        }
        else if (window.End.HasValue && window.EndOffset.HasValue)
        {
            effOut = window.End.Value;
            outOffset = window.EndOffset.Value;
            outFallback = true;
            outFormat = BuildTimeFormat(FallbackIcon, "no punch");
            outNote = $"Effective Out: {Hm(window.End.Value)} = Fallback to {projectId} End (no CheckOut punch!)";
        }
        else
        {
            effOut = null;
            outOffset = null;
            outFallback = false;
            outFormat = string.Empty;
            outNote = "No CheckOut punch and no Project End time yet.";
        }

        return new ResolvedWorkerSpan(projectId, span, effIn, inOffset, inFallback, inFormat, inNote,
                                      effOut, outOffset, outFallback, outFormat, outNote);
    }

    /// <summary>Every span of <paramref name="staffId"/> across all projects, clamped to its window, in chronological order.</summary>
    internal static List<ResolvedWorkerSpan> ResolveWorkerSpans(
        long staffId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        ProjectWindowIndex windows) =>
        spans.Where(kv => kv.Key.StaffId == staffId)
             .SelectMany(kv => kv.Value.Select(sp => ResolveWorkerSpan(sp, windows.For(staffId, kv.Key.ProjectId), kv.Key.ProjectId)))
             .OrderBy(r => r.Raw.InOffset)
             .ToList();

    /// <summary>
    /// The worker's overall span on one project (first effective in -> last effective out across multiple spans),
    /// e.g. for the board's "Check in and out Time" cell. Null when the worker never checked into the project.
    /// </summary>
    internal static ResolvedWorkerSpan? ResolveWorkerProjectSpan(
        long staffId, string projectId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        ProjectWindowIndex windows) =>
        ResolveWorkerProjectSpan(staffId, projectId, spans, windows.For(staffId, NormalizeProjectId(projectId)));

    /// <summary>Same as the <see cref="ProjectWindowIndex"/> overload, against an already-computed <paramref name="window"/> (board slot).</summary>
    internal static ResolvedWorkerSpan? ResolveWorkerProjectSpan(
        long staffId, string projectId,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        ProjectWindow window)
    {
        projectId = NormalizeProjectId(projectId);
        if (!spans.TryGetValue((staffId, projectId), out var mine) || mine.Count == 0) return null;

        var resolved = mine.Select(sp => ResolveWorkerSpan(sp, window, projectId)).OrderBy(r => r.Raw.InOffset).ToList();
        if (resolved.Count == 1) return resolved[0];

        var first = resolved[0];
        var last = resolved[^1];
        return first with
        {
            EffectiveOut = last.EffectiveOut,
            OutOffset = last.OutOffset,
            OutIsFallback = last.OutIsFallback,
            OutFormat = last.OutFormat,
            OutNote = last.OutNote
        };
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

        // Per-project spans and windows (manual Start / End overrides from the board, else latest-in / earliest-out fallback).
        var spans = BuildSpans(approved, date, dayStartTime, nextDayStartTime, slots, otStartTime);
        var windows = new ProjectWindowIndex(slots, spans, dayStartTime);

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
            decimal additionalShifts = Math.Max(0, projectsDone.Count - 2) * 0.5m;
            // +1.0 OT hour per premix hole on Road_Reinstatement projects, no thresholds - but only for a worker who actually clocked in.
            decimal additionalOt = present
                ? mySlots.Where(s => Roster.SupportsPremixHoles(s.ProjectID)).Sum(s => Math.Max(0, s.PremixHoles)) * 1.0m
                : 0m;

            // ---- displayed clock-in / clock-out + OT (raw punches by default) ----
            TimeOnly? shownIn = clockIn;
            TimeOnly? shownOut = clockOut;
            string inFormat = clockInEntry != null ? BuildTimeFormat(SourceIcon(clockInEntry.Source)) : string.Empty;
            string outFormat = clockOutEntry != null ? BuildTimeFormat(SourceIcon(clockOutEntry.Source)) : string.Empty;
            bool inFallback = false, outFallback = false;
            string inNote = clockInEntry != null
                ? $"Clock-In: {Hm(clockInEntry.Time)} = {SourceIcon(clockInEntry.Source)} {clockInEntry.Source} punch{(IsManualOverride(clockInEntry) ? " (manual override)" : string.Empty)}"
                : "No CheckIn punch.";
            string outNote = clockOutEntry != null
                ? $"Clock-Out: {Hm(clockOutEntry.Time)} = {SourceIcon(clockOutEntry.Source)} {clockOutEntry.Source} punch{(IsManualOverride(clockOutEntry) ? " (manual override)" : string.Empty)}"
                : present ? "No CheckOut punch." : string.Empty;
            decimal ot;
            string otNote;

            if (isOffice || Roster.IsAdminStaff(id))
            {
                ot = 0m;
                additionalOt = 0m;
                otNote = "OT: 0.0 hrs (Office / Admin staff do not earn OT)";
            }
            else
            {
                var resolved = ResolveWorkerSpans(id, spans, windows);
                if (resolved.Count > 0)
                {
                    var first = resolved[0];
                    var last = resolved[^1];
                    shownIn = first.EffectiveIn;
                    inFormat = first.InFormat;
                    inFallback = first.InIsFallback;
                    inNote = first.InNote;
                    shownOut = last.EffectiveOut;
                    outFormat = last.OutFormat;
                    outFallback = last.OutIsFallback;
                    outNote = last.OutNote;

                    if (nightBand)
                    {
                        ot = 0m;
                        otNote = "OT: 0.0 hrs (Night shift — no regular OT)";
                    }
                    else
                    {
                        var lines = new List<string>();
                        ot = 0m;
                        foreach (var rs in resolved)
                        {
                            if (!rs.OutOffset.HasValue || rs.OutOffset.Value <= rs.InOffset) continue;
                            decimal spanOt = OvertimeHours(rs.InOffset, rs.OutOffset.Value, otStartTime, dayStartTime);
                            ot += spanOt;
                            lines.Add(OvertimeEquation(rs.InOffset, rs.OutOffset.Value, otStartTime, dayStartTime, spanOt, rs.ProjectId));
                        }
                        otNote = lines.Count == 0
                            ? "OT: 0.0 hrs (no project span with an effective end yet)"
                            : lines.Count == 1
                                ? lines[0]
                                : $"OT: {Hrs(ot)} hrs total\n" + string.Join("\n", lines);
                    }
                }
                else if (!present)
                {
                    ot = 0m;
                    otNote = "OT: 0.0 hrs (no CheckIn punch)";
                }
                else if (nightBand)
                {
                    ot = 0m;
                    otNote = "OT: 0.0 hrs (Night shift — no regular OT)";
                }
                else if (!clockOut.HasValue)
                {
                    ot = 0m;
                    otNote = "OT: 0.0 hrs (no CheckOut punch and no project to fall back to)";
                }
                else
                {
                    ot = OvertimeHours(inOffset, outOffset, otStartTime, dayStartTime);
                    otNote = OvertimeEquation(inOffset, outOffset, otStartTime, dayStartTime, ot, null);
                }
            }

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
                ClockInTime = shownIn,
                ClockOutTime = shownOut,
                ClockInDisplayFormat = inFormat,
                ClockOutDisplayFormat = outFormat,
                ClockInIsFallback = inFallback,
                ClockOutIsFallback = outFallback,
                ClockInNote = inNote,
                ClockOutNote = outNote,
                OTHoursNote = otNote,
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

    /// <summary>Human-readable OT math for one span, e.g. <c>OT: 1.0 hrs = floor((19:00 − 18:00) / 0.5h) × 0.5h [PROJ-A]</c>.</summary>
    private static string OvertimeEquation(TimeSpan start, TimeSpan end, TimeOnly otStartTime, TimeOnly dayStartTime, decimal hours, string? projectId)
    {
        TimeSpan otStart = ShiftOffset(otStartTime, dayStartTime);
        TimeSpan otBegin = start > otStart ? start : otStart;
        string tag = string.IsNullOrEmpty(projectId) ? string.Empty : $" [{projectId}]";
        if (end <= otBegin)
        {
            return start > otStart
                ? $"OT: 0.0 hrs (effective end {HmOffset(end, dayStartTime)} <= effective start {HmOffset(start, dayStartTime)}){tag}"
                : $"OT: 0.0 hrs (effective end {HmOffset(end, dayStartTime)} <= OT start {Hm(otStartTime)}){tag}";
        }
        return $"OT: {Hrs(hours)} hrs = floor(({HmOffset(end, dayStartTime)} − {HmOffset(otBegin, dayStartTime)}) / 0.5h) × 0.5h{tag}";
    }

    // =========================================================================
    // 4. VALIDATION WARNINGS
    // =========================================================================

    /// <summary>
    /// Warnings for <c>ValidateSchedule</c>: on-board project slots whose Start / End was auto-filled from a fallback,
    /// slots whose window is inverted (End before Start), and workers with no CheckOut punch whose Clock-Out fell back
    /// to the project End.
    /// </summary>
    internal static List<string> CollectTimeWarnings(
        DateOnly date,
        IReadOnlyList<DailyProjectSchedule> slots,
        IReadOnlyList<AttendanceEntry> entries,
        TimeOnly dayStartTime,
        TimeOnly? nextDayStartTime,
        TimeOnly otStartTime)
    {
        var warnings = new List<string>();
        var spans = BuildSpans(entries, date, dayStartTime, nextDayStartTime, slots, otStartTime);
        var windows = new ProjectWindowIndex(slots, spans, dayStartTime);

        static string Band(Shifts s) => s == Shifts.Night ? "Night Work" : "Day Work";

        foreach (var kv in windows.BySlot.OrderBy(k => k.Key.Shift == Shifts.Morning ? 0 : 1).ThenBy(k => k.Key.ProjectId, StringComparer.OrdinalIgnoreCase))
        {
            var (shift, projectId) = kv.Key;
            var w = kv.Value;
            if (w.StartIsFallback)
                warnings.Add($"{FallbackIcon} [{Band(shift)}] {projectId}: Start {Hm(w.Start)} auto-filled from the latest check-in ({string.Join(", ", w.StartWorkers)}) — no manual Start time.");
            if (w.EndIsFallback)
                warnings.Add($"{FallbackIcon} [{Band(shift)}] {projectId}: End {Hm(w.End)} auto-filled from the earliest exit ({string.Join(", ", w.EndContributors.Select(c => c.WorkerName))}) — no manual End time.");
            if (w.IsInverted)
                warnings.Add($"{FallbackIcon} [{Band(shift)}] {projectId}: End {Hm(w.End)} is before Start {Hm(w.Start)} — please correct the project times.");
        }

        foreach (var kv in spans.OrderBy(k => k.Key.StaffId).ThenBy(k => k.Key.ProjectId, StringComparer.OrdinalIgnoreCase))
        {
            var (staffId, projectId) = kv.Key;
            var window = windows.For(staffId, projectId);
            if (!window.End.HasValue || !window.EndOffset.HasValue) continue;
            foreach (var sp in kv.Value.Where(s => !s.Out.HasValue || !s.OutOffset.HasValue))
            {
                warnings.Add($"{FallbackIcon} {Roster.NameOf(staffId, $"#{staffId}")} checked into {projectId} at {Hm(sp.In)} with no CheckOut punch — Clock-Out fell back to {projectId} End {Hm(window.End.Value)}.");
            }
        }

        return warnings;
    }
}
