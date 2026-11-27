namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

/// <summary>
/// Reads and renders the 'Daily Schedule' dispatch board: a dynamic number of four-column slots (B:E, F:I, J:M, ...),
/// one per active project of the shown date, for Day Work (rows 4-18) and Night Work (rows 20-34).
///   row 4 / 20 : Col 0 Project Code (+ requirement flag) | Col 1 Premix Hole Amount / Type icons | Col 2 Project Start Time | Col 3 Project End Time
///   workers    : Col 0 Staff Name | Col 1 Staff Skills (icons) | Col 2 Effective Check-In | Col 3 Effective Check-Out   (10 rows: 5-14 / 21-30)
///   vehicles   : Col 0 Vehicle Plate | Col 1 Vehicle Type (icon) | Col 2 Current Odometer | Col 3 blank                   (4 rows: 15-18 / 31-34)
/// Project Start / End are real Excel time values: a manual time typed / picked by the user (kept in <see cref="_manualTimes"/>,
/// persisted as <see cref="DailyProjectSchedule.ActualStartTime"/> / <see cref="DailyProjectSchedule.ActualEndTime"/>), else an
/// auto fallback (latest check-in / earliest exit) drawn in amber with a "⚠" prefix that is never read back as manual.
/// The template carries its own static labels; the board only writes the header row and the data blocks.
/// The board is not a ListObject, so reads and writes are single 2D-array COM calls per block.
/// </summary>
internal static class ScheduleBoardMapper
{
    // =========================================================================
    // 1. LAYOUT
    // =========================================================================
    internal static readonly string[] SheetAliases = { "Daily Schedule", "Daily_Schedule", "Schedule" };
    internal const string DefaultSheetName = "Daily Schedule";
    internal const string DateCell = "A2";
    internal const string DayCell = "B2";

    internal const int SlotWidth = 4;
    internal const int FirstSlotColumn = 2;                 // B
    internal const int WorkerRowsPerBand = 10;
    internal const int VehicleRowsPerBand = 4;
    private const int FirstBoardRow = 4;                    // Day Work project header row
    private const int LastBoardRow = 34;
    private const int MaxScanSlots = 100;                   // upper bound when detecting the rendered board width
    private const string EmptyHeaderText = "Project Code";
    private static readonly HashSet<string> PlaceholderTexts = new(StringComparer.OrdinalIgnoreCase) { "Staff Name", "Vehicle Plate" };

    /// <summary>Slots currently drawn across the board (refreshed by <see cref="Read"/> / <see cref="Render"/>).</summary>
    private static int _slotCount = 1;
    internal static void ResetSlotCount()
    {
        _slotCount = 1;
        ClearManualTimes();
    }

    /// <summary>OLE (BGR) colour of auto-fallback times: RGB(180, 101, 0), a readable dark amber.</summary>
    internal const int FallbackAmberColor = 0x0065B4;

    private const string TimeFormat = "hh:mm";
    private static readonly string FallbackTimeFormat = $"\"{AttendanceRollupService.FallbackIcon} \"hh:mm";
    private const string ClearTimeToken = "__CLEAR_TIME__";
    private const string ClearTimeDisplay = "✕ Clear (Use Auto Fallback)";

    /// <summary>
    /// Manual project Start / End per (band, project) for the shown date. The board cells also show auto-fallback times,
    /// so <see cref="Read"/> takes manual times from here only - an amber fallback is never promoted to a manual override.
    /// Hydrated on a date load (<see cref="Render"/> with autoPopulate), edited by <see cref="CommitManualTimeEdit"/> /
    /// <see cref="ApplyProjectTimePick"/>.
    /// </summary>
    private static readonly Dictionary<(Shifts Shift, string ProjectId), (TimeOnly? Start, TimeOnly? End)> _manualTimes = new();

    private static string TimeKey(string? projectId) => (projectId ?? string.Empty).Trim().ToUpperInvariant();

    internal static void ClearManualTimes() => _manualTimes.Clear();

    private static (TimeOnly? Start, TimeOnly? End) ManualTimesFor(Shifts shift, string projectId) =>
        _manualTimes.TryGetValue((shift, TimeKey(projectId)), out var t) ? t : (null, null);

    private static void SetManualTime(Shifts shift, string projectId, bool isStart, TimeOnly? value)
    {
        var key = (shift, TimeKey(projectId));
        var current = _manualTimes.TryGetValue(key, out var t) ? t : (Start: (TimeOnly?)null, End: (TimeOnly?)null);
        current = isStart ? (value, current.End) : (current.Start, value);
        if (current.Start == null && current.End == null) _manualTimes.Remove(key);
        else _manualTimes[key] = current;
    }

    /// <summary>Raw value of the "✕ Remove Slot" Project Header option.</summary>
    private const string RemoveSlotToken = "__REMOVE_SLOT__";
    private const string RemoveSlotDisplay = "✕ Remove Slot";

    /// <summary>
    /// Auto-scheduled projects removed from the shown date's board ("✕ Remove Slot"). Persisted per date as
    /// <see cref="ScheduleRecord.DismissedProjects"/> so a date load never auto-populates them again.
    /// </summary>
    private static HashSet<string> _dismissedProjects = new(StringComparer.OrdinalIgnoreCase);

    internal static void SetDismissedProjects(IEnumerable<string>? projects) =>
        _dismissedProjects = new HashSet<string>(
            (projects ?? Enumerable.Empty<string>()).Where(p => !string.IsNullOrWhiteSpace(p)).Select(p => p.Trim().ToUpperInvariant()),
            StringComparer.OrdinalIgnoreCase);

    internal static List<string> GetDismissedProjects() =>
        _dismissedProjects.OrderBy(p => p, StringComparer.OrdinalIgnoreCase).ToList();
    private static int LastBoardColumn => FirstSlotColumn + (_slotCount * SlotWidth) - 1;

    private sealed record Band(Shifts Shift, string Label, int HeaderRow, int FirstWorkerRow, int FirstVehicleRow)
    {
        public int LastWorkerRow => FirstWorkerRow + WorkerRowsPerBand - 1;
        public int LastVehicleRow => FirstVehicleRow + VehicleRowsPerBand - 1;
    }

    private static Band MakeBand(Shifts shift, string label, int offset) =>
        new(shift, label, 4 + offset, 5 + offset, 15 + offset);

    private static readonly Band[] Bands = { MakeBand(Shifts.Morning, "Day Work", 0), MakeBand(Shifts.Night, "Night Work", 16) };

    internal static int SlotColumn(int slot) => FirstSlotColumn + (slot * SlotWidth);

    private static readonly Regex HeaderCode = new(@"^\s*(\S+)", RegexOptions.Compiled);
    private static readonly Regex LeadingInt = new(@"^\s*(\d+)", RegexOptions.Compiled);

    /// <summary>Premix header cell keeps a numeric value when &gt; 0 ("1 Hole", "5 Holes"); when 0 it displays the project type instead. Numeric premix holes only apply on <c>Road_Reinstatement</c> projects.</summary>
    private const string PremixHoleFormat = "[=1]0 \"Hole\";0 \"Holes\"";
    private const string EmptyOdometerDash = "—";

    internal enum CellKind { None, Header, Premix, Worker, WorkerSkill, Vehicle, VehicleType, ProjectStart, ProjectEnd }
    internal readonly record struct CellLocation(CellKind Kind, Shifts Shift, int Slot, int Row);

    internal static CellLocation Locate(int row, int column)
    {
        var none = new CellLocation(CellKind.None, Shifts.None, 0, row);
        if (column < FirstSlotColumn || column > LastBoardColumn) return none;
        int slot = (column - FirstSlotColumn) / SlotWidth;
        int within = (column - FirstSlotColumn) % SlotWidth;

        foreach (var b in Bands)
        {
            if (row == b.HeaderRow)
            {
                if (within == 0) return new(CellKind.Header, b.Shift, slot, row);
                if (within == 1) return new(CellKind.Premix, b.Shift, slot, row);
                if (within == 2) return new(CellKind.ProjectStart, b.Shift, slot, row);
                if (within == 3) return new(CellKind.ProjectEnd, b.Shift, slot, row);
            }
            if (row >= b.FirstWorkerRow && row <= b.LastWorkerRow)
            {
                if (within == 0) return new(CellKind.Worker, b.Shift, slot, row);
                if (within == 1) return new(CellKind.WorkerSkill, b.Shift, slot, row);
            }
            if (row >= b.FirstVehicleRow && row <= b.LastVehicleRow)
            {
                if (within == 0) return new(CellKind.Vehicle, b.Shift, slot, row);
                if (within == 1) return new(CellKind.VehicleType, b.Shift, slot, row);
            }
        }
        return none;
    }

    /// <summary>Any cell right of the date column inside the board rows (the width is dynamic, so there is no upper column bound).</summary>
    internal static bool IsBoardArea(int row, int column) =>
        row >= FirstBoardRow && row <= LastBoardRow && column >= FirstSlotColumn;

    /// <summary>
    /// Counts the slots actually drawn on the sheet: last populated project header cell (starting at B) on the
    /// header row of either band. Always at least 1.
    /// </summary>
    private static int DetectSlotCount(object wsObj)
    {
        int best = 1;
        try
        {
            dynamic ws = wsObj;
            int cols = MaxScanSlots * SlotWidth;
            foreach (var band in Bands)
            {
                dynamic range = ws.Range[ws.Cells[band.HeaderRow, FirstSlotColumn], ws.Cells[band.HeaderRow, FirstSlotColumn + cols - 1]];
                object[,] v = TableMapper.GetRange2DValues(range, 1, cols);
                // Last populated header wins, so a header the user just cleared does not truncate the board.
                for (int slot = MaxScanSlots - 1; slot >= best; slot--)
                    if (!string.IsNullOrWhiteSpace(v[1, (slot * SlotWidth) + 1]?.ToString())) { best = slot + 1; break; }
            }
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] slot detection failed: {ex.Message}"); }
        return best;
    }

    internal static bool IsBoardSheet(object sheetObj)
    {
        try
        {
            dynamic sheet = sheetObj;
            string name = sheet.Name?.ToString() ?? string.Empty;
            return SheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));
        }
        catch { return false; }
    }

    // =========================================================================
    // 2. DATE CELL
    // =========================================================================
    internal static DateOnly? ReadDate(object wsObj)
    {
        dynamic ws = wsObj;
        try
        {
            object? raw = ws.Range[DateCell].Value2;
            if (raw is double d && d > 1) return DateOnly.FromDateTime(DateTime.FromOADate(d));
            if (raw != null && DateTime.TryParse(raw.ToString(), CultureInfo.CurrentCulture, DateTimeStyles.None, out var dt))
                return DateOnly.FromDateTime(dt);
        }
        catch { }
        return null;
    }

    internal static void WriteDate(object wsObj, DateOnly date)
    {
        dynamic ws = wsObj;
        ws.Range[DateCell].Value2 = date.ToDateTime(TimeOnly.MinValue).ToOADate();
        try { ws.Range[DateCell].NumberFormat = "dd/mm/yyyy"; } catch { }
        ws.Range[DayCell].Value2 = DayName(date);
    }

    private const string IsHolidayHeader = "Is Holiday";
    private const int IsHolidayScanColumns = 26; // A1:Z1

    /// <summary>
    /// Writes the SQL Payroll public-holiday flag under the Row 1 "Is Holiday" header (A1:Z1, case-insensitive) of
    /// <paramref name="ws"/> ('Daily Schedule' / 'Daily Attendances'): TRUE / FALSE in Row 2, the holiday description
    /// as a cell comment, and a read-only <c>=FALSE</c> validation (auto-evaluated, never typed). No-op when the sheet
    /// has no such header. Call inside the caller's Excel freeze (macro thread); never selects or activates.
    /// </summary>
    internal static void WriteIsHolidayCell(dynamic ws, bool isHoliday, string holidayDescription)
    {
        try
        {
            object[,] header = TableMapper.GetRange2DValues(ws.Range["A1"].Resize[1, IsHolidayScanColumns], 1, IsHolidayScanColumns);
            int r0 = header.GetLowerBound(0), c0 = header.GetLowerBound(1);
            int column = 0;
            for (int c = 0; c < IsHolidayScanColumns; c++)
            {
                string text = (Convert.ToString(header[r0, c0 + c], CultureInfo.InvariantCulture) ?? string.Empty).Trim();
                if (text.Equals(IsHolidayHeader, StringComparison.OrdinalIgnoreCase)) { column = c + 1; break; }
            }
            if (column == 0)
            {
                string sheetName = string.Empty;
                try { sheetName = Convert.ToString(ws.Name) ?? string.Empty; } catch { }
                AppLogger.Info($"[ScheduleManager] '{sheetName}' has no '{IsHolidayHeader}' header in A1:Z1; holiday flag not written.");
                return;
            }

            dynamic cell = ws.Cells[2, column];
            cell.Value2 = isHoliday;

            try { cell.ClearComments(); } catch { }
            if (isHoliday)
            {
                string note = string.IsNullOrWhiteSpace(holidayDescription) ? "Public holiday" : holidayDescription.Trim();
                try { cell.AddComment($"{note}\n(SQL Payroll calendar)"); } catch { }
            }

            // Read-only: auto-evaluated from SQL Payroll (template locking convention: custom validation "=FALSE").
            try
            {
                cell.Validation.Delete();
                cell.Validation.Add(7 /* xlValidateCustom */, 1 /* xlValidAlertStop */, 1 /* xlBetween */, "=FALSE");
                cell.Validation.ErrorTitle = IsHolidayHeader;
                cell.Validation.ErrorMessage = Loc.T("schedule.calendar.isHolidayLocked");
            }
            catch { }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] Writing '{IsHolidayHeader}' failed: {ex.Message}");
        }
    }

    /// <summary>English day name (Monday..Sunday) - exactly the values of the <see cref="Days"/> enum list that backs the B2 dropdown.</summary>
    internal static string DayName(DateOnly date) => date.DayOfWeek.ToString();

    private static bool IsDateCell(dynamic cell) => (int)cell.Row == 2 && (int)cell.Column == 1;
    private static bool IsDayCell(dynamic cell) => (int)cell.Row == 2 && (int)cell.Column == 2;

    // =========================================================================
    // 3. READ (one COM call)
    // =========================================================================
    private static object[,] ReadGrid(object wsObj)
    {
        dynamic ws = wsObj;
        dynamic range = ws.Range[ws.Cells[FirstBoardRow, FirstSlotColumn], ws.Cells[LastBoardRow, LastBoardColumn]];
        return TableMapper.GetRange2DValues(range, LastBoardRow - FirstBoardRow + 1, LastBoardColumn - FirstSlotColumn + 1);
    }

    private static string Text(object[,] grid, int row, int col) =>
        grid[row - FirstBoardRow + 1, col - FirstSlotColumn + 1]?.ToString()?.Trim() ?? string.Empty;

    internal static string ParseProjectCode(string headerText)
    {
        if (string.IsNullOrWhiteSpace(headerText) || headerText.StartsWith("Project Code", StringComparison.OrdinalIgnoreCase)) return string.Empty;
        var m = HeaderCode.Match(headerText);
        return m.Success ? m.Groups[1].Value.Trim().ToUpperInvariant() : string.Empty;
    }

    /// <summary>Reads every non-empty slot of the rendered width. Rank prefixes ("[95%] ") pasted into a cell are stripped.</summary>
    internal static List<DailyProjectSchedule> Read(object wsObj, DateOnly date)
    {
        _slotCount = DetectSlotCount(wsObj);
        object[,] grid = ReadGrid(wsObj);
        var result = new List<DailyProjectSchedule>();

        foreach (var band in Bands)
        {
            for (int slot = 0; slot < _slotCount; slot++)
            {
                int col = SlotColumn(slot);
                string projectId = ParseProjectCode(Text(grid, band.HeaderRow, col));
                var record = new DailyProjectSchedule
                {
                    DailyScheduleDate = date,
                    Shift = band.Shift,
                    SlotIndex = slot,
                    ProjectID = projectId,
                    PremixHoles = Roster.SupportsPremixHoles(projectId) ? ParsePremixHoles(Text(grid, band.HeaderRow, col + 1)) : 0
                };
                if (projectId.Length > 0)
                {
                    // Manual times only (never the amber auto-fallback values drawn in the cells).
                    var manual = ManualTimesFor(band.Shift, projectId);
                    record.ActualStartTime = manual.Start;
                    record.ActualEndTime = manual.End;
                }

                for (int r = band.FirstWorkerRow; r <= band.LastWorkerRow; r++)
                {
                    string worker = ScheduleMatchEngine.StripPercent(Text(grid, r, col));
                    if (worker.Length > 0 && !PlaceholderTexts.Contains(worker)) record.Workers.Add(worker);
                }
                for (int r = band.FirstVehicleRow; r <= band.LastVehicleRow; r++)
                {
                    string plate = Text(grid, r, col).ToUpperInvariant();
                    if (plate.Length > 0 && !PlaceholderTexts.Contains(plate)) record.Vehicles.Add(plate);
                }

                if (record.ProjectID.Length > 0 || record.Workers.Count > 0 || record.Vehicles.Count > 0 || record.PremixHoles > 0
                    || record.ActualStartTime.HasValue || record.ActualEndTime.HasValue)
                    result.Add(record);
            }
        }
        return result;
    }

    /// <summary>Accepts a raw number ("5") or typed text with a unit ("5 Holes"); anything else is 0.</summary>
    private static int ParsePremixHoles(string text)
    {
        if (int.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out int holes)) return Math.Max(0, holes);
        var m = LeadingInt.Match(text);
        return m.Success && int.TryParse(m.Groups[1].Value, NumberStyles.Integer, CultureInfo.InvariantCulture, out holes) ? Math.Max(0, holes) : 0;
    }

    /// <summary>One duplicated worker inside a shift band: the slots involved and the exact worker cells (slot, index in slot.Workers).</summary>
    private sealed record WorkerConflict(Shifts Shift, string Worker, List<int> Slots, List<(int Slot, int Index)> Cells);

    /// <summary>
    /// Two check-in spans conflict when they start together, are both still open, or their time ranges overlap.
    /// Back-to-back spans (08:00-12:00, 12:00-17:00) do not conflict. Times are compared as exact (date-aware) offsets
    /// from the working-day start so night spans crossing midnight (22:00-02:00) and >24 h working days compare correctly.
    /// </summary>
    private static bool SpansConflict(WorkerProjectSpan a, WorkerProjectSpan b)
    {
        TimeSpan aIn = a.InOffset;
        TimeSpan bIn = b.InOffset;
        if (aIn == bIn) return true;
        if (a.OutOffset == null && b.OutOffset == null) return true;
        if (a.OutOffset == null) return aIn < b.OutOffset!.Value;
        TimeSpan aOut = a.OutOffset.Value;
        if (b.OutOffset == null) return bIn < aOut;
        TimeSpan bOut = b.OutOffset.Value;
        return aIn < bOut && bIn < aOut;
    }

    private static List<WorkerProjectSpan> SpansFor(Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans, long staffId, string projectId) =>
        spans.TryGetValue((staffId, projectId.Trim().ToUpperInvariant()), out var list) ? list : new List<WorkerProjectSpan>();

    /// <summary>
    /// Time-aware worker duplicates per band. The same worker twice in one slot always counts; across slots of a band
    /// only when both slots have check-in spans for them and any pair of spans conflicts.
    /// </summary>
    private static List<WorkerConflict> ComputeWorkerConflicts(
        IEnumerable<DailyProjectSchedule> slots,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans)
    {
        var result = new List<WorkerConflict>();
        var all = slots.ToList();
        foreach (var band in Bands)
        {
            var groups = all.Where(s => s.Shift == band.Shift)
                .SelectMany(s => s.Workers.Select((w, i) => (Slot: s, Name: w, Index: i)))
                .Where(x => !string.IsNullOrWhiteSpace(x.Name))
                .GroupBy(x => Roster.FindStaff(x.Name) is { } st ? $"#{st.Id}" : "n:" + x.Name.Trim().ToUpperInvariant());

            foreach (var g in groups)
            {
                long? staffId = Roster.FindStaff(g.First().Name)?.Id;
                var bySlot = g.GroupBy(x => x.Slot.SlotIndex).ToList();
                var hit = new HashSet<int>(bySlot.Where(s => s.Count() > 1).Select(s => s.Key));

                if (staffId.HasValue)
                {
                    var withSpans = bySlot
                        .Select(s => (Slot: s.Key, Spans: SpansFor(spans, staffId.Value, s.First().Slot.ProjectID)))
                        .Where(x => x.Spans.Count > 0).ToList();
                    for (int i = 0; i < withSpans.Count; i++)
                        for (int j = i + 1; j < withSpans.Count; j++)
                            if (withSpans[i].Spans.Any(a => withSpans[j].Spans.Any(b => SpansConflict(a, b))))
                            {
                                hit.Add(withSpans[i].Slot);
                                hit.Add(withSpans[j].Slot);
                            }
                }

                if (hit.Count == 0) continue;
                result.Add(new WorkerConflict(band.Shift, g.First().Name, hit.OrderBy(x => x).ToList(),
                    g.Where(x => hit.Contains(x.Slot.SlotIndex)).Select(x => (x.Slot.SlotIndex, x.Index)).ToList()));
            }
        }
        return result;
    }

    /// <summary>Workers double-booked within a shift band (time-aware, see <see cref="ComputeWorkerConflicts"/>), as save-blocking messages.</summary>
    internal static List<string> FindDuplicateWorkers(IEnumerable<DailyProjectSchedule> slots, DateOnly date, IEnumerable<AttendanceEntry> entries)
    {
        var list = slots.ToList();
        var spans = AttendanceRollupService.BuildSpans(entries, date, BoardContext.DayStartTime, BoardContext.NextDayStartTime, list, BoardContext.OTStartTime);
        var messages = new List<string>();
        foreach (var c in ComputeWorkerConflicts(list, spans))
        {
            string label = Bands.First(b => b.Shift == c.Shift).Label;
            string where = string.Join(", ", c.Slots.Select(s => $"slot {s + 1}"));
            messages.Add($"  [Daily Schedule - {label}]: Worker '{c.Worker}' is assigned more than once ({where}).");
        }
        return messages;
    }

    /// <summary>(shift, slot) of every slot whose project also appears in another slot of the same band.</summary>
    private static HashSet<(Shifts Shift, int Slot)> DuplicateProjectSlots(IEnumerable<DailyProjectSchedule> slots)
    {
        var result = new HashSet<(Shifts, int)>();
        foreach (var g in slots.Where(s => !string.IsNullOrWhiteSpace(s.ProjectID))
                     .GroupBy(s => (s.Shift, Id: s.ProjectID.Trim().ToUpperInvariant()))
                     .Where(g => g.Count() > 1))
            foreach (var s in g) result.Add((s.Shift, s.SlotIndex));
        return result;
    }

    /// <summary>The same project in more than one slot of a shift band, as save-blocking messages.</summary>
    internal static List<string> FindDuplicateProjects(IEnumerable<DailyProjectSchedule> slots)
    {
        var messages = new List<string>();
        foreach (var band in Bands)
        {
            var dups = slots.Where(s => s.Shift == band.Shift && !string.IsNullOrWhiteSpace(s.ProjectID))
                .GroupBy(s => s.ProjectID.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1);
            foreach (var g in dups)
            {
                string where = string.Join(", ", g.Select(x => $"slot {x.SlotIndex + 1}"));
                messages.Add($"  [Daily Schedule - {band.Label}]: Project '{g.Key}' appears more than once ({where}).");
            }
        }
        return messages;
    }

    // =========================================================================
    // 4. TEAM AUTO-EXPANSION
    // =========================================================================
    /// <summary>
    /// Replaces any worker entry that names a team (and is not a staff short name) with the team leader plus every
    /// member whose Team list contains it, and copies the team's default vehicles into free vehicle rows.
    /// Duplicates across slots are intentionally allowed here; they are highlighted and block the save instead.
    /// </summary>
    internal static bool ExpandTeams(List<DailyProjectSchedule> slots, out string? notice)
    {
        notice = null;
        bool changed = false;
        var teams = BoardContext.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)).ToList();
        if (teams.Count == 0) return false;

        foreach (var slot in slots)
        {
            var expanded = new List<string>();
            foreach (var entry in slot.Workers)
            {
                var team = Roster.FindStaff(entry) == null
                    ? teams.FirstOrDefault(t => t.TeamName.Trim().Equals(entry, StringComparison.OrdinalIgnoreCase))
                    : null;
                if (team == null) { expanded.Add(entry); continue; }

                changed = true;
                var members = new List<string>();
                if (!string.IsNullOrWhiteSpace(team.TeamLeader) && Roster.FindStaff(team.TeamLeader) is { } lead &&
                    Roster.IsActiveOn(lead, slot.DailyScheduleDate))
                    members.Add(lead.ShortName);
                members.AddRange(BoardContext.Preferences
                    .Where(p => p.Team.Contains(team.TeamName.Trim(), StringComparer.OrdinalIgnoreCase))
                    .Where(p => Roster.IsActiveOn(p.StaffID, slot.DailyScheduleDate))   // employed on the slot's date
                    .Select(p => Roster.NameOf(p.StaffID))
                    .Where(n => !string.IsNullOrWhiteSpace(n))
                    .OrderBy(n => n, StringComparer.OrdinalIgnoreCase));
                expanded.AddRange(members);

                foreach (var plate in team.DefaultVehicles)
                {
                    string p = plate.Trim().ToUpperInvariant();
                    if (p.Length > 0 && slot.Vehicles.Count < VehicleRowsPerBand && !slot.Vehicles.Contains(p, StringComparer.OrdinalIgnoreCase))
                        slot.Vehicles.Add(p);
                }
            }

            // De-duplicate inside the slot only (same worker twice in one slot is never meaningful).
            var distinct = expanded.Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (distinct.Count > WorkerRowsPerBand)
            {
                notice = $"Slot {slot.SlotIndex + 1} holds {WorkerRowsPerBand} workers; {distinct.Count - WorkerRowsPerBand} team member(s) did not fit.";
                distinct = distinct.Take(WorkerRowsPerBand).ToList();
            }
            slot.Workers = distinct;
        }
        return changed;
    }

    // =========================================================================
    // 5. RENDER (2D array writes, caller owns the Excel freeze)
    // =========================================================================
    /// <summary>
    /// Slots to draw for one band, left to right. With <paramref name="autoPopulate"/> (a date load) every project that is
    /// Active on the date gets a slot (iJob -> MWO -> PVT -> Other) merged with its saved row; saved rows for other projects
    /// follow. Auto-scheduled projects the user removed for this date (<see cref="_dismissedProjects"/>) are skipped.
    /// Without it (a live refresh) the slots the user currently has are kept in place.
    /// </summary>
    private static List<DailyProjectSchedule> LayoutSlots(Shifts shift, DateOnly date, IReadOnlyList<DailyProjectSchedule> slots, bool autoPopulate)
    {
        var saved = slots.Where(s => s.Shift == shift).OrderBy(s => s.SlotIndex).ToList();
        if (!autoPopulate) return saved;

        var result = new List<DailyProjectSchedule>();
        var used = new HashSet<int>();
        foreach (var project in Roster.ActiveProjectsOn(date))
        {
            if (_dismissedProjects.Contains(project.Id)) continue;

            int idx = -1;
            for (int i = 0; i < saved.Count && idx < 0; i++)
                if (!used.Contains(i) && saved[i].ProjectID.Trim().Equals(project.Id, StringComparison.OrdinalIgnoreCase)) idx = i;
            if (idx >= 0) { used.Add(idx); result.Add(saved[idx]); }
            else result.Add(new DailyProjectSchedule { DailyScheduleDate = date, Shift = shift, ProjectID = project.Id });
        }
        for (int i = 0; i < saved.Count; i++)
            if (!used.Contains(i)) result.Add(saved[i]);
        return result;
    }

    internal static void Render(object wsObj, DateOnly date, IReadOnlyList<DailyProjectSchedule> slots, IReadOnlyList<AttendanceEntry> entries, bool autoPopulate = false)
    {
        dynamic ws = wsObj;

        // A date load re-seeds the manual project times from the saved rows (live refreshes keep the in-memory edits).
        if (autoPopulate)
        {
            ClearManualTimes();
            foreach (var s in slots.Where(s => !string.IsNullOrWhiteSpace(s.ProjectID)))
            {
                if (s.ActualStartTime.HasValue) SetManualTime(s.Shift, s.ProjectID, isStart: true, s.ActualStartTime);
                if (s.ActualEndTime.HasValue) SetManualTime(s.Shift, s.ProjectID, isStart: false, s.ActualEndTime);
            }
        }

        var layouts = Bands.ToDictionary(b => b.Shift, b => LayoutSlots(b.Shift, date, slots, autoPopulate));

        // Always keep one blank slot on the right so a new project can be picked without widening the board by hand.
        // Only the board's rightmost column counts (a shorter band already has blank slots on its right).
        int maxCount = layouts.Values.Max(l => l.Count);
        bool lastOccupied = layouts.Values.Any(l => l.Count > 0 && l.Count == maxCount &&
            (!string.IsNullOrWhiteSpace(l[^1].ProjectID) || l[^1].Workers.Count > 0 || l[^1].Vehicles.Count > 0));
        int slotCount = Math.Max(1, lastOccupied ? maxCount + 1 : maxCount);

        // Laid-out board (position = slot): feeds per-worker match breakdowns and the conflict checks below.
        var board = new List<DailyProjectSchedule>();
        foreach (var band in Bands)
        {
            var layout = layouts[band.Shift];
            for (int i = 0; i < layout.Count; i++)
            {
                var manual = ManualTimesFor(band.Shift, layout[i].ProjectID);
                board.Add(new DailyProjectSchedule
                {
                    DailyScheduleDate = date, Shift = band.Shift, SlotIndex = i, ProjectID = layout[i].ProjectID,
                    Workers = layout[i].Workers, Vehicles = layout[i].Vehicles,
                    ActualStartTime = manual.Start ?? layout[i].ActualStartTime,
                    ActualEndTime = manual.End ?? layout[i].ActualEndTime
                });
            }
        }

        // Blank-project (thumbprint-only) punches are attributed in memory to the worker's on-board planned project.
        var spans = AttendanceRollupService.BuildSpans(entries, date, BoardContext.DayStartTime, BoardContext.NextDayStartTime, board, BoardContext.OTStartTime);

        int drawn = DetectSlotCount(wsObj);
        int previous = Math.Max(drawn, _slotCount);

        // Live refresh: snapshot the current board (old bounds, before PrepareColumns widens it) so a pending
        // Staff Skills filter on an empty worker row survives the rewrite.
        object[,]? existingGrid = null;
        int existingLastColumn = LastBoardColumn;
        if (!autoPopulate)
        {
            try { existingGrid = ReadGrid(wsObj); }
            catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] reading pending skill filters failed: {ex.Message}"); }
        }

        _slotCount = slotCount;
        PrepareColumns(ws, slotCount, drawn, previous);

        int cols = slotCount * SlotWidth;
        foreach (var band in Bands)
        {
            var layout = layouts[band.Shift];
            var header = new object?[1, cols];
            var workers = new object?[WorkerRowsPerBand, cols];
            var vehicles = new object?[VehicleRowsPerBand, cols];
            var headerComments = new string?[slotCount];
            var typeComments = new string?[slotCount];
            var startComments = new string?[slotCount];
            var endComments = new string?[slotCount];
            var startFallback = new bool[slotCount];
            var endFallback = new bool[slotCount];
            var workerComments = new string?[WorkerRowsPerBand, slotCount];
            var vehicleComments = new string?[VehicleRowsPerBand, slotCount];
            var inFormats = new string?[WorkerRowsPerBand, slotCount];
            var outFormats = new string?[WorkerRowsPerBand, slotCount];
            var inFallback = new bool[WorkerRowsPerBand, slotCount];
            var outFallback = new bool[WorkerRowsPerBand, slotCount];
            var inNotes = new string?[WorkerRowsPerBand, slotCount];
            var outNotes = new string?[WorkerRowsPerBand, slotCount];

            for (int slot = 0; slot < slotCount; slot++)
            {
                var data = slot < layout.Count ? layout[slot] : new DailyProjectSchedule { DailyScheduleDate = date, Shift = band.Shift };
                int c0 = slot * SlotWidth;

                // Same manual-time resolution as the laid-out board above (in-memory edits win over the loaded row).
                var manual = ManualTimesFor(band.Shift, data.ProjectID);
                ProjectWindow? window = data.ProjectID.Length > 0
                    ? AttendanceRollupService.WindowFor(data.ProjectID, spans, manual.Start ?? data.ActualStartTime, manual.End ?? data.ActualEndTime, BoardContext.DayStartTime)
                    : null;
                SlotContext? slotContext = data.Workers.Count > 0
                    ? ReadSlotContext(board, date, new CellLocation(CellKind.Worker, band.Shift, slot, 0))
                    : null;

                var assignedSkills = new List<string>();
                for (int i = 0; i < WorkerRowsPerBand && i < data.Workers.Count; i++)
                {
                    string name = data.Workers[i];
                    var staff = Roster.FindStaff(name);
                    var pref = staff != null ? BoardContext.PreferenceFor(staff.Id) : null;

                    var skills = SkillsFor(staff, pref);
                    assignedSkills.AddRange(skills);

                    if (staff != null && slotContext != null)
                    {
                        var profile = new WorkerProfile(staff, pref, BoardContext.DaysWorkedEarlierThisWeek(staff.Id, date), false);
                        workerComments[i, slot] = ScheduleMatchEngine.FormatBreakdown(profile, slotContext);
                    }

                    workers[i, c0] = name;
                    workers[i, c0 + 1] = string.Join(" ", skills.Select(BoardContext.IconFor).Where(s => s.Length > 0));
                    var resolved = staff != null && window != null
                        ? AttendanceRollupService.ResolveWorkerProjectSpan(staff.Id, data.ProjectID, spans, window)
                        : null;
                    if (resolved != null)
                    {
                        workers[i, c0 + 2] = resolved.EffectiveIn.ToTimeSpan().TotalDays;
                        workers[i, c0 + 3] = resolved.EffectiveOut.HasValue ? resolved.EffectiveOut.Value.ToTimeSpan().TotalDays : null;
                        inFormats[i, slot] = resolved.InFormat;
                        outFormats[i, slot] = resolved.OutFormat;
                        inFallback[i, slot] = resolved.InIsFallback;
                        outFallback[i, slot] = resolved.OutIsFallback;
                        inNotes[i, slot] = resolved.InNote;
                        outNotes[i, slot] = resolved.OutNote;
                    }
                }

                // Keep pending skill filters (Staff Skills set, Staff Name still empty) on unassigned rows,
                // even when no worker currently matches them.
                if (existingGrid != null && SlotColumn(slot) + 1 <= existingLastColumn)
                {
                    for (int i = data.Workers.Count; i < WorkerRowsPerBand; i++)
                    {
                        int r = band.FirstWorkerRow + i;
                        if (Text(existingGrid, r, SlotColumn(slot)).Length > 0) continue;
                        string pending = Text(existingGrid, r, SlotColumn(slot) + 1);
                        if (pending.Length > 0) workers[i, c0 + 1] = pending;
                    }
                }

                var vehicleTypes = new List<VehicleType>();
                for (int i = 0; i < VehicleRowsPerBand && i < data.Vehicles.Count; i++)
                {
                    var v = Roster.FindVehicle(data.Vehicles[i]);
                    vehicles[i, c0] = data.Vehicles[i];
                    if (v != null)
                    {
                        vehicleTypes.Add(v.Type);
                        vehicles[i, c0 + 1] = BoardContext.IconFor(v.Type);
                        vehicles[i, c0 + 2] = v.Odometer > 0 ? (double)v.Odometer : EmptyOdometerDash;
                        var eval = Roster.EvaluateVehicleOn(v, date);
                        if (!eval.IsValid) vehicleComments[i, slot] = eval.FormatComment(date);
                    }
                    else
                    {
                        vehicles[i, c0 + 2] = EmptyOdometerDash;
                        vehicleComments[i, slot] = $"⚠ Vehicle '{data.Vehicles[i]}' was not found in vehicle master.";
                    }
                }

                // Keep pending type filters (Vehicle Type set, Vehicle Plate still empty) on unassigned rows,
                // even when no vehicle currently matches them.
                if (existingGrid != null && SlotColumn(slot) + 1 <= existingLastColumn)
                {
                    for (int i = data.Vehicles.Count; i < VehicleRowsPerBand; i++)
                    {
                        int r = band.FirstVehicleRow + i;
                        if (Text(existingGrid, r, SlotColumn(slot)).Length > 0) continue;
                        string pendingType = Text(existingGrid, r, SlotColumn(slot) + 1);
                        if (pendingType.Length > 0) vehicles[i, c0 + 1] = pendingType;
                    }
                }

                var h = BuildHeader(data, assignedSkills, vehicleTypes);
                header[0, c0] = h.Text;
                header[0, c0 + 1] = FormatPremixCell(data);
                headerComments[slot] = h.Comment;
                typeComments[slot] = BuildProjectTypeComment(data);
                if (window != null)
                {
                    header[0, c0 + 2] = window.Start.HasValue ? window.Start.Value.ToTimeSpan().TotalDays : null;
                    header[0, c0 + 3] = window.End.HasValue ? window.End.Value.ToTimeSpan().TotalDays : null;
                    startComments[slot] = window.StartNote;
                    endComments[slot] = window.EndNote;
                    startFallback[slot] = window.StartIsFallback;
                    endFallback[slot] = window.EndIsFallback;
                }
            }

            WriteBlock(ws, band.HeaderRow, band.HeaderRow, header);
            WriteBlock(ws, band.FirstWorkerRow, band.LastWorkerRow, workers);
            WriteBlock(ws, band.FirstVehicleRow, band.LastVehicleRow, vehicles);

            try { ws.Range[ws.Cells[band.FirstWorkerRow, FirstSlotColumn], ws.Cells[band.LastWorkerRow, FirstSlotColumn + cols - 1]].ClearComments(); } catch { }
            for (int slot = 0; slot < slotCount; slot++)
                for (int i = 0; i < WorkerRowsPerBand; i++)
                {
                    if (!string.IsNullOrEmpty(workerComments[i, slot]))
                        try { ApplySizedComment(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot)], workerComments[i, slot]!); } catch { }
                    if (!string.IsNullOrEmpty(inNotes[i, slot]))
                        try { ApplySizedComment(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot) + 2], inNotes[i, slot]!); } catch { }
                    if (!string.IsNullOrEmpty(outNotes[i, slot]))
                        try { ApplySizedComment(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot) + 3], outNotes[i, slot]!); } catch { }
                }

            // Effective check-in / check-out: source-icon NumberFormats, amber font on auto-fallback times.
            for (int slot = 0; slot < slotCount; slot++)
            {
                try
                {
                    dynamic timeBlock = ws.Range[ws.Cells[band.FirstWorkerRow, SlotColumn(slot) + 2], ws.Cells[band.LastWorkerRow, SlotColumn(slot) + 3]];
                    try { timeBlock.NumberFormat = TimeFormat; } catch { }
                    try { timeBlock.Font.ColorIndex = -4105; } catch { }
                }
                catch { }
                for (int i = 0; i < WorkerRowsPerBand; i++)
                {
                    if (!string.IsNullOrEmpty(inFormats[i, slot]) || inFallback[i, slot])
                        StyleTimeCell(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot) + 2], inFormats[i, slot], inFallback[i, slot]);
                    if (!string.IsNullOrEmpty(outFormats[i, slot]) || outFallback[i, slot])
                        StyleTimeCell(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot) + 3], outFormats[i, slot], outFallback[i, slot]);
                }
            }

            try { ws.Range[ws.Cells[band.FirstVehicleRow, FirstSlotColumn], ws.Cells[band.LastVehicleRow, FirstSlotColumn + cols - 1]].ClearComments(); } catch { }
            for (int slot = 0; slot < slotCount; slot++)
                for (int i = 0; i < VehicleRowsPerBand; i++)
                {
                    if (string.IsNullOrEmpty(vehicleComments[i, slot])) continue;
                    try { ApplySizedComment(ws.Cells[band.FirstVehicleRow + i, SlotColumn(slot)], vehicleComments[i, slot]!); } catch { }
                }

            try { ws.Range[ws.Cells[band.HeaderRow, FirstSlotColumn], ws.Cells[band.HeaderRow, FirstSlotColumn + cols - 1]].ClearComments(); } catch { }
            for (int slot = 0; slot < slotCount; slot++)
            {
                try { ws.Cells[band.HeaderRow, SlotColumn(slot) + 1].NumberFormat = PremixHoleFormat; } catch { }
                if (!string.IsNullOrEmpty(headerComments[slot]))
                    try { ApplySizedComment(ws.Cells[band.HeaderRow, SlotColumn(slot)], headerComments[slot]!); } catch { /* comments are advisory */ }
                if (!string.IsNullOrEmpty(typeComments[slot]))
                    try { ApplySizedComment(ws.Cells[band.HeaderRow, SlotColumn(slot) + 1], typeComments[slot]!); } catch { }

                // Project Start / End: plain "hh:mm" when manual (or empty), "⚠ hh:mm" in amber when auto-filled.
                StyleTimeCell(ws.Cells[band.HeaderRow, SlotColumn(slot) + 2], startFallback[slot] ? FallbackTimeFormat : TimeFormat, startFallback[slot]);
                StyleTimeCell(ws.Cells[band.HeaderRow, SlotColumn(slot) + 3], endFallback[slot] ? FallbackTimeFormat : TimeFormat, endFallback[slot]);
                if (!string.IsNullOrEmpty(startComments[slot]))
                    try { ApplySizedComment(ws.Cells[band.HeaderRow, SlotColumn(slot) + 2], startComments[slot]!); } catch { }
                if (!string.IsNullOrEmpty(endComments[slot]))
                    try { ApplySizedComment(ws.Cells[band.HeaderRow, SlotColumn(slot) + 3], endComments[slot]!); } catch { }
            }
        }

        // Conflicts are computed from the laid-out board (position = slot), reusing the save-time logic.
        var workerCells = new HashSet<(int Row, int Col)>();
        foreach (var c in ComputeWorkerConflicts(board, spans))
        {
            var band = Bands.First(b => b.Shift == c.Shift);
            foreach (var (slot, index) in c.Cells)
                if (index < WorkerRowsPerBand) workerCells.Add((band.FirstWorkerRow + index, SlotColumn(slot)));
        }
        var headerCells = new HashSet<(int Row, int Col)>(DuplicateProjectSlots(board)
            .Select(d => (Bands.First(b => b.Shift == d.Shift).HeaderRow, SlotColumn(d.Slot))));

        EnsureConditionalFormats(ws, slotCount, workerCells, headerCells);
    }

    /// <summary>Sets a time cell's NumberFormat (default "hh:mm") and its font: amber when auto-fallback, automatic otherwise.</summary>
    internal static void StyleTimeCell(dynamic cell, string? numberFormat, bool isFallback)
    {
        try { cell.NumberFormat = string.IsNullOrEmpty(numberFormat) ? TimeFormat : numberFormat; } catch { }
        if (isFallback) { try { cell.Font.Color = FallbackAmberColor; } catch { } }
        else { try { cell.Font.ColorIndex = -4105; } catch { } }
    }

    // Excel does not reflow AutoSize for hidden comments while ScreenUpdating = false, so the box
    // is sized deterministically from text metrics (points), padded for >100% DPI font scaling.
    internal static void ApplySizedComment(dynamic cell, string text)
    {
        string normalized = text.Replace("\r\n", "\n");
        dynamic comment = cell.AddComment(normalized);

        double maxLinePt = 0.0;
        var linePts = new List<double>();
        double maxLineHeightPt = 0.0;

        // Exact GDI measurement with Excel's classic Note font (Tahoma 9pt).
        using (var bmp = new System.Drawing.Bitmap(1, 1))
        using (var g = System.Drawing.Graphics.FromImage(bmp))
        using (var noteFont = new System.Drawing.Font("Tahoma", 9f))
        {
            double ptPerPxX = 72.0 / (g.DpiX > 0 ? g.DpiX : 96f);
            double ptPerPxY = 72.0 / (g.DpiY > 0 ? g.DpiY : 96f);
            var unbounded = new System.Drawing.Size(int.MaxValue, int.MaxValue);
            const System.Windows.Forms.TextFormatFlags flags =
                System.Windows.Forms.TextFormatFlags.NoPrefix | System.Windows.Forms.TextFormatFlags.SingleLine;

            foreach (string line in normalized.Split('\n'))
            {
                string sample = line.Length == 0 ? " " : line;
                var measured = System.Windows.Forms.TextRenderer.MeasureText(g, sample, noteFont, unbounded, flags);
                double linePt = measured.Width * ptPerPxX * 1.08;
                linePts.Add(linePt);
                if (linePt > maxLinePt) maxLinePt = linePt;
                maxLineHeightPt = Math.Max(maxLineHeightPt, measured.Height * ptPerPxY);
            }
        }

        double lineHeightPt = Math.Max(11.5, maxLineHeightPt);
        const double padX = 16.0, padY = 12.0;
        double targetWidth = Math.Clamp(maxLinePt + padX, 110.0, 380.0);
        double innerWidth = Math.Max(90.0, targetWidth - padX);
        int wrappedLines = 0;
        foreach (double lp in linePts)
            wrappedLines += Math.Max(1, (int)Math.Ceiling(lp / innerWidth));
        double targetHeight = Math.Max(36.0, (wrappedLines * lineHeightPt * 1.12) + padY);

        try
        {
            dynamic shape = comment.Shape;
            shape.TextFrame.AutoSize = false;
            shape.Width = targetWidth;
            shape.Height = targetHeight;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] ApplySizedComment sizing failed: {ex.Message}");
        }
    }

    private static void WriteBlock(dynamic ws, int firstRow, int lastRow, object?[,] values)
    {
        dynamic range = ws.Range[ws.Cells[firstRow, FirstSlotColumn], ws.Cells[lastRow, LastBoardColumn]];
        range.Value2 = values;
    }

    /// <summary>
    /// Clones the prototype B:D formatting / widths (Range B4:D18 and B20:D34) into slots that are not drawn yet and
    /// clears columns left over from a previously loaded date that had more slots.
    /// </summary>
    private static void PrepareColumns(dynamic ws, int slotCount, int drawn, int previous)
    {
        try
        {
            for (int slot = Math.Max(1, drawn); slot < slotCount; slot++)
            {
                foreach (var band in Bands)
                {
                    dynamic proto = ws.Range[ws.Cells[band.HeaderRow, FirstSlotColumn], ws.Cells[band.LastVehicleRow, FirstSlotColumn + SlotWidth - 1]];
                    proto.Copy(ws.Cells[band.HeaderRow, SlotColumn(slot)]);
                }
                for (int k = 0; k < SlotWidth; k++)
                {
                    try { ws.Columns[SlotColumn(slot) + k].ColumnWidth = ws.Columns[FirstSlotColumn + k].ColumnWidth; } catch { }
                }
            }
            try { dynamic xlApp = ExcelDna.Integration.ExcelDnaUtil.Application; xlApp.CutCopyMode = false; } catch { }

            if (previous > slotCount)
            {
                dynamic leftover = ws.Range[ws.Cells[FirstBoardRow, SlotColumn(slotCount)], ws.Cells[LastBoardRow, SlotColumn(previous) - 1]];
                try { leftover.ClearComments(); } catch { }
                leftover.Clear();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] preparing board columns failed: {ex.Message}");
        }
    }

    private static object? FormatPremixCell(DailyProjectSchedule data)
    {
        if (data.ProjectID.Length == 0) return null;
        var info = Roster.FindProject(data.ProjectID);
        bool allowsPremix = info != null && info.Types.Contains(ProjectType.Road_Reinstatement);
        if (allowsPremix && data.PremixHoles > 0) return data.PremixHoles;
        if (info != null && info.Types.Count > 0)
        {
            var icons = string.Join(" ", info.Types.Select(BoardContext.IconFor).Where(s => s.Length > 0));
            if (icons.Length > 0) return icons;
        }
        return null;
    }

    private static string TypeText(VehicleType type) =>
        $"{BoardContext.IconFor(type)} {type.ToString().Replace('_', ' ')}".Trim();

    internal static List<string> SkillsOf(WorkerPreference? pref)
    {
        if (pref == null) return new List<string>();
        var list = new List<string>();
        string? primary = BoardContext.EffectivePrimarySkill(pref);
        if (primary != null) list.Add(primary);
        list.AddRange(pref.LicensedSkills);
        list.AddRange(pref.OtherSkills);
        return list.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    /// <summary>
    /// Skills from the worker's preference row; when that row is missing or lists no skills, falls back to the
    /// default skill of their staff_master department so every field worker has at least one skill.
    /// </summary>
    internal static List<string> SkillsFor(StaffInfo? staff, WorkerPreference? pref)
    {
        var skills = SkillsOf(pref);
        if (skills.Count == 0 && staff != null)
            skills.Add(SkillCatalog.GeneralWorker);
        return skills;
    }

    // ---- header: "CODE ⚠" + requirements note (Start / End live in their own columns) -----------
    private static (string Text, string? Comment) BuildHeader(
        DailyProjectSchedule data, List<string> assignedSkills, List<VehicleType> assignedVehicleTypes)
    {
        if (data.ProjectID.Length == 0) return (EmptyHeaderText, null);

        var info = Roster.FindProject(data.ProjectID);
        var (reqSkills, reqVehicles) = BoardContext.RequirementsFor(info?.Types ?? Array.Empty<ProjectType>());
        var missingSkills = reqSkills.Where(r => !assignedSkills.Contains(r, StringComparer.OrdinalIgnoreCase)).ToList();
        var missingVehicles = reqVehicles.Where(r => !assignedVehicleTypes.Contains(r)).ToList();
        bool unmet = missingSkills.Count > 0 || missingVehicles.Count > 0;

        string? comment = null;
        if (reqSkills.Count > 0 || reqVehicles.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine($"Requirements ({string.Join(", ", info?.Types ?? Array.Empty<ProjectType>())})");
            sb.AppendLine("Skills:   " + (reqSkills.Count == 0 ? "-" : string.Join(", ", reqSkills.Select(r => $"{BoardContext.IconFor(r)} {r} {(missingSkills.Contains(r) ? "✘" : "✔")}"))));
            sb.Append("Vehicles: " + (reqVehicles.Count == 0 ? "-" : string.Join(", ", reqVehicles.Select(r => $"{BoardContext.IconFor(r)} {r.ToString().Replace('_', ' ')} {(missingVehicles.Contains(r) ? "✘" : "✔")}"))));
            comment = sb.ToString();
        }

        return ($"{data.ProjectID}{(unmet ? " ⚠" : string.Empty)}", comment);
    }

    private static string? BuildProjectTypeComment(DailyProjectSchedule data)
    {
        var info = data.ProjectID.Length > 0 ? Roster.FindProject(data.ProjectID) : null;
        var types = (info?.Types ?? Array.Empty<ProjectType>()).Where(t => t != ProjectType.None).ToList();
        bool showPremix = types.Contains(ProjectType.Road_Reinstatement) && data.PremixHoles > 0;
        if (types.Count == 0 && !showPremix) return null;

        var lines = new List<string>();
        if (types.Count > 0)
        {
            lines.Add(types.Count == 1 ? "Project Type:" : "Project Types:");
            foreach (var t in types)
            {
                string icon = BoardContext.IconFor(t);
                string name = t.ToString().Replace('_', ' ');
                lines.Add(icon.Length > 0 ? $"{icon} {name}" : name);
            }
        }
        if (showPremix)
            lines.Add($"Premix Holes: {data.PremixHoles} {(data.PremixHoles == 1 ? "Hole" : "Holes")}");
        return string.Join("\n", lines);
    }

    // =========================================================================
    // 6. CONDITIONAL FORMATS (duplicate worker = red, late check-in = amber)
    // =========================================================================
    private static void AddDuplicateHighlight(dynamic ws, int row, int col)
    {
        try
        {
            dynamic fc = ws.Cells[row, col].FormatConditions.Add(2, Type.Missing, "=TRUE");   // xlExpression
            fc.Interior.Color = 13551615;   // RGB(255,199,206)
            fc.Font.Color = 393372;         // RGB(156,0,6)
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] duplicate highlight rule failed: {ex.Message}"); }
    }

    private static void EnsureConditionalFormats(dynamic ws, int slotCount,
        HashSet<(int Row, int Col)> workerCells, HashSet<(int Row, int Col)> headerCells)
    {
        const int xlExpression = 2;
        var cutOff = BoardContext.LateCutOff;
        foreach (var band in Bands)
        {
            // Duplicate project headers: clear the header row rules, then flag only the conflicting header cells.
            try
            {
                dynamic headers = ws.Range[ws.Cells[band.HeaderRow, FirstSlotColumn], ws.Cells[band.HeaderRow, FirstSlotColumn + (slotCount * SlotWidth) - 1]];
                headers.FormatConditions.Delete();
            }
            catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] clearing header rules failed: {ex.Message}"); }
            for (int slot = 0; slot < slotCount; slot++)
                if (headerCells.Contains((band.HeaderRow, SlotColumn(slot)))) AddDuplicateHighlight(ws, band.HeaderRow, SlotColumn(slot));

            for (int slot = 0; slot < slotCount; slot++)
            {
                string nameCol = ColumnLetter(SlotColumn(slot));
                string inCol = ColumnLetter(SlotColumn(slot) + 2);
                string outCol = ColumnLetter(SlotColumn(slot) + 3);

                // Duplicate worker: clear stale rules, then flag only the conflicting name cells.
                try
                {
                    dynamic names = ws.Range[$"{nameCol}{band.FirstWorkerRow}:{nameCol}{band.LastWorkerRow}"];
                    names.FormatConditions.Delete();
                }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] clearing duplicate rules failed: {ex.Message}"); }
                for (int r = band.FirstWorkerRow; r <= band.LastWorkerRow; r++)
                    if (workerCells.Contains((r, SlotColumn(slot)))) AddDuplicateHighlight(ws, r, SlotColumn(slot));

                // Effective In / Out are real time values now: clear stale rules on both (also left over from the
                // 3-column layout), then flag a late Day Work check-in on the In column only. A late check-in always
                // keeps the worker's own punch, so this never overlaps the amber auto-fallback font.
                try { ws.Range[$"{inCol}{band.FirstWorkerRow}:{outCol}{band.LastWorkerRow}"].FormatConditions.Delete(); }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] clearing time rules failed: {ex.Message}"); }

                if (band.Shift != Shifts.Morning) continue;
                try
                {
                    dynamic times = ws.Range[$"{inCol}{band.FirstWorkerRow}:{inCol}{band.LastWorkerRow}"];
                    string first = $"{inCol}{band.FirstWorkerRow}";
                    dynamic late = times.FormatConditions.Add(xlExpression, Type.Missing,
                        $"=AND(ISNUMBER({first}),{first}>TIME({cutOff.Hour},{cutOff.Minute},0))");
                    late.Interior.Color = 10284031;  // RGB(255,235,156)
                    late.Font.Color = 22428;         // RGB(156,87,0)
                }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] late highlight rule failed: {ex.Message}"); }
            }
        }
    }

    internal static string ColumnLetter(int column)
    {
        var sb = new StringBuilder();
        for (int c = column; c > 0; c = (c - 1) / 26) sb.Insert(0, (char)('A' + ((c - 1) % 26)));
        return sb.ToString();
    }

    // =========================================================================
    // 7. MATCH-% RANKED DROPDOWNS
    // =========================================================================
    internal static void RegisterDropdowns()
    {
        CustomDropdowns.Register(AppliesTo, BuildDropdown);
    }

    private static bool AppliesTo(dynamic cell)
    {
        if (!IsBoardSheet((object)cell.Worksheet)) return false;
        if (IsDateCell(cell) || IsDayCell(cell)) return true;   // A2: quick date picker, B2: day-of-week list

        int row = (int)cell.Row, col = (int)cell.Column;
        // The board width is dynamic: re-detect when the cell lies right of what we last saw.
        if (col > LastBoardColumn && IsBoardArea(row, col)) _slotCount = Math.Max(_slotCount, DetectSlotCount((object)cell.Worksheet));

        var loc = Locate(row, col);
        return loc.Kind is CellKind.Header or CellKind.Worker or CellKind.WorkerSkill or CellKind.Vehicle or CellKind.VehicleType
            or CellKind.ProjectStart or CellKind.ProjectEnd;
    }

    private static CustomDropdownResult? BuildDropdown(dynamic cell)
    {
        var loc = Locate((int)cell.Row, (int)cell.Column);
        object ws = cell.Worksheet;
        DateOnly date = ReadDate(ws) ?? DateOnly.FromDateTime(DateTime.Today);

        if (IsDateCell(cell)) return BuildDateDropdown(date);
        if (IsDayCell(cell)) return BuildDayDropdown();

        switch (loc.Kind)
        {
            case CellKind.Header:
                return BuildHeaderDropdown(cell, date, loc);

            case CellKind.Vehicle:
                return BuildVehicleDropdown(ws, date, loc, (int)cell.Row, (int)cell.Column);

            case CellKind.Worker:
                return BuildWorkerDropdown(ws, date, loc, (int)cell.Row, (int)cell.Column);

            case CellKind.WorkerSkill:
                return BuildSkillDropdown(ws, date, loc);

            case CellKind.VehicleType:
                return BuildVehicleTypeDropdown(ws, date, loc);

            case CellKind.ProjectStart:
            case CellKind.ProjectEnd:
                return BuildProjectTimeDropdown(ws, date, loc, loc.Kind == CellKind.ProjectStart);

            default:
                return null;
        }
    }

    // ---- Project Start / End -------------------------------------------------------------------
    /// <summary>
    /// Project Start / End picks: "✕ Clear (Use Auto Fallback)" (only while a manual time is set), the punch times of the
    /// staff checked into the project (latest check-in / earliest exit marked), then the band's default and standard
    /// 30-minute times. Returns null for a slot without a project.
    /// </summary>
    private static CustomDropdownResult? BuildProjectTimeDropdown(object wsObj, DateOnly date, CellLocation loc, bool isStart)
    {
        var band = Bands.FirstOrDefault(b => b.Shift == loc.Shift);
        if (band == null) return null;
        string projectId = SlotProjectId(wsObj, band, loc.Slot);
        if (projectId.Length == 0) return null;

        var options = new List<DropdownOption>();
        var listed = new HashSet<TimeOnly>();
        int rank = 0;

        var manual = ManualTimesFor(loc.Shift, projectId);
        if ((isStart ? manual.Start : manual.End).HasValue)
            options.Add(new DropdownOption { Display = ClearTimeDisplay, Raw = ClearTimeToken, Section = rank++ });

        // Section 1: punch times of the staff on this project.
        try
        {
            var spans = AttendanceRollupService.BuildSpans(ReadEntries(wsObj), date, BoardContext.DayStartTime,
                BoardContext.NextDayStartTime, Read(wsObj, date), BoardContext.OTStartTime);
            string key = AttendanceRollupService.NormalizeProjectId(projectId);
            var punches = spans.Where(kv => kv.Key.ProjectId.Equals(key, StringComparison.OrdinalIgnoreCase))
                .SelectMany(kv => kv.Value.Select(sp => (Name: Roster.NameOf(kv.Key.StaffId, $"#{kv.Key.StaffId}"), Span: sp)))
                .ToList();

            var times = isStart
                ? punches.Select(p => (Time: p.Span.In, Offset: p.Span.InOffset, p.Name)).ToList()
                : punches.Where(p => p.Span.Out.HasValue && p.Span.OutOffset.HasValue)
                         .Select(p => (Time: p.Span.Out!.Value, Offset: p.Span.OutOffset!.Value, p.Name)).ToList();

            if (times.Count > 0)
            {
                // Start: latest check-in first (the auto fallback); End: earliest exit first.
                var groups = times.GroupBy(t => t.Time)
                    .Select(g => (Time: g.Key, Offset: g.Min(x => x.Offset), Names: g.Select(x => x.Name).Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(n => n, StringComparer.OrdinalIgnoreCase).ToList()))
                    .ToList();
                groups = isStart ? groups.OrderByDescending(g => g.Offset).ToList() : groups.OrderBy(g => g.Offset).ToList();
                for (int i = 0; i < groups.Count; i++)
                {
                    var g = groups[i];
                    string tag = i == 0 ? (isStart ? " (latest)" : " (earliest)") : string.Empty;
                    if (listed.Add(g.Time))
                        options.Add(new DropdownOption { Display = $"{Hm(g.Time)} — {string.Join(", ", g.Names)}{tag}", Raw = Hm(g.Time), Section = rank++ });
                }
            }
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] project time punch options failed: {ex.Message}"); }

        // Section 2: band default, then standard 30-minute times.
        bool night = loc.Shift == Shifts.Night;
        (TimeOnly def, int fromMin, int toMin) = (night, isStart) switch
        {
            (false, true) => (new TimeOnly(9, 0), 7 * 60, 12 * 60),
            (false, false) => (new TimeOnly(18, 0), 16 * 60, (23 * 60) + 30),
            (true, true) => (new TimeOnly(20, 0), 19 * 60, (23 * 60) + 30),
            (true, false) => (new TimeOnly(6, 0), 0, 8 * 60)
        };
        if (listed.Add(def))
            options.Add(new DropdownOption { Display = $"{Hm(def)} (Default)", Raw = Hm(def), Section = rank++ });
        for (int m = fromMin; m <= toMin; m += 30)
        {
            var t = new TimeOnly(m / 60, m % 60);
            if (listed.Add(t))
                options.Add(new DropdownOption { Display = Hm(t), Raw = Hm(t), Section = rank++ });
        }

        return new CustomDropdownResult(options, (c, opt) => ApplyProjectTimePick(c, loc, projectId, isStart, opt.Raw));
    }

    /// <summary>Stores the picked manual time (or clears it) and writes the time value so the change observer refreshes cleanly.</summary>
    private static void ApplyProjectTimePick(dynamic cell, CellLocation loc, string projectId, bool isStart, string raw)
    {
        if (raw == ClearTimeToken)
        {
            SetManualTime(loc.Shift, projectId, isStart, null);
            cell.Value2 = null;
            return;
        }
        if (!TryParseTime(raw, out var t)) return;
        SetManualTime(loc.Shift, projectId, isStart, t);
        cell.Value2 = t.ToTimeSpan().TotalDays;
        try { cell.NumberFormat = TimeFormat; } catch { }
    }

    /// <summary>True when (row, col) is a project-header Start / End cell (re-detects the board width when needed).</summary>
    internal static bool IsProjectTimeCell(object wsObj, int row, int col, out CellLocation loc)
    {
        if (col > LastBoardColumn && IsBoardArea(row, col)) _slotCount = Math.Max(_slotCount, DetectSlotCount(wsObj));
        loc = Locate(row, col);
        return loc.Kind is CellKind.ProjectStart or CellKind.ProjectEnd;
    }

    /// <summary>
    /// Commits a typed / cleared Project Start / End cell into <see cref="_manualTimes"/> (macro thread). Blank reverts
    /// that side to the auto fallback; an unparseable value is ignored. False when the cell is not a time cell or the
    /// slot has no project.
    /// </summary>
    internal static bool CommitManualTimeEdit(object wsObj, int row, int col)
    {
        try
        {
            if (!IsProjectTimeCell(wsObj, row, col, out var loc)) return false;
            var band = Bands.FirstOrDefault(b => b.Shift == loc.Shift);
            if (band == null) return false;
            string projectId = SlotProjectId(wsObj, band, loc.Slot);
            if (projectId.Length == 0) return false;

            bool isStart = loc.Kind == CellKind.ProjectStart;
            dynamic ws = wsObj;
            object? raw = ws.Cells[row, col].Value2;
            if (raw == null || string.IsNullOrWhiteSpace(raw.ToString()))
            {
                SetManualTime(loc.Shift, projectId, isStart, null);
                return true;
            }

            TimeOnly t;
            bool ok = raw is double d ? TryFromOADate(d, out t) : TryParseTime(raw.ToString()!, out t);
            if (ok) SetManualTime(loc.Shift, projectId, isStart, t);
            else AppLogger.Warn($"[ScheduleBoard] '{raw}' is not a valid time for {projectId} {(isStart ? "Start" : "End")}; ignored.");
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] committing project time failed: {ex.Message}");
            return false;
        }
    }

    private static string SlotProjectId(object wsObj, Band band, int slot)
    {
        try
        {
            dynamic ws = wsObj;
            return ParseProjectCode((ws.Cells[band.HeaderRow, SlotColumn(slot)].Value2 as object)?.ToString() ?? string.Empty);
        }
        catch { return string.Empty; }
    }

    private static List<AttendanceEntry> ReadEntries(object wsObj)
    {
        dynamic ws = wsObj;
        object wbObj = ws.Parent;
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetMapper.EntriesAliases, "Attendance Entries");
        return entriesSheet != null
            ? TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, SheetMapper.EntriesTableName)
            : new List<AttendanceEntry>();
    }

    private static string Hm(TimeOnly t) => t.ToString("HH:mm", CultureInfo.InvariantCulture);

    /// <summary>Time of day from an Excel serial (the fractional day; the date part is ignored).</summary>
    private static bool TryFromOADate(double d, out TimeOnly time)
    {
        time = default;
        double frac = d - Math.Floor(d);
        if (double.IsNaN(frac) || frac < 0 || frac >= 1) return false;
        long ticks = (long)Math.Round(frac * TimeSpan.TicksPerDay / TimeSpan.TicksPerMinute) * TimeSpan.TicksPerMinute;
        time = new TimeOnly(ticks % TimeSpan.TicksPerDay);
        return true;
    }

    /// <summary>Accepts "HH:mm", "H:mm", "HHmm" ("0900", "1800"), and anything <see cref="TimeOnly"/> / <see cref="DateTime"/> can parse.</summary>
    private static bool TryParseTime(string text, out TimeOnly time)
    {
        time = default;
        string s = text.Trim();
        if (s.Length == 0) return false;
        if (s.Length is 3 or 4 && s.All(char.IsDigit))
        {
            int v = int.Parse(s, CultureInfo.InvariantCulture);
            int h = v / 100, m = v % 100;
            if (h > 23 || m > 59) return false;
            time = new TimeOnly(h, m);
            return true;
        }
        if (TimeOnly.TryParse(s, CultureInfo.InvariantCulture, DateTimeStyles.None, out time)) return true;
        if (TimeOnly.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out time)) return true;
        if (DateTime.TryParse(s, CultureInfo.CurrentCulture, DateTimeStyles.None, out var dt)) { time = TimeOnly.FromDateTime(dt); return true; }
        return false;
    }

    private static string FormatProjectDisplay(ProjectInfo p) =>
        string.IsNullOrWhiteSpace(p.Name) ? p.Id : $"{p.Id} — {p.Name}";

    /// <summary>
    /// Project Header options, in order: "✕ Remove Slot" (only when the slot holds a project), auto-scheduled projects
    /// Active on the date, manual-only Active projects (Internal / Auto-Schedule = FALSE), then every other project
    /// greyed out. Section is a running rank so the order above is preserved.
    /// </summary>
    private static CustomDropdownResult BuildHeaderDropdown(dynamic cell, DateOnly date, CellLocation loc)
    {
        var options = new List<DropdownOption>();
        int rank = 0;

        string currentCode = string.Empty;
        try { currentCode = ParseProjectCode((cell.Value2 as object)?.ToString() ?? string.Empty); } catch { }
        if (currentCode.Length > 0)
            options.Add(new DropdownOption { Display = RemoveSlotDisplay, Raw = RemoveSlotToken, Section = rank++, IsValid = true });

        var listed = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void AddProjects(IEnumerable<ProjectInfo> projects, bool isValid)
        {
            foreach (var p in projects)
                if (listed.Add(p.Id))
                    options.Add(new DropdownOption { Display = FormatProjectDisplay(p), Raw = p.Id, Section = rank++, IsValid = isValid });
        }

        AddProjects(Roster.ActiveProjectsOn(date), isValid: true);
        AddProjects(Roster.ActiveManualProjectsOn(date), isValid: true);
        AddProjects(Roster.Projects, isValid: false);

        return new CustomDropdownResult(options, (c, opt) => ApplyHeaderPick(c, loc, opt.Raw));
    }

    /// <summary>
    /// Project Header pick. "✕ Remove Slot" removes the slot (see <see cref="RemoveSlot"/>). When the picked project already
    /// sits in another slot of the same band, the two whole slot columns (header, premix, window, workers, vehicles) are
    /// swapped with two bulk 2D writes instead of creating a duplicate. Otherwise the project is written and, when the slot
    /// has no workers yet, the workers whose Planned Projects still list it (removed earlier) are restored.
    /// Runs inside the dropdown commit (events suppressed; observers are notified afterwards).
    /// </summary>
    private static void ApplyHeaderPick(dynamic cell, CellLocation loc, string pickedProjectId)
    {
        object wsObj = cell.Worksheet;
        dynamic ws = wsObj;
        var band = Bands.FirstOrDefault(b => b.Shift == loc.Shift);

        if (pickedProjectId == RemoveSlotToken)
        {
            if (band != null) RemoveSlot(ws, band, loc);
            return;
        }

        pickedProjectId = pickedProjectId.Trim();
        _dismissedProjects.Remove(pickedProjectId);
        if (band == null) { cell.Value2 = pickedProjectId; return; }

        int otherSlot = -1;
        try
        {
            _slotCount = Math.Max(_slotCount, DetectSlotCount(wsObj));
            int cols = _slotCount * SlotWidth;
            dynamic headerRange = ws.Range[ws.Cells[band.HeaderRow, FirstSlotColumn], ws.Cells[band.HeaderRow, FirstSlotColumn + cols - 1]];
            object[,] headers = TableMapper.GetRange2DValues(headerRange, 1, cols);
            for (int s = 0; s < _slotCount && otherSlot < 0; s++)
            {
                if (s == loc.Slot) continue;
                string code = ParseProjectCode(headers[1, (s * SlotWidth) + 1]?.ToString() ?? string.Empty);
                if (code.Equals(pickedProjectId.Trim(), StringComparison.OrdinalIgnoreCase)) otherSlot = s;
            }
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] duplicate project lookup failed: {ex.Message}"); }

        if (otherSlot < 0)
        {
            cell.Value2 = pickedProjectId;
            RestorePlannedWorkers(ws, band, loc.Slot, pickedProjectId);
            return;
        }

        try
        {
            int rows = band.LastVehicleRow - band.HeaderRow + 1;
            dynamic mine = ws.Range[ws.Cells[band.HeaderRow, SlotColumn(loc.Slot)], ws.Cells[band.LastVehicleRow, SlotColumn(loc.Slot) + SlotWidth - 1]];
            dynamic theirs = ws.Range[ws.Cells[band.HeaderRow, SlotColumn(otherSlot)], ws.Cells[band.LastVehicleRow, SlotColumn(otherSlot) + SlotWidth - 1]];
            object[,] mineValues = TableMapper.GetRange2DValues(mine, rows, SlotWidth);
            object[,] theirValues = TableMapper.GetRange2DValues(theirs, rows, SlotWidth);
            mine.Value2 = theirValues;
            theirs.Value2 = mineValues;
            SheetOperations.ShowStatusConfirmation($"{pickedProjectId} was already in {band.Label} slot {otherSlot + 1}; slots {loc.Slot + 1} and {otherSlot + 1} were swapped.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] slot swap failed: {ex.Message}");
            cell.Value2 = pickedProjectId;
        }
    }

    /// <summary>
    /// "✕ Remove Slot": asks first when workers / vehicles are assigned, records the project as dismissed for the date
    /// (so a reload does not auto-populate it again), then shifts every slot right of it one slot left in this band with a
    /// single bulk 2D read / write and clears the vacated rightmost slot. Workers keep the project in Planned Projects
    /// (grey) so re-adding it restores them (<see cref="RestorePlannedWorkers"/>).
    /// </summary>
    private static void RemoveSlot(dynamic ws, Band band, CellLocation loc)
    {
        try
        {
            _slotCount = Math.Max(_slotCount, DetectSlotCount((object)ws));
            int slots = Math.Max(1, _slotCount - loc.Slot);
            int rows = band.LastVehicleRow - band.HeaderRow + 1;
            int cols = slots * SlotWidth;
            dynamic range = ws.Range[ws.Cells[band.HeaderRow, SlotColumn(loc.Slot)], ws.Cells[band.LastVehicleRow, SlotColumn(loc.Slot) + cols - 1]];
            object[,] values = TableMapper.GetRange2DValues(range, rows, cols);

            string Cell(int sheetRow, int c) => values[sheetRow - band.HeaderRow + 1, c]?.ToString()?.Trim() ?? string.Empty;
            bool IsEntry(string s) => s.Length > 0 && !PlaceholderTexts.Contains(s);

            string currentCode = ParseProjectCode(Cell(band.HeaderRow, 1));
            int workerCount = 0, vehicleCount = 0;
            for (int r = band.FirstWorkerRow; r <= band.LastWorkerRow; r++) if (IsEntry(ScheduleMatchEngine.StripPercent(Cell(r, 1)))) workerCount++;
            for (int r = band.FirstVehicleRow; r <= band.LastVehicleRow; r++) if (IsEntry(Cell(r, 1))) vehicleCount++;

            if ((workerCount > 0 || vehicleCount > 0) && !ConfirmRemoveSlot(currentCode, band, workerCount, vehicleCount)) return;

            if (currentCode.Length > 0) _dismissedProjects.Add(currentCode);

            var shifted = new object?[rows, cols];
            for (int r = 0; r < rows; r++)
                for (int c = 0; c < cols - SlotWidth; c++)
                    shifted[r, c] = values[r + 1, c + SlotWidth + 1];

            try { range.ClearComments(); } catch { }   // comments are re-derived by the refresh that follows the commit
            range.Value2 = shifted;

            SheetOperations.ShowStatusConfirmation(currentCode.Length > 0
                ? $"Removed {currentCode} from {band.Label} for this date."
                : $"Removed {band.Label} slot {loc.Slot + 1}.");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] remove slot failed: {ex.Message}");
        }
    }

    /// <summary>Yes / No confirmation with the modal-dialog exception: screen updating + events on, dropdown arrows hidden.</summary>
    private static bool ConfirmRemoveSlot(string currentCode, Band band, int workerCount, int vehicleCount)
    {
        dynamic xlApp = ExcelDna.Integration.ExcelDnaUtil.Application;
        bool? prevScreen = null, prevEvents = null;
        try { prevScreen = (bool)xlApp.ScreenUpdating; } catch { }
        try { prevEvents = (bool)xlApp.EnableEvents; } catch { }
        try
        {
            try { xlApp.ScreenUpdating = true; } catch { }
            try { xlApp.EnableEvents = true; } catch { }
            try { Arrow.HideAll(); } catch { }

            string label = currentCode.Length > 0 ? currentCode : $"slot {band.Label}";
            var answer = System.Windows.Forms.MessageBox.Show(
                $"Remove '{label}' from {band.Label}?\n({workerCount} worker(s), {vehicleCount} vehicle(s) assigned)\n\n" +
                $"Workers will remain greyed out in Daily Attendances (Planned Projects) so re-adding '{label}' restores them.",
                "Remove Slot",
                System.Windows.Forms.MessageBoxButtons.YesNo,
                System.Windows.Forms.MessageBoxIcon.Question,
                System.Windows.Forms.MessageBoxDefaultButton.Button2);
            return answer == System.Windows.Forms.DialogResult.Yes;
        }
        finally
        {
            if (prevEvents.HasValue) try { xlApp.EnableEvents = prevEvents.Value; } catch { }
            if (prevScreen.HasValue) try { xlApp.ScreenUpdating = prevScreen.Value; } catch { }
        }
    }

    /// <summary>
    /// After a project is (re-)added to an empty slot, writes the workers whose 'Daily Attendances' Planned Projects still
    /// list it (e.g. its slot was removed earlier) into the slot's Staff Name column with one bulk 2D write.
    /// No-op when the slot already has workers or no worker is planned on the project.
    /// </summary>
    private static void RestorePlannedWorkers(dynamic ws, Band band, int slot, string projectId)
    {
        try
        {
            string code = AttendanceRollupService.NormalizeProjectId(projectId);
            if (code.Length == 0) return;

            int col = SlotColumn(slot);
            dynamic workerRange = ws.Range[ws.Cells[band.FirstWorkerRow, col], ws.Cells[band.LastWorkerRow, col]];
            object[,] current = TableMapper.GetRange2DValues(workerRange, WorkerRowsPerBand, 1);
            for (int r = 1; r <= WorkerRowsPerBand; r++)
            {
                string w = ScheduleMatchEngine.StripPercent(current[r, 1]?.ToString()).Trim();
                if (w.Length > 0 && !PlaceholderTexts.Contains(w)) return;
            }

            object wbObj = ws.Parent;
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetMapper.MasterSheetAliases, "Daily Attendances");
            if (attendanceSheet == null) return;
            dynamic? table = TableMapper.ResolveTable(attendanceSheet, SheetMapper.MasterTableName);
            dynamic? body = table?.DataBodyRange;
            if (table == null || body == null) return;

            int rowCount = (int)body!.Rows.Count;
            int colCount = (int)table!.ListColumns.Count;
            object[,] header = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, colCount);
            int idCol = -1, nameCol = -1, plannedCol = -1;
            for (int c = 1; c <= colCount; c++)
            {
                string h = header[1, c]?.ToString()?.Trim() ?? string.Empty;
                if (h.Equals("Staff ID", StringComparison.OrdinalIgnoreCase)) idCol = c;
                else if (h.Equals("Staff Name", StringComparison.OrdinalIgnoreCase)) nameCol = c;
                else if (h.Equals("Planned Projects", StringComparison.OrdinalIgnoreCase)) plannedCol = c;
            }
            if (plannedCol < 0 || (idCol < 0 && nameCol < 0)) return;

            object[,] data = TableMapper.GetRange2DValues(body, rowCount, colCount);
            var names = new List<string>();
            for (int r = 1; r <= rowCount && names.Count < WorkerRowsPerBand; r++)
            {
                if (!AttendanceRollupService.SplitProjects(data[r, plannedCol]?.ToString()).Contains(code, StringComparer.OrdinalIgnoreCase)) continue;

                StaffInfo? staff = null;
                if (idCol > 0)
                {
                    object? raw = data[r, idCol];
                    long id = raw is double d ? (long)d : long.TryParse(raw?.ToString()?.Trim(), out long parsed) ? parsed : 0;
                    if (id > 0) staff = Roster.FindStaff(id);
                }
                if (staff == null && nameCol > 0) staff = Roster.FindStaff(data[r, nameCol]?.ToString());
                if (staff == null || staff.ShortName.Length == 0) continue;   // field workers only
                if (!names.Contains(staff.ShortName, StringComparer.OrdinalIgnoreCase)) names.Add(staff.ShortName);
            }
            if (names.Count == 0) return;

            var grid = new object?[WorkerRowsPerBand, 1];
            for (int i = 0; i < names.Count; i++) grid[i, 0] = names[i];
            workerRange.Value2 = grid;
            SheetOperations.ShowStatusConfirmation($"Restored {names.Count} planned worker(s) to {code} ({band.Label}).");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] restoring planned workers failed: {ex.Message}");
        }
    }

    /// <summary>A2 quick picks: the shown date +/- one week and today, oldest first. The pick fires the normal date-change flow.</summary>
    private static CustomDropdownResult BuildDateDropdown(DateOnly shown)
    {
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);
        var options = Enumerable.Range(-7, 15).Select(o => shown.AddDays(o)).Append(today).Distinct().OrderBy(d => d)
            .Select((d, i) => new DropdownOption
            {
                Display = d.ToString("ddd dd MMM yyyy", CultureInfo.InvariantCulture) + (d == today ? "  (Today)" : d == shown ? "  (Shown)" : string.Empty),
                Raw = d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
                Section = i
            }).ToList();

        return new CustomDropdownResult(options, (c, opt) =>
            c.Value2 = DateOnly.ParseExact(opt.Raw, "yyyy-MM-dd", CultureInfo.InvariantCulture).ToDateTime(TimeOnly.MinValue).ToOADate());
    }

    /// <summary>B2 picks: Monday..Sunday (the <see cref="Days"/> enum). The pick fires the normal day-change flow.</summary>
    private static CustomDropdownResult BuildDayDropdown()
    {
        var options = Enum.GetValues<Days>().Where(d => d != Days.None)
            .Select((d, i) => new DropdownOption { Display = d.ToString(), Raw = d.ToString(), Section = i }).ToList();
        return new CustomDropdownResult(options, (c, opt) => c.Value2 = opt.Raw);
    }

    /// <summary>
    /// Vehicle Plate picks. Pending type filter: Vehicle Type (col + 1) holds a type while Vehicle Plate is still empty.
    /// Vehicles of another type are greyed out and sorted last, but remain pickable.
    /// </summary>
    private static CustomDropdownResult BuildVehicleDropdown(object wsObj, DateOnly date, CellLocation loc, int row, int col)
    {
        var slot = ReadSlotContext(wsObj, date, loc);

        string? typeFilter = null;
        try
        {
            dynamic ws = wsObj;
            string plateText = ((ws.Cells[row, col].Value2 as object)?.ToString() ?? string.Empty).Trim();
            string typeText = ((ws.Cells[row, col + 1].Value2 as object)?.ToString() ?? string.Empty).Trim();
            if (plateText.Length == 0 && typeText.Length > 0) typeFilter = typeText;
        }
        catch { /* filter is advisory */ }

        var ranked = Roster.Vehicles
            .Select(v =>
            {
                var eval = Roster.EvaluateVehicleOn(v, date);
                string label = v.Plate;
                if (v.Type != VehicleType.None)
                {
                    string typeName = v.Type.ToString().Replace('_', ' ');
                    string icon = BoardContext.IconFor(v.Type);
                    label = icon.Length > 0 ? $"{icon} {v.Plate} — {typeName}" : $"{v.Plate} — {typeName}";
                }
                if (!eval.IsValid) label = $"{label} ({eval.ShortNote})";
                bool isMismatch = typeFilter != null && !MatchesTypeFilter(v.Type, typeFilter);
                return (Vehicle: v, Display: label, IsMismatch: isMismatch, IsInvalid: !eval.IsValid || isMismatch);
            })
            .OrderBy(x => x.IsInvalid ? 1 : 0)
            .ThenBy(x => x.IsMismatch ? 1 : 0)
            .ThenBy(x => slot.RequiredVehicleTypes.Contains(x.Vehicle.Type) ? 0 : 1)
            .ThenBy(x => x.Vehicle.Plate, NaturalStringComparer.Instance)
            .Select((x, i) => new DropdownOption { Display = x.Display, Raw = x.Vehicle.Plate, Section = i, IsValid = !x.IsInvalid })
            .ToList();
        return new CustomDropdownResult(ranked, (c, opt) => c.Value2 = opt.Raw);
    }

    private const string TeamPrefix = "T:";
    private const string StaffPrefix = "S:";

    /// <summary>Icon of the worker's first available skill (falls back to General Worker); used in the Staff Name dropdown.</summary>
    private static string PrimarySkillIconFor(StaffInfo staff, WorkerPreference? pref)
    {
        var skills = SkillsFor(staff, pref);
        string skill = skills.Count > 0 ? skills[0] : SkillCatalog.GeneralWorker;
        string icon = BoardContext.IconFor(skill);
        return icon.Length > 0 ? icon : SkillCatalog.GeneralWorkerIcon;
    }

    private static CustomDropdownResult BuildWorkerDropdown(object wsObj, DateOnly date, CellLocation loc, int row, int col)
    {
        var slot = ReadSlotContext(wsObj, date, loc);
        var scored = new List<(string Display, string Raw, int Percent, bool Blocked, string SortName)>();

        // Pending skill filter: Staff Skills (col + 1) holds a skill while Staff Name is still empty. Workers / teams
        // without that skill are greyed out and sorted last, but remain pickable.
        string? skillFilter = null;
        try
        {
            dynamic ws = wsObj;
            string nameText = ((ws.Cells[row, col].Value2 as object)?.ToString() ?? string.Empty).Trim();
            string skillText = ((ws.Cells[row, col + 1].Value2 as object)?.ToString() ?? string.Empty).Trim();
            if (nameText.Length == 0 && skillText.Length > 0) skillFilter = skillText;
        }
        catch { /* filter is advisory */ }

        var profiles = new Dictionary<long, WorkerProfile>();
        foreach (var s in Roster.ActiveOn(date).Where(s => s.ShortName.Length > 0))
        {
            profiles[s.Id] = new WorkerProfile(s, BoardContext.PreferenceFor(s.Id), BoardContext.DaysWorkedEarlierThisWeek(s.Id, date), false);
            var r = ScheduleMatchEngine.Score(profiles[s.Id], slot);
            bool blocked = r.Blocked || (skillFilter != null && !MatchesSkillFilter(s, skillFilter));
            var inactive = BoardContext.DescribeInactivePrimarySkillLicense(profiles[s.Id].Preference, date);
            string icon = PrimarySkillIconFor(s, profiles[s.Id].Preference);
            string label = icon.Length > 0 ? $"{icon} {s.ShortName}" : s.ShortName;
            string display = ScheduleMatchEngine.WithPercent(label, r.Percent) + (inactive != null ? $" ({inactive.Value.ShortNote})" : string.Empty);
            scored.Add((display, StaffPrefix + s.ShortName, r.Percent, blocked, s.ShortName));
        }

        foreach (var team in BoardContext.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)))
        {
            var members = BoardContext.Preferences
                .Where(p => p.Team.Contains(team.TeamName.Trim(), StringComparer.OrdinalIgnoreCase))
                .Select(p => profiles.TryGetValue(p.StaffID, out var wp) ? wp : null)
                .Where(wp => wp != null).Select(wp => wp!).ToList();
            var r = ScheduleMatchEngine.ScoreTeam(members, slot);
            bool blocked = r.Blocked || (skillFilter != null && !members.Any(m => MatchesSkillFilter(m.Staff, skillFilter)));
            scored.Add((ScheduleMatchEngine.WithPercent(team.TeamName.Trim(), r.Percent), TeamPrefix + team.TeamName.Trim(), r.Percent, blocked, team.TeamName.Trim()));
        }

        var ordered = scored
            .OrderBy(x => x.Blocked ? 1 : 0)
            .ThenByDescending(x => x.Percent)
            .ThenBy(x => x.SortName, StringComparer.OrdinalIgnoreCase)
            .Select((x, i) => new DropdownOption { Display = x.Display, Raw = x.Raw, Section = i, IsValid = !x.Blocked })
            .ToList();

        // Commit strips the "[XX%] " prefix; a team name is expanded by the sheet-change observer.
        return new CustomDropdownResult(ordered, (c, opt) =>
            c.Value2 = ScheduleMatchEngine.StripPercent(opt.Raw.StartsWith(TeamPrefix, StringComparison.Ordinal) || opt.Raw.StartsWith(StaffPrefix, StringComparison.Ordinal)
                ? opt.Raw[2..]
                : opt.Raw));
    }

    // ---- Staff Skills <-> Staff Name -------------------------------------------------------------
    private static bool HasSkill(StaffInfo staff, string skill) =>
        SkillsFor(staff, BoardContext.PreferenceFor(staff.Id)).Contains(skill, StringComparer.OrdinalIgnoreCase);

    /// <summary>True when the worker holds a skill whose name or unicode icon equals the pending Staff Skills filter text.</summary>
    private static bool MatchesSkillFilter(StaffInfo staff, string filter) =>
        SkillsFor(staff, BoardContext.PreferenceFor(staff.Id)).Any(skill =>
            skill.Equals(filter, StringComparison.OrdinalIgnoreCase) ||
            BoardContext.IconFor(skill).Equals(filter, StringComparison.OrdinalIgnoreCase));

    private static CustomDropdownResult BuildSkillDropdown(object wsObj, DateOnly date, CellLocation loc)
    {
        var slot = ReadSlotContext(wsObj, date, loc);
        var configured = BoardContext.AllConfiguredSkills();
        var allowed = new HashSet<string>(configured, StringComparer.OrdinalIgnoreCase);
        var required = slot.RequiredSkills.Where(s => !string.IsNullOrWhiteSpace(s) && allowed.Contains(s.Trim())).ToList();
        var others = configured
            .Where(s => !required.Contains(s, StringComparer.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(s => s, StringComparer.OrdinalIgnoreCase);

        var options = required.Concat(others).Distinct(StringComparer.OrdinalIgnoreCase)
            .Select((s, i) => new DropdownOption { Display = $"{BoardContext.IconFor(s)} {s}".Trim(), Raw = s, Section = i })
            .ToList();

        return new CustomDropdownResult(options, (c, opt) => ApplySkillPick(c, date, loc, opt.Raw));
    }

    /// <summary>
    /// A skill pick never auto-picks a worker. If the Staff Name cell (col - 1) already holds a worker with the skill it
    /// is kept; otherwise Staff Name is cleared and the skill stays as a pending filter for the Staff Name dropdown
    /// (<see cref="BuildWorkerDropdown"/>). The caller's change notification queues the refresh that derives icons.
    /// </summary>
    private static void ApplySkillPick(dynamic cell, DateOnly date, CellLocation loc, string skill)
    {
        _suppressNextHelperClear = true;   // this commit must never be read as the user clearing the row
        object ws = cell.Worksheet;
        int row = (int)cell.Row, col = (int)cell.Column;
        dynamic nameCell = ((dynamic)ws).Cells[row, col - 1];

        string current = ScheduleMatchEngine.StripPercent((nameCell.Value2 as object)?.ToString());
        var existing = Roster.FindStaff(current);
        if (existing == null || !HasSkill(existing, skill))
            nameCell.ClearContents();

        // The clicked cell must stay non-empty or the change observer reads it as "cleared".
        string icon = BoardContext.IconFor(skill);
        cell.Value2 = icon.Length > 0 ? icon : skill;
    }

    // ---- Vehicle Type <-> Vehicle Plate ----------------------------------------------------------
    /// <summary>True when the pending Vehicle Type filter text is the type's unicode icon, its name, or legacy "Icon Name" text.</summary>
    private static bool MatchesTypeFilter(VehicleType type, string filter) =>
        BoardContext.IconFor(type).Equals(filter, StringComparison.OrdinalIgnoreCase) ||
        type.ToString().Replace('_', ' ').Equals(filter.Replace('_', ' '), StringComparison.OrdinalIgnoreCase) ||
        TypeText(type).Equals(filter, StringComparison.OrdinalIgnoreCase);

    private static CustomDropdownResult BuildVehicleTypeDropdown(object wsObj, DateOnly date, CellLocation loc)
    {
        var slot = ReadSlotContext(wsObj, date, loc);
        var options = Enum.GetValues<VehicleType>().Where(t => t != VehicleType.None)
            .OrderBy(t => slot.RequiredVehicleTypes.Contains(t) ? 0 : 1)
            .ThenBy(t => t.ToString(), StringComparer.OrdinalIgnoreCase)
            .Select((t, i) => new DropdownOption { Display = TypeText(t), Raw = t.ToString(), Section = i })
            .ToList();

        return new CustomDropdownResult(options, (c, opt) => ApplyTypePick(c, date, loc, opt.Raw));
    }

    /// <summary>
    /// Mirrors <see cref="ApplySkillPick"/>: a type pick never auto-picks a vehicle. If the Vehicle Plate cell (col - 1)
    /// already holds a vehicle of that type it is kept; otherwise Vehicle Plate is cleared and the type stays as a
    /// pending filter for the Vehicle Plate dropdown (<see cref="BuildVehicleDropdown"/>).
    /// </summary>
    private static void ApplyTypePick(dynamic cell, DateOnly date, CellLocation loc, string typeName)
    {
        _suppressNextHelperClear = true;   // this commit must never be read as the user clearing the row
        if (!Enum.TryParse<VehicleType>(typeName, true, out var type) || type == VehicleType.None) return;

        object ws = cell.Worksheet;
        int row = (int)cell.Row, col = (int)cell.Column;
        dynamic plateCell = ((dynamic)ws).Cells[row, col - 1];

        string current = ((plateCell.Value2 as object)?.ToString() ?? string.Empty).Trim();
        var existing = Roster.FindVehicle(current);
        if (existing == null || existing.Type != type)
            plateCell.ClearContents();

        // The clicked cell must stay non-empty or the change observer reads it as "cleared".
        string icon = BoardContext.IconFor(type);
        cell.Value2 = icon.Length > 0 ? icon : TypeText(type);
    }

    // ---- clearing a helper cell clears its lead cell ---------------------------------------------
    /// <summary>Set by <see cref="ApplySkillPick"/> / <see cref="ApplyTypePick"/>; the next helper-cell change is a dropdown commit, not a clear.</summary>
    private static bool _suppressNextHelperClear;

    /// <summary>True when (row, col) is a project-header Premix cell on the board.</summary>
    internal static bool IsPremixCell(object wsObj, int row, int col)
    {
        if (col > LastBoardColumn && IsBoardArea(row, col)) _slotCount = Math.Max(_slotCount, DetectSlotCount(wsObj));
        return Locate(row, col).Kind == CellKind.Premix;
    }

    /// <summary>Notice text when a positive premix value was typed on a project that is not Road_Reinstatement; otherwise null.</summary>
    internal static string? GetRejectedPremixNotice(object wsObj, int row, int col)
    {
        try
        {
            dynamic ws = wsObj;
            string premixText = (ws.Cells[row, col].Value2 as object)?.ToString() ?? string.Empty;
            if (ParsePremixHoles(premixText) <= 0) return null;
            string headerText = (ws.Cells[row, col - 1].Value2 as object)?.ToString() ?? string.Empty;
            string projectId = ParseProjectCode(headerText);
            if (Roster.SupportsPremixHoles(projectId)) return null;
            return projectId.Length > 0
                ? $"Premix Holes is only allowed for Road Reinstatement projects ('{projectId}' is not Road Reinstatement)."
                : "Premix Holes is only allowed for Road Reinstatement projects (select a Road Reinstatement project first).";
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleBoard] premix validation failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>True when a single Staff Skills / Vehicle Type cell was just emptied (the user is clearing the row).</summary>
    internal static bool IsClearedHelperCell(dynamic target, out int row, out int col)
    {
        row = 0; col = 0;
        try
        {
            if ((int)target.Cells.Count != 1) return false;
            row = (int)target.Row; col = (int)target.Column;
            if (Locate(row, col).Kind is not (CellKind.WorkerSkill or CellKind.VehicleType)) return false;
            if (_suppressNextHelperClear)
            {
                _suppressNextHelperClear = false;
                return false;
            }
            return string.IsNullOrWhiteSpace((target.Value2 as object)?.ToString());
        }
        catch { return false; }
    }

    /// <summary>Empties the Staff Name / Vehicle Plate left of a cleared helper cell. Caller owns the Excel freeze.</summary>
    internal static void ClearLeadCell(object wsObj, int row, int helperColumn)
    {
        dynamic ws = wsObj;
        ws.Cells[row, helperColumn - 1].ClearContents();
    }

    private static SlotContext ReadSlotContext(object wsObj, DateOnly date, CellLocation loc) =>
        ReadSlotContext(Read(wsObj, date), date, loc);

    private static SlotContext ReadSlotContext(List<DailyProjectSchedule> slots, DateOnly date, CellLocation loc)
    {
        var mine = slots.FirstOrDefault(s => s.Shift == loc.Shift && s.SlotIndex == loc.Slot) ?? new DailyProjectSchedule { Shift = loc.Shift, SlotIndex = loc.Slot };

        var info = Roster.FindProject(mine.ProjectID);
        var (skills, vehicleTypes) = BoardContext.RequirementsFor(info?.Types ?? Array.Empty<ProjectType>());

        var sameShiftCounts = new Dictionary<long, int>();
        var otherShiftCounts = new Dictionary<long, int>();
        foreach (var other in slots)
        {
            if (other.Shift == loc.Shift && other.SlotIndex == loc.Slot) continue;
            var target = other.Shift == loc.Shift ? sameShiftCounts : otherShiftCounts;
            var seenInSlot = new HashSet<long>();
            foreach (var w in other.Workers)
                if (Roster.FindStaff(w) is { } st && seenInSlot.Add(st.Id))
                    target[st.Id] = target.GetValueOrDefault(st.Id) + 1;
        }

        var plates = mine.Vehicles;
        var types = plates.Select(p => Roster.FindVehicle(p)?.Type ?? VehicleType.None).Where(t => t != VehicleType.None).ToList();
        return new SlotContext(date, loc.Shift, skills, vehicleTypes, plates, types, sameShiftCounts, otherShiftCounts);
    }
}
