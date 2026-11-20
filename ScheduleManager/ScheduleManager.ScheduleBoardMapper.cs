namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

/// <summary>
/// Reads and renders the 'Daily Schedule' dispatch board: a dynamic number of three-column slots (B:D, E:G, H:J, ...),
/// one per active project of the shown date, for Day Work (rows 4-18) and Night Work (rows 20-34).
///   row 4 / 20 : Project Code (+ requirement flag) | Premix Hole Amount | Project start time - end time
///   workers    : Staff Name | Staff Skills (icons) | Check in and out Time        (10 rows: 5-14 / 21-30)
///   vehicles   : Vehicle Plate | Vehicle Type (icon) | Current Odometer             (4 rows: 15-18 / 31-34)
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

    internal const int SlotWidth = 3;
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

    /// <summary>Premix header cell keeps a numeric value when &gt; 0 ("1 Hole", "5 Holes"); when 0 it displays the project type instead.</summary>
    private const string PremixHoleFormat = "[=1]0 \"Hole\";0 \"Holes\"";
    private const string EmptyOdometerDash = "—";

    internal enum CellKind { None, Header, Premix, Worker, WorkerSkill, Vehicle, VehicleType }
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
                var record = new DailyProjectSchedule
                {
                    DailyScheduleDate = date,
                    Shift = band.Shift,
                    SlotIndex = slot,
                    ProjectID = ParseProjectCode(Text(grid, band.HeaderRow, col)),
                    PremixHoles = ParsePremixHoles(Text(grid, band.HeaderRow, col + 1))
                };

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

                if (record.ProjectID.Length > 0 || record.Workers.Count > 0 || record.Vehicles.Count > 0 || record.PremixHoles > 0)
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
        var spans = AttendanceRollupService.BuildSpans(entries, date, BoardContext.DayStartTime, BoardContext.NextDayStartTime);
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
    /// Active on the date gets a slot (iJob -> MWO -> PVT) merged with its saved row; saved rows for other projects follow.
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
        var spans = AttendanceRollupService.BuildSpans(entries, date, BoardContext.DayStartTime, BoardContext.NextDayStartTime);
        string cutOffText = BoardContext.LateCutOff.ToString("HH:mm", CultureInfo.InvariantCulture);

        var layouts = Bands.ToDictionary(b => b.Shift, b => LayoutSlots(b.Shift, date, slots, autoPopulate));
        int slotCount = Math.Max(1, layouts.Values.Max(l => l.Count));

        // Laid-out board (position = slot): feeds per-worker match breakdowns and the conflict checks below.
        var board = new List<DailyProjectSchedule>();
        foreach (var band in Bands)
        {
            var layout = layouts[band.Shift];
            for (int i = 0; i < layout.Count; i++)
                board.Add(new DailyProjectSchedule { DailyScheduleDate = date, Shift = band.Shift, SlotIndex = i, ProjectID = layout[i].ProjectID, Workers = layout[i].Workers, Vehicles = layout[i].Vehicles });
        }

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
            var comments = new string?[slotCount];
            var workerComments = new string?[WorkerRowsPerBand, slotCount];

            for (int slot = 0; slot < slotCount; slot++)
            {
                var data = slot < layout.Count ? layout[slot] : new DailyProjectSchedule { DailyScheduleDate = date, Shift = band.Shift };
                int c0 = slot * SlotWidth;
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
                    workers[i, c0 + 2] = staff != null && data.ProjectID.Length > 0 && spans.TryGetValue((staff.Id, data.ProjectID), out var list) && list.Count > 0
                        ? FormatSpan(list[0].In, list.Any(s => s.Out == null) ? null : list.Max(s => s.Out))
                        : FormatSpan(null, null);
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
                    }
                    else
                    {
                        vehicles[i, c0 + 2] = EmptyOdometerDash;
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

                var h = BuildHeader(data, spans, assignedSkills, vehicleTypes);
                header[0, c0] = h.Text;
                header[0, c0 + 1] = FormatPremixCell(data);
                header[0, c0 + 2] = h.Window;
                comments[slot] = h.Comment;
            }

            WriteBlock(ws, band.HeaderRow, band.HeaderRow, header);
            WriteBlock(ws, band.FirstWorkerRow, band.LastWorkerRow, workers);
            WriteBlock(ws, band.FirstVehicleRow, band.LastVehicleRow, vehicles);

            try { ws.Range[ws.Cells[band.FirstWorkerRow, FirstSlotColumn], ws.Cells[band.LastWorkerRow, FirstSlotColumn + cols - 1]].ClearComments(); } catch { }
            for (int slot = 0; slot < slotCount; slot++)
                for (int i = 0; i < WorkerRowsPerBand; i++)
                {
                    if (string.IsNullOrEmpty(workerComments[i, slot])) continue;
                    try { ApplySizedComment(ws.Cells[band.FirstWorkerRow + i, SlotColumn(slot)], workerComments[i, slot]!); } catch { }
                }

            for (int slot = 0; slot < slotCount; slot++)
            {
                try { ws.Cells[band.HeaderRow, SlotColumn(slot) + 1].NumberFormat = PremixHoleFormat; } catch { }
                try
                {
                    dynamic cell = ws.Cells[band.HeaderRow, SlotColumn(slot)];
                    cell.ClearComments();
                    if (!string.IsNullOrEmpty(comments[slot])) ApplySizedComment(cell, comments[slot]!);
                }
                catch { /* comments are advisory */ }
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

        EnsureConditionalFormats(ws, cutOffText, slotCount, workerCells, headerCells);
    }

    // Excel does not reflow AutoSize for hidden comments while ScreenUpdating = false, so the box
    // is sized deterministically from text metrics (points), padded for >100% DPI font scaling.
    private static void ApplySizedComment(dynamic cell, string text)
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
        if (data.PremixHoles > 0) return data.PremixHoles;
        if (data.ProjectID.Length == 0) return null;
        var info = Roster.FindProject(data.ProjectID);
        if (info != null && info.Types.Count > 0)
        {
            var icons = string.Join(" ", info.Types.Select(BoardContext.IconFor).Where(s => s.Length > 0));
            if (icons.Length > 0) return icons;
        }
        return 0;
    }

    private static string FormatSpan(TimeOnly? start, TimeOnly? end)
    {
        string s = start.HasValue ? start.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "No Start";
        string e = end.HasValue ? end.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "No End";
        return $"{s} - {e}";
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

    // ---- header: "CODE ⚠" + derived "HH:mm - HH:mm" window + requirements note --------------------
    private static (string Text, string? Window, string? Comment) BuildHeader(
        DailyProjectSchedule data,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        List<string> assignedSkills, List<VehicleType> assignedVehicleTypes)
    {
        if (data.ProjectID.Length == 0) return (EmptyHeaderText, null, null);

        var window = AttendanceRollupService.WindowFor(data.ProjectID, spans);
        string windowText = FormatSpan(window.Start, window.End);

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

        return ($"{data.ProjectID}{(unmet ? " ⚠" : string.Empty)}", windowText, comment);
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

    private static void EnsureConditionalFormats(dynamic ws, string cutOffText, int slotCount,
        HashSet<(int Row, int Col)> workerCells, HashSet<(int Row, int Col)> headerCells)
    {
        const int xlExpression = 2;
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
                string timeCol = ColumnLetter(SlotColumn(slot) + 2);

                // Duplicate worker: clear stale rules, then flag only the conflicting name cells.
                try
                {
                    dynamic names = ws.Range[$"{nameCol}{band.FirstWorkerRow}:{nameCol}{band.LastWorkerRow}"];
                    names.FormatConditions.Delete();
                }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] clearing duplicate rules failed: {ex.Message}"); }
                for (int r = band.FirstWorkerRow; r <= band.LastWorkerRow; r++)
                    if (workerCells.Contains((r, SlotColumn(slot)))) AddDuplicateHighlight(ws, r, SlotColumn(slot));

                if (band.Shift != Shifts.Morning) continue;
                try
                {
                    dynamic times = ws.Range[$"{timeCol}{band.FirstWorkerRow}:{timeCol}{band.LastWorkerRow}"];
                    times.FormatConditions.Delete();
                    string first = $"{timeCol}{band.FirstWorkerRow}";
                    dynamic late = times.FormatConditions.Add(xlExpression, Type.Missing, $"=AND(LEN({first})>=5,ISNUMBER(--LEFT({first},2)),LEFT({first},5)>\"{cutOffText}\")");
                    late.Interior.Color = 10284031;  // RGB(255,235,156)
                    late.Font.Color = 22428;         // RGB(156,87,0)
                }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] late highlight rule failed: {ex.Message}"); }
            }
        }
    }

    private static string ColumnLetter(int column)
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
        return loc.Kind is CellKind.Header or CellKind.Worker or CellKind.WorkerSkill or CellKind.Vehicle or CellKind.VehicleType;
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
            {
                var active = Roster.ActiveProjectsOn(date).Select(p => p.Id).ToList();
                var ordered = active.Concat(Roster.Projects.Select(p => p.Id).Where(id => !active.Contains(id, StringComparer.OrdinalIgnoreCase)))
                    .Select((id, i) => new DropdownOption { Display = id, Raw = id, Section = i }).ToList();
                return new CustomDropdownResult(ordered, (c, opt) => ApplyHeaderPick(c, loc, opt.Raw));
            }

            case CellKind.Vehicle:
                return BuildVehicleDropdown(ws, date, loc, (int)cell.Row, (int)cell.Column);

            case CellKind.Worker:
                return BuildWorkerDropdown(ws, date, loc, (int)cell.Row, (int)cell.Column);

            case CellKind.WorkerSkill:
                return BuildSkillDropdown(ws, date, loc);

            case CellKind.VehicleType:
                return BuildVehicleTypeDropdown(ws, date, loc);

            default:
                return null;
        }
    }

    /// <summary>
    /// Project Header pick. When the project already sits in another slot of the same band, the two whole slot columns
    /// (header, premix, window, workers, vehicles) are swapped with two bulk 2D writes instead of creating a duplicate.
    /// Runs inside the dropdown commit (events suppressed; observers are notified afterwards).
    /// </summary>
    private static void ApplyHeaderPick(dynamic cell, CellLocation loc, string pickedProjectId)
    {
        object wsObj = cell.Worksheet;
        dynamic ws = wsObj;
        var band = Bands.FirstOrDefault(b => b.Shift == loc.Shift);
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

        if (otherSlot < 0) { cell.Value2 = pickedProjectId; return; }

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
            .Select(v => (Vehicle: v, IsMismatch: typeFilter != null && !MatchesTypeFilter(v.Type, typeFilter)))
            .OrderBy(x => x.IsMismatch ? 1 : 0)
            .ThenBy(x => slot.RequiredVehicleTypes.Contains(x.Vehicle.Type) ? 0 : 1)
            .ThenBy(x => x.Vehicle.Plate, NaturalStringComparer.Instance)
            .Select((x, i) => new DropdownOption { Display = x.Vehicle.Plate, Raw = x.Vehicle.Plate, Section = i, IsValid = !x.IsMismatch })
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
