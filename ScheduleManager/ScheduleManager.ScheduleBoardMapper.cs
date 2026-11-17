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
using SKSB_App.Core.WorkbookSave.SheetMapper;

/// <summary>
/// Reads and renders the 'Daily Schedule' dispatch board: 8 slots x (Day Work rows 3-18, Night Work rows 19-34).
/// Every slot is three columns wide (B:D, E:G, ... W:Y):
///   header   : project code + live "(HH:mm - HH:mm)" window                    (anchor of the merged header)
///   premix   : "Premix Holes" label / count one row above the header
///   workers  : Staff Name | Staff Skills (icons) | Check in and out Time        (9 rows)
///   vehicles : Vehicle Plate | Vehicle Type (icon) | Current Odometer            (3 rows)
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

    internal const int SlotCount = 8;
    internal const int SlotWidth = 3;
    internal const int FirstSlotColumn = 2;                 // B
    internal const int WorkerRowsPerBand = 9;
    internal const int VehicleRowsPerBand = 3;
    private const int FirstBoardRow = 3;                    // Day premix row
    private const int LastBoardRow = 34;
    private const int LastBoardColumn = FirstSlotColumn + (SlotCount * SlotWidth) - 1; // Y
    private const string EmptyHeaderText = "Project Code (Time start - Time end)";
    private const string PremixLabel = "Premix Holes";

    private sealed record Band(Shifts Shift, string Label, int PremixRow, int HeaderRow, int FirstWorkerRow, int VehicleHeaderRow, int FirstVehicleRow)
    {
        public int LastWorkerRow => FirstWorkerRow + WorkerRowsPerBand - 1;
        public int LastVehicleRow => FirstVehicleRow + VehicleRowsPerBand - 1;
    }

    private static Band MakeBand(Shifts shift, string label, int offset) =>
        new(shift, label, 3 + offset, 4 + offset, 6 + offset, 15 + offset, 16 + offset);

    private static readonly Band[] Bands = { MakeBand(Shifts.Morning, "Day Work", 0), MakeBand(Shifts.Night, "Night Work", 16) };

    internal static int SlotColumn(int slot) => FirstSlotColumn + (slot * SlotWidth);

    private static readonly Regex HeaderCode = new(@"^\s*(\S+)", RegexOptions.Compiled);

    internal enum CellKind { None, Header, Worker, Vehicle, Premix }
    internal readonly record struct CellLocation(CellKind Kind, Shifts Shift, int Slot, int Row);

    internal static CellLocation Locate(int row, int column)
    {
        var none = new CellLocation(CellKind.None, Shifts.None, 0, row);
        if (column < FirstSlotColumn || column > LastBoardColumn) return none;
        int slot = (column - FirstSlotColumn) / SlotWidth;
        int within = (column - FirstSlotColumn) % SlotWidth;

        foreach (var b in Bands)
        {
            if (row == b.HeaderRow && within == 0) return new(CellKind.Header, b.Shift, slot, row);
            if (row == b.PremixRow && within == 1) return new(CellKind.Premix, b.Shift, slot, row);
            if (row >= b.FirstWorkerRow && row <= b.LastWorkerRow && within == 0) return new(CellKind.Worker, b.Shift, slot, row);
            if (row >= b.FirstVehicleRow && row <= b.LastVehicleRow && within == 0) return new(CellKind.Vehicle, b.Shift, slot, row);
        }
        return none;
    }

    internal static bool IsBoardArea(int row, int column) =>
        row >= FirstBoardRow && row <= LastBoardRow && column >= FirstSlotColumn && column <= LastBoardColumn;

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

    private static Band BandOf(Shifts shift) => Bands.First(b => b.Shift == shift);

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
        ws.Range[DayCell].Value2 = date.ToString("dddd", CultureInfo.CurrentCulture);
    }

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

    /// <summary>Reads every non-empty slot. Rank prefixes ("[95%] ") pasted into a cell are stripped.</summary>
    internal static List<DailyProjectSchedule> Read(object wsObj, DateOnly date)
    {
        object[,] grid = ReadGrid(wsObj);
        var result = new List<DailyProjectSchedule>();

        foreach (var band in Bands)
        {
            for (int slot = 0; slot < SlotCount; slot++)
            {
                int col = SlotColumn(slot);
                var record = new DailyProjectSchedule
                {
                    DailyScheduleDate = date,
                    Shift = band.Shift,
                    SlotIndex = slot,
                    ProjectID = ParseProjectCode(Text(grid, band.HeaderRow, col)),
                    PremixHoles = int.TryParse(Text(grid, band.PremixRow, col + 1), NumberStyles.Any, CultureInfo.InvariantCulture, out int holes) ? Math.Max(0, holes) : 0
                };

                for (int r = band.FirstWorkerRow; r <= band.LastWorkerRow; r++)
                {
                    string worker = ScheduleMatchEngine.StripPercent(Text(grid, r, col));
                    if (worker.Length > 0) record.Workers.Add(worker);
                }
                for (int r = band.FirstVehicleRow; r <= band.LastVehicleRow; r++)
                {
                    string plate = Text(grid, r, col).ToUpperInvariant();
                    if (plate.Length > 0) record.Vehicles.Add(plate);
                }

                if (record.ProjectID.Length > 0 || record.Workers.Count > 0 || record.Vehicles.Count > 0 || record.PremixHoles > 0)
                    result.Add(record);
            }
        }
        return result;
    }

    /// <summary>Workers listed in more than one slot of the same band, as save-blocking messages.</summary>
    internal static List<string> FindDuplicateWorkers(IEnumerable<DailyProjectSchedule> slots)
    {
        var messages = new List<string>();
        foreach (var band in Bands)
        {
            var seen = slots.Where(s => s.Shift == band.Shift)
                .SelectMany(s => s.Workers.Select(w => (Name: w, s.SlotIndex)))
                .GroupBy(x => x.Name, StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1);

            foreach (var g in seen)
            {
                string where = string.Join(", ", g.Select(x => $"slot {x.SlotIndex + 1}"));
                messages.Add($"  [Daily Schedule - {band.Label}]: Worker '{g.Key}' is assigned more than once ({where}).");
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
                if (!string.IsNullOrWhiteSpace(team.TeamLeader)) members.Add(team.TeamLeader.Trim());
                members.AddRange(BoardContext.Preferences
                    .Where(p => p.Team.Contains(team.TeamName.Trim(), StringComparer.OrdinalIgnoreCase))
                    .Select(p => Roster.FindStaff(p.StaffID)?.ShortName ?? p.StaffName)
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
    internal static void Render(object wsObj, DateOnly date, IReadOnlyList<DailyProjectSchedule> slots, IReadOnlyList<AttendanceEntry> entries)
    {
        dynamic ws = wsObj;
        var spans = AttendanceRollupService.BuildSpans(entries, date, slots);
        string cutOffText = BoardContext.LateCutOff.ToString("HH:mm", CultureInfo.InvariantCulture);

        foreach (var band in Bands)
        {
            int cols = SlotCount * SlotWidth;
            var workers = new object?[WorkerRowsPerBand, cols];
            var vehicles = new object?[VehicleRowsPerBand, cols];
            var premix = new object?[1, cols];

            for (int slot = 0; slot < SlotCount; slot++)
            {
                var data = slots.FirstOrDefault(s => s.Shift == band.Shift && s.SlotIndex == slot) ?? new DailyProjectSchedule { Shift = band.Shift, SlotIndex = slot };
                int c0 = slot * SlotWidth;

                premix[0, c0] = PremixLabel;
                premix[0, c0 + 1] = data.PremixHoles;

                var assignedSkills = new List<string>();
                for (int i = 0; i < WorkerRowsPerBand && i < data.Workers.Count; i++)
                {
                    string name = data.Workers[i];
                    var staff = Roster.FindStaff(name);
                    var pref = staff != null ? BoardContext.PreferenceFor(staff.Id) : null;

                    var skills = SkillsOf(pref);
                    assignedSkills.AddRange(skills);

                    workers[i, c0] = name;
                    workers[i, c0 + 1] = string.Join(" ", skills.Select(BoardContext.IconFor).Where(s => s.Length > 0));
                    workers[i, c0 + 2] = staff != null && data.ProjectID.Length > 0 && spans.TryGetValue((staff.Id, data.ProjectID), out var list) && list.Count > 0
                        ? FormatSpan(list[0].In, list.Any(s => s.Out == null) ? null : list.Max(s => s.Out))
                        : null;
                }

                var vehicleTypes = new List<VehicleType>();
                for (int i = 0; i < VehicleRowsPerBand && i < data.Vehicles.Count; i++)
                {
                    var v = Roster.FindVehicle(data.Vehicles[i]);
                    vehicles[i, c0] = data.Vehicles[i];
                    if (v != null)
                    {
                        vehicleTypes.Add(v.Type);
                        vehicles[i, c0 + 1] = $"{BoardContext.IconFor(v.Type)} {v.Type.ToString().Replace('_', ' ')}".Trim();
                        vehicles[i, c0 + 2] = v.Odometer > 0 ? (double)v.Odometer : null;
                    }
                }

                WriteHeader(ws, band, slot, data, spans, assignedSkills, vehicleTypes);
            }

            dynamic workerRange = ws.Range[ws.Cells[band.FirstWorkerRow, FirstSlotColumn], ws.Cells[band.LastWorkerRow, LastBoardColumn]];
            workerRange.Value2 = workers;
            dynamic vehicleRange = ws.Range[ws.Cells[band.FirstVehicleRow, FirstSlotColumn], ws.Cells[band.LastVehicleRow, LastBoardColumn]];
            vehicleRange.Value2 = vehicles;
            dynamic premixRange = ws.Range[ws.Cells[band.PremixRow, FirstSlotColumn], ws.Cells[band.PremixRow, LastBoardColumn]];
            premixRange.Value2 = premix;
        }

        EnsureConditionalFormats(ws, cutOffText);
    }

    private static string FormatSpan(TimeOnly start, TimeOnly? end) =>
        $"{start:HH:mm} - {(end.HasValue ? end.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "…")}";

    internal static List<string> SkillsOf(WorkerPreference? pref)
    {
        if (pref == null) return new List<string>();
        var list = new List<string>();
        if (!string.IsNullOrWhiteSpace(pref.PrimarySkill)) list.Add(pref.PrimarySkill.Trim());
        list.AddRange(pref.LicensedSkills);
        list.AddRange(pref.OtherSkills);
        return list.Where(s => !string.IsNullOrWhiteSpace(s)).Select(s => s.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
    }

    // ---- header: "CODE (08:05 - 17:30) ⚠" + requirements note -------------------------------
    private static void WriteHeader(
        dynamic ws, Band band, int slot, DailyProjectSchedule data,
        Dictionary<(long StaffId, string ProjectId), List<WorkerProjectSpan>> spans,
        List<string> assignedSkills, List<VehicleType> assignedVehicleTypes)
    {
        dynamic cell = ws.Cells[band.HeaderRow, SlotColumn(slot)];
        if (data.ProjectID.Length == 0)
        {
            cell.Value2 = EmptyHeaderText;
            try { cell.ClearComments(); } catch { }
            return;
        }

        var window = AttendanceRollupService.WindowFor(data.ProjectID, spans);
        string timeText = window.Start.HasValue
            ? $" ({window.Start:HH:mm} - {(window.End.HasValue ? window.End.Value.ToString("HH:mm", CultureInfo.InvariantCulture) : "…")})"
            : string.Empty;

        var info = Roster.FindProject(data.ProjectID);
        var (reqSkills, reqVehicles) = BoardContext.RequirementsFor(info?.Types ?? Array.Empty<ProjectType>());
        var missingSkills = reqSkills.Where(r => !assignedSkills.Contains(r, StringComparer.OrdinalIgnoreCase)).ToList();
        var missingVehicles = reqVehicles.Where(r => !assignedVehicleTypes.Contains(r)).ToList();
        bool unmet = missingSkills.Count > 0 || missingVehicles.Count > 0;

        cell.Value2 = $"{data.ProjectID}{timeText}{(unmet ? " ⚠" : string.Empty)}";

        try
        {
            cell.ClearComments();
            if (reqSkills.Count > 0 || reqVehicles.Count > 0)
            {
                var sb = new StringBuilder();
                sb.AppendLine($"Requirements ({string.Join(", ", info?.Types ?? Array.Empty<ProjectType>())})");
                sb.AppendLine("Skills:   " + (reqSkills.Count == 0 ? "-" : string.Join(", ", reqSkills.Select(r => $"{BoardContext.IconFor(r)} {r} {(missingSkills.Contains(r) ? "✘" : "✔")}"))));
                sb.Append("Vehicles: " + (reqVehicles.Count == 0 ? "-" : string.Join(", ", reqVehicles.Select(r => $"{BoardContext.IconFor(r)} {r.ToString().Replace('_', ' ')} {(missingVehicles.Contains(r) ? "✘" : "✔")}"))));
                cell.AddComment(sb.ToString());
            }
        }
        catch { /* comments are advisory */ }
    }

    // =========================================================================
    // 6. CONDITIONAL FORMATS (duplicate worker = red, late check-in = amber)
    // =========================================================================
    private static void EnsureConditionalFormats(dynamic ws, string cutOffText)
    {
        const int xlExpression = 2;
        foreach (var band in Bands)
        {
            // Duplicate worker: same name anywhere in the band's 8 name columns.
            var nameCols = Enumerable.Range(0, SlotCount).Select(s => ColumnLetter(SlotColumn(s))).ToList();
            string counts = string.Join("+", nameCols.Select(c => $"COUNTIF(${c}${band.FirstWorkerRow}:${c}${band.LastWorkerRow},{{0}})"));

            for (int slot = 0; slot < SlotCount; slot++)
            {
                string nameCol = ColumnLetter(SlotColumn(slot));
                string timeCol = ColumnLetter(SlotColumn(slot) + 2);

                try
                {
                    dynamic names = ws.Range[$"{nameCol}{band.FirstWorkerRow}:{nameCol}{band.LastWorkerRow}"];
                    names.FormatConditions.Delete();
                    string first = $"{nameCol}{band.FirstWorkerRow}";
                    dynamic dup = names.FormatConditions.Add(xlExpression, Type.Missing, $"=AND({first}<>\"\",({string.Format(counts, first)})>1)");
                    dup.Interior.Color = 13551615;   // RGB(255,199,206)
                    dup.Font.Color = 393372;         // RGB(156,0,6)
                }
                catch (Exception ex) { AppLogger.Warn($"[ScheduleBoard] duplicate highlight rule failed: {ex.Message}"); }

                if (band.Shift != Shifts.Morning) continue;
                try
                {
                    dynamic times = ws.Range[$"{timeCol}{band.FirstWorkerRow}:{timeCol}{band.LastWorkerRow}"];
                    times.FormatConditions.Delete();
                    string first = $"{timeCol}{band.FirstWorkerRow}";
                    dynamic late = times.FormatConditions.Add(xlExpression, Type.Missing, $"=AND(LEN({first})>=5,LEFT({first},5)>\"{cutOffText}\")");
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
        var loc = Locate((int)cell.Row, (int)cell.Column);
        return loc.Kind is CellKind.Header or CellKind.Worker or CellKind.Vehicle;
    }

    private static CustomDropdownResult? BuildDropdown(dynamic cell)
    {
        var loc = Locate((int)cell.Row, (int)cell.Column);
        object ws = cell.Worksheet;
        DateOnly date = ReadDate(ws) ?? DateOnly.FromDateTime(DateTime.Today);

        switch (loc.Kind)
        {
            case CellKind.Header:
                return new CustomDropdownResult(
                    Roster.Projects.Select((p, i) => new DropdownOption { Display = p.Id, Raw = p.Id, Section = i }).ToList(),
                    (c, opt) => c.Value2 = opt.Raw);

            case CellKind.Vehicle:
            {
                var slot = ReadSlotContext(ws, date, loc);
                var ranked = Roster.Vehicles
                    .OrderBy(v => slot.RequiredVehicleTypes.Contains(v.Type) ? 0 : 1)
                    .ThenBy(v => v.Plate, StringComparer.OrdinalIgnoreCase)
                    .Select((v, i) => new DropdownOption { Display = v.Plate, Raw = v.Plate, Section = i })
                    .ToList();
                return new CustomDropdownResult(ranked, (c, opt) => c.Value2 = opt.Raw);
            }

            case CellKind.Worker:
                return BuildWorkerDropdown(ws, date, loc);

            default:
                return null;
        }
    }

    private const string TeamPrefix = "T:";
    private const string StaffPrefix = "S:";

    private static CustomDropdownResult BuildWorkerDropdown(object wsObj, DateOnly date, CellLocation loc)
    {
        var slot = ReadSlotContext(wsObj, date, loc);
        var scored = new List<(string Display, string Raw, int Percent, bool Blocked)>();

        var profiles = new Dictionary<long, WorkerProfile>();
        foreach (var s in Roster.Staff.Where(s => s.ShortName.Length > 0 && s.IsActive))
        {
            profiles[s.Id] = new WorkerProfile(s, BoardContext.PreferenceFor(s.Id), BoardContext.DaysWorkedEarlierThisWeek(s.Id, date), false);
            var r = ScheduleMatchEngine.Score(profiles[s.Id], slot);
            scored.Add((ScheduleMatchEngine.WithPercent(s.ShortName, r.Percent), StaffPrefix + s.ShortName, r.Percent, r.Blocked));
        }

        foreach (var team in BoardContext.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)))
        {
            var members = BoardContext.Preferences
                .Where(p => p.Team.Contains(team.TeamName.Trim(), StringComparer.OrdinalIgnoreCase))
                .Select(p => profiles.TryGetValue(p.StaffID, out var wp) ? wp : null)
                .Where(wp => wp != null).Select(wp => wp!).ToList();
            var r = ScheduleMatchEngine.ScoreTeam(members, slot);
            scored.Add((ScheduleMatchEngine.WithPercent(team.TeamName.Trim(), r.Percent), TeamPrefix + team.TeamName.Trim(), r.Percent, r.Blocked));
        }

        var ordered = scored
            .OrderBy(x => x.Blocked ? 1 : 0)
            .ThenByDescending(x => x.Percent)
            .ThenBy(x => x.Display, StringComparer.OrdinalIgnoreCase)
            .Select((x, i) => new DropdownOption { Display = x.Display, Raw = x.Raw, Section = i, IsValid = !x.Blocked })
            .ToList();

        // Commit strips the "[XX%] " prefix; a team name is expanded by the sheet-change observer.
        return new CustomDropdownResult(ordered, (c, opt) =>
            c.Value2 = ScheduleMatchEngine.StripPercent(opt.Raw.StartsWith(TeamPrefix, StringComparison.Ordinal) || opt.Raw.StartsWith(StaffPrefix, StringComparison.Ordinal)
                ? opt.Raw[2..]
                : opt.Raw));
    }

    private static SlotContext ReadSlotContext(object wsObj, DateOnly date, CellLocation loc)
    {
        var slots = Read(wsObj, date);
        var mine = slots.FirstOrDefault(s => s.Shift == loc.Shift && s.SlotIndex == loc.Slot) ?? new DailyProjectSchedule { Shift = loc.Shift, SlotIndex = loc.Slot };

        var info = Roster.FindProject(mine.ProjectID);
        var (skills, vehicleTypes) = BoardContext.RequirementsFor(info?.Types ?? Array.Empty<ProjectType>());

        var elsewhere = new HashSet<long>();
        foreach (var other in slots.Where(s => s.Shift == loc.Shift && s.SlotIndex != loc.Slot))
            foreach (var w in other.Workers)
                if (Roster.FindStaff(w) is { } st) elsewhere.Add(st.Id);

        var plates = mine.Vehicles;
        var types = plates.Select(p => Roster.FindVehicle(p)?.Type ?? VehicleType.None).Where(t => t != VehicleType.None).ToList();
        return new SlotContext(date, loc.Shift, skills, vehicleTypes, plates, types, elsewhere);
    }
}
