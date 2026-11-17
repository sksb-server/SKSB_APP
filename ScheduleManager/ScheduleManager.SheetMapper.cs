namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;

internal static partial class SheetMapper
{
    // =========================================================================
    // 1. SHEET / TABLE IDENTITY
    // =========================================================================
    internal static readonly string[] MasterSheetAliases = { "Daily Attendances", "Daily Attendance", "Attendance" };
    internal const string MasterTableName = "DailyAttendanceTable";
    internal static readonly string[] PreferenceAliases = { "Preference", "Staff Preference", "Preferences" };
    internal static readonly string[] EntriesAliases = { "Attendance Entries", "Attendance_Entries" };
    internal const string EntriesTableName = "AttendanceEntriesTable";

    /// <summary>Late Checked In Cut Off Time parameter (moved here from the retired 'Daily Salary' sheet).</summary>
    internal const string LateCutOffCell = "D2";

    private static bool _isRefreshing;
    private static bool _refreshQueued;

    /// <summary>Date currently shown on the board; lets a date change be reverted when the user declines.</summary>
    private static DateOnly? _loadedDate;

    // =========================================================================
    // 2. EXTRACTION & VALIDATION PIPELINE
    // =========================================================================
    internal static (ScheduleRecord? Data, List<string> Errors, List<string> Warnings) ExtractScheduleRecordWithDiagnostics()
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<string> { "No active workbook found." }, warnings);
            object wbObj = wb;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (board == null || attendanceSheet == null || entriesSheet == null || prefSheet == null)
            {
                errors.Add("  The workbook is missing one of: 'Daily Schedule', 'Daily Attendances', 'Attendance Entries', 'Preference'.");
                return (null, errors, warnings);
            }

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board);
            if (date == null)
            {
                errors.Add($"  [Daily Schedule]: Date cell {ScheduleBoardMapper.DateCell} is empty or not a valid date.");
                return (null, errors, warnings);
            }

            var record = new ScheduleRecord { DailyScheduleDate = date.Value };

            // Entries first: default approvals (WhatsApp starts unapproved) must be on the sheet before it is read.
            NormalizeApprovals(entriesSheet);

            RunFrozen(() =>
            {
                record.Preferences = TableMapper.ExtractFromTable<WorkerPreference>((object)prefSheet, "WorkerPreferencesTable");
                record.Teams = TableMapper.ExtractFromTable<TeamRecord>((object)prefSheet, "TeamTable");
                record.AttendanceEntries = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet, EntriesTableName);
                var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet, MasterTableName);
                foreach (var e in existing) e.DailyScheduleDate = date.Value;
                record.WorkerAttendance = existing;
            });
            record.Projects = ScheduleBoardMapper.Read((object)board, date.Value);

            BoardContext.SetPreferences(record.Preferences, record.Teams);
            foreach (var t in record.Teams) t.TeamLeaderStaffID = Roster.FindStaff(t.TeamLeader)?.Id ?? 0;

            // ---- table rule validation (required cells, dropdown sources, formats) ----
            var (ruleErrors, ruleWarnings) = RulesExtractor.ValidateAllWorkbookTables(wbObj);
            errors.AddRange(ruleErrors);
            warnings.AddRange(ruleWarnings);

            ValidateSchedule(record, errors, warnings);

            // ---- roll punches + board up into the daily attendance rows that get persisted ----
            var punchesForDate = record.AttendanceEntries.Where(e => e.Date == date.Value).ToList();
            record.WorkerAttendance = AttendanceRollupService.Run(date.Value, punchesForDate, record.Projects, record.WorkerAttendance, BoardContext.LateCutOff);

            return (record, errors, warnings);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed extracting schedule record", ex);
            return (null, new List<string> { $"Failed extracting schedule: {ex.Message}" }, warnings);
        }
    }

    private static void ValidateSchedule(ScheduleRecord record, List<string> errors, List<string> warnings)
    {
        // Duplicate workers within a shift band block the save.
        errors.AddRange(ScheduleBoardMapper.FindDuplicateWorkers(record.Projects));

        foreach (var slot in record.Projects)
        {
            string label = $"{(slot.Shift == Shifts.Night ? "Night" : "Day")} Work slot {slot.SlotIndex + 1}";
            if (slot.Workers.Count > 0 && slot.ProjectID.Length == 0)
                warnings.Add($"  [Daily Schedule - {label}]: workers are assigned but no project code is set.");
            if (slot.ProjectID.Length > 0 && Roster.FindProject(slot.ProjectID) == null)
                warnings.Add($"  [Daily Schedule - {label}]: project '{slot.ProjectID}' was not found in project master.");

            foreach (var worker in slot.Workers.Where(w => Roster.FindStaff(w) == null).Distinct(StringComparer.OrdinalIgnoreCase))
                errors.Add($"  [Daily Schedule - {label}]: '{worker}' is not a known staff short name (a team name must be expanded first).");
            foreach (var plate in slot.Vehicles.Where(p => Roster.FindVehicle(p) == null).Distinct(StringComparer.OrdinalIgnoreCase))
                warnings.Add($"  [Daily Schedule - {label}]: vehicle '{plate}' was not found in vehicle master.");
        }

        foreach (var g in record.Preferences.Where(p => p.StaffID > 0).GroupBy(p => p.StaffID).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate Staff ID '{g.Key}' in 'WorkerPreferencesTable'.");
        foreach (var g in record.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)).GroupBy(t => t.TeamName.Trim(), StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate team '{g.Key}' in 'TeamTable'.");

        foreach (var g in record.AttendanceEntries.GroupBy(e => (e.Date, e.Time, e.StaffID, e.EntryType, e.Source)).Where(g => g.Count() > 1))
            errors.Add($"  [Attendance Entries]: Staff {g.Key.StaffID} has {g.Count()} identical {g.Key.Source} {g.Key.EntryType} punches at {g.Key.Date:dd/MM/yyyy} {g.Key.Time:HH:mm}.");

        foreach (var day in record.AttendanceEntries.Where(e => e.IsApproved).GroupBy(e => (e.StaffID, e.Date)))
        {
            var ordered = day.OrderBy(e => e.Time).ToList();
            if (ordered[0].EntryType == AttendanceEntryType.CheckOut)
                warnings.Add($"  [Attendance Entries]: Staff {day.Key.StaffID} on {day.Key.Date:dd/MM/yyyy} starts with a CheckOut and has no earlier CheckIn.");
        }

        int pending = record.AttendanceEntries.Count(e => !e.IsApproved);
        if (pending > 0)
            warnings.Add($"  [Attendance Entries]: {pending} unapproved entr{(pending == 1 ? "y is" : "ies are")} excluded from the attendance rollup until 'Is Approved' is TRUE.");
    }

    // =========================================================================
    // 3. APPROVAL DEFAULTS (Thumbprint/Manual/SKSB_App => TRUE, WhatsApp => FALSE until an admin confirms)
    // =========================================================================
    private static void NormalizeApprovals(dynamic entriesSheet)
    {
        try
        {
            dynamic? table = TableMapper.ResolveTable(entriesSheet, EntriesTableName);
            dynamic? body = table?.DataBodyRange;
            if (table == null || body == null) return;

            int rows = (int)body.Rows.Count;
            int cols = (int)table.ListColumns.Count;
            object[,] header = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, cols);
            int sourceCol = -1, approvedCol = -1;
            for (int c = 1; c <= cols; c++)
            {
                string h = header[1, c]?.ToString()?.Trim() ?? string.Empty;
                if (h.Equals("Source", StringComparison.OrdinalIgnoreCase)) sourceCol = c;
                else if (h.Equals("Is Approved", StringComparison.OrdinalIgnoreCase)) approvedCol = c;
            }
            if (sourceCol < 0 || approvedCol < 0) return;

            object[,] data = TableMapper.GetRange2DValues(body, rows, cols);
            var column = new object?[rows, 1];
            bool dirty = false;
            for (int r = 1; r <= rows; r++)
            {
                object? current = data[r, approvedCol];
                string source = data[r, sourceCol]?.ToString()?.Trim() ?? string.Empty;
                bool rowHasData = Enumerable.Range(1, cols).Any(c => c != approvedCol && !string.IsNullOrWhiteSpace(data[r, c]?.ToString()));

                if ((current == null || string.IsNullOrWhiteSpace(current.ToString())) && rowHasData && source.Length > 0)
                {
                    column[r - 1, 0] = !source.Equals(nameof(AttendanceSource.WhatsApp), StringComparison.OrdinalIgnoreCase);
                    dirty = true;
                }
                else
                {
                    column[r - 1, 0] = current;
                }
            }

            if (!dirty) return;
            RunFrozen(() =>
            {
                dynamic target = entriesSheet.Range[body.Cells[1, approvedCol], body.Cells[rows, approvedCol]];
                target.Value2 = column;
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] NormalizeApprovals warning: {ex.Message}");
        }
    }

    // =========================================================================
    // 4. POPULATE
    // =========================================================================
    internal static void PopulateWorksheet(ScheduleRecord record)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            object wbObj = wb;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (board == null) { AppLogger.Warn("[ScheduleManager] Populate skipped: 'Daily Schedule' sheet not found."); return; }

            RunFrozen(() =>
            {
                Roster.Invalidate();
                DateOnly date = record.DailyScheduleDate;
                ScheduleBoardMapper.WriteDate((object)board, date);
                _loadedDate = date;

                // Preferences & teams (names are UI-only helpers resolved from the roster)
                foreach (var p in record.Preferences) p.StaffName = Roster.FindStaff(p.StaffID)?.ShortName ?? p.StaffName;
                foreach (var t in record.Teams)
                    if (t.TeamLeaderStaffID > 0) t.TeamLeader = Roster.FindStaff(t.TeamLeaderStaffID)?.ShortName ?? t.TeamLeader;

                if (prefSheet != null)
                {
                    TableMapper.PopulateTableSafe<TeamRecord>((object)prefSheet, "TeamTable", record.Teams);
                    TableMapper.PopulateTableSafe<WorkerPreference>((object)prefSheet, "WorkerPreferencesTable", record.Preferences);
                }

                // Context (requirements, icons, cut-off) comes from the workbook; prefs/teams from the record just loaded.
                BoardContext.Reload(wbObj);
                BoardContext.SetPreferences(record.Preferences, record.Teams);

                var entries = record.AttendanceEntries.OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();
                foreach (var e in entries) e.StaffName = Roster.FindStaff(e.StaffID)?.ShortName ?? e.StaffName;
                if (entriesSheet != null)
                    TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet, EntriesTableName, entries);

                ScheduleBoardMapper.Render((object)board, date, record.Projects, entries);

                if (attendanceSheet != null)
                {
                    var rows = AttendanceRollupService.Run(date, entries.Where(e => e.Date == date).ToList(), record.Projects, record.WorkerAttendance, BoardContext.LateCutOff);
                    WriteAttendanceRows(attendanceSheet, rows);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed populating schedule workbook", ex);
        }
    }

    private static void WriteAttendanceRows(dynamic attendanceSheet, List<DailyWorkerAttendance> rows)
    {
        foreach (var r in rows) r.StaffName = Roster.FindStaff(r.StaffID)?.ShortName ?? r.StaffName;
        TableMapper.PopulateTableSafe<DailyWorkerAttendance>((object)attendanceSheet, MasterTableName, rows);
    }

    // =========================================================================
    // 5. LIVE REFRESH (board render + attendance rollup)
    // =========================================================================
    /// <summary>
    /// Re-renders the board (team expansion, icons, check-in/out, project windows, requirement flags) and,
    /// optionally, refreshes the Daily Attendances rollup from the current punches. Safe to call repeatedly:
    /// events are suppressed while it writes and it never re-enters itself.
    /// </summary>
    internal static void RefreshDerived(bool reloadContext, bool rollup)
    {
        if (_isRefreshing) return;
        _isRefreshing = true;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null || SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;
            object wbObj = wb;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            if (board == null) return;

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board);
            if (date == null) return;

            if (entriesSheet != null) NormalizeApprovals(entriesSheet);

            RunFrozen(() =>
            {
                if (reloadContext) BoardContext.Reload(wbObj);

                var slots = ScheduleBoardMapper.Read((object)board, date.Value);
                if (ScheduleBoardMapper.ExpandTeams(slots, out string? notice) && notice != null)
                    SheetOperations.ShowStatusConfirmation(notice);

                var entries = entriesSheet != null
                    ? TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet, EntriesTableName)
                    : new List<AttendanceEntry>();

                ScheduleBoardMapper.Render((object)board, date.Value, slots, entries);

                if (rollup && attendanceSheet != null)
                {
                    var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet, MasterTableName);
                    var rows = AttendanceRollupService.Run(date.Value, entries.Where(e => e.Date == date.Value).ToList(), slots, existing, BoardContext.LateCutOff);
                    if (Signature(rows) != Signature(existing)) WriteAttendanceRows(attendanceSheet, rows);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] RefreshDerived failed", ex);
        }
        finally
        {
            _isRefreshing = false;
        }
    }

    private static string Signature(IEnumerable<DailyWorkerAttendance> rows) =>
        string.Join("|", rows.OrderBy(r => r.StaffID).Select(r =>
            $"{r.StaffID}:{r.AttendanceStatus}:{r.ClockInTime:HH:mm}:{r.ClockOutTime:HH:mm}:{r.ProjectCheckIns}:{r.RegularShift}:{r.AdditionalShifts}:{r.OTHours}:{r.AdditionalOTHours}:{r.TotalFlattenedHours}:{r.ManualOverride}:{r.OverrideReason}:{r.IsAllocated}"));

    internal static void QueueRefresh(bool reloadContext, bool rollup)
    {
        if (_refreshQueued) return;
        _refreshQueued = true;
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _refreshQueued = false;
            RefreshDerived(reloadContext, rollup);
        });
    }

    // =========================================================================
    // 6. SHEET-CHANGE OBSERVER (registered once through ParameterCells)
    // =========================================================================
    internal static void OnSheetChanged(dynamic sheet, dynamic target)
    {
        try
        {
            if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode || _isRefreshing) return;

            string name = sheet.Name?.ToString() ?? string.Empty;
            int row = (int)target.Row;
            int col = (int)target.Column;

            if (ScheduleBoardMapper.IsBoardSheet((object)sheet))
            {
                if (target.Cells.Count == 1 && row == 2 && col == 1) { ExcelAsyncUtil.QueueAsMacro(OnDateCellChanged); return; }
                if (ScheduleBoardMapper.IsBoardArea(row, col) || target.Cells.Count > 1) QueueRefresh(reloadContext: false, rollup: true);
                return;
            }

            if (EntriesAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase))) QueueRefresh(reloadContext: false, rollup: true);
            else if (PreferenceAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) ||
                     name.Equals("Lookups", StringComparison.OrdinalIgnoreCase)) QueueRefresh(reloadContext: true, rollup: false);
            else if (MasterSheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) && row <= 2) QueueRefresh(reloadContext: true, rollup: true);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] OnSheetChanged warning: {ex.Message}");
        }
    }

    /// <summary>Moves the board to the target date (confirming first) and loads that day.</summary>
    internal static void GoToDate(DateOnly target)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return;
        dynamic? board = SheetMapperHelper.ResolveSheetSafely((object)wb, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
        if (board == null) return;

        RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board, target));
        OnDateCellChanged();
    }

    private static void OnDateCellChanged()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            dynamic? board = SheetMapperHelper.ResolveSheetSafely((object)wb, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            if (board == null) return;

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board);
            if (date == null || date == _loadedDate) return;

            var answer = MessageBox.Show(
                $"Load the schedule for {date:dddd, dd MMM yyyy}?\n\nUnsaved changes to {(_loadedDate.HasValue ? _loadedDate.Value.ToString("dd MMM yyyy") : "the current date")} will be discarded.",
                "Schedule Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (answer != DialogResult.Yes)
            {
                if (_loadedDate.HasValue) RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board, _loadedDate.Value));
                return;
            }

            var record = ServerSyncCommands.FetchSchedule(date.Value);
            PopulateWorksheet(record);
            SheetOperations.ShowStatusConfirmation($"Schedule for {date:yyyy-MM-dd} loaded.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Date change failed", ex);
        }
    }

    // =========================================================================
    // 6b. THUMBPRINT TERMINAL APPEND
    // =========================================================================
    /// <summary>
    /// Appends terminal punches for the board date to the entries table (Source = Thumbprint), skipping punches
    /// already present. Must run on the macro thread. Returns (added, duplicates, unknown or other-date).
    /// </summary>
    internal static (int Added, int Duplicates, int Unknown) AppendTerminalPunches(IReadOnlyList<TerminalPunch> punches)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return (0, 0, punches.Count);
        object wbObj = wb;

        dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
        DateOnly? date = board != null ? ScheduleBoardMapper.ReadDate((object)board) : null;
        if (entriesSheet == null || date == null) return (0, 0, punches.Count);

        var forDate = punches.Where(p => DateOnly.FromDateTime(p.Timestamp) == date.Value).ToList();
        int added = 0, duplicates = 0, unknown = punches.Count - forDate.Count;

        NormalizeApprovals(entriesSheet);
        RunFrozen(() =>
        {
            var existing = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet, EntriesTableName);
            var result = DeviceSyncManager.BuildEntries(forDate, existing, id => Roster.FindStaff(id) != null);
            added = result.NewEntries.Count;
            duplicates = result.Duplicates;
            unknown += result.UnknownStaff;
            if (added == 0) return;

            var all = existing.Concat(result.NewEntries).OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();
            foreach (var e in all) e.StaffName = Roster.FindStaff(e.StaffID)?.ShortName ?? e.StaffName;
            TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet, EntriesTableName, all);
        });

        if (added > 0) RefreshDerived(reloadContext: false, rollup: true);
        return (added, duplicates, unknown);
    }

    // =========================================================================
    // 7. COM FREEZE (Rule 2): events, screen updating and calculation restored in finally
    // =========================================================================
    internal static void RunFrozen(Action body)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        bool origUpdating = true;
        int origCalculation = -4105; // xlCalculationAutomatic
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }
        try
        {
            body();
        }
        finally
        {
            try { xlApp.Calculation = origCalculation; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }
}
