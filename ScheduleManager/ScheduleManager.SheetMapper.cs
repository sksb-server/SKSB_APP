namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
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
    internal const string LateCutOffCell = "C2";

    /// <summary>OT Start Time parameter: day-band time clocked after this counts as overtime.</summary>
    internal const string OtStartCell = "D2";

    private static bool _isRefreshing;
    private static bool _refreshQueued;

    /// <summary>Session key for <see cref="PastRecordGuard"/> edit approvals.</summary>
    internal const string GuardSession = "ScheduleManager";

    // True once the user changed a date-bound sheet (board, entries, attendance) since the last load / save.
    // Refresh and populate writes run with events off, so they never set it.
    private static bool _hasUnsavedRecordEdits;
    private static bool _editPromptPending;
    internal static bool HasUnsavedRecordEdits => _hasUnsavedRecordEdits;

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
            object wbObj = wb!;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (board == null || attendanceSheet == null || entriesSheet == null || prefSheet == null)
            {
                errors.Add("  The workbook is missing one of: 'Daily Schedule', 'Daily Attendances', 'Attendance Entries', 'Preference'.");
                return (null, errors, warnings);
            }

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board!);
            if (date == null)
            {
                errors.Add($"  [Daily Schedule]: Date cell {ScheduleBoardMapper.DateCell} is empty or not a valid date.");
                return (null, errors, warnings);
            }

            var record = new ScheduleRecord
            {
                DailyScheduleDate = date.Value,
                LateCutOffTime = BoardContext.ReadTimeCell((object)attendanceSheet!, LateCutOffCell, BoardContext.DefaultLateCutOff),
                OTStartTime = BoardContext.ReadTimeCell((object)attendanceSheet!, OtStartCell, BoardContext.DefaultOtStartTime)
            };

            // Entries first: default approvals (WhatsApp starts unapproved) must be on the sheet before it is read.
            NormalizeApprovals(entriesSheet);

            RunFrozen(() =>
            {
                record.Preferences = TableMapper.ExtractFromTable<WorkerPreference>((object)prefSheet!, "WorkerPreferencesTable");
                record.Teams = TableMapper.ExtractFromTable<TeamRecord>((object)prefSheet!, "TeamTable");
                record.AttendanceEntries = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
                var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet!, MasterTableName);
                foreach (var e in existing) e.DailyScheduleDate = date.Value;
                record.WorkerAttendance = existing;
            });
            record.Projects = ScheduleBoardMapper.Read((object)board!, date.Value);

            BoardContext.SetPreferences(record.Preferences, record.Teams);
            foreach (var t in record.Teams) t.TeamLeaderStaffID = Roster.FindStaff(t.TeamLeader)?.Id ?? 0;

            // ---- table rule validation (required cells, dropdown sources, formats) ----
            var (ruleErrors, ruleWarnings) = RulesExtractor.ValidateAllWorkbookTables(wbObj);
            errors.AddRange(ruleErrors);
            warnings.AddRange(ruleWarnings);

            ValidateSchedule(record, errors, warnings);

            // ---- roll punches + board up into the daily attendance rows that get persisted ----
            var punchesForDate = record.AttendanceEntries.Where(e => e.Date == date.Value).ToList();
            record.WorkerAttendance = AttendanceRollupService.Run(date.Value, punchesForDate, record.Projects, record.WorkerAttendance, record.LateCutOffTime, record.OTStartTime);

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
        DateOnly scheduleDate = record.DailyScheduleDate;

        // Duplicate workers within a shift band block the save.
        errors.AddRange(ScheduleBoardMapper.FindDuplicateWorkers(record.Projects));

        foreach (var slot in record.Projects)
        {
            string label = $"{(slot.Shift == Shifts.Night ? "Night" : "Day")} Work slot {slot.SlotIndex + 1}";
            if (slot.Workers.Count > 0 && slot.ProjectID.Length == 0)
                warnings.Add($"  [Daily Schedule - {label}]: workers are assigned but no project code is set.");
            if (slot.ProjectID.Length > 0 && Roster.FindProject(slot.ProjectID) == null)
                warnings.Add($"  [Daily Schedule - {label}]: project '{slot.ProjectID}' was not found in project master.");

            // Resolution is against ALL field workers (a worker who resigned after this date, or who is already on a
            // historical row, resolves cleanly). Only someone not employed on this date is flagged.
            foreach (var worker in slot.Workers.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var staff = Roster.FindStaff(worker);
                if (staff == null)
                    errors.Add($"  [Daily Schedule - {label}]: '{worker}' is not a known field worker (a team name must be expanded first).");
                else if (!Roster.IsActiveOn(staff, scheduleDate))
                    errors.Add($"  [Daily Schedule - {label}]: '{worker}' was not employed on {scheduleDate:yyyy-MM-dd} ({Roster.DescribeInactive(staff, scheduleDate)}).");
            }
            foreach (var plate in slot.Vehicles.Where(p => Roster.FindVehicle(p) == null).Distinct(StringComparer.OrdinalIgnoreCase))
                warnings.Add($"  [Daily Schedule - {label}]: vehicle '{plate}' was not found in vehicle master.");
        }

        DateOnly today = DateOnly.FromDateTime(DateTime.Today);
        foreach (var p in record.Preferences.Where(p => p.StaffID > 0 && !Roster.IsActiveOn(p.StaffID, today)))
            errors.Add($"  [Preference]: Staff ID '{p.StaffID}' is not a currently employed field worker (office / admin / resigned staff cannot be scheduled).");
        foreach (var id in record.AttendanceEntries
                     .Where(e => e.StaffID > 0 && !Roster.IsActiveOn(e.StaffID, e.Date))
                     .Select(e => e.StaffID).Distinct())
            warnings.Add($"  [Attendance Entries]: Staff ID '{id}' was not an employed field worker on the punch date; those punches are excluded from the attendance rollup.");

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

            int rows = (int)body!.Rows.Count;
            int cols = (int)table!.ListColumns.Count;
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
            object wbObj = wb!;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (board == null) { AppLogger.Warn("[ScheduleManager] Populate skipped: 'Daily Schedule' sheet not found."); return; }

            RunFrozen(() =>
            {
                Roster.Invalidate();
                DateOnly date = record.DailyScheduleDate;
                ScheduleBoardMapper.WriteDate((object)board!, date);
                if (_loadedDate != date) PastRecordGuard.ResetEditSession(GuardSession);   // edit approvals are per date
                _loadedDate = date;
                _hasUnsavedRecordEdits = false;

                // Staff dropdowns: workers employed on this date + anyone this record already refers to.
                var referenced = record.AttendanceEntries.Select(e => e.StaffID)
                    .Concat(record.WorkerAttendance.Select(a => a.StaffID))
                    .Concat(record.Projects.SelectMany(p => p.Workers).Select(w => Roster.FindStaff(w)?.Id ?? 0))
                    .Where(id => id > 0).ToList();
                ServerSyncCommands.SyncLookupLists(date, referenced);

                // Preferences & teams (names are UI-only helpers resolved from the roster)
                // Preferences are global, not dated: only currently employed field workers are shown (and re-saved).
                DateOnly todayDate = DateOnly.FromDateTime(DateTime.Today);
                record.Preferences = record.Preferences.Where(p => Roster.IsActiveOn(p.StaffID, todayDate)).ToList();
                foreach (var p in record.Preferences) p.StaffName = Roster.NameOf(p.StaffID, p.StaffName);
                foreach (var t in record.Teams)
                    if (t.TeamLeaderStaffID > 0) t.TeamLeader = Roster.NameOf(t.TeamLeaderStaffID, t.TeamLeader);

                if (prefSheet != null)
                {
                    TableMapper.PopulateTableSafe<TeamRecord>((object)prefSheet!, "TeamTable", record.Teams);
                    TableMapper.PopulateTableSafe<WorkerPreference>((object)prefSheet!, "WorkerPreferencesTable", record.Preferences);
                }

                // The loaded date's cut-off / OT start go on the sheet first so Reload (and the rollup below) use them.
                if (attendanceSheet != null) WriteAttendanceParameters((object)attendanceSheet!, record.LateCutOffTime, record.OTStartTime);

                // Context (requirements, icons, cut-off, OT start) comes from the workbook; prefs/teams from the record just loaded.
                BoardContext.Reload(wbObj);
                BoardContext.SetPreferences(record.Preferences, record.Teams);

                var entries = record.AttendanceEntries.OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();
                foreach (var e in entries) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
                if (entriesSheet != null)
                    TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, entries);

                ScheduleBoardMapper.Render((object)board!, date, record.Projects, entries, autoPopulate: true);

                if (attendanceSheet != null)
                {
                    var rows = AttendanceRollupService.Run(date, entries.Where(e => e.Date == date).ToList(), record.Projects, record.WorkerAttendance, BoardContext.LateCutOff, BoardContext.OTStartTime);
                    WriteAttendanceRows(attendanceSheet, rows);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed populating schedule workbook", ex);
        }
    }

    /// <summary>Writes the per-date Late Cut-Off (C2) and OT Start (D2) times as Excel time serials. Call inside RunFrozen.</summary>
    private static void WriteAttendanceParameters(dynamic attendanceSheet, TimeOnly lateCutOff, TimeOnly otStart)
    {
        try
        {
            dynamic late = attendanceSheet.Range[LateCutOffCell];
            late.Value2 = lateCutOff.ToTimeSpan().TotalDays;
            late.NumberFormat = "hh:mm";
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Writing {LateCutOffCell} failed: {ex.Message}"); }
        try
        {
            dynamic ot = attendanceSheet.Range[OtStartCell];
            ot.Value2 = otStart.ToTimeSpan().TotalDays;
            ot.NumberFormat = "hh:mm";
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Writing {OtStartCell} failed: {ex.Message}"); }
    }

    private static void WriteAttendanceRows(dynamic attendanceSheet, List<DailyWorkerAttendance> rows)
    {
        foreach (var r in rows) r.StaffName = Roster.NameOf(r.StaffID, r.StaffName);
        TableMapper.PopulateTableSafe<DailyWorkerAttendance>((object)attendanceSheet!, MasterTableName, rows);
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
            object wbObj = wb!;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            if (board == null) return;

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board!);
            if (date == null) return;

            if (entriesSheet != null) NormalizeApprovals(entriesSheet);

            RunFrozen(() =>
            {
                if (reloadContext) BoardContext.Reload(wbObj);

                var slots = ScheduleBoardMapper.Read((object)board!, date.Value);
                if (ScheduleBoardMapper.ExpandTeams(slots, out string? notice) && notice != null)
                    SheetOperations.ShowStatusConfirmation(notice);

                var entries = entriesSheet != null
                    ? TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName)
                    : new List<AttendanceEntry>();

                ScheduleBoardMapper.Render((object)board!, date.Value, slots, entries);

                if (rollup && attendanceSheet != null)
                {
                    var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet!, MasterTableName);
                    var rows = AttendanceRollupService.Run(date.Value, entries.Where(e => e.Date == date.Value).ToList(), slots, existing, BoardContext.LateCutOff, BoardContext.OTStartTime);
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
            $"{r.StaffID}:{r.AttendanceStatus}:{r.ClockInTime:HH:mm}:{r.ClockOutTime:HH:mm}:{r.ProjectCheckIns}:{r.RegularShift}:{r.AdditionalShifts}:{r.OTHours}:{r.AdditionalOTHours}:{r.TotalFlattenedHours}:{r.IsAllocated}"));

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

            // The observer is global: ignore every workbook that is not a Schedule Manager workbook.
            if (!SheetOperations.IsModuleWorkbook(null, Events.ScheduleSyncDescriptor)) return;

            string name = sheet.Name?.ToString() ?? string.Empty;
            int row = (int)target.Row;
            int col = (int)target.Column;

            if (ScheduleBoardMapper.IsBoardSheet((object)sheet))
            {
                // Top bar: A2 (date) and B2 (day) drive each other. Only decide here; the work is queued for the macro thread.
                dynamic xlApp = ExcelDnaUtil.Application;
                bool hitsDate = Hits(xlApp, target, sheet.Range[ScheduleBoardMapper.DateCell]);
                bool hitsDay = Hits(xlApp, target, sheet.Range[ScheduleBoardMapper.DayCell]);
                if (hitsDate || hitsDay) { QueueDateChange(hitsDate ? DateSource.Date : DateSource.Day); return; }
                RegisterRecordEdit();

                // Clearing Staff Skills / Vehicle Type clears its Staff Name / Vehicle Plate (queued: COM writes belong on the macro thread).
                if (ScheduleBoardMapper.IsClearedHelperCell((object)target, out int clearRow, out int clearCol))
                {
                    object boardObj = sheet;
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { RunFrozen(() => ScheduleBoardMapper.ClearLeadCell(boardObj, clearRow, clearCol)); }
                        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] clearing linked cell failed: {ex.Message}"); }
                        RefreshDerived(reloadContext: false, rollup: true);
                    });
                    return;
                }

                if (ScheduleBoardMapper.IsBoardArea(row, col) || target.Cells.Count > 1) QueueRefresh(reloadContext: false, rollup: true);
                return;
            }

            bool isEntries = EntriesAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));
            bool isAttendance = MasterSheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));

            // C2 (Late Cut-Off) / D2 (OT Start): per-date parameters. Decide here; confirm + refresh on the macro thread.
            if (isAttendance && row <= 2)
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                bool hitsLate = Hits(xlApp, target, sheet.Range[LateCutOffCell]);
                bool hitsOt = Hits(xlApp, target, sheet.Range[OtStartCell]);
                if (hitsLate || hitsOt)
                {
                    object attObj = sheet;
                    ExcelAsyncUtil.QueueAsMacro(() => OnAttendanceParameterChanged(attObj, hitsLate, hitsOt));
                    return;
                }
            }

            if (isEntries || (isAttendance && row >= 4)) RegisterRecordEdit();

            if (isEntries) QueueRefresh(reloadContext: false, rollup: true);
            else if (PreferenceAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) ||
                     name.Equals("Lookups", StringComparison.OrdinalIgnoreCase)) QueueRefresh(reloadContext: true, rollup: false);
            else if (MasterSheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) && row <= 2) QueueRefresh(reloadContext: true, rollup: true);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] OnSheetChanged warning: {ex.Message}");
        }
    }

    /// <summary>
    /// Macro thread. A non-default cut-off / OT start is confirmed first (with a stronger note on past dates);
    /// declining restores the previous values. Accepted changes mark the record edited (past-record guard) and
    /// re-run the rollup. Blank / unreadable input falls back to the previous value.
    /// </summary>
    private static void OnAttendanceParameterChanged(object attendanceObj, bool hitsLate, bool hitsOt)
    {
        try
        {
            object attendanceSheet = attendanceObj;
            TimeOnly prevLate = BoardContext.LateCutOff;
            TimeOnly prevOt = BoardContext.OTStartTime;
            TimeOnly newLate = hitsLate ? BoardContext.ReadTimeCell(attendanceSheet, LateCutOffCell, prevLate) : prevLate;
            TimeOnly newOt = hitsOt ? BoardContext.ReadTimeCell(attendanceSheet, OtStartCell, prevOt) : prevOt;

            if (newLate == prevLate && newOt == prevOt)
            {
                // Same value (or junk / blank): just put the normalized times back.
                RunFrozen(() => WriteAttendanceParameters(attendanceSheet, prevLate, prevOt));
                return;
            }

            var notes = new List<string>();
            if (newLate != prevLate && newLate != BoardContext.DefaultLateCutOff)
                notes.Add($"  Late Cut-Off Time: {newLate:HH:mm} (default {BoardContext.DefaultLateCutOff:HH:mm})");
            if (newOt != prevOt && newOt != BoardContext.DefaultOtStartTime)
                notes.Add($"  OT Start Time: {newOt:HH:mm} (default {BoardContext.DefaultOtStartTime:HH:mm})");

            if (notes.Count > 0)
            {
                bool isPast = _loadedDate.HasValue && _loadedDate.Value < DateOnly.FromDateTime(DateTime.Today);
                string when = _loadedDate.HasValue ? _loadedDate.Value.ToString("dddd, dd MMM yyyy") : "this date";
                string pastNote = isPast
                    ? $"\n\nWARNING: {when} is a PAST record. Saving will recalculate Late status and OT Hours for every worker on that day."
                    : string.Empty;

                PrepareForModal();
                var answer = MessageBox.Show(
                    $"You set a non-default value for {when}:\n\n{string.Join("\n", notes)}{pastNote}\n\nKeep this value?",
                    "Schedule Manager", MessageBoxButtons.YesNo, isPast ? MessageBoxIcon.Warning : MessageBoxIcon.Question);

                if (answer != DialogResult.Yes)
                {
                    RunFrozen(() => WriteAttendanceParameters(attendanceSheet, prevLate, prevOt));
                    SheetOperations.ShowStatusConfirmation("Attendance parameter change reverted.");
                    return;
                }
            }

            RunFrozen(() => WriteAttendanceParameters(attendanceSheet, newLate, newOt));
            RegisterRecordEdit();
            RefreshDerived(reloadContext: true, rollup: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Attendance parameter change failed", ex);
        }
    }

    /// <summary>Rule 2 modal exception: make sure Excel repaints and no dropdown arrow floats over the dialog.</summary>
    private static void PrepareForModal()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        try { if (!(bool)xlApp.ScreenUpdating) xlApp.ScreenUpdating = true; } catch { }
        try { if (!(bool)xlApp.EnableEvents) xlApp.EnableEvents = true; } catch { }
        try { Arrow.HideAll(); } catch { }
    }

    /// <summary>Moves the board to the target date (confirming first) and loads that day.</summary>
    // =========================================================================
    // 6a. TOP BAR: DATE (A2) <-> DAY (B2)
    // =========================================================================
    private enum DateSource { Date, Day }

    private static bool _dateChangeQueued;
    private static DateSource _pendingDateSource;

    private static bool Hits(dynamic xlApp, dynamic target, dynamic range)
    {
        try { return range != null && xlApp.Intersect(target, range) != null; }
        catch { return false; }
    }

    /// <summary>Coalesces A2 / B2 edits into one macro. A date edit wins over a day edit made in the same burst.</summary>
    private static void QueueDateChange(DateSource source)
    {
        if (!_dateChangeQueued || source == DateSource.Date) _pendingDateSource = source;
        if (_dateChangeQueued) return;
        _dateChangeQueued = true;
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _dateChangeQueued = false;
            OnDateCellChanged(_pendingDateSource);
        });
    }

    /// <summary>
    /// B2 holds a weekday. It resolves to that weekday inside the Monday-Sunday week of the date currently shown
    /// (the loaded date), so picking "Friday" while viewing Tuesday 18 Nov moves to Friday 21 Nov.
    /// </summary>
    private static DateOnly? ResolveDayPick(object boardObj)
    {
        dynamic board = boardObj;
        DayOfWeek? picked = ParseDayName((board.Range[ScheduleBoardMapper.DayCell].Value2 as object)?.ToString());
        if (picked == null) return null;

        DateOnly anchor = _loadedDate ?? ScheduleBoardMapper.ReadDate(boardObj) ?? DateOnly.FromDateTime(DateTime.Today);
        return WeekdayInWeekOf(anchor, picked.Value);
    }

    /// <summary>The given weekday inside the Monday-Sunday week that contains <paramref name="anchor"/>.</summary>
    internal static DateOnly WeekdayInWeekOf(DateOnly anchor, DayOfWeek picked)
    {
        DateOnly monday = anchor.AddDays(-(((int)anchor.DayOfWeek + 6) % 7));
        return monday.AddDays(((int)picked + 6) % 7);
    }

    private static DayOfWeek? ParseDayName(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return null;
        string t = text.Trim();
        // Days enum: None = 0, Monday = 1 .. Sunday = 7  =>  modulo 7 is DayOfWeek (Sunday = 0).
        if (Enum.TryParse<Days>(t, true, out var d) && d != Days.None) return (DayOfWeek)((int)d % 7);
        var fmt = CultureInfo.CurrentCulture.DateTimeFormat;
        for (int i = 0; i < 7; i++)
        {
            var dow = (DayOfWeek)i;
            if (t.Equals(fmt.GetDayName(dow), StringComparison.CurrentCultureIgnoreCase) ||
                t.Equals(fmt.GetAbbreviatedDayName(dow), StringComparison.CurrentCultureIgnoreCase)) return dow;
        }
        return null;
    }

    internal static void GoToDate(DateOnly target)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return;
        dynamic? board = SheetMapperHelper.ResolveSheetSafely((object)wb!, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
        if (board == null) return;

        RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board!, target));
        OnDateCellChanged(DateSource.Date);
    }

    // =========================================================================
    // 6a. PAST-RECORD EDIT FAILSAFE
    // =========================================================================
    /// <summary>
    /// Marks the loaded date as edited and, for a date before today, asks once (per workbook + date) whether the user
    /// really wants to change history. Declining discards the edit by reloading the saved copy.
    /// </summary>
    private static void RegisterRecordEdit()
    {
        _hasUnsavedRecordEdits = true;

        DateOnly? loaded = _loadedDate;
        if (loaded == null || loaded.Value >= DateOnly.FromDateTime(DateTime.Today) || _editPromptPending) return;

        DateOnly day = loaded.Value;
        _editPromptPending = true;
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                if (PastRecordGuard.ConfirmPastRecordEdit(GuardSession, day, "Schedule Manager")) return;

                PopulateWorksheet(ServerSyncCommands.FetchSchedule(day));
                SheetOperations.ShowStatusConfirmation($"Edit of past schedule {day:yyyy-MM-dd} discarded.");
            }
            catch (Exception ex) { AppLogger.Error("[ScheduleManager] Past-record edit guard failed", ex); }
            finally { _editPromptPending = false; }
        });
    }

    private static void OnDateCellChanged(DateSource source)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            dynamic? board = SheetMapperHelper.ResolveSheetSafely((object)wb!, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            if (board == null) return;

            DateOnly? date = source == DateSource.Day ? ResolveDayPick((object)board!) : ScheduleBoardMapper.ReadDate((object)board!);
            if (date == null)
            {
                // Unreadable date / unknown weekday: put the loaded date back rather than leave junk in the top bar.
                if (_loadedDate.HasValue) RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board!, _loadedDate.Value));
                SheetOperations.ShowStatusConfirmation(source == DateSource.Day ? "Pick a day of the week (Monday - Sunday)." : "Enter a valid date.");
                return;
            }

            // A2 and B2 always mirror each other. Events are off inside RunFrozen, so this write cannot re-enter the observer.
            RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board!, date.Value));
            if (date == _loadedDate) return;

            var answer = MessageBox.Show(
                $"Load the schedule for {date:dddd, dd MMM yyyy}?\n\nUnsaved changes to {(_loadedDate.HasValue ? _loadedDate.Value.ToString("dd MMM yyyy") : "the current date")} will be discarded.",
                "Schedule Manager", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (answer != DialogResult.Yes)
            {
                if (_loadedDate.HasValue) RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board!, _loadedDate.Value));
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
        object wbObj = wb!;

        dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
        DateOnly? date = board != null ? ScheduleBoardMapper.ReadDate((object)board!) : null;
        if (entriesSheet == null || date == null) return (0, 0, punches.Count);

        DateOnly day = date!.Value;
        var forDate = punches.Where(p => DateOnly.FromDateTime(p.Timestamp) == day).ToList();
        int added = 0, duplicates = 0, unknown = punches.Count - forDate.Count;

        NormalizeApprovals(entriesSheet);
        RunFrozen(() =>
        {
            var existing = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
            var result = DeviceSyncManager.BuildEntries(forDate, existing, id => Roster.IsActiveOn(id, day));
            added = result.NewEntries.Count;
            duplicates = result.Duplicates;
            unknown += result.UnknownStaff;
            if (added == 0) return;

            var all = existing.Concat(result.NewEntries).OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();
            foreach (var e in all) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
            TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, all);
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
