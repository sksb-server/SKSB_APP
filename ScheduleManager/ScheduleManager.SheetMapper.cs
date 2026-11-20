namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
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

    /// <summary>Day Start Time parameter: working-day boundary; punches before it belong to the previous working day.</summary>
    internal const string DayStartCell = "C2";

    /// <summary>Late Checked In Cut Off Time parameter (moved here from the retired 'Daily Salary' sheet).</summary>
    internal const string LateCutOffCell = "D2";

    /// <summary>OT Start Time parameter: day-band time clocked after this counts as overtime.</summary>
    internal const string OtStartCell = "E2";

    private static bool _isRefreshing;
    private static bool _refreshQueued;

    // Fingerprints of the attendance entries as last loaded / saved: PastRecordGuard only flags entries not in this set.
    private static HashSet<string> _loadedEntryFingerprints = new(StringComparer.Ordinal);

    // Fingerprints of entries known to match the PostgreSQL copy (excludes unpersisted Admin auto entries and
    // auto-derived ProjectID changes). Save upserts only entries NOT in this set; without a baseline it upserts all.
    private static HashSet<string> _persistedEntryFingerprints = new(StringComparer.Ordinal);
    private static bool _hasPersistedBaseline;

    // Tracks entries whose ProjectID was auto-derived from the schedule (entry identity -> last auto-derived ProjectID).
    // Manual edits to a different ProjectID stop matching and are preserved; clearing ProjectID re-enables auto-derivation.
    private static readonly Dictionary<string, string> _autoDerivedEntryProjects = new(StringComparer.Ordinal);

    private static string EntryIdentityKey(AttendanceEntry e) =>
        $"{e.Date:yyyy-MM-dd}|{e.Time:HH:mm:ss}|{e.StaffID}|{e.EntryType}|{e.Source}";

    // Full-precision date/time: the default TimeOnly.ToString() drops seconds, which would let a new punch hide behind an
    // existing one in the same minute and be skipped by the delta upsert.
    private static string EntryFingerprint(AttendanceEntry e) =>
        string.Join("|",
            e.Date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture),
            e.Time.ToString("HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture),
            e.StaffID,
            (e.ProjectID ?? string.Empty).Trim().ToUpperInvariant(),
            e.EntryType, e.Source,
            (e.OverrideReason ?? string.Empty).Trim(),
            e.IsApproved,
            (e.WhatsAppPicture ?? string.Empty).Trim());

    internal static IEnumerable<DateOnly> GetModifiedAttendanceEntryDates(IEnumerable<AttendanceEntry> current) =>
        current.Where(e => !_loadedEntryFingerprints.Contains(EntryFingerprint(e))).Select(e => e.Date).Distinct().ToList();

    /// <param name="persistedToServer">
    /// True when <paramref name="entries"/> were committed to PostgreSQL (becomes the delta-upsert baseline);
    /// false for offline saves, which leave the server baseline untouched.
    /// </param>
    internal static void MarkSaved(IEnumerable<AttendanceEntry> entries, bool persistedToServer = true)
    {
        var list = entries as IList<AttendanceEntry> ?? entries.ToList();
        _loadedEntryFingerprints = new HashSet<string>(list.Select(EntryFingerprint), StringComparer.Ordinal);
        _hasUnsavedRecordEdits = false;
        _hasUnsavedGlobalEdits = false;

        if (persistedToServer)
        {
            _persistedEntryFingerprints = new HashSet<string>(list.Select(EntryFingerprint), StringComparer.Ordinal);
            _hasPersistedBaseline = true;
        }
    }

    private static void ClearTableFilter(dynamic sheet, string tableName)
    {
        try
        {
            dynamic? table = TableMapper.ResolveTable(sheet, tableName);
            if (table?.AutoFilter != null && (bool)table.AutoFilter.FilterMode)
                table.AutoFilter.ShowAllData();
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Could not clear filter on {tableName}: {ex.Message}"); }
    }

    private static void EnsureAutoFilterEnabled(dynamic sheet, string tableName)
    {
        try
        {
            dynamic? table = TableMapper.ResolveTable(sheet, tableName);
            if (table == null) return;
            table.ShowAutoFilter = true;
            table.ShowAutoFilterDropDown = true;
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Could not enable AutoFilter on {tableName}: {ex.Message}"); }
    }

    private static void ApplyDateFilter(dynamic sheet, string tableName, DateOnly date)
    {
        try
        {
            dynamic? table = TableMapper.ResolveTable(sheet, tableName);
            if (table == null) return;
            table.ShowAutoFilter = true;
            table.ShowAutoFilterDropDown = true;
            if (table.DataBodyRange == null) return;

            int dateColIndex = -1;
            int idx = 0;
            foreach (dynamic col in table.ListColumns)
            {
                idx++;
                string header = ((string?)col.Name ?? string.Empty).Trim();
                if (header.Equals("Date", StringComparison.OrdinalIgnoreCase)) { dateColIndex = idx; break; }
            }
            if (dateColIndex < 0) return;

            string oa = ((long)date.ToDateTime(TimeOnly.MinValue).ToOADate()).ToString(CultureInfo.InvariantCulture);
            table.Range.AutoFilter(dateColIndex, $">={oa}", 1 /* xlAnd */, $"<={oa}");
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Could not filter {tableName} to {date:yyyy-MM-dd}: {ex.Message}"); }
    }

    /// <summary>Session key for <see cref="PastRecordGuard"/> edit approvals.</summary>
    internal const string GuardSession = "ScheduleManager";

    // True once the user changed a date-bound sheet (board, entries, attendance) since the last load / save.
    // Refresh and populate writes run with events off, so they never set it.
    private static bool _hasUnsavedRecordEdits;
    private static bool _editPromptPending;
    internal static bool HasUnsavedRecordEdits => _hasUnsavedRecordEdits;

    // Edits that are not tied to a single date (preferences / lookups) or rolled-up attendance for the current day.
    private static bool _hasUnsavedGlobalEdits;
    private static bool _isSwitchingDate;
    internal static bool HasUnsavedChanges => _hasUnsavedRecordEdits || _hasUnsavedGlobalEdits;

    /// <summary>Date currently shown on the board; lets a date change be reverted when the user declines.</summary>
    private static DateOnly? _loadedDate;
    internal static DateOnly? LoadedDate => _loadedDate;

    // =========================================================================
    // 2. EXTRACTION & VALIDATION PIPELINE
    // =========================================================================
    internal static (ScheduleRecord? Data, List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ExtractScheduleRecordWithDiagnostics()
    {
        using var _ = LoadingOverlay.BeginScope("Schedule Manager — Validating Schedule", "Reading dispatch board and attendance tables...");
        var errors = new List<ValidationIssue>();
        var warnings = new List<ValidationIssue>();
        static ValidationIssue GeneralError(string message) =>
            new(ValidationSeverity.Error, string.Empty, string.Empty, 0, 0, string.Empty, message);
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<ValidationIssue> { GeneralError("No active workbook found.") }, warnings);
            object wbObj = wb!;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");
            dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (board == null || attendanceSheet == null || entriesSheet == null || prefSheet == null)
            {
                errors.Add(GeneralError("The workbook is missing one of: 'Daily Schedule', 'Daily Attendances', 'Attendance Entries', 'Preference'."));
                return (null, errors, warnings);
            }

            DateOnly? date = ScheduleBoardMapper.ReadDate((object)board!);
            if (date == null)
            {
                var (dateRow, dateCol) = ParseA1Address(ScheduleBoardMapper.DateCell, defaultRow: 2, defaultCol: 2);
                string boardName = ScheduleBoardMapper.DefaultSheetName;
                try { boardName = Convert.ToString(board!.Name) ?? boardName; } catch { }
                errors.Add(new ValidationIssue(
                    ValidationSeverity.Error, boardName, string.Empty, dateRow, dateCol, "Date",
                    $"Date cell {ScheduleBoardMapper.DateCell} is empty or not a valid date."));
                return (null, errors, warnings);
            }

            var record = new ScheduleRecord
            {
                DailyScheduleDate = date.Value,
                DayStartTime = BoardContext.ReadTimeCell((object)attendanceSheet!, DayStartCell, BoardContext.DefaultDayStartTime),
                NextDayStartTime = BoardContext.NextDayStartTime,
                LateCutOffTime =BoardContext.ReadTimeCell((object)attendanceSheet!, LateCutOffCell, BoardContext.DefaultLateCutOff),
                OTStartTime = BoardContext.ReadTimeCell((object)attendanceSheet!, OtStartCell, BoardContext.DefaultOtStartTime)
            };

            // Entries first: default approvals (WhatsApp starts unapproved) must be on the sheet before it is read.
            NormalizeApprovals(entriesSheet);

            LoadingOverlay.Report("Reading preferences, teams & attendance entries...");
            RunFrozen(() =>
            {
                record.Preferences = TableMapper.ExtractFromTable<WorkerPreference>((object)prefSheet!, "WorkerPreferencesTable");
                record.Teams = TableMapper.ExtractFromTable<TeamRecord>((object)prefSheet!, "TeamTable");
                record.Skills = TableMapper.ExtractFromTable<ScheduleSkillRecord>((object)prefSheet!, "SkillsTable");
                record.ProjectSkillRequirements = TableMapper.ExtractFromTable<ProjectSkillRequirement>((object)prefSheet!, "ProjectRequirementsTable");
                ClearTableFilter(entriesSheet, EntriesTableName);
                try
                {
                    record.AttendanceEntries = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
                }
                finally
                {
                    ApplyDateFilter(entriesSheet, EntriesTableName, date.Value);
                }
                var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet!, MasterTableName);
                foreach (var e in existing) e.DailyScheduleDate = date.Value;
                record.WorkerAttendance = existing;
            });
            LoadingOverlay.Report("Reading daily schedule board...");
            record.Projects = ScheduleBoardMapper.Read((object)board!, date.Value);
            RunFrozen(() => SyncEntryProjectsWithSchedule(entriesSheet, record.AttendanceEntries, date.Value, record.Projects, record.DayStartTime, record.NextDayStartTime, record.OTStartTime, resetTracking: false));

            // Delta upsert: only entries that differ from the known server copy. No baseline => null => full upsert.
            record.ModifiedAttendanceEntries = _hasPersistedBaseline
                ? record.AttendanceEntries.Where(e => !_persistedEntryFingerprints.Contains(EntryFingerprint(e))).ToList()
                : null;

            BoardContext.SetSkills(record.Skills);
            ServerSyncCommands.EnsureLicensesLoaded();
            PruneOrphanedTeamAndSkillReferences(prefSheet!, record);
            BoardContext.SetPreferences(record.Preferences, record.Teams);
            BoardContext.SetProjectSkillRequirements(record.ProjectSkillRequirements);
            RunFrozen(() => ServerSyncCommands.SyncSkillList());
            foreach (var t in record.Teams) t.TeamLeaderStaffID = Roster.FindStaff(t.TeamLeader)?.Id ?? 0;

            // Dropdown lists strictly match workers employed on the selected date.
            RunFrozen(() => ServerSyncCommands.SyncStaffLookupLists(date.Value, synchronous: true));

            // ---- table rule validation (required cells, dropdown sources, formats) ----
            var warningDateFilters = new Dictionary<string, Func<DateOnly, TimeOnly?, bool>>(StringComparer.OrdinalIgnoreCase)
            {
                [EntriesTableName] = (entryDate, entryTime) => entryTime.HasValue
                    ? DeviceSyncManager.IsInWorkingDay(entryDate, entryTime.Value, date.Value, record.DayStartTime, record.NextDayStartTime)
                    : entryDate == date.Value
            };
            // These six tables were just healed + deduplicated by TableMapper.ExtractFromTable above: validate only.
            var alreadyHealed = new HashSet<string>(StringComparer.OrdinalIgnoreCase)
            {
                EntriesTableName, MasterTableName, "WorkerPreferencesTable", "TeamTable", "SkillsTable", "ProjectRequirementsTable"
            };
            var (ruleErrors, ruleWarnings) = RulesExtractor.ValidateAllWorkbookTables(wbObj, warningDateFilters, alreadyHealed);
            errors.AddRange(ruleErrors);
            warnings.AddRange(ruleWarnings);

            LoadingOverlay.Report("Validating schedule & worker availability...");
            var scheduleErrors = new List<string>();
            var scheduleWarnings = new List<string>();
            ValidateSchedule(record, scheduleErrors, scheduleWarnings);
            string boardSheetName = ScheduleBoardMapper.DefaultSheetName;
            try { boardSheetName = Convert.ToString(board!.Name) ?? boardSheetName; } catch { }
            errors.AddRange(scheduleErrors.Select(m => ValidationIssue.FromLegacyMessage(ValidationSeverity.Error, m, boardSheetName)));
            warnings.AddRange(scheduleWarnings.Select(m => ValidationIssue.FromLegacyMessage(ValidationSeverity.Warning, m, boardSheetName)));

            // ---- roll punches + board up into the daily attendance rows that get persisted ----
            LoadingOverlay.Report("Rolling up daily worker attendance...");
            var punchesForDate = record.AttendanceEntries
                .Where(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date.Value, record.DayStartTime, record.NextDayStartTime)).ToList();
            record.WorkerAttendance = AttendanceRollupService.Run(date.Value, punchesForDate, record.Projects, record.WorkerAttendance, record.DayStartTime, record.LateCutOffTime, record.OTStartTime, record.NextDayStartTime);

            return (record, errors, warnings);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed extracting schedule record", ex);
            return (null, new List<ValidationIssue> { GeneralError($"Failed extracting schedule: {ex.Message}") }, warnings);
        }
    }

    /// <summary>Parses an A1-style address (e.g. "A2") into 1-based row / column; falls back to the defaults.</summary>
    private static (int Row, int Col) ParseA1Address(string? address, int defaultRow, int defaultCol)
    {
        if (string.IsNullOrWhiteSpace(address)) return (defaultRow, defaultCol);
        string a = address.Replace("$", string.Empty).Trim().ToUpperInvariant();
        int i = 0, col = 0;
        while (i < a.Length && a[i] >= 'A' && a[i] <= 'Z')
        {
            col = col * 26 + (a[i] - 'A' + 1);
            i++;
        }
        return col > 0 && i < a.Length && int.TryParse(a.Substring(i), NumberStyles.Integer, CultureInfo.InvariantCulture, out int row) && row > 0
            ? (row, col)
            : (defaultRow, defaultCol);
    }

    private static void ValidateSchedule(ScheduleRecord record, List<string> errors, List<string> warnings)
    {
        DateOnly scheduleDate = record.DailyScheduleDate;
        var dayEntries = record.AttendanceEntries
            .Where(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, scheduleDate, record.DayStartTime, record.NextDayStartTime)).ToList();

        // Duplicate workers within a shift band block the save.
        errors.AddRange(ScheduleBoardMapper.FindDuplicateWorkers(record.Projects, record.DailyScheduleDate, record.AttendanceEntries));
        errors.AddRange(ScheduleBoardMapper.FindDuplicateProjects(record.Projects));

        foreach (var slot in record.Projects)
        {
            string label = $"{(slot.Shift == Shifts.Night ? "Night" : "Day")} Work slot {slot.SlotIndex + 1}";
            if (slot.Workers.Count > 0 && slot.ProjectID.Length == 0)
                warnings.Add($"  [Daily Schedule - {label}]: workers are assigned but no project code is set.");
            if (slot.ProjectID.Length > 0 && Roster.FindProject(slot.ProjectID) == null)
                warnings.Add($"  [Daily Schedule - {label}]: project '{slot.ProjectID}' was not found in project master.");

            if (slot.PremixHoles > 0 && !Roster.SupportsPremixHoles(slot.ProjectID))
                errors.Add($"  [Daily Schedule - {label}]: Premix Holes is only allowed for Road Reinstatement projects.");

            // Resolution is against ALL attendance staff (a worker who resigned after this date, or who is already on a
            // historical row, resolves cleanly). Office / Admin staff resolve but cannot be scheduled; only field
            // workers employed on this date pass.
            foreach (var worker in slot.Workers.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                var staff = Roster.FindAttendanceStaff(worker);
                if (staff == null)
                    errors.Add($"  [Daily Schedule - {label}]: '{worker}' is not a known field worker (a team name must be expanded first).");
                else if (!Roster.IsFieldDepartment(staff.Department))
                    errors.Add($"  [Daily Schedule - {label}]: '{worker}' is {staff.Department} staff and cannot be assigned to Daily Schedule (only field workers can be scheduled).");
                else if (!Roster.IsActiveOn(staff, scheduleDate))
                    errors.Add($"  [Daily Schedule - {label}]: '{worker}' was not employed on {scheduleDate:yyyy-MM-dd} ({Roster.DescribeInactive(staff, scheduleDate)}).");
            }
            foreach (var plate in slot.Vehicles.Where(p => Roster.FindVehicle(p) == null).Distinct(StringComparer.OrdinalIgnoreCase))
                warnings.Add($"  [Daily Schedule - {label}]: vehicle '{plate}' was not found in vehicle master.");
        }

        foreach (var p in record.Preferences.Where(p => p.StaffID > 0 && Roster.FindStaff(p.StaffID) == null))
            errors.Add($"  [Preference]: Staff ID '{p.StaffID}' is not a known field worker.");
        foreach (var t in record.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamLeader) && Roster.FindStaff(t.TeamLeader) == null))
            errors.Add($"  [Team]: Team leader '{t.TeamLeader}' is not a known field worker.");
        ServerSyncCommands.EnsureLicensesLoaded();
        foreach (var p in record.Preferences)
        {
            var illegal = p.OtherSkills.Where(s => BoardContext.IsLicensedSkill(s))
                .Select(s => s.Trim()).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
            if (illegal.Count > 0)
                errors.Add($"  [Preference]: Staff ID '{p.StaffID}' has licensed skill(s) in 'Other Skills' ({string.Join(", ", illegal)}) — licensed skills are populated automatically in 'Licensed Skills' from Staff Manager.");

            if (!string.IsNullOrWhiteSpace(p.PrimarySkill))
            {
                string primary = p.PrimarySkill.Trim();
                bool held = p.LicensedSkills
                    .Where(s => !s.StartsWith("License ", StringComparison.OrdinalIgnoreCase))
                    .Concat(p.OtherSkills)
                    .Any(s => s.Trim().Equals(primary, StringComparison.OrdinalIgnoreCase));
                var inactive = held ? null : BoardContext.DescribeInactivePrimarySkillLicense(p, scheduleDate);
                if (!held && inactive != null)
                {
                    if (Roster.IsActiveOn(p.StaffID, scheduleDate))
                        warnings.Add($"  [Preference]: Staff ID '{p.StaffID}' has Primary Skill '{primary}', but their {inactive.Value.Detail} — they will not count as a qualified {primary} on {scheduleDate:dd/MM/yyyy}.");
                }
                else if (!held)
                    errors.Add($"  [Preference]: Staff ID '{p.StaffID}' has Primary Skill '{primary}', which is not in their 'Licensed Skills' or 'Other Skills'.");
            }
        }
        if (record.AttendanceEntries.Any(e => e.StaffID <= 0))
            errors.Add("  [Attendance Entries]: Field 'Staff ID' is required and cannot be empty.");
        foreach (var id in dayEntries.Where(e => e.StaffID > 0).Select(e => e.StaffID).Distinct())
        {
            var staff = Roster.FindAttendanceStaff(id);
            if (staff == null)
                warnings.Add($"  [Attendance Entries]: Staff ID '{id}' is not a known staff member; those punches are excluded from the attendance rollup.");
            else if (!Roster.IsActiveOn(staff, scheduleDate))
                warnings.Add($"  [Attendance Entries]: Staff ID '{id}' ({staff.ShortName}) was not employed on {scheduleDate:yyyy-MM-dd} ({Roster.DescribeInactive(staff, scheduleDate)}); those punches are excluded from the attendance rollup.");
        }

        foreach (var g in record.Preferences.Where(p => p.StaffID > 0).GroupBy(p => p.StaffID).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate Staff ID '{g.Key}' in 'WorkerPreferencesTable'.");
        foreach (var g in record.Teams.Where(t => !string.IsNullOrWhiteSpace(t.TeamName)).GroupBy(t => t.TeamName.Trim(), StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate team '{g.Key}' in 'TeamTable'.");
        foreach (var g in record.Skills.Where(s => !string.IsNullOrWhiteSpace(s.SkillName)).GroupBy(s => s.SkillName.Trim(), StringComparer.OrdinalIgnoreCase).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate skill '{g.Key}' in 'SkillsTable'.");
        foreach (var g in record.ProjectSkillRequirements.Where(r => r.ProjectType != ProjectType.None).GroupBy(r => r.ProjectType).Where(g => g.Count() > 1))
            errors.Add($"  [Preference]: Duplicate project type '{g.Key}' in 'ProjectRequirementsTable'.");

        foreach (var g in record.AttendanceEntries.GroupBy(e => (e.Date, e.Time, e.StaffID, e.EntryType, e.Source)).Where(g => g.Count() > 1))
            errors.Add($"  [Attendance Entries]: Staff {g.Key.StaffID} has {g.Count()} identical {g.Key.Source} {g.Key.EntryType} punches at {g.Key.Date:dd/MM/yyyy} {g.Key.Time:HH:mm}.");

        foreach (var day in dayEntries.Where(e => e.IsApproved).GroupBy(e => e.StaffID))
        {
            var ordered = day.OrderBy(e => e.Date).ThenBy(e => e.Time).ToList();
            if (ordered[0].EntryType == AttendanceEntryType.CheckOut)
                warnings.Add($"  [Attendance Entries]: Staff {day.Key} on {scheduleDate:dd/MM/yyyy} starts with a CheckOut and has no earlier CheckIn.");
        }

        int pending = dayEntries.Count(e => !e.IsApproved);
        if (pending > 0)
            warnings.Add($"  [Attendance Entries]: {pending} unapproved entr{(pending == 1 ? "y is" : "ies are")} excluded from the attendance rollup until 'Is Approved' is TRUE.");
    }

    // =========================================================================
    // 2b. PRIMARY SKILL DROPDOWN (per-row: Licensed Skills + Other Skills)
    // =========================================================================
    internal static void RegisterPreferenceDropdowns() =>
        CustomDropdowns.Register(IsPrimarySkillCell, BuildPrimarySkillDropdown);

    private static bool IsPrimarySkillCell(dynamic cell)
    {
        try
        {
            string sheetName = cell.Worksheet.Name?.ToString() ?? string.Empty;
            if (!PreferenceAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase))) return false;

            dynamic? table = cell.ListObject;
            if (table == null) return false;
            if (!string.Equals((string?)table.Name, "WorkerPreferencesTable", StringComparison.OrdinalIgnoreCase)) return false;
            dynamic? body = table.DataBodyRange;
            if (body == null) return false;

            int row = (int)cell.Row;
            int bodyRow = (int)body.Row;
            int bodyRows = (int)body.Rows.Count;
            if (row < bodyRow || row >= bodyRow + bodyRows) return false;

            int colIndex = (int)cell.Column - (int)table.Range.Column + 1;
            string header = ((string?)table.ListColumns[colIndex].Name ?? string.Empty).Trim();
            return header.Equals("Primary Skill", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static CustomDropdownResult? BuildPrimarySkillDropdown(dynamic cell)
    {
        try
        {
            dynamic table = cell.ListObject;
            dynamic sheet = cell.Worksheet;
            int row = (int)cell.Row;
            int licCol = -1, otherCol = -1;
            foreach (dynamic col in table.ListColumns)
            {
                string name = ((string?)col.Name ?? string.Empty).Trim();
                int abs = (int)col.Range.Column;
                if (name.Equals("Licensed Skills", StringComparison.OrdinalIgnoreCase)) licCol = abs;
                else if (name.Equals("Other Skills", StringComparison.OrdinalIgnoreCase)) otherCol = abs;
            }

            string licensed = licCol > 0 ? sheet.Cells[row, licCol].Value2?.ToString() ?? string.Empty : string.Empty;
            string other = otherCol > 0 ? sheet.Cells[row, otherCol].Value2?.ToString() ?? string.Empty : string.Empty;

            var options = new List<DropdownOption>();
            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            void Add(string s, int sec)
            {
                if (!seen.Add(s)) return;
                options.Add(new DropdownOption { Display = s, Raw = s, Section = sec, IsValid = true });
            }

            foreach (var s in RuleEvaluator.SplitMultiSelectTokens(licensed)
                         .Where(s => !s.StartsWith("License ", StringComparison.OrdinalIgnoreCase)))
                Add(s, 0);
            foreach (var s in RuleEvaluator.SplitMultiSelectTokens(other))
                Add(s, 1);

            return new CustomDropdownResult(options, (c, opt) => c.Value2 = opt.Raw);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] Primary Skill dropdown failed: {ex.Message}");
            return null;
        }
    }

    // =========================================================================
    // 2b. PROJECT ID AUTO-DERIVATION (Daily Schedule -> Attendance Entries)
    // =========================================================================
    private static void SyncEntryProjectsWithSchedule(
        dynamic? entriesSheet, List<AttendanceEntry> entries, DateOnly date,
        IReadOnlyList<DailyProjectSchedule> slots, TimeOnly dayStartTime, TimeOnly nextDayStartTime, TimeOnly otStartTime, bool resetTracking)
    {
        if (resetTracking) _autoDerivedEntryProjects.Clear();
        var changedByKey = new Dictionary<string, string>(StringComparer.Ordinal);

        foreach (var e in entries)
        {
            if (e.StaffID <= 0 || !DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date, dayStartTime, nextDayStartTime)) continue;

            string key = EntryIdentityKey(e);
            string current = AttendanceRollupService.NormalizeProjectId(e.ProjectID);
            string scheduled = AttendanceRollupService.ScheduledProjectFor(e.StaffID, e.Date, e.Time, date, slots, dayStartTime, otStartTime);

            if (resetTracking)
            {
                if (current.Length == 0)
                {
                    if (scheduled.Length > 0)
                    {
                        e.ProjectID = scheduled;
                        _autoDerivedEntryProjects[key] = scheduled;
                        changedByKey[key] = scheduled;
                    }
                }
                else if (scheduled.Length > 0 && current.Equals(scheduled, StringComparison.OrdinalIgnoreCase))
                {
                    _autoDerivedEntryProjects[key] = scheduled;
                }
                // else: manual override already persisted; not tracked.
            }
            else
            {
                bool wasAuto = _autoDerivedEntryProjects.TryGetValue(key, out string? lastAuto);
                if (current.Length == 0)
                {
                    if (scheduled.Length > 0)
                    {
                        e.ProjectID = scheduled;
                        _autoDerivedEntryProjects[key] = scheduled;
                        changedByKey[key] = scheduled;
                    }
                    else
                    {
                        _autoDerivedEntryProjects.Remove(key);
                    }
                }
                else if (wasAuto && current.Equals(lastAuto, StringComparison.OrdinalIgnoreCase))
                {
                    if (!current.Equals(scheduled, StringComparison.OrdinalIgnoreCase))
                    {
                        e.ProjectID = scheduled;
                        if (scheduled.Length > 0) _autoDerivedEntryProjects[key] = scheduled;
                        else _autoDerivedEntryProjects.Remove(key);
                        changedByKey[key] = scheduled;
                    }
                }
                else
                {
                    _autoDerivedEntryProjects.Remove(key); // manual override: preserve e.ProjectID
                }
            }
        }

        if (entriesSheet != null && changedByKey.Count > 0)
            WriteDerivedProjectIdsToSheet(entriesSheet, changedByKey);
    }

    private static void WriteDerivedProjectIdsToSheet(dynamic entriesSheet, Dictionary<string, string> changedByKey)
    {
        try
        {
            dynamic? table = TableMapper.ResolveTable(entriesSheet, EntriesTableName);
            dynamic? body = table?.DataBodyRange;
            if (table == null || body == null) return;

            int rows = (int)body!.Rows.Count;
            int cols = (int)table!.ListColumns.Count;
            object[,] header = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, cols);
            int dateCol = -1, timeCol = -1, staffCol = -1, projectCol = -1, typeCol = -1, sourceCol = -1;
            for (int c = 1; c <= cols; c++)
            {
                string h = header[1, c]?.ToString()?.Trim() ?? string.Empty;
                if (h.Equals("Date", StringComparison.OrdinalIgnoreCase)) dateCol = c;
                else if (h.Equals("Time", StringComparison.OrdinalIgnoreCase)) timeCol = c;
                else if (h.Equals("Staff ID", StringComparison.OrdinalIgnoreCase)) staffCol = c;
                else if (h.Equals("Project ID", StringComparison.OrdinalIgnoreCase)) projectCol = c;
                else if (h.Equals("Entry Type", StringComparison.OrdinalIgnoreCase)) typeCol = c;
                else if (h.Equals("Source", StringComparison.OrdinalIgnoreCase)) sourceCol = c;
            }
            if (dateCol < 0 || timeCol < 0 || staffCol < 0 || projectCol < 0 || typeCol < 0 || sourceCol < 0) return;

            object[,] data = TableMapper.GetRange2DValues(body, rows, cols);
            var column = new object?[rows, 1];
            bool dirty = false;
            for (int r = 1; r <= rows; r++)
            {
                column[r - 1, 0] = data[r, projectCol];

                object? dv = data[r, dateCol], tv = data[r, timeCol], sv = data[r, staffCol];
                DateOnly rowDate;
                if (dv is double dd) rowDate = DateOnly.FromDateTime(DateTime.FromOADate(dd));
                else if (dv is DateTime ddt) rowDate = DateOnly.FromDateTime(ddt);
                else if (!DateOnly.TryParse(dv?.ToString(), out rowDate)) continue;

                TimeOnly rowTime;
                if (tv is double td)
                {
                    int secs = (int)Math.Round((td - Math.Floor(td)) * 86400.0);
                    rowTime = TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(secs >= 86400 ? 0 : secs));
                }
                else if (tv is DateTime tdt) rowTime = TimeOnly.FromDateTime(tdt);
                else if (!TimeOnly.TryParse(tv?.ToString(), out rowTime)) continue;

                long staffId;
                if (sv is double sd) staffId = (long)sd;
                else if (!long.TryParse(sv?.ToString(), out staffId)) continue;

                if (!Enum.TryParse<AttendanceEntryType>(data[r, typeCol]?.ToString()?.Trim(), true, out var et)) continue;
                if (!Enum.TryParse<AttendanceSource>(data[r, sourceCol]?.ToString()?.Trim(), true, out var src)) continue;
                if (staffId <= 0) continue;

                string key = $"{rowDate:yyyy-MM-dd}|{rowTime:HH:mm:ss}|{staffId}|{et}|{src}";
                if (changedByKey.TryGetValue(key, out string? newProject))
                {
                    column[r - 1, 0] = newProject;
                    dirty = true;
                }
            }

            if (!dirty) return;
            dynamic target = entriesSheet.Range[body.Cells[1, projectCol], body.Cells[rows, projectCol]];
            target.Value2 = column;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] WriteDerivedProjectIdsToSheet warning: {ex.Message}");
        }
    }

    // =========================================================================
    // 3. APPROVAL DEFAULTS (Thumbprint/Manual/SKSB_App/Admin => TRUE, WhatsApp => FALSE until an admin confirms)
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
        using var _ = LoadingOverlay.BeginScope($"Schedule Manager — Loading {record.DailyScheduleDate:dd MMM yyyy}", "Synchronizing staff & vehicle lookup lists...");
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
                _hasUnsavedGlobalEdits = false;

                // SkillsTable: an empty record (first load / fresh server) keeps the template-seeded rows instead of wiping them.
                bool writeSkills = record.Skills.Count > 0;
                if (!writeSkills && prefSheet != null)
                    record.Skills = TableMapper.ExtractFromTable<ScheduleSkillRecord>((object)prefSheet!, "SkillsTable");
                BoardContext.SetSkills(record.Skills);   // StaffSkillList below includes these
                bool writeProjectSkillReqs = record.ProjectSkillRequirements.Count > 0;
                if (!writeProjectSkillReqs && prefSheet != null)
                    record.ProjectSkillRequirements = TableMapper.ExtractFromTable<ProjectSkillRequirement>((object)prefSheet!, "ProjectRequirementsTable");
                BoardContext.SetProjectSkillRequirements(record.ProjectSkillRequirements);
                ServerSyncCommands.SyncLookupLists(date);

                // Preferences & teams (names are UI-only helpers resolved from the roster)
                // Preferences are global, not dated: resigned field workers are kept visible and preserved on save.
                foreach (var p in record.Preferences) p.StaffName = Roster.NameOf(p.StaffID, p.StaffName);
                foreach (var t in record.Teams)
                    if (t.TeamLeaderStaffID > 0) t.TeamLeader = Roster.NameOf(t.TeamLeaderStaffID, t.TeamLeader);

                LoadingOverlay.Report("Populating teams & worker preferences...");
                if (prefSheet != null)
                {
                    if (writeSkills)
                        TableMapper.PopulateTableSafe<ScheduleSkillRecord>((object)prefSheet!, "SkillsTable", record.Skills);
                    if (writeProjectSkillReqs)
                        TableMapper.PopulateTableSafe<ProjectSkillRequirement>((object)prefSheet!, "ProjectRequirementsTable", record.ProjectSkillRequirements);
                    TableMapper.PopulateTableSafe<TeamRecord>((object)prefSheet!, "TeamTable", record.Teams);
                    TableMapper.PopulateTableSafe<WorkerPreference>((object)prefSheet!, "WorkerPreferencesTable", record.Preferences);
                }

                // The loaded date's day start / cut-off / OT start go on the sheet first so Reload (and the rollup below) use them.
                if (attendanceSheet != null) WriteAttendanceParameters((object)attendanceSheet!, record.DayStartTime, record.LateCutOffTime, record.OTStartTime);

                // Context (requirements, icons, day start, cut-off, OT start) comes from the workbook; prefs/teams from the record just loaded.
                BoardContext.Reload(wbObj);
                BoardContext.SetNextDayStartTime(record.NextDayStartTime);
                BoardContext.SetPreferences(record.Preferences, record.Teams);

                // Server baseline for the delta upsert: captured BEFORE ProjectID auto-derivation and Admin auto entries,
                // so both still count as changes on the next save. Offline / pending-upload data is never a server baseline.
                if (ServerEnvironment.IsOnline && !OfflineSnapshot.HasPendingSnapshot<ScheduleRecord>())
                {
                    _persistedEntryFingerprints = new HashSet<string>(record.AttendanceEntries.Select(EntryFingerprint), StringComparer.Ordinal);
                    _hasPersistedBaseline = true;
                }
                else
                {
                    _persistedEntryFingerprints.Clear();
                    _hasPersistedBaseline = false;
                }

                SyncEntryProjectsWithSchedule(null, record.AttendanceEntries, date, record.Projects, BoardContext.DayStartTime, BoardContext.NextDayStartTime, BoardContext.OTStartTime, resetTracking: true);

                // Active Admin staff with no punch that working day get approved auto CheckIn/CheckOut entries.
                // They are fingerprinted as loaded below, and _hasUnsavedGlobalEdits stays false, so merely browsing
                // dates never triggers a save; they persist with the next real save of this date.
                var adminAutoEntries = AttendanceRollupService.BuildAdminAutoEntries(date, record.AttendanceEntries, BoardContext.DayStartTime, BoardContext.OTStartTime, BoardContext.NextDayStartTime);
                if (adminAutoEntries.Count > 0) record.AttendanceEntries.AddRange(adminAutoEntries);

                var entries = record.AttendanceEntries
                    .OrderByDescending(e => e.Date).ThenByDescending(e => e.Time).ToList();
                var dayEntries = entries
                    .Where(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date, BoardContext.DayStartTime, BoardContext.NextDayStartTime)).ToList();
                foreach (var e in entries) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
                _loadedEntryFingerprints = new HashSet<string>(entries.Select(EntryFingerprint), StringComparer.Ordinal);
                LoadingOverlay.Report("Populating attendance entries...");
                if (entriesSheet != null)
                {
                    ClearTableFilter(entriesSheet, EntriesTableName);
                    TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, entries);
                    ApplyDateFilter(entriesSheet, EntriesTableName, date);
                }

                LoadingOverlay.Report("Rendering daily schedule board...");
                ScheduleBoardMapper.Render((object)board!, date, record.Projects, dayEntries, autoPopulate: true);

                if (attendanceSheet != null)
                {
                    LoadingOverlay.Report("Rolling up daily attendances...");
                    var rows = AttendanceRollupService.Run(date,
                        dayEntries,
                        record.Projects, record.WorkerAttendance, BoardContext.DayStartTime, BoardContext.LateCutOff, BoardContext.OTStartTime, BoardContext.NextDayStartTime);
                    WriteAttendanceRows(attendanceSheet, rows);
                }
            });
            RibbonController.RefreshWorkingDayDisplay();
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed populating schedule workbook", ex);
        }
    }

    /// <summary>
    /// Lightweight post-save refresh: re-renders the board and rewrites <c>DailyAttendanceTable</c> from the saved record.
    /// Deliberately leaves <c>AttendanceEntriesTable</c> (10k+ historical rows, already current on the sheet) and the
    /// Preference tables untouched.
    /// </summary>
    internal static void RefreshAfterSave(ScheduleRecord record)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            object wbObj = wb!;

            dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
            dynamic? attendanceSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, "Daily Attendances");

            RunFrozen(() =>
            {
                var dayEntries = record.AttendanceEntries
                    .Where(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, record.DailyScheduleDate, record.DayStartTime, record.NextDayStartTime))
                    .ToList();

                if (board != null) ScheduleBoardMapper.Render((object)board!, record.DailyScheduleDate, record.Projects, dayEntries, autoPopulate: true);
                if (attendanceSheet != null) WriteAttendanceRows(attendanceSheet, record.WorkerAttendance);
            });
            RibbonController.RefreshWorkingDayDisplay();
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Failed refreshing schedule workbook after save", ex);
        }
    }

    /// <summary>
    /// Writes the per-date Day Start (C2), Late Cut-Off (D2) and OT Start (E2) times as Excel time serials in strict
    /// 24-hour "hh:mm". Call inside RunFrozen.
    /// </summary>
    private static void WriteAttendanceParameters(dynamic attendanceSheet, TimeOnly dayStart, TimeOnly lateCutOff, TimeOnly otStart)
    {
        WriteTimeParameter(attendanceSheet, DayStartCell, dayStart);
        WriteTimeParameter(attendanceSheet, LateCutOffCell, lateCutOff);
        WriteTimeParameter(attendanceSheet, OtStartCell, otStart);
    }

    private static void WriteTimeParameter(dynamic attendanceSheet, string cellAddress, TimeOnly value)
    {
        try
        {
            dynamic cell = attendanceSheet.Range[cellAddress];
            cell.Value2 = value.ToTimeSpan().TotalDays;
            cell.NumberFormat = "hh:mm";
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Writing {cellAddress} failed: {ex.Message}"); }
    }

    /// <summary>Re-evaluates WorkerPreferencesTable[Licensed Skills] from the cached StaffManager licences (bulk read/write).</summary>
    private static void RefreshLicensedSkills(object wbObj, DateOnly scheduleDate)
    {
        try
        {
            ServerSyncCommands.EnsureLicensesLoaded();

            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, PreferenceAliases, "Preference");
            if (prefSheet == null) return;

            dynamic? table = null;
            try { table = prefSheet.ListObjects["WorkerPreferencesTable"]; } catch { }
            if (table == null || table.DataBodyRange == null) return;

            dynamic? idCol = null, nameCol = null, licCol = null;
            foreach (dynamic col in table.ListColumns)
            {
                string header = ((string?)col.Name ?? string.Empty).Trim();
                if (header.Equals("Staff ID", StringComparison.OrdinalIgnoreCase)) idCol = col;
                else if (header.Equals("Staff Name", StringComparison.OrdinalIgnoreCase)) nameCol = col;
                else if (header.Equals("Licensed Skills", StringComparison.OrdinalIgnoreCase)) licCol = col;
            }
            if (idCol == null || licCol == null || idCol.DataBodyRange == null || licCol.DataBodyRange == null) return;

            int rowCount = (int)table.DataBodyRange.Rows.Count;
            if (rowCount <= 0) return;

            object[,] ids = TableMapper.GetRange2DValues(idCol.DataBodyRange, rowCount, 1);
            object[,] lics = TableMapper.GetRange2DValues(licCol.DataBodyRange, rowCount, 1);
            object[,]? names = nameCol != null && nameCol.DataBodyRange != null
                ? TableMapper.GetRange2DValues(nameCol.DataBodyRange, rowCount, 1)
                : null;

            var skillsByStaff = ServerSyncCommands.BuildLicensedSkillsMap(BoardContext.VehicleLicenses, BoardContext.SkillLicenses, scheduleDate);
            var grid = new object[rowCount, 1];
            bool anyChanged = false;

            for (int r = 1; r <= rowCount; r++)
            {
                object? rawId = ids[r, 1];
                long staffId = 0;
                if (rawId is double d) staffId = (long)d;
                else if (rawId != null) long.TryParse(rawId.ToString()?.Trim(), out staffId);

                if (staffId <= 0 && names != null)
                {
                    string staffName = (names[r, 1]?.ToString() ?? string.Empty).Trim();
                    if (staffName.Length > 0) staffId = Roster.FindStaff(staffName)?.Id ?? 0;
                }

                string expected = staffId > 0 && skillsByStaff.TryGetValue(staffId, out var list) ? string.Join(", ", list) : string.Empty;
                string current = (lics[r, 1]?.ToString() ?? string.Empty).Trim();
                grid[r - 1, 0] = expected;
                if (!string.Equals(current, expected, StringComparison.Ordinal)) anyChanged = true;
            }

            if (anyChanged) licCol.DataBodyRange.Value2 = grid;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] RefreshLicensedSkills failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Strips Team / Skill references that no longer exist in TeamTable / SkillsTable (and licensed skills) from
    /// WorkerPreferencesTable and ProjectRequirementsTable, both in the record and on the sheet (bulk column writes).
    /// </summary>
    private static void PruneOrphanedTeamAndSkillReferences(dynamic prefSheet, ScheduleRecord record)
    {
        try
        {
            var validTeams = new HashSet<string>(
                record.Teams.Select(t => (t.TeamName ?? string.Empty).Trim()).Where(n => n.Length > 0),
                StringComparer.OrdinalIgnoreCase);
            bool pruneTeams = validTeams.Count > 0;
            var validSkills = new HashSet<string>(BoardContext.AllConfiguredSkills(), StringComparer.OrdinalIgnoreCase);
            bool pruneSkills = record.Skills.Any(s => !string.IsNullOrWhiteSpace(s.SkillName));
            if (!pruneTeams && !pruneSkills) return;

            int removed = 0;

            // ---- in-memory record ----
            foreach (var p in record.Preferences)
            {
                if (pruneTeams) removed += p.Team.RemoveAll(s => !validTeams.Contains(s.Trim()));
                if (pruneSkills)
                {
                    removed += p.OtherSkills.RemoveAll(s => !validSkills.Contains(s.Trim()));
                    removed += p.BlacklistedSkills.RemoveAll(s => !validSkills.Contains(s.Trim()));
                    if (!string.IsNullOrWhiteSpace(p.PrimarySkill) && !validSkills.Contains(p.PrimarySkill.Trim()))
                    {
                        p.PrimarySkill = string.Empty;
                        removed++;
                    }
                }
            }
            if (pruneSkills)
                foreach (var r in record.ProjectSkillRequirements)
                    removed += r.Skills.RemoveAll(s => !validSkills.Contains(s.Trim()));

            // ---- worksheet (bulk 2D column writes) ----
            int sheetRemoved = 0;
            RunFrozen(() =>
            {
                if (pruneTeams) sheetRemoved += PruneMultiSelectColumn(prefSheet, "WorkerPreferencesTable", "Team", validTeams);
                if (pruneSkills)
                {
                    sheetRemoved += PruneMultiSelectColumn(prefSheet, "WorkerPreferencesTable", "Other Skills", validSkills);
                    sheetRemoved += PruneMultiSelectColumn(prefSheet, "WorkerPreferencesTable", "Blacklisted Skills", validSkills);
                    sheetRemoved += PruneSingleValueColumn(prefSheet, "WorkerPreferencesTable", "Primary Skill", validSkills);
                    sheetRemoved += PruneMultiSelectColumn(prefSheet, "ProjectRequirementsTable", "Required Skills", validSkills);
                }
            });

            removed = Math.Max(removed, sheetRemoved);
            if (removed > 0)
            {
                AppLogger.Info($"[ScheduleManager] Cleaned {removed} orphaned team/skill reference(s) from Preference tables.");
                SheetOperations.ShowStatusConfirmation($"Cleaned {removed} orphaned team/skill reference(s) from Preference tables.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] PruneOrphanedTeamAndSkillReferences failed: {ex.Message}");
        }
    }

    private static dynamic? GetTableColumnBody(dynamic sheet, string tableName, string header, out int rowCount)
    {
        rowCount = 0;
        dynamic? table = null;
        try { table = sheet.ListObjects[tableName]; } catch { }
        if (table == null) return null;
        dynamic tbl = table!;
        if (tbl.DataBodyRange == null) return null;
        foreach (dynamic col in tbl.ListColumns)
        {
            string name = ((string?)col.Name ?? string.Empty).Trim();
            if (!name.Equals(header, StringComparison.OrdinalIgnoreCase)) continue;
            if (col.DataBodyRange == null) return null;
            rowCount = (int)tbl.DataBodyRange.Rows.Count;
            return col.DataBodyRange;
        }
        return null;
    }

    private static int PruneMultiSelectColumn(dynamic sheet, string tableName, string header, HashSet<string> valid)
    {
        try
        {
            dynamic? body = GetTableColumnBody(sheet, tableName, header, out int rowCount);
            if (body == null || rowCount <= 0) return 0;
            object[,] values = TableMapper.GetRange2DValues(body, rowCount, 1);
            var grid = new object[rowCount, 1];
            int removed = 0;
            bool anyChanged = false;
            for (int r = 1; r <= rowCount; r++)
            {
                object? raw = values[r, 1];
                string current = raw?.ToString() ?? string.Empty;
                var tokens = RuleEvaluator.SplitMultiSelectTokens(current);
                var kept = tokens.Where(t => valid.Contains(t.Trim())).ToList();
                if (kept.Count != tokens.Count)
                {
                    removed += tokens.Count - kept.Count;
                    anyChanged = true;
                    grid[r - 1, 0] = string.Join(", ", kept);
                }
                else grid[r - 1, 0] = raw ?? string.Empty;
            }
            if (anyChanged) ((dynamic)body!).Value2 = grid;
            return removed;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] Pruning {tableName}[{header}] failed: {ex.Message}");
            return 0;
        }
    }

    private static int PruneSingleValueColumn(dynamic sheet, string tableName, string header, HashSet<string> valid)
    {
        try
        {
            dynamic? body = GetTableColumnBody(sheet, tableName, header, out int rowCount);
            if (body == null || rowCount <= 0) return 0;
            object[,] values = TableMapper.GetRange2DValues(body, rowCount, 1);
            var grid = new object[rowCount, 1];
            int removed = 0;
            for (int r = 1; r <= rowCount; r++)
            {
                object? raw = values[r, 1];
                string current = (raw?.ToString() ?? string.Empty).Trim();
                if (current.Length > 0 && !valid.Contains(current))
                {
                    grid[r - 1, 0] = string.Empty;
                    removed++;
                }
                else grid[r - 1, 0] = raw ?? string.Empty;
            }
            if (removed > 0) ((dynamic)body!).Value2 = grid;
            return removed;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] Pruning {tableName}[{header}] failed: {ex.Message}");
            return 0;
        }
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
    /// With <paramref name="ensureAdminEntries"/> (explicit Roll Up) the roster is reloaded and missing Admin auto
    /// CheckIn / CheckOut entries are added; newly generated entries mark the day as unsaved so a date switch persists them.
    /// </summary>
    internal static void RefreshDerived(bool reloadContext, bool rollup, bool ensureAdminEntries = false)
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
                if (ensureAdminEntries)
                {
                    Roster.Invalidate();
                    ServerSyncCommands.SyncStaffLookupLists(date.Value, synchronous: true);
                    // The next date's Day Start Time (end of this working day) may have been edited elsewhere.
                    if (ServerEnvironment.IsOnline)
                        BoardContext.SetNextDayStartTime(ServerSyncCommands.FetchDayStartTime(date.Value.AddDays(1)));
                }

                if (reloadContext)
                {
                    RefreshLicensedSkills(wbObj, date.Value);
                    BoardContext.Reload(wbObj);
                    // A SkillsTable edit must show up in the StaffSkillList dropdowns straight away.
                    ServerSyncCommands.SyncSkillList();
                }

                var slots = ScheduleBoardMapper.Read((object)board!, date.Value);
                if (ScheduleBoardMapper.ExpandTeams(slots, out string? notice) && notice != null)
                    SheetOperations.ShowStatusConfirmation(notice);

                var entries = entriesSheet != null
                    ? TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName)
                    : new List<AttendanceEntry>();

                if (ensureAdminEntries && rollup && entriesSheet != null)
                {
                    var adminAutoEntries = AttendanceRollupService.BuildAdminAutoEntries(date.Value, entries, BoardContext.DayStartTime, BoardContext.OTStartTime, BoardContext.NextDayStartTime);
                    if (adminAutoEntries.Count > 0)
                    {
                        entries.AddRange(adminAutoEntries);
                        entries = entries.OrderByDescending(e => e.Date).ThenByDescending(e => e.Time).ToList();
                        foreach (var e in entries) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
                        ClearTableFilter(entriesSheet, EntriesTableName);
                        TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, entries);
                        ApplyDateFilter(entriesSheet, EntriesTableName, date.Value);
                        foreach (var e in adminAutoEntries) _loadedEntryFingerprints.Add(EntryFingerprint(e));
                        // Generated entries are not in the database yet: persist them on the next date switch.
                        _hasUnsavedGlobalEdits = true;
                    }
                }

                SyncEntryProjectsWithSchedule(entriesSheet, entries, date.Value, slots, BoardContext.DayStartTime, BoardContext.NextDayStartTime, BoardContext.OTStartTime, resetTracking: false);

                ScheduleBoardMapper.Render((object)board!, date.Value, slots, entries);

                if (rollup && attendanceSheet != null)
                {
                    var existing = TableMapper.ExtractFromTable<DailyWorkerAttendance>((object)attendanceSheet!, MasterTableName);
                    var rows = AttendanceRollupService.Run(date.Value,
                        entries.Where(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, date.Value, BoardContext.DayStartTime, BoardContext.NextDayStartTime)).ToList(),
                        slots, existing, BoardContext.DayStartTime, BoardContext.LateCutOff, BoardContext.OTStartTime, BoardContext.NextDayStartTime);
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

                if ((int)target.Cells.Count == 1 && ScheduleBoardMapper.IsPremixCell((object)sheet, row, col))
                {
                    object premixObj = sheet; int premixRow = row, premixCol = col;
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        string? notice = ScheduleBoardMapper.GetRejectedPremixNotice(premixObj, premixRow, premixCol);
                        RefreshDerived(reloadContext: false, rollup: true);
                        if (notice != null) SheetOperations.ShowStatusConfirmation(notice);
                    });
                    return;
                }

                if (ScheduleBoardMapper.IsBoardArea(row, col) || target.Cells.Count > 1) QueueRefresh(reloadContext: false, rollup: true);
                return;
            }

            bool isEntries = EntriesAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));
            bool isAttendance = MasterSheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase));

            // C2 (Day Start) / D2 (Late Cut-Off) / E2 (OT Start): per-date parameters. Decide here; confirm + refresh on the macro thread.
            if (isAttendance && row <= 2)
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                bool hitsDayStart = Hits(xlApp, target, sheet.Range[DayStartCell]);
                bool hitsLate = Hits(xlApp, target, sheet.Range[LateCutOffCell]);
                bool hitsOt = Hits(xlApp, target, sheet.Range[OtStartCell]);
                if (hitsDayStart || hitsLate || hitsOt)
                {
                    object attObj = sheet;
                    ExcelAsyncUtil.QueueAsMacro(() => OnAttendanceParameterChanged(attObj, hitsDayStart, hitsLate, hitsOt));
                    return;
                }
            }

            if (isEntries || (isAttendance && row >= 4)) RegisterRecordEdit();

            if (isEntries) QueueRefresh(reloadContext: false, rollup: true);
            else if (PreferenceAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) ||
                     name.Equals("Lookups", StringComparison.OrdinalIgnoreCase))
            {
                _hasUnsavedGlobalEdits = true;
                QueueRefresh(reloadContext: true, rollup: false);
            }
            else if (MasterSheetAliases.Any(a => a.Equals(name, StringComparison.OrdinalIgnoreCase)) && row <= 2) QueueRefresh(reloadContext: true, rollup: true);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager] OnSheetChanged warning: {ex.Message}");
        }
    }

    /// <summary>
    /// Macro thread. A non-default day start / cut-off / OT start is confirmed first (with a stronger note on past dates);
    /// declining restores the previous values. Accepted changes mark the record edited (past-record guard) and
    /// re-run the rollup. Blank / unreadable input falls back to the previous value.
    /// </summary>
    private static void OnAttendanceParameterChanged(object attendanceObj, bool hitsDayStart, bool hitsLate, bool hitsOt)
    {
        try
        {
            object attendanceSheet = attendanceObj;
            TimeOnly prevDayStart = BoardContext.DayStartTime;
            TimeOnly prevLate = BoardContext.LateCutOff;
            TimeOnly prevOt = BoardContext.OTStartTime;
            TimeOnly newDayStart = hitsDayStart ? BoardContext.ReadTimeCell(attendanceSheet, DayStartCell, prevDayStart) : prevDayStart;
            TimeOnly newLate = hitsLate ? BoardContext.ReadTimeCell(attendanceSheet, LateCutOffCell, prevLate) : prevLate;
            TimeOnly newOt = hitsOt ? BoardContext.ReadTimeCell(attendanceSheet, OtStartCell, prevOt) : prevOt;

            if (newDayStart == prevDayStart && newLate == prevLate && newOt == prevOt)
            {
                // Same value (or junk / blank): just put the normalized times back.
                RunFrozen(() => WriteAttendanceParameters(attendanceSheet, prevDayStart, prevLate, prevOt));
                return;
            }

            var notes = new List<string>();
            if (newDayStart != prevDayStart && newDayStart != BoardContext.DefaultDayStartTime)
                notes.Add($"  Day Start Time: {newDayStart:HH:mm} (default {BoardContext.DefaultDayStartTime:HH:mm})");
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
                    RunFrozen(() => WriteAttendanceParameters(attendanceSheet, prevDayStart, prevLate, prevOt));
                    SheetOperations.ShowStatusConfirmation("Attendance parameter change reverted.");
                    return;
                }
            }

            RunFrozen(() => WriteAttendanceParameters(attendanceSheet, newDayStart, newLate, newOt));
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

        try { SwitchToDate((object)board!, target); }
        catch (Exception ex) { AppLogger.Error("[ScheduleManager] Date change failed", ex); }
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

            if (date == _loadedDate)
            {
                // A2 and B2 always mirror each other. Events are off inside RunFrozen, so this write cannot re-enter the observer.
                RunFrozen(() => ScheduleBoardMapper.WriteDate((object)board!, date.Value));
                return;
            }

            SwitchToDate((object)board!, date.Value);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ScheduleManager] Date change failed", ex);
        }
    }

    /// <summary>
    /// Switches the board to <paramref name="target"/>. Unsaved changes on the outgoing date are saved first through the
    /// normal save pipeline (errors / warnings / past-record guard / offline conflicts); a blocked save keeps the outgoing date.
    /// </summary>
    private static void SwitchToDate(object boardObj, DateOnly target)
    {
        if (_isSwitchingDate) return;
        _isSwitchingDate = true;
        try
        {
            if (target == _loadedDate)
            {
                RunFrozen(() => ScheduleBoardMapper.WriteDate(boardObj, target));
                return;
            }

            DateOnly? previousDate = _loadedDate;
            bool savedPrevious = false;

            if (HasUnsavedChanges && previousDate.HasValue)
            {
                // Extraction reads A2/B2, so put the outgoing date back first or it would be saved under the new date.
                RunFrozen(() => ScheduleBoardMapper.WriteDate(boardObj, previousDate.Value));

                bool saved = SaveEntry.TrySave(Events.ScheduleSyncDescriptor, showSuccessNotification: false, repopulateAfterSave: false);
                if (!saved)
                {
                    RunFrozen(() => ScheduleBoardMapper.WriteDate(boardObj, previousDate.Value));
                    SheetOperations.ShowStatusConfirmation($"Could not switch to {target:yyyy-MM-dd}; stayed on {previousDate.Value:yyyy-MM-dd}.");
                    return;
                }
                savedPrevious = true;
            }

            using (LoadingOverlay.BeginScope($"Schedule Manager — Loading {target:dd MMM yyyy}", "Fetching schedule from system..."))
            {
                var record = ServerSyncCommands.FetchSchedule(target);
                PopulateWorksheet(record);
            }
            SheetOperations.ShowStatusConfirmation(savedPrevious && previousDate.HasValue
                ? $"Saved {previousDate.Value:yyyy-MM-dd} and loaded {target:yyyy-MM-dd}."
                : $"Schedule for {target:yyyy-MM-dd} loaded.");
        }
        finally { _isSwitchingDate = false; }
    }

    // =========================================================================
    // 6b. THUMBPRINT TERMINAL APPEND
    // =========================================================================
    /// <summary>
    /// Appends terminal punches for the board date to the entries table (Source = Thumbprint), skipping punches
    /// already present. Must run on the macro thread. Returns (added, duplicates, unknown or other-date).
    /// Punches from any attendance staff (Worker, Office, Admin) are accepted.
    /// </summary>
    internal static (int Added, int Duplicates, int Unknown) AppendTerminalPunches(IReadOnlyList<TerminalPunch> punches, DateOnly? targetDate = null)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return (0, 0, punches.Count);
        object wbObj = wb!;

        dynamic? board = SheetMapperHelper.ResolveSheetSafely(wbObj, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, EntriesAliases, "Attendance Entries");
        DateOnly? date = targetDate ?? (board != null ? ScheduleBoardMapper.ReadDate((object)board!) : null);
        if (entriesSheet == null || date == null) return (0, 0, punches.Count);

        DateOnly day = date!.Value;
        TimeOnly dayStart = DeviceSyncManager.DefaultDayStartTime;
        int added = 0, duplicates = 0, unknown = 0;
        bool wrote = false;

        NormalizeApprovals(entriesSheet);
        RunFrozen(() =>
        {
            var existing = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
            var (clean, removedExistingDuplicates) = DeviceSyncManager.DeduplicateExisting(existing, dayStart);
            var result = DeviceSyncManager.BuildEntries(punches, clean, id => Roster.FindAttendanceStaff(id) != null, dayStart);
            added = result.NewEntries.Count;
            duplicates = result.Duplicates + removedExistingDuplicates;
            unknown += result.UnknownStaff;
            if (added == 0 && removedExistingDuplicates == 0) { ApplyDateFilter(entriesSheet, EntriesTableName, day); return; }

            var all = clean.Concat(result.NewEntries).OrderByDescending(e => e.Date).ThenByDescending(e => e.Time).ToList();
            ServerSyncCommands.SyncLookupLists(day);
            foreach (var e in all) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
            ClearTableFilter(entriesSheet, EntriesTableName);
            TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, all);
            ApplyDateFilter(entriesSheet, EntriesTableName, day);
            wrote = true;
        });

        if (wrote)
        {
            RegisterRecordEdit();
            RefreshDerived(reloadContext: false, rollup: true);
        }
        return (added, duplicates, unknown);
    }

    /// <summary>Thumbprint/any entries currently on the Attendance Entries sheet. Macro thread only.</summary>
    internal static List<AttendanceEntry> ReadCurrentSheetEntries()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return new List<AttendanceEntry>();
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely((object)wb!, EntriesAliases, "Attendance Entries");
        if (entriesSheet == null) return new List<AttendanceEntry>();
        return TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
    }

    /// <summary>
    /// Merges already-persisted terminal entries (all dates) into the open sheet and filters the Date column to
    /// <paramref name="selectedDate"/>. Does not mark the record edited (rows are already in the database). Macro thread only.
    /// Returns (rows added in total, rows added within the selected working day).
    /// </summary>
    internal static (int TotalAdded, int AddedForSelectedDate) MergePersistedEntries(IReadOnlyList<AttendanceEntry> persistedEntries, DateOnly selectedDate)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return (0, 0);
        dynamic? entriesSheet = SheetMapperHelper.ResolveSheetSafely((object)wb!, EntriesAliases, "Attendance Entries");
        if (entriesSheet == null) return (0, 0);

        if (persistedEntries.Count == 0)
        {
            RunFrozen(() => ApplyDateFilter(entriesSheet, EntriesTableName, selectedDate));
            return (0, 0);
        }

        int totalAdded = 0, addedForDate = 0;
        bool wrote = false;
        NormalizeApprovals(entriesSheet);
        RunFrozen(() =>
        {
            var existing = TableMapper.ExtractFromTable<AttendanceEntry>((object)entriesSheet!, EntriesTableName);
            var (clean, removedExistingDuplicates) = DeviceSyncManager.DeduplicateExisting(existing, DeviceSyncManager.DefaultDayStartTime);
            var toAdd = persistedEntries.Where(n => !clean.Any(c =>
                (c.Date == n.Date && c.Time == n.Time && c.StaffID == n.StaffID && c.EntryType == n.EntryType && c.Source == n.Source) ||
                (c.Source == AttendanceSource.Thumbprint && n.Source == AttendanceSource.Thumbprint && c.StaffID == n.StaffID &&
                 Math.Abs((c.Date.ToDateTime(c.Time) - n.Date.ToDateTime(n.Time)).TotalSeconds) <= 60))).ToList();
            if (toAdd.Count == 0 && removedExistingDuplicates == 0)
            {
                ApplyDateFilter(entriesSheet, EntriesTableName, selectedDate);
                return;
            }

            var all = clean.Concat(toAdd).OrderByDescending(e => e.Date).ThenByDescending(e => e.Time).ToList();
            ServerSyncCommands.SyncLookupLists(selectedDate);
            foreach (var e in all) e.StaffName = Roster.NameOf(e.StaffID, e.StaffName);
            ClearTableFilter(entriesSheet, EntriesTableName);
            TableMapper.PopulateTableSafe<AttendanceEntry>((object)entriesSheet!, EntriesTableName, all);
            ApplyDateFilter(entriesSheet, EntriesTableName, selectedDate);
            // Already persisted: keep PastRecordGuard from treating them as unsaved manual edits.
            foreach (var e in toAdd)
            {
                string fp = EntryFingerprint(e);
                _loadedEntryFingerprints.Add(fp);
                _persistedEntryFingerprints.Add(fp);
            }
            totalAdded = toAdd.Count;
            addedForDate = toAdd.Count(e => DeviceSyncManager.IsInWorkingDay(e.Date, e.Time, selectedDate, BoardContext.DayStartTime, BoardContext.NextDayStartTime));
            wrote = true;
        });

        if (wrote) RefreshDerived(reloadContext: false, rollup: true);
        // Rolled-up attendance for the current day changed: persist on date switch without tripping PastRecordGuard.
        if (addedForDate > 0) _hasUnsavedGlobalEdits = true;
        return (totalAdded, addedForDate);
    }

    /// <summary>Called after the ribbon's "Remove Duplicates" deleted rows: mark the record edited and rebuild the derived views.</summary>
    internal static void OnDuplicateRowsCleaned()
    {
        RegisterRecordEdit();
        RefreshDerived(reloadContext: true, rollup: true);
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
