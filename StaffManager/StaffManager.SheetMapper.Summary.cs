namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;

public static partial class SheetMapper
{
    // =========================================================================
    // 1. CONSTANTS & CACHED STATE
    // =========================================================================
    private static readonly string[] SummarySheetAliases = { "Staff Summary", "StaffSummary" };
    private const string SummaryTableName = "StaffSummaryTable";
    private static readonly string[] SummaryTableFallbackNames = { "StaffSummaryTable", "SkillLicensesTable58" };
    private const string SummaryStaffIdCell = "A2";
    private const string SummaryStaffNameCell = "B2";
    private const string SummaryPrimarySkillCell = "C2";

    private static List<StaffRecord> _summaryStaffCache = new();
    private static Dictionary<long, string> _primarySkillsByStaffId = new();
    private static Dictionary<string, string> _skillIcons = new(StringComparer.OrdinalIgnoreCase);

    private static readonly Action<dynamic, dynamic> _summaryObserver = OnSummarySheetChanged;
    private static bool _summaryHooked;
    private static bool _isRefreshingSummary;
    private static bool _summaryChangeQueued;
    private static string _pendingSummarySource = SummaryPrimarySkillCell;

    // =========================================================================
    // 2. SKILL ICON HELPERS
    // =========================================================================
    private static string FormatSkillWithIcon(string? skill)
    {
        string s = (skill ?? string.Empty).Trim();
        if (s.Length == 0) return string.Empty;
        return _skillIcons.TryGetValue(s, out var icon) && !string.IsNullOrWhiteSpace(icon) ? $"{icon} {s}" : s;
    }

    private static string StripSkillIcon(string? text)
    {
        string t = (text ?? string.Empty).Trim();
        if (t.Length == 0) return string.Empty;

        foreach (var kv in _skillIcons)
        {
            if (string.IsNullOrWhiteSpace(kv.Value)) continue;
            if (t.Equals($"{kv.Value} {kv.Key}", StringComparison.OrdinalIgnoreCase)) return kv.Key;
        }

        int space = t.IndexOf(' ');
        if (space > 0 && t.Take(space).Any(c => c > 127)) return t[(space + 1)..].Trim();
        return t;
    }

    private static void RebuildSkillIcons()
    {
        var map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        foreach (var kv in SkillCatalog.VehicleLicenseRoleIcons) map[kv.Key] = kv.Value;
        map[SkillCatalog.GeneralWorker] = SkillCatalog.GeneralWorkerIcon;

        try
        {
            foreach (var t in ServerSyncCommands.FetchSkillLicenseTypesFromDb() ?? new List<SkillLicenseType>())
                if (!string.IsNullOrWhiteSpace(t.SkillLicenseName) && !string.IsNullOrWhiteSpace(t.DisplayIcon))
                    map[t.SkillLicenseName.Trim()] = t.DisplayIcon.Trim();
        }
        catch { }

        try
        {
            foreach (var s in OfflineSnapshot.Load<ScheduleRecord>()?.Skills ?? new List<ScheduleSkillRecord>())
                if (!string.IsNullOrWhiteSpace(s.SkillName) && !string.IsNullOrWhiteSpace(s.DisplayIcon))
                    map[s.SkillName.Trim()] = s.DisplayIcon.Trim();
        }
        catch { }

        _skillIcons = map;
    }

    private static string PrimarySkillOf(StaffRecord staff) =>
        _primarySkillsByStaffId.TryGetValue(staff.StaffID, out var s) ? s : string.Empty;

    // =========================================================================
    // 3. CONTEXT INITIALISATION (populate-time)
    // =========================================================================
    private static void InitializeStaffSummaryContext(object wbObj, List<StaffRecord> masters)
    {
        try
        {
            _summaryStaffCache = (masters ?? new List<StaffRecord>()).Where(s => s.StaffID > 0).OrderBy(s => s.StaffID).ToList();
            _primarySkillsByStaffId = ServerSyncCommands.FetchWorkerPrimarySkills();
            RebuildSkillIcons();

            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SummarySheetAliases, "Staff Summary");
            if (ws == null) return;

            bool hasSelection = ParseStaffId(ws.Range[SummaryStaffIdCell].Value2) != null
                                || FindStaffByName(ws.Range[SummaryStaffNameCell].Value2) != null;
            if (!hasSelection) return;

            RunFrozen(() => ApplySummarySelection(ws, SummaryStaffIdCell));
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Staff Summary init warning: {ex.Message}");
        }
    }

    private static void EnsureSummaryCache()
    {
        if (_summaryStaffCache.Count > 0) return;
        try
        {
            var list = ExtractStaffRecordsWithDiagnostics().Data;
            if (list == null || list.Count == 0) list = ServerSyncCommands.FetchStaffRecords();
            if (list != null)
                _summaryStaffCache = list.Where(s => s.StaffID > 0).OrderBy(s => s.StaffID).ToList();
            if (_primarySkillsByStaffId.Count == 0) _primarySkillsByStaffId = ServerSyncCommands.FetchWorkerPrimarySkills();
            if (_skillIcons.Count == 0) RebuildSkillIcons();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Staff Summary cache load failed: {ex.Message}");
        }
    }

    // =========================================================================
    // 4. HOOKS (observer + dropdown providers)
    // =========================================================================
    internal static void RegisterSummaryHooks()
    {
        if (_summaryHooked) return;
        _summaryHooked = true;
        ParameterCells.RegisterChangeObserver(_summaryObserver);

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryStaffIdCell), _ =>
        {
            EnsureSummaryCache();
            // Every staff is listed; those outside the selected Primary Skill (C2) are greyed out below the matching ones.
            string selectedSkill = SelectedPrimarySkill();
            var options = _summaryStaffCache.Select(s =>
            {
                bool matches = StaffMatchesSkill(s, selectedSkill);
                return new DropdownOption
                {
                    Raw = s.StaffID.ToString("D4"),
                    Display = string.IsNullOrWhiteSpace(s.StaffDetails.StaffName) ? s.StaffID.ToString("D4") : $"{s.StaffID:D4} — {s.StaffDetails.StaffName}",
                    IsValid = matches,
                    Section = matches ? 0 : 1
                };
            }).ToList();
            return new CustomDropdownResult(options, (cell, opt) =>
            {
                if (long.TryParse(opt.Raw, out long id)) { cell.NumberFormat = "0000"; cell.Value2 = (double)id; }
            });
        });

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryStaffNameCell), _ =>
        {
            EnsureSummaryCache();
            string selectedSkill = SelectedPrimarySkill();
            var options = _summaryStaffCache
                .Where(s => !string.IsNullOrWhiteSpace(s.StaffDetails.StaffName))
                .Select(s =>
                {
                    bool matches = StaffMatchesSkill(s, selectedSkill);
                    return new DropdownOption
                    {
                        Raw = s.StaffDetails.StaffName,
                        Display = s.StaffDetails.StaffName,
                        IsValid = matches,
                        Section = matches ? 0 : 1
                    };
                })
                .ToList();
            return new CustomDropdownResult(options, (cell, opt) => { cell.Value2 = opt.Raw; });
        });

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryPrimarySkillCell), _ =>
        {
            EnsureSummaryCache();
            var options = _primarySkillsByStaffId.Values
                .Where(v => !string.IsNullOrWhiteSpace(v)).Select(v => v.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .Select(v => new DropdownOption { Raw = FormatSkillWithIcon(v), Display = FormatSkillWithIcon(v), Section = 0 })
                .ToList();
            return new CustomDropdownResult(options, (cell, opt) => { cell.Value2 = opt.Raw; });
        });
    }

    internal static void UnregisterSummaryHooks()
    {
        ParameterCells.UnregisterChangeObserver(_summaryObserver);
        _summaryHooked = false;
    }

    private static string SelectedPrimarySkill()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? ws = FindSummarySheet(xlApp.ActiveWorkbook);
            if (ws != null) return StripSkillIcon(ToText(ws.Range[SummaryPrimarySkillCell].Value2));
        }
        catch { }
        return string.Empty;
    }

    private static bool StaffMatchesSkill(StaffRecord s, string skill) =>
        skill.Length == 0 || PrimarySkillOf(s).Equals(skill, StringComparison.OrdinalIgnoreCase);

    private static bool IsSummaryContext() =>
        !SheetOperations.IsMasterTemplateFile() && !SheetOperations.IsTemplateEditMode
        && SheetOperations.IsModuleWorkbook(null, Events.StaffSyncDescriptor);

    private static bool IsSummaryCell(dynamic cell, string address)
    {
        try
        {
            if (!IsSummaryContext()) return false;
            string sheetName = (string)cell.Worksheet.Name;
            if (!SummarySheetAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase))) return false;
            return string.Equals((string)cell.Address(false, false), address, StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static dynamic? FindSummarySheet(dynamic? wb)
    {
        if (wb == null) return null;
        foreach (var alias in SummarySheetAliases)
        {
            try { dynamic ws = wb.Sheets[alias]; if (ws != null) return ws; } catch { }
        }
        return null;
    }

    // =========================================================================
    // 5. CHANGE OBSERVER (3-way sync A2 <-> B2 <-> C2)
    // =========================================================================
    private static bool Hits(dynamic xlApp, dynamic target, dynamic range)
    {
        try { return range != null && xlApp.Intersect(target, range) != null; }
        catch { return false; }
    }

    private static void OnSummarySheetChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (_isRefreshingSummary) return;
            if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;
            if (!SheetOperations.IsModuleWorkbook(null, Events.StaffSyncDescriptor)) return;

            string sheetName = (string)sh.Name;
            if (!SummarySheetAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase))) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            string? source = null;
            if (Hits(xlApp, target, sh.Range[SummaryStaffIdCell])) source = SummaryStaffIdCell;
            else if (Hits(xlApp, target, sh.Range[SummaryStaffNameCell])) source = SummaryStaffNameCell;
            else if (Hits(xlApp, target, sh.Range[SummaryPrimarySkillCell])) source = SummaryPrimarySkillCell;
            if (source == null) return;

            // Coalesce a burst of edits: A2 beats B2 beats C2.
            if (!_summaryChangeQueued || source == SummaryStaffIdCell
                || (source == SummaryStaffNameCell && _pendingSummarySource == SummaryPrimarySkillCell))
                _pendingSummarySource = source;
            if (_summaryChangeQueued) return;
            _summaryChangeQueued = true;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                _summaryChangeQueued = false;
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic? ws = FindSummarySheet(app.ActiveWorkbook);
                    if (ws == null) return;
                    EnsureSummaryCache();
                    RunFrozen(() => ApplySummarySelection(ws, _pendingSummarySource));
                }
                catch (Exception ex)
                {
                    AppLogger.Error("[StaffManager] Staff Summary refresh failed", ex);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Staff Summary observer warning: {ex.Message}");
        }
    }

    // =========================================================================
    // 6. SELECTION LOGIC
    // =========================================================================
    private static string ToText(object? v) => v?.ToString()?.Trim() ?? string.Empty;

    private static long? ParseStaffId(object? raw)
    {
        switch (raw)
        {
            case null: return null;
            case double d: return d > 0 ? (long)d : null;
            case long l: return l > 0 ? l : null;
        }
        string text = ToText(raw);
        int dash = text.IndexOfAny(new[] { '—', '-', ' ' });
        if (dash > 0) text = text[..dash];
        return long.TryParse(text, out long id) && id > 0 ? id : null;
    }

    private static StaffRecord? FindStaffById(object? raw)
    {
        long? id = ParseStaffId(raw);
        return id == null ? null : _summaryStaffCache.FirstOrDefault(s => s.StaffID == id.Value);
    }

    private static StaffRecord? FindStaffByName(object? raw)
    {
        string name = ToText(raw);
        if (name.Length == 0) return null;
        return _summaryStaffCache.FirstOrDefault(s =>
            name.Equals(s.StaffDetails.StaffName?.Trim(), StringComparison.OrdinalIgnoreCase) ||
            name.Equals(s.StaffDetails.StaffShortName?.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    private static void ApplySummarySelection(dynamic ws, string source)
    {
        bool prev = _isRefreshingSummary;
        _isRefreshingSummary = true;
        try
        {
            StaffRecord? staff;
            if (source == SummaryStaffIdCell)
            {
                staff = FindStaffById(ws.Range[SummaryStaffIdCell].Value2);
                if (staff == null) { ClearSummary(ws, clearId: false, clearName: true, clearSkill: true); return; }
            }
            else if (source == SummaryStaffNameCell)
            {
                staff = FindStaffByName(ws.Range[SummaryStaffNameCell].Value2);
                if (staff == null) { ClearSummary(ws, clearId: true, clearName: false, clearSkill: true); return; }
            }
            else
            {
                string skill = StripSkillIcon(ToText(ws.Range[SummaryPrimarySkillCell].Value2));
                if (skill.Length == 0) return; // all staff selectable again; keep the current selection

                try { ws.Range[SummaryPrimarySkillCell].Value2 = FormatSkillWithIcon(skill); } catch { }
                var matching = _summaryStaffCache
                    .Where(s => PrimarySkillOf(s).Equals(skill, StringComparison.OrdinalIgnoreCase)).ToList();

                var current = FindStaffById(ws.Range[SummaryStaffIdCell].Value2);
                if (current != null && matching.Any(m => m.StaffID == current.StaffID)) return;
                if (matching.Count == 1) staff = matching[0];
                else { ClearSummary(ws, clearId: true, clearName: true, clearSkill: false); return; }
            }

            WriteTopBar(ws, staff);
            TableMapper.PopulateTableSafe<StaffSummaryRow>((object)ws, ServerSyncCommands.FetchStaffSummaryRows(staff), SummaryTableFallbackNames);
        }
        finally
        {
            _isRefreshingSummary = prev;
        }
    }

    private static void WriteTopBar(dynamic ws, StaffRecord staff)
    {
        try { ws.Range[SummaryStaffIdCell].NumberFormat = "0000"; } catch { }
        try { ws.Range[SummaryStaffIdCell].Value2 = (double)staff.StaffID; } catch { }
        try { ws.Range[SummaryStaffNameCell].Value2 = staff.StaffDetails.StaffName; } catch { }
        try { ws.Range[SummaryPrimarySkillCell].Value2 = FormatSkillWithIcon(PrimarySkillOf(staff)); } catch { }
    }

    private static void ClearSummary(dynamic ws, bool clearId, bool clearName, bool clearSkill)
    {
        if (clearId) { try { ws.Range[SummaryStaffIdCell].ClearContents(); } catch { } }
        if (clearName) { try { ws.Range[SummaryStaffNameCell].ClearContents(); } catch { } }
        if (clearSkill) { try { ws.Range[SummaryPrimarySkillCell].ClearContents(); } catch { } }
        TableMapper.PopulateTableSafe<StaffSummaryRow>((object)ws, new List<StaffSummaryRow>(), SummaryTableFallbackNames);
    }

    // =========================================================================
    // 7. COM FREEZE (Rule 2)
    // =========================================================================
    private static void RunFrozen(Action body)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        bool origUpdating = true;
        int origCalculation = -4105;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; } catch { }
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
