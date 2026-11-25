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
    private const string SummaryDepartmentCell = "C2";

    private static List<StaffRecord> _summaryStaffCache = new();

    private static readonly Action<dynamic, dynamic> _summaryObserver = OnSummarySheetChanged;
    private static bool _summaryHooked;
    private static bool _isRefreshingSummary;
    private static bool _summaryChangeQueued;
    private static string _pendingSummarySource = SummaryDepartmentCell;

    // =========================================================================
    // 2. DEPARTMENT HELPERS
    // =========================================================================
    private static string DepartmentOf(StaffRecord staff) => (staff.StaffDetails?.Department ?? string.Empty).Trim();

    // =========================================================================
    // 3. CONTEXT INITIALISATION (populate-time)
    // =========================================================================
    private static void InitializeStaffSummaryContext(object wbObj, List<StaffRecord> masters)
    {
        try
        {
            _summaryStaffCache = (masters ?? new List<StaffRecord>()).Where(s => s.StaffID > 0).OrderBy(s => s.StaffID).ToList();

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
            // Every staff is listed; those outside the selected Department (C2) are greyed out below the matching ones.
            string selectedDepartment = SelectedDepartment();
            var options = _summaryStaffCache.Select(s =>
            {
                bool matches = StaffMatchesDepartment(s, selectedDepartment);
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
            string selectedDepartment = SelectedDepartment();
            var options = _summaryStaffCache
                .Where(s => !string.IsNullOrWhiteSpace(s.StaffDetails.StaffName))
                .Select(s =>
                {
                    bool matches = StaffMatchesDepartment(s, selectedDepartment);
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

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryDepartmentCell), _ =>
        {
            EnsureSummaryCache();
            var options = _summaryStaffCache.Select(DepartmentOf)
                .Where(v => v.Length > 0)
                .Distinct(StringComparer.OrdinalIgnoreCase).OrderBy(v => v, StringComparer.OrdinalIgnoreCase)
                .Select(v => new DropdownOption { Raw = v, Display = v, Section = 0 })
                .ToList();
            return new CustomDropdownResult(options, (cell, opt) => { cell.Value2 = opt.Raw; });
        });
    }

    internal static void UnregisterSummaryHooks()
    {
        ParameterCells.UnregisterChangeObserver(_summaryObserver);
        _summaryHooked = false;
    }

    private static string SelectedDepartment()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? ws = FindSummarySheet(xlApp.ActiveWorkbook);
            if (ws != null) return ToText(ws.Range[SummaryDepartmentCell].Value2);
        }
        catch { }
        return string.Empty;
    }

    private static bool StaffMatchesDepartment(StaffRecord s, string department) =>
        department.Length == 0 || DepartmentOf(s).Equals(department, StringComparison.OrdinalIgnoreCase);

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
            else if (Hits(xlApp, target, sh.Range[SummaryDepartmentCell])) source = SummaryDepartmentCell;
            if (source == null) return;

            // Coalesce a burst of edits: A2 beats B2 beats C2.
            if (!_summaryChangeQueued || source == SummaryStaffIdCell
                || (source == SummaryStaffNameCell && _pendingSummarySource == SummaryDepartmentCell))
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
                if (staff == null) { ClearSummary(ws, clearId: false, clearName: true, clearDepartment: true); return; }
            }
            else if (source == SummaryStaffNameCell)
            {
                staff = FindStaffByName(ws.Range[SummaryStaffNameCell].Value2);
                if (staff == null) { ClearSummary(ws, clearId: true, clearName: false, clearDepartment: true); return; }
            }
            else
            {
                string department = ToText(ws.Range[SummaryDepartmentCell].Value2);
                if (department.Length == 0) return; // all staff selectable again; keep the current selection

                var matching = _summaryStaffCache
                    .Where(s => DepartmentOf(s).Equals(department, StringComparison.OrdinalIgnoreCase)).ToList();

                var current = FindStaffById(ws.Range[SummaryStaffIdCell].Value2);
                if (current != null && matching.Any(m => m.StaffID == current.StaffID)) return;
                if (matching.Count == 1) staff = matching[0];
                else { ClearSummary(ws, clearId: true, clearName: true, clearDepartment: false); return; }
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
        try { ws.Range[SummaryDepartmentCell].Value2 = DepartmentOf(staff); } catch { }
    }

    private static void ClearSummary(dynamic ws, bool clearId, bool clearName, bool clearDepartment)
    {
        if (clearId) { try { ws.Range[SummaryStaffIdCell].ClearContents(); } catch { } }
        if (clearName) { try { ws.Range[SummaryStaffNameCell].ClearContents(); } catch { } }
        if (clearDepartment) { try { ws.Range[SummaryDepartmentCell].ClearContents(); } catch { } }
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
