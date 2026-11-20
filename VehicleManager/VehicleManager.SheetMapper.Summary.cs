namespace SKSB_App.VehicleManager;

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

internal static partial class SheetMapper
{
    // =========================================================================
    // 1. CONSTANTS & CACHED STATE
    // =========================================================================
    private static readonly string[] SummarySheetAliases = { "Vehicle Summary", "VehicleSummary", "Vehicles Summary" };
    private static readonly string[] SummaryTableFallbackNames = { "VehicleSummaryTable", "AllVehiclesTable12" };
    private const string SummaryPlateCell = "A2";
    private const string SummaryTypeCell = "B2";

    private static List<VehicleRecord> _summaryVehicleCache = new();

    private static readonly Action<dynamic, dynamic> _summaryObserver = OnSummarySheetChanged;
    private static bool _summaryHooked;
    private static bool _isRefreshingSummary;
    private static bool _summaryChangeQueued;
    private static string _pendingSummarySource = SummaryTypeCell;

    // =========================================================================
    // 2. CONTEXT INITIALISATION (populate-time)
    // =========================================================================
    private static void InitializeVehicleSummaryContext(object wbObj, List<VehicleRecord> masters)
    {
        try
        {
            _summaryVehicleCache = (masters ?? new List<VehicleRecord>())
                .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                .OrderBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase).ToList();

            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SummarySheetAliases, "Vehicle Summary");
            if (ws == null) return;
            if (FindVehicleByPlate(ws.Range[SummaryPlateCell].Value2) == null) return;

            RunFrozen(() => ApplySummarySelection(ws, SummaryPlateCell));
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[VehicleManager] Vehicle Summary init warning: {ex.Message}");
        }
    }

    private static void EnsureSummaryCache()
    {
        if (_summaryVehicleCache.Count > 0) return;
        try
        {
            var list = ExtractVehicleRecordsWithDiagnostics().Data;
            if (list == null || list.Count == 0) list = ServerSyncCommands.FetchVehicleRecords();
            if (list != null)
                _summaryVehicleCache = list.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                    .OrderBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase).ToList();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[VehicleManager] Vehicle Summary cache load failed: {ex.Message}");
        }
    }

    // =========================================================================
    // 3. HOOKS (observer + dropdown providers)
    // =========================================================================
    internal static void RegisterSummaryHooks()
    {
        if (_summaryHooked) return;
        _summaryHooked = true;
        ParameterCells.RegisterChangeObserver(_summaryObserver);

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryPlateCell), _ =>
        {
            EnsureSummaryCache();
            // Every vehicle is listed; those outside the selected Vehicle Type (B2) are greyed out below the matching ones.
            VehicleType selectedType = SelectedVehicleType();
            var options = _summaryVehicleCache.Select(v =>
            {
                bool matches = selectedType == VehicleType.None || v.VehicleDetail.Type == selectedType;
                return new DropdownOption
                {
                    Raw = v.VehiclePlate,
                    Display = v.VehicleDetail.Type != VehicleType.None ? $"{v.VehiclePlate} — {v.VehicleDetail.Type}" : v.VehiclePlate,
                    IsValid = matches,
                    Section = matches ? 0 : 1
                };
            }).ToList();
            return new CustomDropdownResult(options, (cell, opt) => { cell.Value2 = opt.Raw; });
        });

        CustomDropdowns.Register(c => IsSummaryCell(c, SummaryTypeCell), _ =>
        {
            EnsureSummaryCache();
            var options = _summaryVehicleCache
                .Select(v => v.VehicleDetail.Type).Where(t => t != VehicleType.None).Distinct()
                .Select(t => t.ToString()).OrderBy(t => t, StringComparer.OrdinalIgnoreCase)
                .Select(t => new DropdownOption { Raw = t, Display = t, Section = 0 })
                .ToList();
            return new CustomDropdownResult(options, (cell, opt) => { cell.Value2 = opt.Raw; });
        });
    }

    internal static void UnregisterSummaryHooks()
    {
        ParameterCells.UnregisterChangeObserver(_summaryObserver);
        _summaryHooked = false;
    }

    private static VehicleType SelectedVehicleType()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? ws = FindSummarySheet(xlApp.ActiveWorkbook);
            if (ws != null) return ParseVehicleType(ToText(ws.Range[SummaryTypeCell].Value2));
        }
        catch { }
        return VehicleType.None;
    }

    private static bool IsSummaryContext() =>
        !SheetOperations.IsMasterTemplateFile() && !SheetOperations.IsTemplateEditMode
        && SheetOperations.IsModuleWorkbook(null, Events.VehicleSyncDescriptor);

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
    // 4. CHANGE OBSERVER (2-way sync A2 <-> B2)
    // =========================================================================
    private static void OnSummarySheetChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (_isRefreshingSummary) return;
            if (!IsSummaryContext()) return;

            string sheetName = (string)sh.Name;
            if (!SummarySheetAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase))) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            string? source = null;
            if (Hits(xlApp, target, sh.Range[SummaryPlateCell])) source = SummaryPlateCell;
            else if (Hits(xlApp, target, sh.Range[SummaryTypeCell])) source = SummaryTypeCell;
            if (source == null) return;

            // Coalesce a burst of edits: A2 beats B2.
            if (!_summaryChangeQueued || source == SummaryPlateCell) _pendingSummarySource = source;
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
                    AppLogger.Error("[VehicleManager] Vehicle Summary refresh failed", ex);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[VehicleManager] Vehicle Summary observer warning: {ex.Message}");
        }
    }

    // =========================================================================
    // 5. SELECTION LOGIC
    // =========================================================================
    private static string ToText(object? v) => v?.ToString()?.Trim() ?? string.Empty;

    private static VehicleType ParseVehicleType(string? text)
    {
        string t = (text ?? string.Empty).Replace(" ", string.Empty).Replace("_", string.Empty);
        if (t.Length == 0) return VehicleType.None;
        foreach (VehicleType vt in Enum.GetValues<VehicleType>())
            if (vt != VehicleType.None && vt.ToString().Replace("_", string.Empty).Equals(t, StringComparison.OrdinalIgnoreCase))
                return vt;
        return VehicleType.None;
    }

    private static VehicleRecord? FindVehicleByPlate(object? raw)
    {
        string text = ToText(raw);
        int dash = text.IndexOf(" — ", StringComparison.Ordinal);
        if (dash > 0) text = text[..dash].Trim();
        if (text.Length == 0) return null;
        return _summaryVehicleCache.FirstOrDefault(v => text.Equals(v.VehiclePlate?.Trim(), StringComparison.OrdinalIgnoreCase));
    }

    private static void ApplySummarySelection(dynamic ws, string source)
    {
        bool prev = _isRefreshingSummary;
        _isRefreshingSummary = true;
        try
        {
            VehicleRecord? vehicle;
            if (source == SummaryPlateCell)
            {
                vehicle = FindVehicleByPlate(ws.Range[SummaryPlateCell].Value2);
                if (vehicle == null) { ClearSummary(ws, clearPlate: false, clearType: true); return; }
            }
            else
            {
                VehicleType type = ParseVehicleType(ToText(ws.Range[SummaryTypeCell].Value2));
                if (type == VehicleType.None) return; // all vehicles selectable again; keep the current selection

                var matching = _summaryVehicleCache.Where(v => v.VehicleDetail.Type == type).ToList();
                var current = FindVehicleByPlate(ws.Range[SummaryPlateCell].Value2);
                if (current != null && current.VehicleDetail.Type == type) return;
                if (matching.Count == 1) vehicle = matching[0];
                else { ClearSummary(ws, clearPlate: true, clearType: false); return; }
            }

            WriteTopBar(ws, vehicle);
            var rows = ServerSyncCommands.BuildVehicleSummaryRows(vehicle);
            TableMapper.PopulateTableSafe<VehicleSummaryRow>((object)ws, rows, SummaryTableFallbackNames);
            ApplyOdometerUsedFormat(ws, vehicle);
        }
        finally
        {
            _isRefreshingSummary = prev;
        }
    }

    private static void ApplyOdometerUsedFormat(dynamic ws, VehicleRecord vehicle)
    {
        try
        {
            dynamic? table = null;
            foreach (var name in SummaryTableFallbackNames)
            {
                table = TableMapper.ResolveTable((object)ws, name);
                if (table != null) break;
            }
            dynamic? body = table?.DataBodyRange;
            if (table == null || body == null) return;

            int cols = (int)table.ListColumns.Count;
            for (int c = 1; c <= cols; c++)
            {
                if (!string.Equals(ToText(table.ListColumns[c].Name), "Odometer Used", StringComparison.OrdinalIgnoreCase)) continue;
                table.ListColumns[c].DataBodyRange.NumberFormat = VehicleRecord.IsHourMeterType(vehicle.VehicleDetail.Type)
                    ? RuleRegistry.OdometerHoursFormat
                    : RuleRegistry.OdometerKmFormat;
                break;
            }
        }
        catch { }
    }

    private static void WriteTopBar(dynamic ws, VehicleRecord vehicle)
    {
        try { ws.Range[SummaryPlateCell].Value2 = vehicle.VehiclePlate; } catch { }
        try { ws.Range[SummaryTypeCell].Value2 = vehicle.VehicleDetail.Type.ToString(); } catch { }
    }

    private static void ClearSummary(dynamic ws, bool clearPlate, bool clearType)
    {
        if (clearPlate) { try { ws.Range[SummaryPlateCell].ClearContents(); } catch { } }
        if (clearType) { try { ws.Range[SummaryTypeCell].ClearContents(); } catch { } }
        TableMapper.PopulateTableSafe<VehicleSummaryRow>((object)ws, new List<VehicleSummaryRow>(), SummaryTableFallbackNames);
    }

    // =========================================================================
    // 6. COM FREEZE (Rule 2)
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
