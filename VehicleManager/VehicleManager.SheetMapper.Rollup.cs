namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// Declarative rollups for the vehicle workbook. 'AllVehiclesTable' expiry / cost columns are evaluated in C#
/// ([ColumnRollup] / [ColumnCalculate] on <see cref="VehicleDetail"/>) as All-Time totals (the per-date breakdown
/// lives on 'Vehicle Summary'). Workshops' cross-table rollups honour the B2 / C2 date filter on 'Workshops'.
/// </summary>
internal static partial class SheetMapper
{
    private const string FilterAnchorCell = "B2";
    private const string FilterScopeCell = "C2";

    // =========================================================================
    // 1. RECORD-LEVEL CALCULATION (extract + populate pipelines)
    // =========================================================================
    private static void ApplyVehicleCalculations(object wbObj, List<VehicleRecord> masters)
    {
        DateOnly anchor = DateOnly.FromDateTime(DateTime.Today);
        foreach (var vehicle in masters)
        {
            if (string.IsNullOrWhiteSpace(vehicle.VehiclePlate)) continue;
            RuleCalculator.ApplyRecordCalculations(vehicle, null, null, anchor, DateFilterScope.All_Time);
        }
    }

    // =========================================================================
    // 2. DATE FILTER CHANGE OBSERVER
    // =========================================================================
    private static bool _rollupHooksRegistered;
    private static bool _workshopRefreshQueued;
    private static readonly Action<dynamic, dynamic> _filterObserver = OnFilterCellChanged;

    internal static void RegisterRollupHooks()
    {
        if (_rollupHooksRegistered) return;
        _rollupHooksRegistered = true;
        ParameterCells.RegisterChangeObserver(_filterObserver);
    }

    internal static void UnregisterRollupHooks()
    {
        ParameterCells.UnregisterChangeObserver(_filterObserver);
        _rollupHooksRegistered = false;
    }

    private static bool Hits(dynamic xlApp, dynamic target, dynamic range)
    {
        try { return range != null && xlApp.Intersect(target, range) != null; }
        catch { return false; }
    }

    /// <summary>Only decides what changed (Excel event thread); COM mutation is deferred to a queued macro.</summary>
    private static void OnFilterCellChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (SheetOperations.IsTemplateEditMode || SheetOperations.IsMasterTemplateFile()) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            string sheetName = sh.Name?.ToString() ?? string.Empty;
            bool isFilterCell = Hits(xlApp, target, sh.Range[FilterAnchorCell]) || Hits(xlApp, target, sh.Range[FilterScopeCell]);
            if (!isFilterCell) return;

            if (WorkshopAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                if (_workshopRefreshQueued) return;
                _workshopRefreshQueued = true;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    _workshopRefreshQueued = false;
                    try { RefreshWorkshopRollups(); }
                    catch (Exception ex) { AppLogger.Error("[VehicleManager] RefreshWorkshopRollups failed", ex); }
                });
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("[VehicleManager] OnFilterCellChanged failed", ex);
        }
    }

    // =========================================================================
    // 3. WORKSHOP REFRESH (macro thread)
    // =========================================================================
    /// <summary>Re-evaluates the cross-table rollups of 'WorkshopsTable' after its B2 / C2 changed.</summary>
    private static void RefreshWorkshopRollups()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return;

        dynamic? ws = SheetMapperHelper.ResolveSheetSafely((object)wb!, WorkshopAliases, "Workshops");
        dynamic? table = ws == null ? null : TableMapper.ResolveTable((object)ws, "WorkshopsTable");
        if (table == null) return;

        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { RuleEvaluator.EvaluateAndHealTable((object)ws!, (object)table!); }
        finally { try { xlApp.EnableEvents = origEvents; } catch { } }
    }
}
