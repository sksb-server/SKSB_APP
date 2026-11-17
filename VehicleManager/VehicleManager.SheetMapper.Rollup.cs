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
/// ([ColumnRollup] / [ColumnCalculate] on <see cref="VehicleDetail"/>) honouring the date filter on
/// 'All Vehicles' (B2 anchor date, C2 scope). Workshops' cross-table rollups use the same cells on 'Workshops'.
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
        var (anchor, scope) = ReadDateFilter(wbObj, AllVehicleAliases, MasterSheetName);
        foreach (var vehicle in masters)
        {
            if (string.IsNullOrWhiteSpace(vehicle.VehiclePlate)) continue;
            RuleCalculator.ApplyRecordCalculations(vehicle, null, null, anchor, scope);
        }
    }

    private static (DateOnly Anchor, DateFilterScope Scope) ReadDateFilter(object wbObj, string[] aliases, string defaultName)
    {
        DateOnly anchor = DateOnly.FromDateTime(DateTime.Today);
        DateFilterScope scope = DateFilterScope.All_Time;
        try
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, aliases, defaultName);
            if (ws != null)
            {
                if (RuleDataCache.TryParseDate((object?)ws.Range[FilterAnchorCell].Value2, out var a)) anchor = a;
                scope = DateFilterHelper.ParseScope(((object?)ws.Range[FilterScopeCell].Value2)?.ToString());
            }
        }
        catch { }
        return (anchor, scope);
    }

    // =========================================================================
    // 2. DATE FILTER CHANGE OBSERVER
    // =========================================================================
    private static bool _rollupHooksRegistered;
    private static bool _masterRefreshQueued;
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

            if (AllVehicleAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
            {
                if (_masterRefreshQueued) return;
                _masterRefreshQueued = true;
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    _masterRefreshQueued = false;
                    try { RefreshMasterRollups(); }
                    catch (Exception ex) { AppLogger.Error("[VehicleManager] RefreshMasterRollups failed", ex); }
                });
            }
            else if (WorkshopAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase)))
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
    // 3. SCOPED REFRESH (macro thread)
    // =========================================================================
    private static readonly (string Header, Func<VehicleRecord, decimal?> Value)[] ScopedMasterColumns =
    {
        ("Fuel Cost", v => v.VehicleDetail.FuelCost),
        ("Maintenance Cost", v => v.VehicleDetail.MaintenanceCost),
        ("Insurance Cost", v => v.VehicleDetail.InsuranceCost),
        ("Permit Cost", v => v.VehicleDetail.PermitCost),
        ("Road Tax and Inspection Cost", v => v.VehicleDetail.RoadTaxAndInspectionCost),
        ("Insurance Claim", v => v.VehicleDetail.InsuranceClaimAmount),
        ("Total Cost", v => v.VehicleDetail.TotalCost)
    };

    /// <summary>Re-evaluates the date-scoped cost columns of 'AllVehiclesTable' after B2 / C2 changed.</summary>
    private static void RefreshMasterRollups()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return;
        object wbObj = (object)wb!;

        var (records, _, _) = ExtractVehicleRecordsWithDiagnostics(); // children attached + calculations applied
        if (records == null || records.Count == 0) return;

        bool origEvents = true;
        bool origUpdating = true;
        int origCalculation = -4105; // xlCalculationAutomatic
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }
        try
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, AllVehicleAliases, MasterSheetName);
            dynamic? table = ws == null ? null : TableMapper.ResolveTable((object)ws, MasterTableName);
            dynamic? body = table?.DataBodyRange;
            if (table == null || body == null) return;

            int rows = (int)body!.Rows.Count;
            int cols = (int)table!.ListColumns.Count;
            object[,] headers = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, cols);
            object[,] data = TableMapper.GetRange2DValues(body, rows, cols);

            int plateCol = Enumerable.Range(1, cols).FirstOrDefault(c =>
                string.Equals(headers[1, c]?.ToString()?.Trim(), "Vehicle Plate", StringComparison.OrdinalIgnoreCase));
            if (plateCol < 1) return;

            var byPlate = records
                .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                .GroupBy(v => v.VehiclePlate.Trim().ToUpperInvariant())
                .ToDictionary(g => g.Key, g => g.First());

            foreach (var (header, value) in ScopedMasterColumns)
            {
                int c = Enumerable.Range(1, cols).FirstOrDefault(i =>
                    string.Equals(headers[1, i]?.ToString()?.Trim(), header, StringComparison.OrdinalIgnoreCase));
                if (c < 1) continue;

                // One COM write per column (bulk 2D array)
                var colData = new object[rows, 1];
                for (int r = 1; r <= rows; r++)
                {
                    string plate = (data[r, plateCol]?.ToString() ?? string.Empty).Trim().ToUpperInvariant();
                    // null (no matching rows) => the em-dash text; a genuine 0 stays numeric 0
                    decimal? amount = byPlate.TryGetValue(plate, out var v) ? value(v) : null;
                    colData[r - 1, 0] = amount.HasValue ? (double)amount.Value : "\u2014";
                }
                table.ListColumns[c].DataBodyRange.Value2 = colData;
            }
        }
        finally
        {
            try { xlApp.Calculation = origCalculation; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

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
