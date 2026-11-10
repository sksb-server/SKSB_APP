namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

internal static partial class SheetMapper
{
    private static void ExtractChildRecords(object wbObj, List<VehicleRecord> vehicles)
    {
        dynamic wb = wbObj;
        var vehicleMap = vehicles
            .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
            .GroupBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        // 1. Maintenances
        dynamic? wsMaint = ResolveSheetSafely((object)wb, MaintenanceAliases, "Maintenances");
        if (wsMaint != null)
        {
            List<Maintenance> maintenances = TableMapper.ExtractFromTable<Maintenance>((object)wsMaint, "MaintenancesTable");
            foreach (var m in maintenances)
            {
                if (vehicleMap.TryGetValue(m.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.Maintenances.Add(m);
            }
        }

        // 2. Road Taxes
        dynamic? wsRt = ResolveSheetSafely((object)wb, RoadTaxAliases, "Road Taxes");
        if (wsRt != null)
        {
            List<RoadTax> roadTaxes = TableMapper.ExtractFromTable<RoadTax>((object)wsRt, "RoadTaxesTable");
            foreach (var rt in roadTaxes)
            {
                if (vehicleMap.TryGetValue(rt.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.RoadTaxes.Add(rt);
            }
        }

        // 3. Inspections
        dynamic? wsInsp = ResolveSheetSafely((object)wb, InspectionAliases, "Inspections");
        if (wsInsp != null)
        {
            List<Inspection> inspections = TableMapper.ExtractFromTable<Inspection>((object)wsInsp, "InspectionsTable");
            foreach (var insp in inspections)
            {
                if (vehicleMap.TryGetValue(insp.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.Inspections.Add(insp);
            }
        }

        // 4. Permits
        dynamic? wsPermit = ResolveSheetSafely((object)wb, PermitAliases, "Permits");
        if (wsPermit != null)
        {
            List<VehiclePermit> permits = TableMapper.ExtractFromTable<VehiclePermit>((object)wsPermit, "PermitsTable");
            foreach (var p in permits)
            {
                if (vehicleMap.TryGetValue(p.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.Permits.Add(p);
            }
        }

        // 5. Insurances
        dynamic? wsIns = ResolveSheetSafely((object)wb, InsuranceAliases, "Insurances");
        if (wsIns != null)
        {
            List<VehicleInsurance> insurances = TableMapper.ExtractFromTable<VehicleInsurance>((object)wsIns, "InsurancesTable");
            foreach (var ins in insurances)
            {
                if (vehicleMap.TryGetValue(ins.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.Insurances.Add(ins);
            }
        }

        // 6. OdometersTable
        dynamic? wsOdo = ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            List<OdometerLog> odoLogs = TableMapper.ExtractFromTable<OdometerLog>((object)wsOdo, "OdometersTable");
            foreach (var o in odoLogs)
            {
                if (vehicleMap.TryGetValue(o.VehiclePlate, out VehicleRecord? v) && v != null)
                    v.OdometerLogs.Add(o);
            }
        }
    }

    private static void PopulateChildTables(object wbObj, List<VehicleRecord> records)
    {
        dynamic wb = wbObj;

        // 1. MaintenancesTable
        dynamic? wsMaint = ResolveSheetSafely((object)wb, MaintenanceAliases, "Maintenances");
        if (wsMaint != null)
        {
            try { wsMaint.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allMaintenances = records.SelectMany(v => v.Maintenances.Select(m =>
            {
                if (string.IsNullOrWhiteSpace(m.VehiclePlate)) m.VehiclePlate = v.VehiclePlate;
                return m;
            })).ToList();

            TableMapper.PopulateTableSafe<Maintenance>((object)wsMaint, allMaintenances, "MaintenancesTable");
            SheetProtections.ApplyStyleDrivenProtection(wsMaint);
        }

        // 2. RoadTaxesTable
        dynamic? wsRt = ResolveSheetSafely((object)wb, RoadTaxAliases, "Road Taxes");
        if (wsRt != null)
        {
            try { wsRt.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allRoadTaxes = records.SelectMany(v => v.RoadTaxes.Select(rt =>
            {
                if (string.IsNullOrWhiteSpace(rt.VehiclePlate)) rt.VehiclePlate = v.VehiclePlate;
                return rt;
            })).ToList();

            TableMapper.PopulateTableSafe<RoadTax>((object)wsRt, allRoadTaxes, "RoadTaxesTable");
            SheetProtections.ApplyStyleDrivenProtection(wsRt);
        }

        // 3. InspectionsTable
        dynamic? wsInsp = ResolveSheetSafely((object)wb, InspectionAliases, "Inspections");
        if (wsInsp != null)
        {
            try { wsInsp.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allInspections = records.SelectMany(v => v.Inspections.Select(insp =>
            {
                if (string.IsNullOrWhiteSpace(insp.VehiclePlate)) insp.VehiclePlate = v.VehiclePlate;
                return insp;
            })).ToList();

            TableMapper.PopulateTableSafe<Inspection>((object)wsInsp, allInspections, "InspectionsTable");
            SheetProtections.ApplyStyleDrivenProtection(wsInsp);
        }

        // 4. PermitsTable
        dynamic? wsPermit = ResolveSheetSafely((object)wb, PermitAliases, "Permits");
        if (wsPermit != null)
        {
            try { wsPermit.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allPermits = records.SelectMany(v => v.Permits.Select(p =>
            {
                if (string.IsNullOrWhiteSpace(p.VehiclePlate)) p.VehiclePlate = v.VehiclePlate;
                return p;
            })).ToList();

            TableMapper.PopulateTableSafe<VehiclePermit>((object)wsPermit, allPermits, "PermitsTable");
            SheetProtections.ApplyStyleDrivenProtection(wsPermit);
        }

        // 5. InsurancesTable
        dynamic? wsIns = ResolveSheetSafely((object)wb, InsuranceAliases, "Insurances");
        if (wsIns != null)
        {
            try { wsIns.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allInsurances = records.SelectMany(v => v.Insurances.Select(ins =>
            {
                if (string.IsNullOrWhiteSpace(ins.VehiclePlate)) ins.VehiclePlate = v.VehiclePlate;
                return ins;
            })).ToList();

            TableMapper.PopulateTableSafe<VehicleInsurance>((object)wsIns, allInsurances, "InsurancesTable");
            SheetProtections.ApplyStyleDrivenProtection(wsIns);
        }

        // 6. OdometersTable
        dynamic? wsOdo = ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            try { wsOdo.Unprotect("SKSB_Internal_Guard"); } catch { }
            var allOdoLogs = records.SelectMany(v => v.OdometerLogs.Select(o =>
            {
                if (string.IsNullOrWhiteSpace(o.VehiclePlate)) o.VehiclePlate = v.VehiclePlate;
                return o;
            })).ToList();

            TableMapper.PopulateTableSafe<OdometerLog>((object)wsOdo, allOdoLogs, "OdometersTable");
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
        }
    }
}