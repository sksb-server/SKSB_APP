namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;

internal static class SheetMapper
{
    public const string MasterSheetName = "All Vehicles";
    public const string MasterTableName = "AllVehiclesTable";

    private static readonly string[] MasterSheetAliases = { "All Vehicles", "AllVehicles", "Vehicle_Master", "Vehicles" };
    private static readonly string[] MaintenanceAliases = { "Maintenances", "Maintenance", "Vehicle_Maintenance" };
    private static readonly string[] WorkshopAliases = { "Workshops", "Workshop" };
    private static readonly string[] RoadTaxAliases = { "Road Taxes", "RoadTaxes", "Road Tax" };
    private static readonly string[] InspectionAliases = { "Inspections", "Inspection" };
    private static readonly string[] PermitAliases = { "Permits", "Permit" };
    private static readonly string[] InsuranceAliases = { "Insurances", "Insurance" };
    private static readonly string[] OdometerAliases = { "Odometers", "Odometer", "Fuel Logs", "Fuel" };

    internal static (List<VehicleRecord>? Data, List<string> Errors, List<string> Warnings) ExtractVehicleRecordsWithDiagnostics()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<string> { "No active workbook found." }, new List<string>());

            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs == null)
            {
                return (null, new List<string> { $"Worksheet '{MasterSheetName}' was not found in workbook." }, new List<string>());
            }

            // 1. Extract Master Vehicle Records
            List<VehicleRecord> vehicles = TableMapper.ExtractFromTable<VehicleRecord>((object)masterWs, MasterTableName);

            // 2. Extract Child Collections into parent vehicles
            ExtractChildRecords((object)wb, vehicles);

            // 3. Universally Validate ALL Tables Across All Worksheets (Statically Bound)
            var (errors, warnings) = RulesExtractor.ValidateAllWorkbookTables((object)wb);

            // Primary Key Integrity Check
            var plateGroups = vehicles
                .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                .GroupBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase);

            foreach (var group in plateGroups)
            {
                if (group.Count() > 1)
                {
                    errors.Add($"Duplicate Vehicle Plate '{group.Key}' detected {group.Count()} times in master table.");
                }
            }

            return (vehicles, errors, warnings);
        }
        catch (Exception ex)
        {
            return (null, new List<string> { $"Failed extracting vehicle records: {ex.Message}" }, new List<string>());
        }
    }

    internal static List<VehicleRecord> ExtractVehicleRecords()
    {
        var (data, _, _) = ExtractVehicleRecordsWithDiagnostics();
        return data ?? new List<VehicleRecord>();
    }

    internal static void PopulateWorksheet(List<VehicleRecord> records)
    {
        if (records == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            // 1. Populate Master Table (AllVehiclesTable)
            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs != null)
            {
                try { masterWs.Unprotect("SKSB_Internal_Guard"); } catch { }

                if (masterWs.Cells[1, 2].Value2 == null)
                    masterWs.Cells[1, 2].Value2 = DateTime.Today.ToOADate();

                if (string.IsNullOrWhiteSpace(masterWs.Cells[1, 3].Value2?.ToString()))
                    masterWs.Cells[1, 3].Value2 = "All_Time";

                TableMapper.PopulateTableSafe<VehicleRecord>((object)masterWs, MasterTableName, records);
                SheetProtections.ApplyStyleDrivenProtection(masterWs);
            }

            // 2. Populate Child Tables (Maintenances, RoadTaxes, Inspections, Permits, Insurances, Odometers)
            PopulateChildTables((object)wb, records);

            // 3. Populate Auxiliary Tables (Workshops, InsuranceCompanies, FuelCards, FuelCompanyNames)
            PopulateAuxiliaryTables((object)wb);

#if DEBUG
            ShowPopulateDiagnostics((object)wb, records);
#endif
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[VehicleSheetMapper.PopulateWorksheet Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    private static void ExtractChildRecords(object wbObj, List<VehicleRecord> vehicles)
    {
        dynamic wb = wbObj;
        var vehicleMap = vehicles
            .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
            .GroupBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase)
            .ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);

        // Maintenances
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

        // Road Taxes
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

        // Inspections
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

        // Permits
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

        // Insurances
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

        // 3. FuelCardsTable & FuelCompanyNamesTable (Sheet: Odometers)
        dynamic? wsOdo = ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            try { wsOdo.Unprotect("SKSB_Internal_Guard"); } catch { }
            var fuelCards = ServerSyncCommands.FetchFuelCards();
            TableMapper.PopulateTableSafe<FuelCardRecord>((object)wsOdo, "FuelCardsTable", fuelCards);

            // Populate directly into FuelCompanyNamesTable using COM (works regardless of active sheet)
            var fuelCompanies = ServerSyncCommands.FetchFuelCompanies();
            var companyRecords = fuelCompanies.Select(name => new FuelCompanyNameRecord { FuelCompanyName = name }).ToList();
            TableMapper.PopulateTableSafe<FuelCompanyNameRecord>((object)wsOdo, companyRecords, "FuelCompanyNamesTable", "FuelCompaniesTable");
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
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
                if (m.VehicleType == VehicleType.None) m.VehicleType = v.VehicleDetail.Type;
                return m;
            })).ToList();

            TableMapper.PopulateTableSafe<Maintenance>((object)wsMaint, "MaintenancesTable", allMaintenances);
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
                if (rt.VehicleType == VehicleType.None) rt.VehicleType = v.VehicleDetail.Type;
                return rt;
            })).ToList();

            TableMapper.PopulateTableSafe<RoadTax>((object)wsRt, "RoadTaxesTable", allRoadTaxes);
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
                if (insp.VehicleType == VehicleType.None) insp.VehicleType = v.VehicleDetail.Type;
                return insp;
            })).ToList();

            TableMapper.PopulateTableSafe<Inspection>((object)wsInsp, "InspectionsTable", allInspections);
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
                if (p.VehicleType == VehicleType.None) p.VehicleType = v.VehicleDetail.Type;
                return p;
            })).ToList();

            TableMapper.PopulateTableSafe<VehiclePermit>((object)wsPermit, "PermitsTable", allPermits);
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
                if (ins.VehicleType == VehicleType.None) ins.VehicleType = v.VehicleDetail.Type;
                return ins;
            })).ToList();

            TableMapper.PopulateTableSafe<VehicleInsurance>((object)wsIns, "InsurancesTable", allInsurances);
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
                if (o.VehicleType == VehicleType.None) o.VehicleType = v.VehicleDetail.Type;
                return o;
            })).ToList();

            TableMapper.PopulateTableSafe<OdometerLog>((object)wsOdo, "OdometersTable", allOdoLogs);
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
        }
    }

    private static void PopulateAuxiliaryTables(object wbObj)
    {
        dynamic wb = wbObj;

        // 1. WorkshopsTable (Sheet: Workshops)
        dynamic? wsWorkshop = ResolveSheetSafely((object)wb, WorkshopAliases, "Workshops");
        if (wsWorkshop != null)
        {
            try { wsWorkshop.Unprotect("SKSB_Internal_Guard"); } catch { }
            var workshops = ServerSyncCommands.FetchWorkshops();
            TableMapper.PopulateTableSafe<Workshop>((object)wsWorkshop, "WorkshopsTable", workshops);
            SheetProtections.ApplyStyleDrivenProtection(wsWorkshop);
        }

        // 2. InsuranceCompaniesTable (Sheet: Insurances)
        dynamic? wsIns = ResolveSheetSafely((object)wb, InsuranceAliases, "Insurances");
        if (wsIns != null)
        {
            try { wsIns.Unprotect("SKSB_Internal_Guard"); } catch { }
            var insuranceCompanies = ServerSyncCommands.FetchInsuranceCompanies();
            TableMapper.PopulateTableSafe<InsuranceCompanyRecord>((object)wsIns, "InsuranceCompaniesTable", insuranceCompanies);
            SheetProtections.ApplyStyleDrivenProtection(wsIns);
        }

        // 3. FuelCardsTable & FuelCompanyNamesTable (Sheet: Odometers)
        dynamic? wsOdo = ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            try { wsOdo.Unprotect("SKSB_Internal_Guard"); } catch { }
            var fuelCards = ServerSyncCommands.FetchFuelCards();
            TableMapper.PopulateTableSafe<FuelCardRecord>((object)wsOdo, "FuelCardsTable", fuelCards);

            var fuelCompanies = ServerSyncCommands.FetchFuelCompanies();
            SheetMapperHelper.SyncReferenceTable((object)wb, "FuelCompanyNamesTable", "", fuelCompanies);
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
        }
    }

    internal static dynamic? ResolveSheetSafely(object wbObj, string[] aliases, string defaultName)
    {
        dynamic wb = wbObj;
        string resolvedName = SheetOperations.ResolveSheetName(aliases, defaultName);
        try
        {
            dynamic s = wb.Sheets[resolvedName];
            if (s != null) return s;
        }
        catch { }

        foreach (dynamic ws in wb.Sheets)
        {
            string name = ws.Name?.ToString() ?? string.Empty;
            if (aliases.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase)))
            {
                return ws;
            }
        }
        return null;
    }

#if DEBUG
    private static void ShowPopulateDiagnostics(object wbObj, List<VehicleRecord> records)
    {
        try
        {
            if (!DebugDialog.DebugEnabled) return;

            dynamic wb = wbObj;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== VEHICLE MANAGER POPULATE AUDIT ===");
            sb.AppendLine($"Total Vehicle Profiles Loaded: {records.Count}");
            sb.AppendLine();

            string[] inspectTables =
            {
                "AllVehiclesTable",
                "MaintenancesTable",
                "WorkshopsTable",
                "RoadTaxesTable",
                "InspectionsTable",
                "PermitsTable",
                "InsurancesTable",
                "InsuranceCompaniesTable",
                "OdometersTable",
                "FuelCardsTable",
                "FuelCompanyNamesTable"
            };

            foreach (var tName in inspectTables)
            {
                dynamic? tbl = null;
                string sName = "Unknown";
                foreach (dynamic ws in wb.Sheets)
                {
                    try
                    {
                        tbl = ws.ListObjects[tName];
                        if (tbl != null) { sName = ws.Name?.ToString() ?? "Unknown"; break; }
                    }
                    catch { }
                }

                if (tbl != null)
                {
                    int cols = (int)tbl.ListColumns.Count;
                    int rows = tbl.DataBodyRange != null ? (int)tbl.DataBodyRange.Rows.Count : 0;
                    int startCol = (int)tbl.HeaderRowRange.Column;
                    int startRow = (int)tbl.HeaderRowRange.Row;
                    sb.AppendLine($"  [{tName}] on Sheet '{sName}':");
                    sb.AppendLine($"      Dimensions : {rows} data row(s) × {cols} column(s)");
                    sb.AppendLine($"      Coordinates: Row {startRow}, Columns {startCol} to {startCol + cols - 1}");
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine($"  [{tName}]: NOT FOUND on any worksheet");
                    sb.AppendLine();
                }
            }

            DebugDialog.ShowLog("Vehicle Manager - Populate Audit", sb.ToString());
        }
        catch { }
    }
#endif
}