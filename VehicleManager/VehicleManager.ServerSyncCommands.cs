namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

internal static class ServerSyncCommands
{
    internal static bool SaveVehiclesSilent(List<VehicleRecord> vehicleList, out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;

            // 1. Extract Auxiliary Tables from active workbook
            var workshops = new List<Workshop>();
            var insuranceCompanies = new List<InsuranceCompanyRecord>();
            var fuelCards = new List<FuelCardRecord>();
            var fuelCompanies = new List<string>();

            if (wb != null)
            {
                dynamic? wsWorkshops = SheetMapper.ResolveSheetSafely(wb, new[] { "Workshops", "Workshop" }, "Workshops");
                if (wsWorkshops != null)
                    workshops = TableMapper.ExtractFromTable<Workshop>((object)wsWorkshops, "WorkshopsTable");

                dynamic? wsInsurances = SheetMapper.ResolveSheetSafely(wb, new[] { "Insurances", "Insurance" }, "Insurances");
                if (wsInsurances != null)
                    insuranceCompanies = TableMapper.ExtractFromTable<InsuranceCompanyRecord>((object)wsInsurances, "InsuranceCompaniesTable");

                dynamic? wsOdometers = SheetMapper.ResolveSheetSafely(wb, new[] { "Odometers", "Odometer" }, "Odometers");
                if (wsOdometers != null)
                {
                    fuelCards = TableMapper.ExtractFromTable<FuelCardRecord>((object)wsOdometers, "FuelCardsTable");

                    // Extract ONLY valid rows inside the ListObject boundary so deleted rows are excluded
                    var companyRecords = TableMapper.ExtractFromTable<FuelCompanyNameRecord>((object)wsOdometers, "FuelCompanyNamesTable", "FuelCompaniesTable");
                    fuelCompanies = companyRecords.Select(c => c.FuelCompanyName).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
                }
            }

            // 2. Persist to PostgreSQL Server
            ServerDispatcher.PersistVehicles(vehicleList, isFullBatch: true);
            ServerDispatcher.SaveAuxiliaryTables(workshops, insuranceCompanies, fuelCards, fuelCompanies);

            // 3. Persist local offline snapshots
            OfflineSnapshot.Save(vehicleList, writeToDatabase: false);
            OfflineSnapshot.Save(workshops, writeToDatabase: false);
            OfflineSnapshot.Save(insuranceCompanies, writeToDatabase: false);
            OfflineSnapshot.Save(fuelCards, writeToDatabase: false);

            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] SaveVehiclesSilent Error: {ex.Message}");
            return false;
        }
    }

    internal static List<VehicleRecord> FetchVehicleRecords()
    {
        using var _ = SKSB_App.Core.Diagnostics.AppLogger.Scope("VehicleManager.FetchVehicleRecords");

        try
        {
            var db = new PostgreSqlService();

            // 1. Fetch Vehicle Master
            const string sqlMaster = "SELECT * FROM vehicle_master ORDER BY vehicle_plate;";
            var rawVehicles = db.Query<dynamic>(sqlMaster).ToList();
            SKSB_App.Core.Diagnostics.AppLogger.Info($"[PostgreSQL Query] vehicle_master returned {rawVehicles.Count} record(s).");

            if (rawVehicles.Any())
            {
                var serverVehicles = new List<VehicleRecord>();

                foreach (var v in rawVehicles)
                {
                    string plate = (string)v.vehicle_plate;

                    var target = new VehicleRecord
                    {
                        VehiclePlate = plate,
                        VehicleStatus = Enum.TryParse<VehicleStatus>((string)v.status, true, out var vs) ? vs : VehicleStatus.Active,
                        VehicleDetail = new VehicleDetail
                        {
                            VIN = (string)(v.vin ?? string.Empty),
                            Type = Enum.TryParse<VehicleType>((string)v.vehicle_type, true, out var vt) ? vt : VehicleType.None,
                            Brand = (string)(v.brand ?? string.Empty),
                            Model = (string)(v.model ?? string.Empty),
                            RoadTaxExpiry = v.road_tax_expiry != null ? DateOnly.FromDateTime((DateTime)v.road_tax_expiry) : default,
                            InspectionExpiry = v.inspection_expiry != null ? DateOnly.FromDateTime((DateTime)v.inspection_expiry) : default,
                            CurrentOdometer = v.current_odometer != null ? Convert.ToInt32(v.current_odometer) : 0,
                            BoughtDate = v.bought_date != null ? DateOnly.FromDateTime((DateTime)v.bought_date) : null,
                            SoldDate = v.sold_date != null ? DateOnly.FromDateTime((DateTime)v.sold_date) : null,
                            BoughtPrice = v.bought_price != null ? Convert.ToDecimal(v.bought_price) : 0.00m,
                            SoldPrice = v.sold_price != null ? Convert.ToDecimal(v.sold_price) : 0.00m
                        },
                        Maintenances = new List<Maintenance>(),
                        RoadTaxes = new List<RoadTax>(),
                        Inspections = new List<Inspection>(),
                        Permits = new List<VehiclePermit>(),
                        Insurances = new List<VehicleInsurance>(),
                        OdometerLogs = new List<OdometerLog>()
                    };

                    serverVehicles.Add(target);
                }

                var maintMap = serverVehicles.ToDictionary(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase);

                // 2. Fetch Maintenances
                try
                {
                    const string sqlMaintenance = "SELECT * FROM vehicle_maintenance ORDER BY maintenance_id;";
                    var rawMaintenances = db.Query<dynamic>(sqlMaintenance).ToList();
                    SKSB_App.Core.Diagnostics.AppLogger.Info($"[PostgreSQL Query] vehicle_maintenance returned {rawMaintenances.Count} record(s).");

                    foreach (var m in rawMaintenances)
                    {
                        string mPlate = (string)m.vehicle_plate;
                        if (maintMap.TryGetValue(mPlate, out var parentVehicle))
                        {
                            Enum.TryParse<ServiceType>((string)(m.service_type ?? "None"), true, out var sType);
                            Enum.TryParse<MaintenanceStatus>((string)(m.status ?? "In_Progress"), true, out var mStatus);

                            parentVehicle.Maintenances.Add(new Maintenance
                            {
                                MaintenanceID = Convert.ToInt64(m.maintenance_id),
                                VehiclePlate = mPlate,
                                ServiceType = sType,
                                DateIn = m.date_in != null ? DateOnly.FromDateTime(Convert.ToDateTime(m.date_in)) : DateOnly.FromDateTime(DateTime.Today),
                                EstimatedReturn = m.estimated_return != null ? DateOnly.FromDateTime(Convert.ToDateTime(m.estimated_return)) : DateOnly.FromDateTime(DateTime.Today),
                                DateCompleted = m.date_completed != null ? DateOnly.FromDateTime(Convert.ToDateTime(m.date_completed)) : null,
                                OdometerIn = m.odometer_in != null ? Convert.ToInt32(m.odometer_in) : 0,
                                OdometerOut = m.odometer_out != null ? Convert.ToInt32(m.odometer_out) : 0,
                                RepairCost = m.repair_cost != null ? Convert.ToDecimal(m.repair_cost) : 0.00m,
                                WorkshopName = m.workshop_name?.ToString() ?? string.Empty,
                                MaintenanceStatus = mStatus,
                                Remarks = m.remarks?.ToString() ?? string.Empty,

                                // ✅ Load all 11 breakdown costs:
                                TyreCost = m.tyre_cost != null ? Convert.ToDecimal(m.tyre_cost) : 0.00m,
                                BrakeCost = m.brake_cost != null ? Convert.ToDecimal(m.brake_cost) : 0.00m,
                                ClutchCost = m.clutch_cost != null ? Convert.ToDecimal(m.clutch_cost) : 0.00m,
                                ServiceCost = m.service_cost != null ? Convert.ToDecimal(m.service_cost) : 0.00m,
                                FuelPumpCost = m.fuel_pump_cost != null ? Convert.ToDecimal(m.fuel_pump_cost) : 0.00m,
                                AircondCost = m.aircond_cost != null ? Convert.ToDecimal(m.aircond_cost) : 0.00m,
                                WiringCost = m.wiring_cost != null ? Convert.ToDecimal(m.wiring_cost) : 0.00m,
                                BatteryCost = m.battery_cost != null ? Convert.ToDecimal(m.battery_cost) : 0.00m,
                                TowingCost = m.towing_cost != null ? Convert.ToDecimal(m.towing_cost) : 0.00m,
                                CushionCost = m.cushion_cost != null ? Convert.ToDecimal(m.cushion_cost) : 0.00m,
                                PartsCost = m.parts_cost != null ? Convert.ToDecimal(m.parts_cost) : 0.00m
                            });
                        }
                    }
                }
                catch (Exception exMaint)
                {
                    SKSB_App.Core.Diagnostics.AppLogger.Error("[PostgreSQL Query] vehicle_maintenance failed", exMaint);
                }

                // 3. Fetch Road Taxes
                try
                {
                    const string sqlRt = "SELECT * FROM vehicle_road_taxes ORDER BY road_tax_id;";
                    var rawRt = db.Query<dynamic>(sqlRt).ToList();
                    foreach (var r in rawRt)
                    {
                        string p = (string)r.vehicle_plate;
                        if (maintMap.TryGetValue(p, out var pv))
                        {
                            Enum.TryParse<RoadTaxStatus>((string)(r.status ?? "Active"), true, out var rStatus);
                            pv.RoadTaxes.Add(new RoadTax
                            {
                                RoadTaxID = Convert.ToInt64(r.road_tax_id),
                                VehiclePlate = p,
                                RenewDate = r.renew_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(r.renew_date)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = r.expiry_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(r.expiry_date)) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                                Amount = r.amount != null ? Convert.ToDecimal(r.amount) : 0.00m,
                                Status = rStatus,
                                Remarks = r.remarks?.ToString() ?? string.Empty
                            });
                        }
                    }
                }
                catch (Exception exRt) { AppLogger.Error("[PostgreSQL Query] vehicle_road_taxes failed", exRt); }

                // 4. Fetch Inspections
                try
                {
                    const string sqlInsp = "SELECT * FROM vehicle_inspections ORDER BY inspection_id;";
                    var rawInsp = db.Query<dynamic>(sqlInsp).ToList();
                    foreach (var i in rawInsp)
                    {
                        string p = (string)i.vehicle_plate;
                        if (maintMap.TryGetValue(p, out var pv))
                        {
                            Enum.TryParse<InspectionStatus>((string)(i.status ?? "Scheduled"), true, out var iStatus);
                            pv.Inspections.Add(new Inspection
                            {
                                InspectionID = Convert.ToInt64(i.inspection_id),
                                VehiclePlate = p,
                                ScheduleDateIn = i.schedule_date_in != null ? DateOnly.FromDateTime(Convert.ToDateTime(i.schedule_date_in)) : DateOnly.FromDateTime(DateTime.Today),
                                DateOut = i.date_out != null ? DateOnly.FromDateTime(Convert.ToDateTime(i.date_out)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = i.expiry_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(i.expiry_date)) : DateOnly.FromDateTime(DateTime.Today.AddMonths(6)),
                                Amount = i.amount != null ? Convert.ToDecimal(i.amount) : 0.00m,
                                Status = iStatus,
                                Address = i.address?.ToString() ?? string.Empty,
                                Remarks = i.remarks?.ToString() ?? string.Empty
                            });
                        }
                    }
                }
                catch (Exception exInsp) { AppLogger.Error("[PostgreSQL Query] vehicle_inspections failed", exInsp); }

                // 5. Fetch Permits
                try
                {
                    const string sqlPermits = "SELECT * FROM vehicle_permits ORDER BY permit_id;";
                    var rawPermits = db.Query<dynamic>(sqlPermits).ToList();
                    foreach (var pRec in rawPermits)
                    {
                        string p = (string)pRec.vehicle_plate;
                        if (maintMap.TryGetValue(p, out var pv))
                        {
                            Enum.TryParse<VehiclePermitType>((string)(pRec.permit_type ?? "APAD_Carrier_A"), true, out var pType);
                            Enum.TryParse<VehiclePermitStatus>((string)(pRec.status ?? "Active"), true, out var pStatus);
                            pv.Permits.Add(new VehiclePermit
                            {
                                PermitID = Convert.ToInt64(pRec.permit_id),
                                VehiclePlate = p,
                                PermitType = pType,
                                RenewDate = pRec.renew_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(pRec.renew_date)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = pRec.expiry_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(pRec.expiry_date)) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                                Amount = pRec.amount != null ? Convert.ToDecimal(pRec.amount) : 0.00m,
                                Status = pStatus,
                                Remarks = pRec.remarks?.ToString() ?? string.Empty
                            });
                        }
                    }
                }
                catch (Exception exPermits) { AppLogger.Error("[PostgreSQL Query] vehicle_permits failed", exPermits); }

                // 6. Fetch Insurances
                try
                {
                    const string sqlIns = "SELECT * FROM vehicle_insurances ORDER BY insurance_id;";
                    var rawIns = db.Query<dynamic>(sqlIns).ToList();
                    foreach (var ins in rawIns)
                    {
                        string p = (string)ins.vehicle_plate;
                        if (maintMap.TryGetValue(p, out var pv))
                        {
                            Enum.TryParse<InsuranceType>((string)(ins.insurance_type ?? "Comprehensive"), true, out var insType);
                            Enum.TryParse<InsuranceStatus>((string)(ins.status ?? "Active"), true, out var insStatus);
                            pv.Insurances.Add(new VehicleInsurance
                            {
                                InsuranceID = Convert.ToInt64(ins.insurance_id),
                                VehiclePlate = p,
                                InsuranceCompany = ins.insurance_company?.ToString() ?? string.Empty,
                                InsuranceType = insType,
                                RenewDate = ins.renew_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(ins.renew_date)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = ins.expiry_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(ins.expiry_date)) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                                Amount = ins.amount != null ? Convert.ToDecimal(ins.amount) : 0.00m,
                                Status = insStatus,
                                Remarks = ins.remarks?.ToString() ?? string.Empty,
                                InsuranceNumber = ins.insurance_number?.ToString() ?? string.Empty // <-- Added
                            });
                        }
                    }
                }
                catch (Exception exIns) { AppLogger.Error("[PostgreSQL Query] vehicle_insurances failed", exIns); }

                // 7. Fetch Odometers
                try
                {
                    const string sqlOdo = "SELECT * FROM vehicle_odometers ORDER BY odometer_id;";
                    var rawOdo = db.Query<dynamic>(sqlOdo).ToList();
                    foreach (var o in rawOdo)
                    {
                        string p = (string)o.vehicle_plate;
                        if (maintMap.TryGetValue(p, out var pv))
                        {
                            // Map staff_id cleanly (supports both legacy driven_by fallback and new staff_id)
                            long sId = 0;
                            if (o.staff_id != null)
                            {
                                long.TryParse(o.staff_id.ToString(), out sId);
                            }
                            else if (o.driven_by != null)
                            {
                                long.TryParse(o.driven_by.ToString(), out sId);
                            }

                            pv.OdometerLogs.Add(new OdometerLog
                            {
                                OdometerID = Convert.ToInt64(o.odometer_id),
                                Date = o.date != null ? DateOnly.FromDateTime(Convert.ToDateTime(o.date)) : DateOnly.FromDateTime(DateTime.Today),
                                VehiclePlate = p,
                                StaffID = sId,
                                Project = o.project?.ToString() ?? string.Empty,
                                OdometerBefore = o.odometer_before != null ? Convert.ToInt32(o.odometer_before) : 0,
                                CurrentOdometer = o.current_odometer != null ? Convert.ToInt32(o.current_odometer) : 0,
                                FuelCardNumber = o.fuel_card_number?.ToString() ?? string.Empty,
                                FuelCompany = o.fuel_company?.ToString() ?? string.Empty,
                                StationName = o.station_name?.ToString() ?? string.Empty,
                                ReceiptNumber = o.receipt_number?.ToString() ?? string.Empty,
                                Amount = o.amount != null ? Convert.ToDecimal(o.amount) : 0.00m,
                                Remarks = o.remarks?.ToString() ?? string.Empty
                            });
                        }
                    }
                }
                catch (Exception exOdo) { AppLogger.Error("[PostgreSQL Query] vehicle_odometers failed", exOdo); }

                // Summary of child records
                int totalRoadTaxes = serverVehicles.Sum(v => v.RoadTaxes.Count);
                int totalInspections = serverVehicles.Sum(v => v.Inspections.Count);
                int totalPermits = serverVehicles.Sum(v => v.Permits.Count);
                int totalInsurances = serverVehicles.Sum(v => v.Insurances.Count);
                int totalOdometers = serverVehicles.Sum(v => v.OdometerLogs.Count);

                SKSB_App.Core.Diagnostics.AppLogger.Info(
                    $"[Server Assembly Summary] Vehicles={serverVehicles.Count}, Maintenances={serverVehicles.Sum(v => v.Maintenances.Count)}, " +
                    $"RoadTaxes={totalRoadTaxes}, Inspections={totalInspections}, Permits={totalPermits}, Insurances={totalInsurances}, Odometers={totalOdometers}");

                try
                {
                    OfflineSnapshot.Save(serverVehicles, writeToDatabase: false);
                }
                catch { }

                return serverVehicles;
            }
        }
        catch (Exception ex)
        {
            SKSB_App.Core.Diagnostics.AppLogger.Error("FetchVehicleRecords server query failed", ex);
        }

        var fallback = OfflineSnapshot.Load<List<VehicleRecord>>() ?? new List<VehicleRecord>();
        SKSB_App.Core.Diagnostics.AppLogger.Warn($"[Server Offline Fallback] Loaded {fallback.Count} vehicle records from local offline snapshot.");
        return fallback;
    }

    internal static List<Workshop> FetchWorkshops()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sql = @"
                SELECT 
                    workshop_id AS WorkshopID, 
                    workshop_name AS WorkshopName, 
                    workshop_address AS WorkshopAddress, 
                    workshop_owner AS WorkshopOwner, 
                    primary_contact AS PrimaryContact, 
                    secondary_contact AS SecondaryContact 
                FROM vehicle_workshops 
                ORDER BY workshop_id;";
            var list = db.Query<Workshop>(sql).ToList();
            if (list.Any()) return list;
        }
        catch { }

        return OfflineSnapshot.Load<List<Workshop>>() ?? new List<Workshop>();
    }

    internal static List<InsuranceCompanyRecord> FetchInsuranceCompanies()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sql = @"
                SELECT 
                    insurance_company AS InsuranceCompany, 
                    company_contact_number AS CompanyContactNumber, 
                    company_address AS CompanyAddress, 
                    insurance_agent AS InsuranceAgent, 
                    agent_contact_number AS AgentContactNumber 
                FROM vehicle_insurance_companies 
                ORDER BY insurance_company;";
            var list = db.Query<InsuranceCompanyRecord>(sql).ToList();
            if (list.Any()) return list;
        }
        catch { }

        return OfflineSnapshot.Load<List<InsuranceCompanyRecord>>() ?? new List<InsuranceCompanyRecord>();
    }

    internal static List<FuelCardRecord> FetchFuelCards()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sql = @"
                SELECT 
                    fuel_card_number AS FuelCardNumber, 
                    fuel_company_name AS FuelCompanyName, 
                    card_owner_id AS CardOwnerID, 
                    card_owner_name AS CardOwnerName 
                FROM vehicle_fuel_cards 
                ORDER BY fuel_card_number;";
            var list = db.Query<FuelCardRecord>(sql).ToList();
            if (list.Any()) return list;
        }
        catch { }

        return OfflineSnapshot.Load<List<FuelCardRecord>>() ?? new List<FuelCardRecord>();
    }

    internal static List<string> FetchFuelCompanies()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sql = "SELECT fuel_company_name FROM lookup_fuel_companies ORDER BY fuel_company_name;";
            var list = db.Query<string>(sql).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            return list; // Return exact records from database (even if empty)
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FetchFuelCompanies DB Error] {ex.Message}");
        }

        var cached = OfflineSnapshot.Load<List<FuelCompanyNameRecord>>();
        return cached?.Select(c => c.FuelCompanyName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList() ?? new List<string>();
    }
}