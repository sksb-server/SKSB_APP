namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;

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
                            VehiclePlate = plate,
                            VIN = (string)(v.vin ?? string.Empty),
                            Type = Enum.TryParse<VehicleType>((string)v.vehicle_type, true, out var vt) ? vt : VehicleType.None,
                            Brand = (string)(v.brand ?? string.Empty),
                            Model = (string)(v.model ?? string.Empty),
                            RoadTaxExpiry = v.road_tax_expiry != null ? DateOnly.FromDateTime((DateTime)v.road_tax_expiry) : default,
                            InspectionExpiry = v.inspection_expiry != null ? DateOnly.FromDateTime((DateTime)v.inspection_expiry) : default,
                            CurrentOdometer = v.current_odometer != null ? Convert.ToInt32(v.current_odometer) : 0
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
                                WorkshopAddress = m.workshop_address?.ToString() ?? string.Empty,
                                MaintenanceStatus = mStatus,
                                Remarks = m.remarks?.ToString() ?? string.Empty
                            });
                        }
                    }
                }
                catch (Exception exMaint)
                {
                    SKSB_App.Core.Diagnostics.AppLogger.Error("[PostgreSQL Query] vehicle_maintenance failed", exMaint);
                }

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
            if (!ServerEnvironment.IsSimulatedOffline())
            {
                var db = new PostgreSqlService();
                const string sql = "SELECT fuel_company_name FROM lookup_fuel_companies ORDER BY fuel_company_name;";
                return db.Query<string>(sql).Where(s => !string.IsNullOrWhiteSpace(s)).Distinct().ToList();
            }
        }
        catch { }

        var cached = OfflineSnapshot.Load<List<FuelCompanyNameRecord>>();
        return cached?.Select(c => c.FuelCompanyName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList() ?? new List<string>();
    }
}