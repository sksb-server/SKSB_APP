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

internal static class ServerSyncCommands
{
    private static (List<Workshop> Workshops, List<InsuranceCompanyRecord> Insurers, List<FuelCardRecord> FuelCards, List<string> FuelCompanies) ExtractAuxiliaryTables()
    {
        var workshops = new List<Workshop>();
        var insuranceCompanies = new List<InsuranceCompanyRecord>();
        var fuelCards = new List<FuelCardRecord>();
        var fuelCompanies = new List<string>();

        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? wb = xlApp.ActiveWorkbook;
        if (wb == null) return (workshops, insuranceCompanies, fuelCards, fuelCompanies);

        dynamic? wsWorkshops = SheetMapperHelper.ResolveSheetSafely((object)wb, new[] { "Workshops", "Workshop" }, "Workshops");
        if (wsWorkshops != null)
            workshops = TableMapper.ExtractFromTable<Workshop>((object)wsWorkshops, "WorkshopsTable");

        dynamic? wsInsurances = SheetMapperHelper.ResolveSheetSafely((object)wb, new[] { "Insurances", "Insurance" }, "Insurances");
        if (wsInsurances != null)
            insuranceCompanies = TableMapper.ExtractFromTable<InsuranceCompanyRecord>((object)wsInsurances, "InsuranceCompaniesTable");

        dynamic? wsOdometers = SheetMapperHelper.ResolveSheetSafely((object)wb, new[] { "Odometers", "Odometer" }, "Odometers");
        if (wsOdometers != null)
        {
            fuelCards = TableMapper.ExtractFromTable<FuelCardRecord>((object)wsOdometers, "FuelCardsTable");

            // Extract ONLY valid rows inside the ListObject boundary so deleted rows are excluded
            var companyRecords = TableMapper.ExtractFromTable<FuelCompanyNameRecord>((object)wsOdometers, "FuelCompanyNamesTable", "FuelCompaniesTable");
            fuelCompanies = companyRecords.Select(c => c.FuelCompanyName).Where(n => !string.IsNullOrWhiteSpace(n)).Distinct(StringComparer.OrdinalIgnoreCase).ToList();
        }

        return (workshops, insuranceCompanies, fuelCards, fuelCompanies);
    }

    private static void SaveAuxiliarySnapshots(List<Workshop> workshops, List<InsuranceCompanyRecord> insuranceCompanies, List<FuelCardRecord> fuelCards, List<string> fuelCompanies)
    {
        OfflineSnapshot.Save(workshops, writeToDatabase: false);
        OfflineSnapshot.Save(insuranceCompanies, writeToDatabase: false);
        OfflineSnapshot.Save(fuelCards, writeToDatabase: false);
        OfflineSnapshot.Save(fuelCompanies.Select(n => new FuelCompanyNameRecord { FuelCompanyName = n }).ToList(), writeToDatabase: false);
    }

    /// <summary>Offline save: snapshots the master list AND every auxiliary table, and queues the pending upload.</summary>
    internal static void SaveVehiclesOffline(List<VehicleRecord> vehicleList)
    {
        var (workshops, insuranceCompanies, fuelCards, fuelCompanies) = ExtractAuxiliaryTables();
        SaveAuxiliarySnapshots(workshops, insuranceCompanies, fuelCards, fuelCompanies);
        OfflineSnapshot.SaveOfflineSnapshot(vehicleList);
    }

    internal static bool SaveVehiclesSilent(List<VehicleRecord> vehicleList, out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            // 1. Extract Auxiliary Tables from active workbook
            var (workshops, insuranceCompanies, fuelCards, fuelCompanies) = ExtractAuxiliaryTables();

            // Startup pending-upload commit runs before PopulateWorksheet: empty sheets must not wipe the
            // auxiliary tables that were saved offline, so fall back to their snapshots.
            if (OfflineSnapshot.HasPendingSnapshot<List<VehicleRecord>>())
            {
                if (workshops.Count == 0) workshops = OfflineSnapshot.Load<List<Workshop>>() ?? workshops;
                if (insuranceCompanies.Count == 0) insuranceCompanies = OfflineSnapshot.Load<List<InsuranceCompanyRecord>>() ?? insuranceCompanies;
                if (fuelCards.Count == 0) fuelCards = OfflineSnapshot.Load<List<FuelCardRecord>>() ?? fuelCards;
                if (fuelCompanies.Count == 0)
                {
                    fuelCompanies = (OfflineSnapshot.Load<List<FuelCompanyNameRecord>>() ?? new())
                        .Select(c => c.FuelCompanyName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
                }
            }

            AppLogger.Info(
                $"[VehicleManager.SaveVehiclesSilent] Vehicles={vehicleList.Count}, Maintenances={vehicleList.Sum(v => v.Maintenances.Count)}, " +
                $"RoadTaxes={vehicleList.Sum(v => v.RoadTaxes.Count)}, Inspections={vehicleList.Sum(v => v.Inspections.Count)}, " +
                $"Permits={vehicleList.Sum(v => v.Permits.Count)}, Insurances={vehicleList.Sum(v => v.Insurances.Count)}, " +
                $"InsuranceClaims={vehicleList.Sum(v => v.InsuranceClaims?.Count ?? 0)}, Odometers={vehicleList.Sum(v => v.OdometerLogs.Count)}");

            // 2. Persist to PostgreSQL Server
            ServerDispatcher.PersistVehicles(vehicleList, isFullBatch: true);
            ServerDispatcher.SaveAuxiliaryTables(workshops, insuranceCompanies, fuelCards, fuelCompanies);

            // 3. Persist local offline snapshots (only after the server commit succeeded)
            SaveAuxiliarySnapshots(workshops, insuranceCompanies, fuelCards, fuelCompanies);
            OfflineSnapshot.Save(vehicleList, writeToDatabase: false);

            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] SaveVehiclesSilent Error: {ex.Message}");
            return false;
        }
    }

    internal static List<VehicleRecord> FetchVehicleRecords() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "VehicleRecord (all vehicle tables)", QueryVehicleRecords);

    /// <summary>Assembles every vehicle (master + all child tables) from PostgreSQL. An empty server returns an empty list.</summary>
    private static List<VehicleRecord> QueryVehicleRecords(PostgreSqlService db)
    {
        // 1. Fetch Vehicle Master
        const string sqlMaster = "SELECT * FROM vehicle_master ORDER BY vehicle_plate;";
        var rawVehicles = db.Query<dynamic>(sqlMaster).ToList();
        AppLogger.Info($"[PostgreSQL Query] vehicle_master returned {rawVehicles.Count} record(s).");

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
                        DrivenBy = ServerSyncHelper.ColumnText(v, "driven_by"),
                        UsedAt = ServerSyncHelper.ColumnText(v, "used_at"),
                        Type = Enum.TryParse<VehicleType>((string)v.vehicle_type, true, out var vt) ? vt : VehicleType.None,
                        Brand = (string)(v.brand ?? string.Empty),
                        Model = (string)(v.model ?? string.Empty),
                        RoadTaxExpiry = ServerSyncHelper.ValidDate(v.road_tax_expiry),
                        InspectionExpiry = ServerSyncHelper.ValidDate(v.inspection_expiry),
                        CurrentOdometer = v.current_odometer != null ? Math.Round(Convert.ToDecimal(v.current_odometer), 1) : (decimal?)null,
                        BoughtDate = ServerSyncHelper.ValidDate(v.bought_date),
                        SoldDate = ServerSyncHelper.ValidDate(v.sold_date),
                        BoughtPrice = v.bought_price != null ? Convert.ToDecimal(v.bought_price) : (decimal?)null,
                        SoldPrice = v.sold_price != null ? Convert.ToDecimal(v.sold_price) : (decimal?)null
                    },
                    Maintenances = new List<Maintenance>(),
                    RoadTaxes = new List<RoadTax>(),
                    Inspections = new List<Inspection>(),
                    Permits = new List<VehiclePermit>(),
                    Insurances = new List<VehicleInsurance>(),
                    InsuranceClaims = new List<InsuranceClaim>(),
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
                            OdometerIn = m.odometer_in != null ? Math.Round(Convert.ToDecimal(m.odometer_in), 1) : 0,
                            OdometerOut = m.odometer_out != null ? Math.Round(Convert.ToDecimal(m.odometer_out), 1) : 0,
                            RepairCost = m.repair_cost != null ? Convert.ToDecimal(m.repair_cost) : (decimal?)null,
                            WorkshopName = m.workshop_name?.ToString() ?? string.Empty,
                            WorkshopAddress = ServerSyncHelper.ColumnText(m, "workshop_address"),
                            MaintenanceStatus = mStatus,
                            Remarks = m.remarks?.ToString() ?? string.Empty,

                            // ✅ Load all 11 breakdown costs:
                            TyreCost = m.tyre_cost != null ? Convert.ToDecimal(m.tyre_cost) : (decimal?)null,
                            BrakeCost = m.brake_cost != null ? Convert.ToDecimal(m.brake_cost) : (decimal?)null,
                            ClutchCost = m.clutch_cost != null ? Convert.ToDecimal(m.clutch_cost) : (decimal?)null,
                            ServiceCost = m.service_cost != null ? Convert.ToDecimal(m.service_cost) : (decimal?)null,
                            FuelPumpCost = m.fuel_pump_cost != null ? Convert.ToDecimal(m.fuel_pump_cost) : (decimal?)null,
                            AircondCost = m.aircond_cost != null ? Convert.ToDecimal(m.aircond_cost) : (decimal?)null,
                            WiringCost = m.wiring_cost != null ? Convert.ToDecimal(m.wiring_cost) : (decimal?)null,
                            BatteryCost = m.battery_cost != null ? Convert.ToDecimal(m.battery_cost) : (decimal?)null,
                            TowingCost = m.towing_cost != null ? Convert.ToDecimal(m.towing_cost) : (decimal?)null,
                            CushionCost = m.cushion_cost != null ? Convert.ToDecimal(m.cushion_cost) : (decimal?)null,
                            PartsCost = m.parts_cost != null ? Convert.ToDecimal(m.parts_cost) : (decimal?)null
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
                            Amount = r.amount != null ? Convert.ToDecimal(r.amount) : (decimal?)null,
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
                            Amount = i.amount != null ? Convert.ToDecimal(i.amount) : (decimal?)null,
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
                            Amount = pRec.amount != null ? Convert.ToDecimal(pRec.amount) : (decimal?)null,
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
                            Amount = ins.amount != null ? Convert.ToDecimal(ins.amount) : (decimal?)null,
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
                            OdometerBefore = o.odometer_before != null ? Math.Round(Convert.ToDecimal(o.odometer_before), 1) : (decimal?)null,
                            CurrentOdometer = o.current_odometer != null ? Math.Round(Convert.ToDecimal(o.current_odometer), 1) : 0,
                            FuelCardNumber = o.fuel_card_number?.ToString() ?? string.Empty,
                            FuelCompany = o.fuel_company?.ToString() ?? string.Empty,
                            StationName = o.station_name?.ToString() ?? string.Empty,
                            ReceiptNumber = o.receipt_number?.ToString() ?? string.Empty,
                            Time = o.time?.ToString() ?? string.Empty,
                            FuelType = o.fuel_type?.ToString() ?? string.Empty,
                            PriceAtFueling = o.price_at_fueling != null ? Convert.ToDecimal(o.price_at_fueling) : 0.00m,
                            FuelAmount = o.fuel_amount != null ? Convert.ToDecimal(o.fuel_amount) : 0.00m,
                            Amount = o.amount != null ? Convert.ToDecimal(o.amount) : 0.00m,
                            Remarks = o.remarks?.ToString() ?? string.Empty
                        });
                    }
                }
            }
            catch (Exception exOdo) { AppLogger.Error("[PostgreSQL Query] vehicle_odometers failed", exOdo); }

            // 8. Fetch Insurance Claims
            try
            {
                const string sqlClaims = "SELECT * FROM vehicle_insurance_claims ORDER BY claim_id;";
                var rawClaims = db.Query<dynamic>(sqlClaims).ToList();
                SKSB_App.Core.Diagnostics.AppLogger.Info($"[PostgreSQL Query] vehicle_insurance_claims returned {rawClaims.Count} record(s).");

                foreach (var c in rawClaims)
                {
                    string p = (string)c.vehicle_plate;
                    if (maintMap.TryGetValue(p, out var pv))
                    {
                        Enum.TryParse<InsuranceClaimStatus>((string)(c.status ?? "Pending"), true, out var cStatus);
                        Enum.TryParse<VehicleType>((string)(c.vehicle_type ?? "None"), true, out var cVehicleType);

                        pv.InsuranceClaims.Add(new InsuranceClaim
                        {
                            ClaimID = Convert.ToInt64(c.claim_id),
                            ClaimDate = c.claim_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(c.claim_date)) : DateOnly.FromDateTime(DateTime.Today),
                            SettlementDate = c.settlement_date != null ? DateOnly.FromDateTime(Convert.ToDateTime(c.settlement_date)) : null,
                            VehiclePlate = p,
                            VehicleType = cVehicleType,
                            StaffID = c.staff_id != null ? Convert.ToInt64(c.staff_id) : 0,
                            MaintenanceID = c.maintenance_id != null ? Convert.ToInt64(c.maintenance_id) : 0,
                            WorkshopName = c.workshop_name?.ToString() ?? string.Empty,
                            Status = cStatus,
                            ClaimAmount = c.claim_amount != null ? Convert.ToDecimal(c.claim_amount) : (decimal?)null,
                            Remarks = c.remarks?.ToString() ?? string.Empty
                        });
                    }
                }
            }
            catch (Exception exClaims) { AppLogger.Error("[PostgreSQL Query] vehicle_insurance_claims failed", exClaims); }

            // Summary of child records
            int totalInsuranceClaims = serverVehicles.Sum(v => v.InsuranceClaims.Count);
            int totalRoadTaxes = serverVehicles.Sum(v => v.RoadTaxes.Count);
            int totalInspections = serverVehicles.Sum(v => v.Inspections.Count);
            int totalPermits = serverVehicles.Sum(v => v.Permits.Count);
            int totalInsurances = serverVehicles.Sum(v => v.Insurances.Count);
            int totalOdometers = serverVehicles.Sum(v => v.OdometerLogs.Count);

            SKSB_App.Core.Diagnostics.AppLogger.Info(
                $"[Server Assembly Summary] Vehicles={serverVehicles.Count}, Maintenances={serverVehicles.Sum(v => v.Maintenances.Count)}, " +
                $"RoadTaxes={totalRoadTaxes}, Inspections={totalInspections}, Permits={totalPermits}, Insurances={totalInsurances}, InsuranceClaims={totalInsuranceClaims}, Odometers={totalOdometers}");

            return serverVehicles;
        }

        // The server answered successfully but holds no vehicles.
        return new List<VehicleRecord>();
    }

    private const string Module = "VehicleManager";

    internal static List<Workshop> FetchWorkshops() => ServerSyncHelper.FetchAuxiliaryTable<Workshop>(Module, "WorkshopsTable", @"
        SELECT
            workshop_id AS WorkshopID,
            workshop_name AS WorkshopName,
            workshop_address AS WorkshopAddress,
            workshop_owner AS WorkshopOwner,
            primary_contact AS PrimaryContact,
            secondary_contact AS SecondaryContact
        FROM vehicle_workshops
        ORDER BY workshop_id;");

    internal static List<InsuranceCompanyRecord> FetchInsuranceCompanies() => ServerSyncHelper.FetchAuxiliaryTable<InsuranceCompanyRecord>(Module, "InsuranceCompaniesTable", @"
        SELECT
            insurance_company AS InsuranceCompany,
            company_contact_number AS CompanyContactNumber,
            company_address AS CompanyAddress,
            insurance_agent AS InsuranceAgent,
            agent_contact_number AS AgentContactNumber
        FROM vehicle_insurance_companies
        ORDER BY insurance_company;");

    internal static List<FuelCardRecord> FetchFuelCards() => ServerSyncHelper.FetchAuxiliaryTable<FuelCardRecord>(Module, "FuelCardsTable", @"
        SELECT
            fuel_card_number AS FuelCardNumber,
            fuel_company_name AS FuelCompanyName,
            card_owner_id AS CardOwnerID,
            card_owner_name AS CardOwnerName
        FROM vehicle_fuel_cards
        ORDER BY fuel_card_number;");

    internal static List<string> FetchFuelCompanies() =>
        ServerSyncHelper.FetchAuthoritativeList<FuelCompanyNameRecord>(Module, "FuelCompanyNamesTable", db =>
                db.Query<FuelCompanyNameRecord>("SELECT fuel_company_name AS FuelCompanyName FROM lookup_fuel_companies ORDER BY fuel_company_name;").ToList())
            .Select(c => c.FuelCompanyName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct()
            .ToList();
}
