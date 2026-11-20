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

    // =========================================================================
    // VEHICLE SUMMARY (READ-ONLY BUILD & SCHEDULE FETCH)
    // =========================================================================
    private static DateOnly ToDateOnly(object? v) => v switch
    {
        DateOnly d => d,
        DateTime dt => DateOnly.FromDateTime(dt),
        _ => default
    };

    private static string JoinDistinct(IEnumerable<string> items) =>
        string.Join(", ", items.Where(i => !string.IsNullOrWhiteSpace(i)).Select(i => i.Trim()).Distinct(StringComparer.OrdinalIgnoreCase));

    /// <summary>Project codes scheduled for the plate, keyed by schedule date.</summary>
    private static Dictionary<DateOnly, List<string>> FetchScheduledProjectsByDate(string plateUpper)
    {
        var result = new Dictionary<DateOnly, List<string>>();
        void Add(DateOnly d, string project)
        {
            if (d.Year <= 1900 || string.IsNullOrWhiteSpace(project)) return;
            if (!result.TryGetValue(d, out var list)) result[d] = list = new List<string>();
            list.Add(project.Trim());
        }

        if (ServerEnvironment.IsOnline)
        {
            try
            {
                var db = new PostgreSqlService();
                var rows = db.Query<dynamic>(
                    "SELECT schedule_date, COALESCE(project_id, '') AS project_id FROM daily_project_schedules WHERE EXISTS (SELECT 1 FROM unnest(vehicles) v WHERE upper(trim(v)) = @PlateUpper) ORDER BY schedule_date ASC;",
                    new { PlateUpper = plateUpper });
                foreach (var r in rows) Add(ToDateOnly((object?)r.schedule_date), (string)(r.project_id ?? string.Empty));
                return result;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[VehicleManager] Vehicle summary schedule query failed: {ex.Message}");
                if (!ServerEnvironment.IsOffline) return result;
                result.Clear();
            }
        }

        try
        {
            var snap = OfflineSnapshot.Load<ScheduleRecord>();
            if (snap == null) return result;
            foreach (var p in snap.Projects)
                if ((p.Vehicles ?? new List<string>()).Any(v => string.Equals(v?.Trim(), plateUpper, StringComparison.OrdinalIgnoreCase)))
                    Add(snap.DailyScheduleDate, p.ProjectID);
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[VehicleManager] Vehicle summary offline schedule load failed: {ex.Message}");
        }
        return result;
    }

    internal static VehicleStatus GetVehicleStatusOn(VehicleRecord vehicle, DateOnly date)
    {
        DateOnly today = DateOnly.FromDateTime(DateTime.Today);

        if (vehicle.VehicleDetail.SoldDate.HasValue && date >= vehicle.VehicleDetail.SoldDate.Value)
            return VehicleStatus.Decommissioned;

        if (vehicle.Maintenances.Any(m => m.MaintenanceStatus != MaintenanceStatus.Cancelled && m.DateIn <= date
            && (m.DateCompleted.HasValue
                ? date <= m.DateCompleted.Value
                : m.MaintenanceStatus == MaintenanceStatus.In_Progress && date <= today)))
            return VehicleStatus.In_Shop;

        static bool Lapsed(IEnumerable<(DateOnly Start, DateOnly Expiry)> items, DateOnly on)
        {
            var started = items.Where(i => i.Start <= on).ToList();
            return started.Count > 0 && started.Max(i => i.Expiry) < on;
        }

        if (Lapsed(vehicle.Insurances.Select(i => (i.RenewDate, i.ExpiryDate)), date)
            || Lapsed(vehicle.Permits.Select(p => (p.RenewDate, p.ExpiryDate)), date)
            || Lapsed(vehicle.RoadTaxes.Select(r => (r.RenewDate, r.ExpiryDate)), date)
            || Lapsed(vehicle.Inspections.Select(i => (i.ScheduleDateIn, i.ExpiryDate)), date))
            return VehicleStatus.Expired;

        if (vehicle.VehicleStatus == VehicleStatus.Suspended && date >= today)
            return VehicleStatus.Suspended;

        return VehicleStatus.Active;
    }

    private static decimal? SumOrNull(IEnumerable<decimal?> values)
    {
        var list = values.Where(v => v.HasValue).Select(v => v!.Value).ToList();
        return list.Count == 0 ? null : list.Sum();
    }

    internal static List<VehicleSummaryRow> BuildVehicleSummaryRows(VehicleRecord vehicle)
    {
        var rows = new List<VehicleSummaryRow>();
        string plateUpper = (vehicle.VehiclePlate ?? string.Empty).Trim().ToUpperInvariant();
        if (plateUpper.Length == 0) return rows;

        var scheduled = FetchScheduledProjectsByDate(plateUpper);

        var staffMap = new Dictionary<long, string>();
        try
        {
            foreach (var s in StaffLookupHelper.LoadCachedStaff() ?? new List<StaffRecord>())
            {
                string name = !string.IsNullOrWhiteSpace(s.StaffDetails.StaffShortName) ? s.StaffDetails.StaffShortName : s.StaffDetails.StaffName;
                if (!string.IsNullOrWhiteSpace(name)) staffMap[s.StaffID] = name.Trim();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[VehicleManager] Vehicle summary staff lookup failed: {ex.Message}");
        }

        string StaffLabel(string? name, long id) =>
            !string.IsNullOrWhiteSpace(name) ? name.Trim() : staffMap.TryGetValue(id, out var n) ? n : string.Empty;

        var dates = new HashSet<DateOnly>();
        dates.UnionWith(vehicle.OdometerLogs.Select(o => o.Date));
        dates.UnionWith(vehicle.Maintenances.Select(m => m.DateIn));
        dates.UnionWith(vehicle.Insurances.Select(i => i.RenewDate));
        dates.UnionWith(vehicle.Permits.Select(p => p.RenewDate));
        dates.UnionWith(vehicle.RoadTaxes.Select(r => r.RenewDate));
        dates.UnionWith(vehicle.Inspections.Select(i => i.ScheduleDateIn));
        dates.UnionWith(vehicle.InsuranceClaims.Select(c => c.ClaimDate));
        dates.UnionWith(scheduled.Keys);

        foreach (var date in dates.Where(d => d.Year > 1900).OrderByDescending(d => d))
        {
            var odos = vehicle.OdometerLogs.Where(o => o.Date == date).ToList();
            var claims = vehicle.InsuranceClaims.Where(c => c.ClaimDate == date).ToList();

            var drivers = odos.Select(o => StaffLabel(o.StaffName, o.StaffID))
                .Concat(claims.Select(c => StaffLabel(c.StaffName, c.StaffID)));
            var projects = (scheduled.TryGetValue(date, out var sp) ? sp : new List<string>())
                .Concat(odos.Select(o => o.Project));

            var deltas = new List<decimal>();
            foreach (var o in odos)
            {
                if (o.OdometerDifference.HasValue) deltas.Add(o.OdometerDifference.Value);
                else if (o.OdometerBefore.HasValue && o.OdometerBefore.Value > 0m)
                    deltas.Add(o.CurrentOdometer > o.OdometerBefore.Value ? o.CurrentOdometer - o.OdometerBefore.Value : 0m);
            }

            decimal? fuel = odos.Count > 0 ? odos.Sum(o => o.Amount) : null;
            decimal? maint = SumOrNull(vehicle.Maintenances.Where(m => m.DateIn == date).Select(m => m.RepairCost));
            decimal? ins = SumOrNull(vehicle.Insurances.Where(i => i.RenewDate == date).Select(i => i.Amount));
            decimal? permit = SumOrNull(vehicle.Permits.Where(p => p.RenewDate == date).Select(p => p.Amount));
            decimal? roadInspect = SumOrNull(vehicle.RoadTaxes.Where(r => r.RenewDate == date).Select(r => r.Amount)
                .Concat(vehicle.Inspections.Where(i => i.ScheduleDateIn == date).Select(i => i.Amount)));
            decimal? claim = SumOrNull(claims.Select(c => c.ClaimAmount));

            decimal? total = null;
            if (fuel.HasValue || maint.HasValue || ins.HasValue || permit.HasValue || roadInspect.HasValue || claim.HasValue)
                total = (fuel ?? 0m) + (maint ?? 0m) + (ins ?? 0m) + (permit ?? 0m) + (roadInspect ?? 0m)
                        - ((claim ?? 0m) > 0m ? claim!.Value : 0m);

            rows.Add(new VehicleSummaryRow
            {
                Date = date,
                Status = GetVehicleStatusOn(vehicle, date),
                DrivenBy = JoinDistinct(drivers),
                UsedAt = JoinDistinct(projects),
                OdometerUsed = deltas.Count > 0 ? Math.Round(deltas.Sum(), 1) : null,
                FuelCost = fuel,
                MaintenanceCost = maint,
                InsuranceCost = ins,
                PermitCost = permit,
                RoadTaxAndInspectionCost = roadInspect,
                InsuranceClaimAmount = claim,
                TotalCost = total
            });
        }
        return rows;
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
