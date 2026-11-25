namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;
using System.Globalization;
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

    // =========================================================================
    // 1. DELTA BASELINES & FINGERPRINTS
    // =========================================================================
    // Rows as last loaded from the server; diffing the sheet against them yields the delta to save. Fingerprints
    // cover the persisted user-editable fields only (rollups / calculated / UI-only lookup columns are excluded).
    private static readonly TableBaseline<string> VehicleBaseline = new();
    private static readonly TableBaseline<long> MaintenanceBaseline = new();
    private static readonly TableBaseline<long> ServiceBaseline = new();
    private static readonly TableBaseline<long> RoadTaxBaseline = new();
    private static readonly TableBaseline<long> InspectionBaseline = new();
    private static readonly TableBaseline<long> PermitBaseline = new();
    private static readonly TableBaseline<long> InsuranceBaseline = new();
    private static readonly TableBaseline<long> InsuranceClaimBaseline = new();
    private static readonly TableBaseline<long> OdometerBaseline = new();
    private static readonly TableBaseline<long> WorkshopBaseline = new();
    private static readonly TableBaseline<string> InsuranceCompanyBaseline = new();
    private static readonly TableBaseline<string> FuelCardBaseline = new();

    private const string Sep = "\u001F";

    private static string PlateKey(string? plate) => (plate ?? string.Empty).Trim().ToUpperInvariant();
    private static string TextKey(string? key) => (key ?? string.Empty).Trim();

    private static string Norm(string? s) => (s ?? string.Empty).Trim();
    private static string Norm(DateOnly d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static string Norm(DateOnly? d) => d.HasValue ? Norm(d.Value) : string.Empty;
    private static string Norm(decimal m) => decimal.Round(m, 2).ToString("F2", CultureInfo.InvariantCulture);
    private static string Norm(decimal? m) => m.HasValue ? Norm(m.Value) : string.Empty;
    private static string Norm(long v) => v.ToString(CultureInfo.InvariantCulture);
    private static string NormOdo(decimal m) => decimal.Round(m, 1).ToString("F1", CultureInfo.InvariantCulture);
    private static string NormOdo(decimal? m) => m.HasValue ? NormOdo(m.Value) : string.Empty;
    private static string Norm4(decimal m) => decimal.Round(m, 4).ToString("F4", CultureInfo.InvariantCulture);

    private static string VehicleFingerprint(VehicleRecord v)
    {
        var d = v.VehicleDetail ?? new VehicleDetail();
        return string.Join(Sep,
            PlateKey(v.VehiclePlate), d.Type.ToString(), Norm(d.Brand), Norm(d.Model), Norm(d.VIN),
            Norm(d.BoughtDate), Norm(d.BoughtPrice), Norm(d.SoldDate), Norm(d.SoldPrice),
            NormOdo(d.CurrentOdometer), Norm(d.DrivenBy), Norm(d.UsedAt), v.VehicleStatus.ToString());
    }

    /// <summary>Repair Total Cost is a formula over the breakdown costs, so it is covered by them.</summary>
    private static string MaintenanceFingerprint(Maintenance m) => string.Join(Sep,
        Norm(m.MaintenanceID), PlateKey(m.VehiclePlate), m.ServiceType.ToString(),
        Norm(m.DateIn), Norm(m.EstimatedReturn), Norm(m.DateCompleted), NormOdo(m.OdometerIn), NormOdo(m.OdometerOut),
        Norm(m.WorkshopName), Norm(m.WorkshopAddress), m.MaintenanceStatus.ToString(),
        Norm(m.TyreCost), Norm(m.BrakeCost), Norm(m.ClutchCost), Norm(m.FuelPumpCost),
        Norm(m.AircondCost), Norm(m.WiringCost), Norm(m.BatteryCost), Norm(m.TowingCost), Norm(m.CushionCost),
        Norm(m.PartsCost), Norm(m.Remarks));

    private static string ServiceFingerprint(VehicleService s) => string.Join(Sep,
        Norm(s.ServiceID), PlateKey(s.VehiclePlate), s.ServiceCategory.ToString(), Norm(s.ServiceDate),
        NormOdo(s.OdometerAtService), NormOdo(s.NextServiceOdometer), Norm(s.NextServiceDate),
        Norm(s.WorkshopName), Norm(s.WorkshopAddress), s.Status.ToString(), Norm(s.Cost), Norm(s.Remarks));

    private static string RenewalFingerprint(VehicleRenewalRecord r) => string.Join(Sep,
        PlateKey(r.VehiclePlate), Norm(r.RenewDate), Norm(r.ExpiryDate), Norm(r.Amount), Norm(r.Remarks));

    private static string RoadTaxFingerprint(RoadTax r) => string.Join(Sep,
        Norm(r.RoadTaxID), RenewalFingerprint(r), r.Status.ToString());

    private static string PermitFingerprint(VehiclePermit p) => string.Join(Sep,
        Norm(p.PermitID), RenewalFingerprint(p), p.PermitType.ToString(), p.Status.ToString());

    private static string InsuranceFingerprint(VehicleInsurance i) => string.Join(Sep,
        Norm(i.InsuranceID), RenewalFingerprint(i), Norm(i.InsuranceNumber), Norm(i.InsuranceCompany),
        i.InsuranceType.ToString(), i.Status.ToString());

    private static string InspectionFingerprint(Inspection i) => string.Join(Sep,
        Norm(i.InspectionID), PlateKey(i.VehiclePlate), Norm(i.ScheduleDateIn), Norm(i.DateOut), Norm(i.ExpiryDate),
        Norm(i.Amount), i.Status.ToString(), Norm(i.Address), Norm(i.Remarks));

    private static string InsuranceClaimFingerprint(InsuranceClaim c) => string.Join(Sep,
        Norm(c.ClaimID), PlateKey(c.VehiclePlate), Norm(c.ClaimDate), Norm(c.SettlementDate), Norm(c.StaffID),
        Norm(c.MaintenanceID), Norm(c.WorkshopName), c.Status.ToString(), Norm(c.ClaimAmount), Norm(c.Remarks));

    /// <summary>Odometer Before is a rollup of earlier rows (recomputed on every load), so it is not fingerprinted.</summary>
    private static string OdometerFingerprint(OdometerLog o) => string.Join(Sep,
        Norm(o.OdometerID), PlateKey(o.VehiclePlate), Norm(o.Date), Norm(o.StaffID), Norm(o.Project),
        NormOdo(o.CurrentOdometer), Norm(o.FuelCardNumber), Norm(o.FuelCompany), Norm(o.StationName),
        Norm(o.ReceiptNumber), Norm(o.FuelType), Norm4(o.PriceAtFueling), Norm4(o.FuelAmount), Norm(o.Amount),
        Norm(o.Time), Norm(o.Remarks));

    private static string WorkshopFingerprint(Workshop w) => string.Join(Sep,
        Norm(w.WorkshopID), Norm(w.WorkshopName), Norm(w.WorkshopAddress), Norm(w.WorkshopOwner),
        Norm(w.PrimaryContact), Norm(w.SecondaryContact));

    private static string InsuranceCompanyFingerprint(InsuranceCompanyRecord c) => string.Join(Sep,
        TextKey(c.InsuranceCompany), Norm(c.CompanyContactNumber), Norm(c.CompanyAddress),
        Norm(c.InsuranceAgent), Norm(c.AgentContactNumber));

    /// <summary>Card Owner Name is derived from Card Owner ID (two-way sync), so it is not fingerprinted.</summary>
    private static string FuelCardFingerprint(FuelCardRecord c) => string.Join(Sep,
        TextKey(c.FuelCardNumber), Norm(c.FuelCompanyName), Norm(c.CardOwnerID));

    private static void CaptureVehicleBaselines(List<VehicleRecord> serverList, ISet<string> failedChildTables)
    {
        VehicleBaseline.Capture(serverList, v => PlateKey(v.VehiclePlate), VehicleFingerprint, v => v.ServerUpdatedAt);

        // A child query that failed must not become an (incomplete) baseline: its rows would look "Added" and be
        // re-inserted with new IDs. Without a baseline the next save upserts them by their existing IDs instead.
        void CaptureChild<T>(TableBaseline<long> baseline, string table, Func<VehicleRecord, IEnumerable<T>?> selector,
            Func<T, long> key, Func<T, string> fingerprint) where T : VehicleChildRecord
        {
            if (failedChildTables.Contains(table)) { baseline.Clear(); return; }
            baseline.Capture(serverList.SelectMany(v => selector(v) ?? Enumerable.Empty<T>()), key, fingerprint, r => r.ServerUpdatedAt);
        }

        CaptureChild(MaintenanceBaseline, "vehicle_maintenance", v => v.Maintenances, m => m.MaintenanceID, MaintenanceFingerprint);
        CaptureChild(ServiceBaseline, "vehicle_services", v => v.Services, s => s.ServiceID, ServiceFingerprint);
        CaptureChild(RoadTaxBaseline, "vehicle_road_taxes", v => v.RoadTaxes, r => r.RoadTaxID, RoadTaxFingerprint);
        CaptureChild(InspectionBaseline, "vehicle_inspections", v => v.Inspections, i => i.InspectionID, InspectionFingerprint);
        CaptureChild(PermitBaseline, "vehicle_permits", v => v.Permits, p => p.PermitID, PermitFingerprint);
        CaptureChild(InsuranceBaseline, "vehicle_insurances", v => v.Insurances, i => i.InsuranceID, InsuranceFingerprint);
        CaptureChild(InsuranceClaimBaseline, "vehicle_insurance_claims", v => v.InsuranceClaims, c => c.ClaimID, InsuranceClaimFingerprint);
        CaptureChild(OdometerBaseline, "vehicle_odometers", v => v.OdometerLogs, o => o.OdometerID, OdometerFingerprint);
    }

    /// <summary>
    /// Server <c>updated_at</c> as a concurrency token. Kept exactly as Npgsql returned it (UTC for <c>timestamptz</c>,
    /// Unspecified for <c>timestamp</c>) so it round-trips against the column's own type.
    /// </summary>
    private static DateTime? ToUtcDateTime(object? raw) => raw switch
    {
        DateTime dt => dt.Kind == DateTimeKind.Local ? DateTime.SpecifyKind(dt.ToUniversalTime(), DateTimeKind.Utc) : dt,
        DateTimeOffset dto => DateTime.SpecifyKind(dto.UtcDateTime, DateTimeKind.Utc),
        _ => null
    };

    private static DateTime? UpdatedAt(object row) => ToUtcDateTime(ServerSyncHelper.Col(row, "updated_at"));

    // =========================================================================
    // 2. SAVE COMMANDS
    // =========================================================================
    /// <summary>
    /// Non-interactive save (startup pending upload / contracts without conflict checks). Upsert-only: with no
    /// baseline every row is upserted; with a baseline only changed rows are written, overwriting the server.
    /// </summary>
    internal static bool SaveVehiclesSilent(List<VehicleRecord> vehicleList, out string errorMessage) =>
        SaveVehiclesWithDelta(vehicleList, forceOverwrite: true, out errorMessage, out _);

    /// <summary>
    /// Delta save: only rows that differ from the last server load are written (no hard deletes) across
    /// vehicle_master, the 7 child tables and the auxiliary directories, in one transaction. Modified rows are
    /// checked against their server <c>updated_at</c>; on conflict nothing is saved and <paramref name="conflicts"/>
    /// lists the affected rows.
    /// </summary>
    internal static bool SaveVehiclesWithDelta(List<VehicleRecord> vehicleList, bool forceOverwrite, out string errorMessage, out IReadOnlyList<ConcurrencyConflict>? conflicts)
    {
        errorMessage = string.Empty;
        conflicts = null;
        vehicleList ??= new List<VehicleRecord>();
        try
        {
            // 1. Auxiliary tables from the active workbook. Startup pending-upload commit runs before
            //    PopulateWorksheet: empty sheets must not hide the auxiliary rows saved offline.
            var (workshops, insuranceCompanies, fuelCards, fuelCompanies) = ExtractAuxiliaryTables();
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

            // 2. Child rows inherit the owning plate when the sheet left it blank.
            foreach (var v in vehicleList)
            {
                void Inherit(IEnumerable<VehicleChildRecord>? rows)
                {
                    foreach (var c in rows ?? Enumerable.Empty<VehicleChildRecord>())
                        if (string.IsNullOrWhiteSpace(c.VehiclePlate)) c.VehiclePlate = v.VehiclePlate;
                }
                Inherit(v.Maintenances); Inherit(v.Services); Inherit(v.RoadTaxes); Inherit(v.Inspections); Inherit(v.Permits);
                Inherit(v.Insurances); Inherit(v.InsuranceClaims); Inherit(v.OdometerLogs);
            }

            var vehicles = vehicleList.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)).ToList();
            IEnumerable<T> Children<T>(Func<VehicleRecord, IEnumerable<T>?> selector) where T : VehicleChildRecord =>
                vehicles.SelectMany(v => selector(v) ?? Enumerable.Empty<T>()).Where(c => !string.IsNullOrWhiteSpace(c.VehiclePlate));

            // 3. Diff every table against its server baseline.
            var payload = new VehicleDeltaPayload
            {
                Vehicles = VehicleBaseline.Diff(vehicles, v => PlateKey(v.VehiclePlate), VehicleFingerprint),
                Maintenances = MaintenanceBaseline.Diff(Children(v => v.Maintenances), m => m.MaintenanceID, MaintenanceFingerprint),
                Services = ServiceBaseline.Diff(Children(v => v.Services), s => s.ServiceID, ServiceFingerprint),
                RoadTaxes = RoadTaxBaseline.Diff(Children(v => v.RoadTaxes), r => r.RoadTaxID, RoadTaxFingerprint),
                Inspections = InspectionBaseline.Diff(Children(v => v.Inspections), i => i.InspectionID, InspectionFingerprint),
                Permits = PermitBaseline.Diff(Children(v => v.Permits), p => p.PermitID, PermitFingerprint),
                Insurances = InsuranceBaseline.Diff(Children(v => v.Insurances), i => i.InsuranceID, InsuranceFingerprint),
                InsuranceClaims = InsuranceClaimBaseline.Diff(Children(v => v.InsuranceClaims), c => c.ClaimID, InsuranceClaimFingerprint),
                OdometerLogs = OdometerBaseline.Diff(Children(v => v.OdometerLogs), o => o.OdometerID, OdometerFingerprint),
                Workshops = WorkshopBaseline.Diff(workshops.Where(w => !string.IsNullOrWhiteSpace(w.WorkshopName)), w => w.WorkshopID, WorkshopFingerprint),
                InsuranceCompanies = InsuranceCompanyBaseline.Diff(
                    insuranceCompanies.Where(c => !string.IsNullOrWhiteSpace(c.InsuranceCompany)), c => TextKey(c.InsuranceCompany), InsuranceCompanyFingerprint),
                FuelCards = FuelCardBaseline.Diff(
                    fuelCards.Where(c => !string.IsNullOrWhiteSpace(c.FuelCardNumber)), c => TextKey(c.FuelCardNumber), FuelCardFingerprint),
                FuelCompanies = fuelCompanies
            };

            AppLogger.Info(
                $"[{Module}] Delta save (force={forceOverwrite}): Vehicles={vehicles.Count}, Maintenances={vehicles.Sum(v => v.Maintenances.Count)}, " +
                $"Services={vehicles.Sum(v => v.Services?.Count ?? 0)}, RoadTaxes={vehicles.Sum(v => v.RoadTaxes.Count)}, Inspections={vehicles.Sum(v => v.Inspections.Count)}, " +
                $"Permits={vehicles.Sum(v => v.Permits.Count)}, Insurances={vehicles.Sum(v => v.Insurances.Count)}, " +
                $"InsuranceClaims={vehicles.Sum(v => v.InsuranceClaims?.Count ?? 0)}, Odometers={vehicles.Sum(v => v.OdometerLogs.Count)}");

            // 4. Single transaction (rolled back on any conflict); local snapshots only after the server accepted it.
            ServerDispatcher.PersistVehicleDelta(payload, forceOverwrite);

            SaveAuxiliarySnapshots(workshops, insuranceCompanies, fuelCards, fuelCompanies);
            OfflineSnapshot.Save(vehicleList, writeToDatabase: false);
            return true;
        }
        catch (ConcurrencyConflictException cex)
        {
            conflicts = cex.Conflicts;
            errorMessage = cex.Message;
            AppLogger.Warn($"[{Module}] Delta save rolled back: {cex.Conflicts.Count} concurrency conflict(s).");
            return false;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error($"[{Module}] SaveVehiclesWithDelta failed", ex);
            return false;
        }
    }

    // =========================================================================
    // 3. CONCURRENCY CONFLICT RESOLUTION ("Keep Server Values")
    // =========================================================================
    /// <summary>
    /// After the pane wrote the server's values into the conflicting cells, re-reads ONLY those rows and moves their
    /// delta baselines to the server's latest fingerprint + <c>updated_at</c>. The next save then treats them as
    /// unchanged, while every other edited row keeps its local edits and its original baseline.
    /// </summary>
    internal static void AcceptServerConflictBaselines(IReadOnlyList<ConcurrencyConflict> conflicts)
    {
        if (conflicts == null || conflicts.Count == 0) return;
        if (!ServerEnvironment.IsOnline)
        {
            AppLogger.Warn($"[{Module}] Keep Server Values: server offline, conflict baselines not refreshed.");
            return;
        }

        IEnumerable<string> KeysFor(string table) => conflicts
            .Where(c => string.Equals(c.TableName, table, StringComparison.OrdinalIgnoreCase))
            .Select(c => (c.RawKeyValue ?? string.Empty).Trim())
            .Where(k => k.Length > 0)
            .Distinct(StringComparer.Ordinal);

        long[] IdsFor(string table) => KeysFor(table)
            .Select(k => long.TryParse(k, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) ? id : 0)
            .Where(id => id > 0)
            .Distinct()
            .ToArray();

        int refreshed = 0;
        try
        {
            var db = new PostgreSqlService();

            void RefreshChild<T>(string table, string idColumn, TableBaseline<long> baseline,
                Func<object, string, T> map, Func<T, long> key, Func<T, string> fingerprint) where T : VehicleChildRecord
            {
                long[] ids = IdsFor(table);
                if (ids.Length == 0) return;
                foreach (var row in db.Query<dynamic>($"SELECT * FROM {table} WHERE {idColumn} = ANY(@Ids);", new { Ids = ids }))
                {
                    object r = row;
                    T rec = map(r, ServerSyncHelper.ColumnText(r, "vehicle_plate"));
                    baseline.UpsertEntry(key(rec), fingerprint(rec), rec.ServerUpdatedAt);
                    refreshed++;
                }
            }

            string[] plates = KeysFor("vehicle_master").Select(PlateKey).Distinct().ToArray();
            if (plates.Length > 0)
            {
                foreach (var row in db.Query<dynamic>("SELECT * FROM vehicle_master WHERE vehicle_plate = ANY(@Plates);", new { Plates = plates }))
                {
                    VehicleRecord v = MapVehicleRow((object)row);
                    VehicleBaseline.UpsertEntry(PlateKey(v.VehiclePlate), VehicleFingerprint(v), v.ServerUpdatedAt);
                    refreshed++;
                }
            }

            RefreshChild("vehicle_maintenance", "maintenance_id", MaintenanceBaseline, MapMaintenanceRow, m => m.MaintenanceID, MaintenanceFingerprint);
            RefreshChild("vehicle_services", "service_id", ServiceBaseline, MapServiceRow, s => s.ServiceID, ServiceFingerprint);
            RefreshChild("vehicle_road_taxes", "road_tax_id", RoadTaxBaseline, MapRoadTaxRow, r => r.RoadTaxID, RoadTaxFingerprint);
            RefreshChild("vehicle_inspections", "inspection_id", InspectionBaseline, MapInspectionRow, i => i.InspectionID, InspectionFingerprint);
            RefreshChild("vehicle_permits", "permit_id", PermitBaseline, MapPermitRow, p => p.PermitID, PermitFingerprint);
            RefreshChild("vehicle_insurances", "insurance_id", InsuranceBaseline, MapInsuranceRow, i => i.InsuranceID, InsuranceFingerprint);
            RefreshChild("vehicle_insurance_claims", "claim_id", InsuranceClaimBaseline, MapInsuranceClaimRow, c => c.ClaimID, InsuranceClaimFingerprint);
            RefreshChild("vehicle_odometers", "odometer_id", OdometerBaseline, MapOdometerRow, o => o.OdometerID, OdometerFingerprint);

            long[] workshopIds = IdsFor("vehicle_workshops");
            if (workshopIds.Length > 0)
            {
                foreach (var w in db.Query<Workshop>(SqlSelectWorkshops + " WHERE workshop_id = ANY(@Ids);", new { Ids = workshopIds }))
                {
                    WorkshopBaseline.UpsertEntry(w.WorkshopID, WorkshopFingerprint(w), w.ServerUpdatedAt);
                    refreshed++;
                }
            }

            string[] insurers = KeysFor("vehicle_insurance_companies").ToArray();
            if (insurers.Length > 0)
            {
                foreach (var c in db.Query<InsuranceCompanyRecord>(SqlSelectInsuranceCompanies + " WHERE insurance_company = ANY(@Keys);", new { Keys = insurers }))
                {
                    InsuranceCompanyBaseline.UpsertEntry(TextKey(c.InsuranceCompany), InsuranceCompanyFingerprint(c), c.ServerUpdatedAt);
                    refreshed++;
                }
            }

            string[] cards = KeysFor("vehicle_fuel_cards").ToArray();
            if (cards.Length > 0)
            {
                foreach (var c in db.Query<FuelCardRecord>(SqlSelectFuelCards + " WHERE fuel_card_number = ANY(@Keys);", new { Keys = cards }))
                {
                    FuelCardBaseline.UpsertEntry(TextKey(c.FuelCardNumber), FuelCardFingerprint(c), c.ServerUpdatedAt);
                    refreshed++;
                }
            }

            AppLogger.Info($"[{Module}] Keep Server Values: refreshed {refreshed} conflict baseline(s) for {conflicts.Count} conflict(s).");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{Module}] AcceptServerConflictBaselines failed", ex);
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

        // Child tables whose query failed: their baselines are cleared instead of captured from partial data.
        var failedChildTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        if (rawVehicles.Any())
        {
            var serverVehicles = new List<VehicleRecord>();
            foreach (var v in rawVehicles) serverVehicles.Add(MapVehicleRow((object)v));

            var maintMap = new Dictionary<string, VehicleRecord>(StringComparer.OrdinalIgnoreCase);
            foreach (var v in serverVehicles) maintMap[v.VehiclePlate] = v;

            // 2-8. Child tables (one query each; rows attach to their vehicle by plate)
            void LoadChildren<T>(string table, string idColumn, Func<VehicleRecord, List<T>> list, Func<object, string, T> map)
            {
                try
                {
                    var rows = db.Query<dynamic>($"SELECT * FROM {table} ORDER BY {idColumn};").ToList();
                    AppLogger.Info($"[PostgreSQL Query] {table} returned {rows.Count} record(s).");
                    foreach (var row in rows)
                    {
                        object r = row;
                        string plate = ServerSyncHelper.ColumnText(r, "vehicle_plate");
                        if (maintMap.TryGetValue(plate.Trim(), out var parent))
                            list(parent).Add(map(r, plate));
                    }
                }
                catch (Exception ex)
                {
                    failedChildTables.Add(table);
                    AppLogger.Error($"[PostgreSQL Query] {table} failed", ex);
                }
            }

            LoadChildren("vehicle_maintenance", "maintenance_id", v => v.Maintenances, MapMaintenanceRow);
            LoadChildren("vehicle_services", "service_id", v => v.Services, MapServiceRow);
            LoadChildren("vehicle_road_taxes", "road_tax_id", v => v.RoadTaxes, MapRoadTaxRow);
            LoadChildren("vehicle_inspections", "inspection_id", v => v.Inspections, MapInspectionRow);
            LoadChildren("vehicle_permits", "permit_id", v => v.Permits, MapPermitRow);
            LoadChildren("vehicle_insurances", "insurance_id", v => v.Insurances, MapInsuranceRow);
            LoadChildren("vehicle_odometers", "odometer_id", v => v.OdometerLogs, MapOdometerRow);
            LoadChildren("vehicle_insurance_claims", "claim_id", v => v.InsuranceClaims, MapInsuranceClaimRow);

            // Summary of child records
            int totalInsuranceClaims = serverVehicles.Sum(v => v.InsuranceClaims.Count);
            int totalRoadTaxes = serverVehicles.Sum(v => v.RoadTaxes.Count);
            int totalInspections = serverVehicles.Sum(v => v.Inspections.Count);
            int totalPermits = serverVehicles.Sum(v => v.Permits.Count);
            int totalInsurances = serverVehicles.Sum(v => v.Insurances.Count);
            int totalOdometers = serverVehicles.Sum(v => v.OdometerLogs.Count);

            SKSB_App.Core.Diagnostics.AppLogger.Info(
                $"[Server Assembly Summary] Vehicles={serverVehicles.Count}, Maintenances={serverVehicles.Sum(v => v.Maintenances.Count)}, " +
                $"Services={serverVehicles.Sum(v => v.Services.Count)}, RoadTaxes={totalRoadTaxes}, Inspections={totalInspections}, Permits={totalPermits}, Insurances={totalInsurances}, InsuranceClaims={totalInsuranceClaims}, Odometers={totalOdometers}");

            // Only reached when the server answered (offline-cache rows must never become a baseline).
            CaptureVehicleBaselines(serverVehicles, failedChildTables);
            return serverVehicles;
        }

        // The server answered successfully but holds no vehicles: the (empty) result is an authoritative baseline.
        var empty = new List<VehicleRecord>();
        CaptureVehicleBaselines(empty, failedChildTables);
        return empty;
    }

    // -------------------------------------------------------------------------
    // Row mapping (shared by the full load and conflict-baseline refresh so fingerprints always agree)
    // -------------------------------------------------------------------------
    private static string Txt(object row, string col) => ServerSyncHelper.ColumnText(row, col);
    private static DateOnly DateOr(object row, string col, DateOnly fallback) => ServerSyncHelper.DateOrDefault(ServerSyncHelper.Col(row, col), fallback);
    private static DateOnly? DateN(object row, string col) => ServerSyncHelper.Col(row, col) is { } raw ? DateOnly.FromDateTime(Convert.ToDateTime(raw)) : null;
    private static decimal? DecN(object row, string col) => ServerSyncHelper.DecN(ServerSyncHelper.Col(row, col));
    private static decimal Dec(object row, string col) => ServerSyncHelper.Dec(ServerSyncHelper.Col(row, col));
    private static decimal? OdoN(object row, string col) => DecN(row, col) is decimal d ? Math.Round(d, 1) : null;
    private static long Lng(object row, string col) => long.TryParse(Txt(row, col), NumberStyles.Integer, CultureInfo.InvariantCulture, out long v) ? v : 0;
    private static TEnum ParseEnum<TEnum>(object row, string col, TEnum fallback) where TEnum : struct, Enum =>
        Enum.TryParse<TEnum>(Txt(row, col), ignoreCase: true, out var e) ? e : fallback;
    private static DateOnly Today => DateOnly.FromDateTime(DateTime.Today);

    private static VehicleRecord MapVehicleRow(object v) => new()
    {
        VehiclePlate = Txt(v, "vehicle_plate"),
        VehicleStatus = ParseEnum(v, "status", VehicleStatus.Active),
        VehicleDetail = new VehicleDetail
        {
            VIN = Txt(v, "vin"),
            DrivenBy = Txt(v, "driven_by"),
            UsedAt = Txt(v, "used_at"),
            Type = ParseEnum(v, "vehicle_type", VehicleType.None),
            Brand = Txt(v, "brand"),
            Model = Txt(v, "model"),
            RoadTaxExpiry = ServerSyncHelper.ValidDate(ServerSyncHelper.Col(v, "road_tax_expiry")),
            InspectionExpiry = ServerSyncHelper.ValidDate(ServerSyncHelper.Col(v, "inspection_expiry")),
            CurrentOdometer = OdoN(v, "current_odometer"),
            BoughtDate = ServerSyncHelper.ValidDate(ServerSyncHelper.Col(v, "bought_date")),
            SoldDate = ServerSyncHelper.ValidDate(ServerSyncHelper.Col(v, "sold_date")),
            BoughtPrice = DecN(v, "bought_price"),
            SoldPrice = DecN(v, "sold_price")
        },
        Maintenances = new List<Maintenance>(),
        Services = new List<VehicleService>(),
        RoadTaxes = new List<RoadTax>(),
        Inspections = new List<Inspection>(),
        Permits = new List<VehiclePermit>(),
        Insurances = new List<VehicleInsurance>(),
        InsuranceClaims = new List<InsuranceClaim>(),
        OdometerLogs = new List<OdometerLog>(),
        ServerUpdatedAt = UpdatedAt(v)
    };

    private static Maintenance MapMaintenanceRow(object m, string plate) => new()
    {
        MaintenanceID = Lng(m, "maintenance_id"),
        VehiclePlate = plate,
        ServiceType = ParseEnum(m, "service_type", ServiceType.None),
        DateIn = DateOr(m, "date_in", Today),
        EstimatedReturn = DateOr(m, "estimated_return", Today),
        DateCompleted = DateN(m, "date_completed"),
        OdometerIn = OdoN(m, "odometer_in") ?? 0m,
        OdometerOut = OdoN(m, "odometer_out") ?? 0m,
        RepairCost = DecN(m, "repair_cost"),
        WorkshopName = Txt(m, "workshop_name"),
        WorkshopAddress = Txt(m, "workshop_address"),
        MaintenanceStatus = ParseEnum(m, "status", MaintenanceStatus.In_Progress),
        Remarks = Txt(m, "remarks"),
        TyreCost = DecN(m, "tyre_cost"),
        BrakeCost = DecN(m, "brake_cost"),
        ClutchCost = DecN(m, "clutch_cost"),
        FuelPumpCost = DecN(m, "fuel_pump_cost"),
        AircondCost = DecN(m, "aircond_cost"),
        WiringCost = DecN(m, "wiring_cost"),
        BatteryCost = DecN(m, "battery_cost"),
        TowingCost = DecN(m, "towing_cost"),
        CushionCost = DecN(m, "cushion_cost"),
        PartsCost = DecN(m, "parts_cost"),
        ServerUpdatedAt = UpdatedAt(m)
    };

    private static VehicleService MapServiceRow(object s, string plate) => new()
    {
        ServiceID = Lng(s, "service_id"),
        VehiclePlate = plate,
        ServiceCategory = ParseEnum(s, "service_category", ServiceCategory.General_Service),
        ServiceDate = DateOr(s, "service_date", Today),
        OdometerAtService = OdoN(s, "odometer_at_service"),
        NextServiceOdometer = OdoN(s, "next_service_odometer"),
        NextServiceDate = DateN(s, "next_service_date"),
        WorkshopName = Txt(s, "workshop_name"),
        WorkshopAddress = Txt(s, "workshop_address"),
        Status = ParseEnum(s, "status", VehicleServiceStatus.Completed),
        Cost = DecN(s, "cost"),
        Remarks = Txt(s, "remarks"),
        ServerUpdatedAt = UpdatedAt(s)
    };

    private static RoadTax MapRoadTaxRow(object r, string plate) => new()
    {
        RoadTaxID = Lng(r, "road_tax_id"),
        VehiclePlate = plate,
        RenewDate = DateOr(r, "renew_date", Today),
        ExpiryDate = DateOr(r, "expiry_date", Today.AddYears(1)),
        Amount = DecN(r, "amount"),
        Status = ParseEnum(r, "status", default(RoadTaxStatus)),
        Remarks = Txt(r, "remarks"),
        ServerUpdatedAt = UpdatedAt(r)
    };

    private static Inspection MapInspectionRow(object i, string plate) => new()
    {
        InspectionID = Lng(i, "inspection_id"),
        VehiclePlate = plate,
        ScheduleDateIn = DateOr(i, "schedule_date_in", Today),
        DateOut = DateOr(i, "date_out", Today),
        ExpiryDate = DateOr(i, "expiry_date", Today.AddMonths(6)),
        Amount = DecN(i, "amount"),
        Status = ParseEnum(i, "status", default(InspectionStatus)),
        Address = Txt(i, "address"),
        Remarks = Txt(i, "remarks"),
        ServerUpdatedAt = UpdatedAt(i)
    };

    private static VehiclePermit MapPermitRow(object p, string plate) => new()
    {
        PermitID = Lng(p, "permit_id"),
        VehiclePlate = plate,
        PermitType = ParseEnum(p, "permit_type", VehiclePermitType.APAD_Carrier_A),
        RenewDate = DateOr(p, "renew_date", Today),
        ExpiryDate = DateOr(p, "expiry_date", Today.AddYears(1)),
        Amount = DecN(p, "amount"),
        Status = ParseEnum(p, "status", default(VehiclePermitStatus)),
        Remarks = Txt(p, "remarks"),
        ServerUpdatedAt = UpdatedAt(p)
    };

    private static VehicleInsurance MapInsuranceRow(object ins, string plate) => new()
    {
        InsuranceID = Lng(ins, "insurance_id"),
        VehiclePlate = plate,
        InsuranceCompany = Txt(ins, "insurance_company"),
        InsuranceType = ParseEnum(ins, "insurance_type", default(InsuranceType)),
        RenewDate = DateOr(ins, "renew_date", Today),
        ExpiryDate = DateOr(ins, "expiry_date", Today.AddYears(1)),
        Amount = DecN(ins, "amount"),
        Status = ParseEnum(ins, "status", default(InsuranceStatus)),
        Remarks = Txt(ins, "remarks"),
        InsuranceNumber = Txt(ins, "insurance_number"),
        ServerUpdatedAt = UpdatedAt(ins)
    };

    private static OdometerLog MapOdometerRow(object o, string plate)
    {
        // staff_id, with the legacy driven_by column as fallback
        long staffId = ServerSyncHelper.Col(o, "staff_id") != null ? Lng(o, "staff_id") : Lng(o, "driven_by");
        return new OdometerLog
        {
            OdometerID = Lng(o, "odometer_id"),
            Date = DateOr(o, "date", Today),
            VehiclePlate = plate,
            StaffID = staffId,
            Project = Txt(o, "project"),
            OdometerBefore = OdoN(o, "odometer_before"),
            CurrentOdometer = OdoN(o, "current_odometer") ?? 0m,
            FuelCardNumber = Txt(o, "fuel_card_number"),
            FuelCompany = Txt(o, "fuel_company"),
            StationName = Txt(o, "station_name"),
            ReceiptNumber = Txt(o, "receipt_number"),
            Time = Txt(o, "time"),
            FuelType = Txt(o, "fuel_type"),
            PriceAtFueling = Dec(o, "price_at_fueling"),
            FuelAmount = Dec(o, "fuel_amount"),
            Amount = Dec(o, "amount"),
            Remarks = Txt(o, "remarks"),
            ServerUpdatedAt = UpdatedAt(o)
        };
    }

    private static InsuranceClaim MapInsuranceClaimRow(object c, string plate) => new()
    {
        ClaimID = Lng(c, "claim_id"),
        ClaimDate = DateOr(c, "claim_date", Today),
        SettlementDate = DateN(c, "settlement_date"),
        VehiclePlate = plate,
        VehicleType = ParseEnum(c, "vehicle_type", VehicleType.None),
        StaffID = Lng(c, "staff_id"),
        MaintenanceID = Lng(c, "maintenance_id"),
        WorkshopName = Txt(c, "workshop_name"),
        Status = ParseEnum(c, "status", default(InsuranceClaimStatus)),
        ClaimAmount = DecN(c, "claim_amount"),
        Remarks = Txt(c, "remarks"),
        ServerUpdatedAt = UpdatedAt(c)
    };

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

        // Service due (lowest priority): latest non-cancelled service per category as of 'date'.
        var latestServices = (vehicle.Services ?? new List<VehicleService>())
            .Where(s => s.Status != VehicleServiceStatus.Cancelled && s.ServiceDate <= date)
            .GroupBy(s => s.ServiceCategory)
            .Select(g => g.OrderByDescending(s => s.ServiceDate).ThenByDescending(s => s.ServiceID).First())
            .ToList();

        if (latestServices.Any(s => s.NextServiceDate.HasValue && s.NextServiceDate.Value.Year > 1900 && s.NextServiceDate.Value < date))
            return VehicleStatus.Service_Needed;

        if (date >= today)
        {
            decimal curOdo = vehicle.VehicleDetail.CurrentOdometer ?? 0m;
            if (latestServices.Any(s => s.NextServiceOdometer.HasValue && s.NextServiceOdometer.Value > 0m && curOdo >= s.NextServiceOdometer.Value)
                || vehicle.VehicleStatus == VehicleStatus.Service_Needed)
                return VehicleStatus.Service_Needed;
        }

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
        dates.UnionWith(vehicle.Services.Select(s => s.ServiceDate));
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
            decimal? service = SumOrNull(vehicle.Services.Where(s => s.ServiceDate == date && s.Status != VehicleServiceStatus.Cancelled).Select(s => s.Cost));
            decimal? maint = SumOrNull(vehicle.Maintenances.Where(m => m.DateIn == date).Select(m => m.RepairCost));
            decimal? ins = SumOrNull(vehicle.Insurances.Where(i => i.RenewDate == date).Select(i => i.Amount));
            decimal? permit = SumOrNull(vehicle.Permits.Where(p => p.RenewDate == date).Select(p => p.Amount));
            decimal? roadInspect = SumOrNull(vehicle.RoadTaxes.Where(r => r.RenewDate == date).Select(r => r.Amount)
                .Concat(vehicle.Inspections.Where(i => i.ScheduleDateIn == date).Select(i => i.Amount)));
            decimal? claim = SumOrNull(claims.Select(c => c.ClaimAmount));

            decimal? total = null;
            if (fuel.HasValue || service.HasValue || maint.HasValue || ins.HasValue || permit.HasValue || roadInspect.HasValue || claim.HasValue)
                total = (fuel ?? 0m) + (service ?? 0m) + (maint ?? 0m) + (ins ?? 0m) + (permit ?? 0m) + (roadInspect ?? 0m)
                        - ((claim ?? 0m) > 0m ? claim!.Value : 0m);

            rows.Add(new VehicleSummaryRow
            {
                Date = date,
                Status = GetVehicleStatusOn(vehicle, date),
                DrivenBy = JoinDistinct(drivers),
                UsedAt = JoinDistinct(projects),
                OdometerUsed = deltas.Count > 0 ? Math.Round(deltas.Sum(), 1) : null,
                FuelCost = fuel,
                ServiceCost = service,
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

    // Auxiliary directories. Baselines are captured inside the server fetcher, i.e. only when PostgreSQL answered
    // (never from the offline cache that FetchAuthoritativeList falls back to).
    private const string SqlSelectWorkshops = @"
        SELECT
            workshop_id AS WorkshopID,
            workshop_name AS WorkshopName,
            workshop_address AS WorkshopAddress,
            workshop_owner AS WorkshopOwner,
            primary_contact AS PrimaryContact,
            secondary_contact AS SecondaryContact,
            updated_at AS ServerUpdatedAt
        FROM vehicle_workshops";

    private const string SqlSelectInsuranceCompanies = @"
        SELECT
            insurance_company AS InsuranceCompany,
            company_contact_number AS CompanyContactNumber,
            company_address AS CompanyAddress,
            insurance_agent AS InsuranceAgent,
            agent_contact_number AS AgentContactNumber,
            updated_at AS ServerUpdatedAt
        FROM vehicle_insurance_companies";

    private const string SqlSelectFuelCards = @"
        SELECT
            fuel_card_number AS FuelCardNumber,
            fuel_company_name AS FuelCompanyName,
            card_owner_id AS CardOwnerID,
            card_owner_name AS CardOwnerName,
            updated_at AS ServerUpdatedAt
        FROM vehicle_fuel_cards";

    internal static List<Workshop> FetchWorkshops() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "WorkshopsTable", db =>
        {
            var list = db.Query<Workshop>(SqlSelectWorkshops + " ORDER BY workshop_id;").ToList();
            WorkshopBaseline.Capture(list, w => w.WorkshopID, WorkshopFingerprint, w => w.ServerUpdatedAt);
            return list;
        });

    internal static List<InsuranceCompanyRecord> FetchInsuranceCompanies() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "InsuranceCompaniesTable", db =>
        {
            var list = db.Query<InsuranceCompanyRecord>(SqlSelectInsuranceCompanies + " ORDER BY insurance_company;").ToList();
            InsuranceCompanyBaseline.Capture(list, c => TextKey(c.InsuranceCompany), InsuranceCompanyFingerprint, c => c.ServerUpdatedAt);
            return list;
        });

    internal static List<FuelCardRecord> FetchFuelCards() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "FuelCardsTable", db =>
        {
            var list = db.Query<FuelCardRecord>(SqlSelectFuelCards + " ORDER BY fuel_card_number;").ToList();
            FuelCardBaseline.Capture(list, c => TextKey(c.FuelCardNumber), FuelCardFingerprint, c => c.ServerUpdatedAt);
            return list;
        });

    internal static List<string> FetchFuelCompanies() =>
        ServerSyncHelper.FetchAuthoritativeList<FuelCompanyNameRecord>(Module, "FuelCompanyNamesTable", db =>
                db.Query<FuelCompanyNameRecord>("SELECT fuel_company_name AS FuelCompanyName FROM lookup_fuel_companies ORDER BY fuel_company_name;").ToList())
            .Select(c => c.FuelCompanyName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Distinct()
            .ToList();
}
