namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Dapper;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types;

public static partial class ServerDispatcher
{
    private static readonly PostgreSqlService Db = new();

    // =========================================================================
    // 1. UNIFIED DISPATCH ROUTER
    // =========================================================================
    public static void Dispatch<T>(T data)
    {
        switch (data)
        {
            // Staff Manager
            case IEnumerable<StaffRecord> staffBatch:
                PersistStaff(staffBatch.ToList(), isFullBatch: true);
                break;

            // Schedule Manager (one working date + reference tables)
            case ScheduleRecord schedule:
                PersistSchedule(schedule);
                break;

            // Standalone Child Entity Collections
            case IEnumerable<SkillLicenseType> sTypes:
                PersistSkillLicenseTypes(sTypes);
                break;
        }
    }

    // =========================================================================
    // 2. VEHICLE MANAGER PERSISTENCE
    // =========================================================================
    // History-preserving: vehicles, their child rows and the auxiliary directories are never hard-deleted. Vehicles
    // are retired via Sold Date / Status; child rows accumulate as history; lookup tables are append-only.

    /// <summary>Set once the vehicle schema migrations ran and the vehicle sequences were synced in this process.</summary>
    private static volatile bool _vehicleSchemaEnsured;

    /// <summary>Expiry dates of 1900 or earlier (0001-01-01 sentinels) are saved as NULL.</summary>
    private static DateOnly? ValidExpiry(DateOnly? d) => d.HasValue && d.Value.Year > 1900 ? d : null;

    private static string NormalizePlate(string? plate) => (plate ?? string.Empty).Trim().ToUpperInvariant();
    private static DateTime ToDbDate(DateOnly d) => d.ToDateTime(TimeOnly.MinValue);
    private static DateTime? ToDbDateN(DateOnly? d) => d.HasValue ? d.Value.ToDateTime(TimeOnly.MinValue) : null;
    private static decimal RoundOdo(decimal v) => Math.Round(v, 1);
    private static decimal? RoundOdo(decimal? v) => v.HasValue ? Math.Round(v.Value, 1) : null;

    private static bool IsValidVehicle(VehicleRecord v) => !string.IsNullOrWhiteSpace(v.VehiclePlate);
    private static bool IsValidVehicleChild(VehicleChildRecord c) => !string.IsNullOrWhiteSpace(c.VehiclePlate);
    private static bool IsValidWorkshop(Workshop w) => !string.IsNullOrWhiteSpace(w.WorkshopName);
    private static bool IsValidInsuranceCompany(InsuranceCompanyRecord c) => !string.IsNullOrWhiteSpace(c.InsuranceCompany);
    private static bool IsValidFuelCard(FuelCardRecord c) => !string.IsNullOrWhiteSpace(c.FuelCardNumber);

    // -------------------------------------------------------------------------
    // 2a. IDEMPOTENT SCHEMA MIGRATIONS & FORWARD-ONLY SEQUENCES
    // -------------------------------------------------------------------------
    /// <summary>Server-assigned PK sequences (new rows always take <c>nextval</c>, so concurrent users never collide).</summary>
    private static readonly (string Table, string IdColumn, string Sequence)[] VehicleSequences =
    {
        ("vehicle_maintenance", "maintenance_id", "vehicle_maintenance_maintenance_id_seq"),
        ("vehicle_road_taxes", "road_tax_id", "vehicle_road_taxes_road_tax_id_seq"),
        ("vehicle_inspections", "inspection_id", "vehicle_inspections_inspection_id_seq"),
        ("vehicle_permits", "permit_id", "vehicle_permits_permit_id_seq"),
        ("vehicle_insurances", "insurance_id", "vehicle_insurances_insurance_id_seq"),
        ("vehicle_insurance_claims", "claim_id", "vehicle_insurance_claims_claim_id_seq"),
        ("vehicle_odometers", "odometer_id", "vehicle_odometers_odometer_id_seq"),
        ("vehicle_workshops", "workshop_id", "vehicle_workshops_workshop_id_seq")
    };

    private static readonly string[] VehicleConcurrencyTables =
    {
        "vehicle_master", "vehicle_maintenance", "vehicle_road_taxes", "vehicle_inspections", "vehicle_permits",
        "vehicle_insurances", "vehicle_insurance_claims", "vehicle_odometers",
        "vehicle_workshops", "vehicle_insurance_companies", "vehicle_fuel_cards"
    };

    /// <summary>
    /// Idempotent migration: legacy vehicle columns / tables, an <c>updated_at</c> concurrency column on every vehicle
    /// table, and the child PK sequences moved forward past existing PKs (never backwards, so in-flight
    /// <c>nextval</c> values from other sessions stay unique). Runs once per process; returns true when it ran.
    /// </summary>
    private static bool EnsureVehicleSchemaAndSequences(System.Data.IDbConnection conn, System.Data.IDbTransaction trans)
    {
        if (_vehicleSchemaEnsured) return false;

        conn.Execute(@"
CREATE TABLE IF NOT EXISTS vehicle_insurance_claims (
    claim_id        BIGINT PRIMARY KEY,
    claim_date      DATE NOT NULL,
    settlement_date DATE NULL,
    vehicle_plate   VARCHAR(32) NOT NULL,
    vehicle_type    VARCHAR(64) DEFAULT '',
    staff_id        BIGINT DEFAULT 0,
    maintenance_id  BIGINT DEFAULT 0,
    workshop_name   VARCHAR(256) DEFAULT '',
    status          VARCHAR(32) NOT NULL DEFAULT 'Pending',
    claim_amount    NUMERIC(14,2) DEFAULT 0,
    remarks         TEXT DEFAULT '',
    updated_at      TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);", transaction: trans);

        // Odometers / engine-hour meters carry 1 decimal place (NUMERIC(12,1)).
        conn.Execute(@"
DO $$
BEGIN
    IF EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name = 'vehicle_master' AND column_name = 'current_odometer' AND data_type = 'integer') THEN
        ALTER TABLE vehicle_master ALTER COLUMN current_odometer TYPE NUMERIC(12,1) USING current_odometer::numeric(12,1);
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name = 'vehicle_maintenance' AND column_name = 'odometer_in' AND data_type = 'integer') THEN
        ALTER TABLE vehicle_maintenance ALTER COLUMN odometer_in TYPE NUMERIC(12,1) USING odometer_in::numeric(12,1);
        ALTER TABLE vehicle_maintenance ALTER COLUMN odometer_out TYPE NUMERIC(12,1) USING odometer_out::numeric(12,1);
    END IF;
    IF EXISTS (SELECT 1 FROM information_schema.columns WHERE table_name = 'vehicle_odometers' AND column_name = 'current_odometer' AND data_type = 'integer') THEN
        ALTER TABLE vehicle_odometers ALTER COLUMN odometer_before TYPE NUMERIC(12,1) USING odometer_before::numeric(12,1);
        ALTER TABLE vehicle_odometers ALTER COLUMN current_odometer TYPE NUMERIC(12,1) USING current_odometer::numeric(12,1);
    END IF;
END $$;", transaction: trans);

        // No 10-year placeholder expiry defaults; Driven By / Used At / Workshop Address / fuel columns.
        conn.Execute(@"
ALTER TABLE vehicle_master ALTER COLUMN road_tax_expiry DROP DEFAULT;
ALTER TABLE vehicle_master ALTER COLUMN inspection_expiry DROP DEFAULT;
ALTER TABLE vehicle_master ADD COLUMN IF NOT EXISTS driven_by VARCHAR(128) DEFAULT '';
ALTER TABLE vehicle_master ADD COLUMN IF NOT EXISTS used_at   VARCHAR(256) DEFAULT '';
ALTER TABLE vehicle_maintenance ADD COLUMN IF NOT EXISTS workshop_address TEXT DEFAULT '';
ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS time VARCHAR(16) DEFAULT '';
ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS fuel_type VARCHAR(64) DEFAULT '';
ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS price_at_fueling NUMERIC(12,4) DEFAULT 0;
ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS fuel_amount NUMERIC(12,4) DEFAULT 0;", transaction: trans);

        foreach (var table in VehicleConcurrencyTables)
            conn.Execute($"ALTER TABLE {table} ADD COLUMN IF NOT EXISTS updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP;", transaction: trans);

        foreach (var (table, idColumn, sequence) in VehicleSequences)
        {
            conn.Execute($"CREATE SEQUENCE IF NOT EXISTS {sequence};", transaction: trans);
            conn.Execute($@"
                SELECT setval('{sequence}',
                    GREATEST(COALESCE((SELECT MAX({idColumn}) FROM {table}), 0),
                             (SELECT last_value FROM {sequence}), 1),
                    (SELECT COUNT(*) > 0 FROM {table}) OR (SELECT is_called FROM {sequence}));", transaction: trans);
        }

        return true;
    }

    // -------------------------------------------------------------------------
    // 2b. ROW PARAMETER MODELS (DateOnly projected to DateTime; odometers rounded to 1 decimal place)
    // -------------------------------------------------------------------------
    private static object ToVehicleMasterParams(VehicleRecord v, DateTime? expectedUpdatedAt = null)
    {
        var d = v.VehicleDetail ?? new VehicleDetail();
        DateOnly? roadTaxExpiry = ValidExpiry(d.RoadTaxExpiry ?? (v.RoadTaxes ?? new List<RoadTax>()).OrderByDescending(r => r.ExpiryDate).FirstOrDefault()?.ExpiryDate);
        DateOnly? inspectionExpiry = ValidExpiry(d.InspectionExpiry ?? (v.Inspections ?? new List<Inspection>()).OrderByDescending(i => i.ExpiryDate).FirstOrDefault()?.ExpiryDate);
        return new
        {
            VehiclePlate = NormalizePlate(v.VehiclePlate),
            Vin = d.VIN ?? string.Empty,
            VehicleType = d.Type.ToString(),
            Brand = d.Brand ?? string.Empty,
            Model = d.Model ?? string.Empty,
            RoadTaxExpiry = ToDbDateN(roadTaxExpiry),
            InspectionExpiry = ToDbDateN(inspectionExpiry),
            CurrentOdometer = RoundOdo(d.CurrentOdometer),
            BoughtDate = ToDbDateN(d.BoughtDate),
            BoughtPrice = d.BoughtPrice,
            SoldDate = ToDbDateN(d.SoldDate),
            SoldPrice = d.SoldPrice,
            DrivenBy = d.DrivenBy ?? string.Empty,
            UsedAt = d.UsedAt ?? string.Empty,
            Status = v.VehicleStatus.ToString(),
            ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
        };
    }

    private static object ToMaintenanceParams(Maintenance m, DateTime? expectedUpdatedAt = null) => new
    {
        MaintenanceId = m.MaintenanceID,
        VehiclePlate = NormalizePlate(m.VehiclePlate),
        ServiceType = m.ServiceType.ToString(),
        DateIn = ToDbDate(m.DateIn),
        EstimatedReturn = ToDbDate(m.EstimatedReturn),
        DateCompleted = ToDbDateN(m.DateCompleted),
        OdometerIn = RoundOdo(m.OdometerIn),
        OdometerOut = RoundOdo(m.OdometerOut),
        RepairCost = m.RepairCost,
        WorkshopName = m.WorkshopName ?? string.Empty,
        WorkshopAddress = m.WorkshopAddress ?? string.Empty,
        Status = m.MaintenanceStatus.ToString(),
        Remarks = m.Remarks ?? string.Empty,
        TyreCost = m.TyreCost,
        BrakeCost = m.BrakeCost,
        ClutchCost = m.ClutchCost,
        ServiceCost = m.ServiceCost,
        FuelPumpCost = m.FuelPumpCost,
        AircondCost = m.AircondCost,
        WiringCost = m.WiringCost,
        BatteryCost = m.BatteryCost,
        TowingCost = m.TowingCost,
        CushionCost = m.CushionCost,
        PartsCost = m.PartsCost,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToInspectionParams(Inspection i, DateTime? expectedUpdatedAt = null) => new
    {
        InspectionId = i.InspectionID,
        VehiclePlate = NormalizePlate(i.VehiclePlate),
        ScheduleDateIn = ToDbDate(i.ScheduleDateIn),
        DateOut = ToDbDate(i.DateOut),
        ExpiryDate = ToDbDate(i.ExpiryDate),
        Amount = i.Amount,
        Status = i.Status.ToString(),
        Address = i.Address ?? string.Empty,
        Remarks = i.Remarks ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToRoadTaxParams(RoadTax rt, DateTime? expectedUpdatedAt = null) => new
    {
        RoadTaxId = rt.RoadTaxID,
        VehiclePlate = NormalizePlate(rt.VehiclePlate),
        VehicleType = rt.VehicleType.ToString(),
        RenewDate = ToDbDate(rt.RenewDate),
        ExpiryDate = ToDbDate(rt.ExpiryDate),
        Amount = rt.Amount,
        Status = rt.Status.ToString(),
        Remarks = rt.Remarks ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToPermitParams(VehiclePermit p, DateTime? expectedUpdatedAt = null) => new
    {
        PermitId = p.PermitID,
        VehiclePlate = NormalizePlate(p.VehiclePlate),
        VehicleType = p.VehicleType.ToString(),
        PermitType = p.PermitType.ToString(),
        RenewDate = ToDbDate(p.RenewDate),
        ExpiryDate = ToDbDate(p.ExpiryDate),
        Amount = p.Amount,
        Status = p.Status.ToString(),
        Remarks = p.Remarks ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToInsuranceParams(VehicleInsurance ins, DateTime? expectedUpdatedAt = null) => new
    {
        InsuranceId = ins.InsuranceID,
        VehiclePlate = NormalizePlate(ins.VehiclePlate),
        VehicleType = ins.VehicleType.ToString(),
        InsuranceCompany = ins.InsuranceCompany ?? string.Empty,
        InsuranceType = ins.InsuranceType.ToString(),
        InsuranceNumber = ins.InsuranceNumber ?? string.Empty,
        RenewDate = ToDbDate(ins.RenewDate),
        ExpiryDate = ToDbDate(ins.ExpiryDate),
        Amount = ins.Amount,
        Status = ins.Status.ToString(),
        Remarks = ins.Remarks ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToInsuranceClaimParams(InsuranceClaim c, DateTime? expectedUpdatedAt = null) => new
    {
        ClaimId = c.ClaimID,
        ClaimDate = ToDbDate(c.ClaimDate),
        SettlementDate = ToDbDateN(c.SettlementDate),
        VehiclePlate = NormalizePlate(c.VehiclePlate),
        VehicleType = c.VehicleType.ToString(),
        StaffId = c.StaffID,
        MaintenanceId = c.MaintenanceID,
        WorkshopName = c.WorkshopName ?? string.Empty,
        Status = c.Status.ToString(),
        ClaimAmount = c.ClaimAmount,
        Remarks = c.Remarks ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToOdometerParams(OdometerLog o, DateTime? expectedUpdatedAt = null) => new
    {
        OdometerId = o.OdometerID,
        Date = ToDbDate(o.Date),
        VehiclePlate = NormalizePlate(o.VehiclePlate),
        VehicleType = o.VehicleType.ToString(),
        StaffId = o.StaffID,
        Project = o.Project ?? string.Empty,
        OdometerBefore = RoundOdo(o.OdometerBefore),
        CurrentOdometer = RoundOdo(o.CurrentOdometer),
        FuelCardNumber = o.FuelCardNumber ?? string.Empty,
        FuelCompany = o.FuelCompany ?? string.Empty,
        StationName = o.StationName ?? string.Empty,
        ReceiptNumber = o.ReceiptNumber ?? string.Empty,
        Amount = o.Amount,
        Remarks = o.Remarks ?? string.Empty,
        Time = o.Time ?? string.Empty,
        FuelType = o.FuelType ?? string.Empty,
        PriceAtFueling = o.PriceAtFueling,
        FuelAmount = o.FuelAmount,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToWorkshopParams(Workshop w, DateTime? expectedUpdatedAt = null) => new
    {
        WorkshopId = w.WorkshopID,
        WorkshopName = (w.WorkshopName ?? string.Empty).Trim(),
        WorkshopAddress = w.WorkshopAddress ?? string.Empty,
        WorkshopOwner = w.WorkshopOwner ?? string.Empty,
        PrimaryContact = w.PrimaryContact ?? string.Empty,
        SecondaryContact = w.SecondaryContact ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToInsuranceCompanyParams(InsuranceCompanyRecord c, DateTime? expectedUpdatedAt = null) => new
    {
        InsuranceCompany = (c.InsuranceCompany ?? string.Empty).Trim(),
        CompanyContactNumber = c.CompanyContactNumber ?? string.Empty,
        CompanyAddress = c.CompanyAddress ?? string.Empty,
        InsuranceAgent = c.InsuranceAgent ?? string.Empty,
        AgentContactNumber = c.AgentContactNumber ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static object ToFuelCardParams(FuelCardRecord c, DateTime? expectedUpdatedAt = null) => new
    {
        FuelCardNumber = (c.FuelCardNumber ?? string.Empty).Trim(),
        FuelCompanyName = c.FuelCompanyName ?? string.Empty,
        CardOwnerId = c.CardOwnerID ?? string.Empty,
        CardOwnerName = c.CardOwnerName ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    // -------------------------------------------------------------------------
    // 2c. SQL (updated_at always strictly increases so concurrency tokens never repeat)
    // -------------------------------------------------------------------------
    private sealed record VehicleTableSql(string InsertNewId, string Upsert, string InsertIfAbsent, string ConditionalUpdate);

    private static string SnakeToPascal(string snake) =>
        string.Concat(snake.Split('_', StringSplitOptions.RemoveEmptyEntries).Select(p => char.ToUpperInvariant(p[0]) + p.Substring(1)));

    /// <summary>
    /// Builds the four statements for one table keyed by <paramref name="keyColumn"/>. Parameter names are the
    /// PascalCase form of each column (matching the <c>To*Params</c> models). <paramref name="sequence"/> enables
    /// <see cref="VehicleTableSql.InsertNewId"/> (server-assigned PK).
    /// </summary>
    private static VehicleTableSql BuildVehicleTableSql(string table, string keyColumn, string? sequence, params string[] columns)
    {
        string keyParam = "@" + SnakeToPascal(keyColumn);
        string columnList = string.Join(", ", columns);
        string valueList = string.Join(", ", columns.Select(c => "@" + SnakeToPascal(c)));
        string insertHead = $"INSERT INTO {table} ({keyColumn}, {columnList}, updated_at) VALUES (";
        string insertTail = $", {valueList}, clock_timestamp())";
        string insert = insertHead + keyParam + insertTail;

        return new VehicleTableSql(
            InsertNewId: sequence == null ? string.Empty : $"{insertHead}nextval('{sequence}'){insertTail};",
            Upsert: $"{insert} ON CONFLICT ({keyColumn}) DO UPDATE SET " +
                    string.Join(", ", columns.Select(c => $"{c} = EXCLUDED.{c}")) +
                    $", updated_at = GREATEST(clock_timestamp(), {table}.updated_at + INTERVAL '1 millisecond');",
            InsertIfAbsent: $"{insert} ON CONFLICT ({keyColumn}) DO NOTHING;",
            ConditionalUpdate: $"UPDATE {table} SET " +
                    string.Join(", ", columns.Select(c => $"{c} = @{SnakeToPascal(c)}")) +
                    ", updated_at = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond') " +
                    $"WHERE {keyColumn} = {keyParam} AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;");
    }

    private const string SqlVehicleMasterInsert = @"
INSERT INTO vehicle_master (
    vehicle_plate, vin, vehicle_type, brand, model,
    road_tax_expiry, inspection_expiry, current_odometer,
    bought_date, bought_price, sold_date, sold_price,
    driven_by, used_at, status, updated_at
) VALUES (
    @VehiclePlate, @Vin, @VehicleType, @Brand, @Model,
    @RoadTaxExpiry, @InspectionExpiry, @CurrentOdometer,
    @BoughtDate, @BoughtPrice, @SoldDate, @SoldPrice,
    @DrivenBy, @UsedAt, @Status, clock_timestamp()
)";

    private static readonly VehicleTableSql VehicleMasterSql = new(
        InsertNewId: string.Empty,
        Upsert: SqlVehicleMasterInsert + @"
ON CONFLICT (vehicle_plate) DO UPDATE SET
    vin               = COALESCE(NULLIF(EXCLUDED.vin, ''), vehicle_master.vin),
    vehicle_type      = EXCLUDED.vehicle_type,
    brand             = COALESCE(NULLIF(EXCLUDED.brand, ''), vehicle_master.brand),
    model             = COALESCE(NULLIF(EXCLUDED.model, ''), vehicle_master.model),
    road_tax_expiry   = EXCLUDED.road_tax_expiry,
    inspection_expiry = EXCLUDED.inspection_expiry,
    current_odometer  = CASE WHEN EXCLUDED.current_odometer > 0 THEN EXCLUDED.current_odometer ELSE vehicle_master.current_odometer END,
    bought_date       = EXCLUDED.bought_date,
    bought_price      = EXCLUDED.bought_price,
    sold_date         = EXCLUDED.sold_date,
    sold_price        = EXCLUDED.sold_price,
    driven_by         = EXCLUDED.driven_by,
    used_at           = EXCLUDED.used_at,
    status            = EXCLUDED.status,
    updated_at        = GREATEST(clock_timestamp(), vehicle_master.updated_at + INTERVAL '1 millisecond');",
        InsertIfAbsent: SqlVehicleMasterInsert + @"
ON CONFLICT (vehicle_plate) DO NOTHING;",
        ConditionalUpdate: @"
UPDATE vehicle_master SET
    vin               = COALESCE(NULLIF(@Vin, ''), vin),
    vehicle_type      = @VehicleType,
    brand             = COALESCE(NULLIF(@Brand, ''), brand),
    model             = COALESCE(NULLIF(@Model, ''), model),
    road_tax_expiry   = @RoadTaxExpiry,
    inspection_expiry = @InspectionExpiry,
    current_odometer  = CASE WHEN @CurrentOdometer > 0 THEN @CurrentOdometer ELSE current_odometer END,
    bought_date       = @BoughtDate,
    bought_price      = @BoughtPrice,
    sold_date         = @SoldDate,
    sold_price        = @SoldPrice,
    driven_by         = @DrivenBy,
    used_at           = @UsedAt,
    status            = @Status,
    updated_at        = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
WHERE vehicle_plate = @VehiclePlate
  AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;");

    private static readonly VehicleTableSql MaintenanceSql = BuildVehicleTableSql(
        "vehicle_maintenance", "maintenance_id", "vehicle_maintenance_maintenance_id_seq",
        "vehicle_plate", "service_type", "date_in", "estimated_return", "date_completed", "odometer_in", "odometer_out",
        "repair_cost", "workshop_name", "workshop_address", "status", "remarks",
        "tyre_cost", "brake_cost", "clutch_cost", "service_cost", "fuel_pump_cost",
        "aircond_cost", "wiring_cost", "battery_cost", "towing_cost", "cushion_cost", "parts_cost");

    private static readonly VehicleTableSql InspectionSql = BuildVehicleTableSql(
        "vehicle_inspections", "inspection_id", "vehicle_inspections_inspection_id_seq",
        "vehicle_plate", "schedule_date_in", "date_out", "expiry_date", "amount", "status", "address", "remarks");

    private static readonly VehicleTableSql RoadTaxSql = BuildVehicleTableSql(
        "vehicle_road_taxes", "road_tax_id", "vehicle_road_taxes_road_tax_id_seq",
        "vehicle_plate", "vehicle_type", "renew_date", "expiry_date", "amount", "status", "remarks");

    private static readonly VehicleTableSql PermitSql = BuildVehicleTableSql(
        "vehicle_permits", "permit_id", "vehicle_permits_permit_id_seq",
        "vehicle_plate", "vehicle_type", "permit_type", "renew_date", "expiry_date", "amount", "status", "remarks");

    private static readonly VehicleTableSql InsuranceSql = BuildVehicleTableSql(
        "vehicle_insurances", "insurance_id", "vehicle_insurances_insurance_id_seq",
        "vehicle_plate", "vehicle_type", "insurance_company", "insurance_type", "insurance_number",
        "renew_date", "expiry_date", "amount", "status", "remarks");

    private static readonly VehicleTableSql InsuranceClaimSql = BuildVehicleTableSql(
        "vehicle_insurance_claims", "claim_id", "vehicle_insurance_claims_claim_id_seq",
        "claim_date", "settlement_date", "vehicle_plate", "vehicle_type", "staff_id", "maintenance_id",
        "workshop_name", "status", "claim_amount", "remarks");

    private static readonly VehicleTableSql OdometerSql = BuildVehicleTableSql(
        "vehicle_odometers", "odometer_id", "vehicle_odometers_odometer_id_seq",
        "date", "vehicle_plate", "vehicle_type", "staff_id", "project", "odometer_before", "current_odometer",
        "fuel_card_number", "fuel_company", "station_name", "receipt_number", "amount", "remarks",
        "time", "fuel_type", "price_at_fueling", "fuel_amount");

    private static readonly VehicleTableSql WorkshopSql = BuildVehicleTableSql(
        "vehicle_workshops", "workshop_id", "vehicle_workshops_workshop_id_seq",
        "workshop_name", "workshop_address", "workshop_owner", "primary_contact", "secondary_contact");

    private static readonly VehicleTableSql InsuranceCompanySql = BuildVehicleTableSql(
        "vehicle_insurance_companies", "insurance_company", null,
        "company_contact_number", "company_address", "insurance_agent", "agent_contact_number");

    private static readonly VehicleTableSql FuelCardSql = BuildVehicleTableSql(
        "vehicle_fuel_cards", "fuel_card_number", null,
        "fuel_company_name", "card_owner_id", "card_owner_name");

    // -------------------------------------------------------------------------
    // 2d. TABLE SPECS (SQL + Excel locator + conflict diff per vehicle table)
    // -------------------------------------------------------------------------
    private sealed class VehicleTableSpec<T>
    {
        public required string DbTable { get; init; }
        public required string KeyColumn { get; init; }
        public required string Label { get; init; }
        public required string ExcelTable { get; init; }
        public required string KeyHeader { get; init; }
        public required VehicleTableSql Sql { get; init; }
        /// <summary>True = numeric PK from a sequence (new rows take <c>nextval</c>); false = natural key.</summary>
        public bool UsesSequence { get; init; }
        public required Func<T, bool> IsValid { get; init; }
        public required Func<T, object> Key { get; init; }
        public required Func<T, DateTime?, object> ToParams { get; init; }
        public required Func<T, string> Describe { get; init; }
        public required Func<System.Data.IDbConnection, System.Data.IDbTransaction, T, List<FieldConflict>> Diff { get; init; }
    }

    private static readonly VehicleTableSpec<VehicleRecord> VehicleMasterSpec = new()
    {
        DbTable = "vehicle_master", KeyColumn = "vehicle_plate", Label = "Vehicle",
        ExcelTable = "AllVehiclesTable", KeyHeader = "Vehicle Plate", Sql = VehicleMasterSql,
        IsValid = IsValidVehicle, Key = v => NormalizePlate(v.VehiclePlate), ToParams = ToVehicleMasterParams,
        Describe = v => $"{v.VehicleDetail?.Brand} {v.VehicleDetail?.Model}".Trim(),
        Diff = ConflictDiff.ForVehicle
    };

    private static readonly VehicleTableSpec<Maintenance> MaintenanceSpec = new()
    {
        DbTable = "vehicle_maintenance", KeyColumn = "maintenance_id", Label = "Maintenance",
        ExcelTable = "MaintenancesTable", KeyHeader = "Maintenance ID", Sql = MaintenanceSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = m => m.MaintenanceID, ToParams = ToMaintenanceParams,
        Describe = m => $"{NormalizePlate(m.VehiclePlate)}, {m.ServiceType}",
        Diff = ConflictDiff.ForMaintenance
    };

    private static readonly VehicleTableSpec<RoadTax> RoadTaxSpec = new()
    {
        DbTable = "vehicle_road_taxes", KeyColumn = "road_tax_id", Label = "Road Tax",
        ExcelTable = "RoadTaxesTable", KeyHeader = "Road Tax ID", Sql = RoadTaxSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = r => r.RoadTaxID, ToParams = ToRoadTaxParams,
        Describe = r => NormalizePlate(r.VehiclePlate),
        Diff = ConflictDiff.ForRoadTax
    };

    private static readonly VehicleTableSpec<Inspection> InspectionSpec = new()
    {
        DbTable = "vehicle_inspections", KeyColumn = "inspection_id", Label = "Inspection",
        ExcelTable = "InspectionsTable", KeyHeader = "Inspection ID", Sql = InspectionSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = i => i.InspectionID, ToParams = ToInspectionParams,
        Describe = i => NormalizePlate(i.VehiclePlate),
        Diff = ConflictDiff.ForInspection
    };

    private static readonly VehicleTableSpec<VehiclePermit> PermitSpec = new()
    {
        DbTable = "vehicle_permits", KeyColumn = "permit_id", Label = "Permit",
        ExcelTable = "PermitsTable", KeyHeader = "Permit ID", Sql = PermitSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = p => p.PermitID, ToParams = ToPermitParams,
        Describe = p => $"{NormalizePlate(p.VehiclePlate)}, {p.PermitType}",
        Diff = ConflictDiff.ForPermit
    };

    private static readonly VehicleTableSpec<VehicleInsurance> InsuranceSpec = new()
    {
        DbTable = "vehicle_insurances", KeyColumn = "insurance_id", Label = "Insurance Renewal",
        ExcelTable = "InsurancesTable", KeyHeader = "Insurance Renewal ID", Sql = InsuranceSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = i => i.InsuranceID, ToParams = ToInsuranceParams,
        Describe = i => $"{NormalizePlate(i.VehiclePlate)}, {i.InsuranceCompany}".TrimEnd(' ', ','),
        Diff = ConflictDiff.ForInsurance
    };

    private static readonly VehicleTableSpec<InsuranceClaim> InsuranceClaimSpec = new()
    {
        DbTable = "vehicle_insurance_claims", KeyColumn = "claim_id", Label = "Insurance Claim",
        ExcelTable = "InsuranceClaimsTable", KeyHeader = "Insurance Claim ID", Sql = InsuranceClaimSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = c => c.ClaimID, ToParams = ToInsuranceClaimParams,
        Describe = c => NormalizePlate(c.VehiclePlate),
        Diff = ConflictDiff.ForInsuranceClaim
    };

    private static readonly VehicleTableSpec<OdometerLog> OdometerSpec = new()
    {
        DbTable = "vehicle_odometers", KeyColumn = "odometer_id", Label = "Fuel Transaction",
        ExcelTable = "OdometersTable", KeyHeader = "Fuel Transaction ID", Sql = OdometerSql, UsesSequence = true,
        IsValid = IsValidVehicleChild, Key = o => o.OdometerID, ToParams = ToOdometerParams,
        Describe = o => $"{NormalizePlate(o.VehiclePlate)}, {o.Date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}",
        Diff = ConflictDiff.ForOdometer
    };

    private static readonly VehicleTableSpec<Workshop> WorkshopSpec = new()
    {
        DbTable = "vehicle_workshops", KeyColumn = "workshop_id", Label = "Workshop",
        ExcelTable = "WorkshopsTable", KeyHeader = "Workshop ID", Sql = WorkshopSql, UsesSequence = true,
        IsValid = IsValidWorkshop, Key = w => w.WorkshopID, ToParams = ToWorkshopParams,
        Describe = w => (w.WorkshopName ?? string.Empty).Trim(),
        Diff = ConflictDiff.ForWorkshop
    };

    private static readonly VehicleTableSpec<InsuranceCompanyRecord> InsuranceCompanySpec = new()
    {
        DbTable = "vehicle_insurance_companies", KeyColumn = "insurance_company", Label = "Insurance Company",
        ExcelTable = "InsuranceCompaniesTable", KeyHeader = "Insurance Company", Sql = InsuranceCompanySql,
        IsValid = IsValidInsuranceCompany, Key = c => (c.InsuranceCompany ?? string.Empty).Trim(), ToParams = ToInsuranceCompanyParams,
        Describe = _ => string.Empty,
        Diff = ConflictDiff.ForInsuranceCompany
    };

    private static readonly VehicleTableSpec<FuelCardRecord> FuelCardSpec = new()
    {
        DbTable = "vehicle_fuel_cards", KeyColumn = "fuel_card_number", Label = "Fuel Card",
        ExcelTable = "FuelCardsTable", KeyHeader = "Fuel Card Number", Sql = FuelCardSql,
        IsValid = IsValidFuelCard, Key = c => (c.FuelCardNumber ?? string.Empty).Trim(), ToParams = ToFuelCardParams,
        Describe = c => (c.FuelCompanyName ?? string.Empty).Trim(),
        Diff = ConflictDiff.ForFuelCard
    };

    /// <summary>
    /// Applies one table's delta. FallbackUpsert: unconditional upsert by key (sequence tables: rows without an ID take
    /// <c>nextval</c>). Added: sequence tables ALWAYS take <c>nextval</c> (client AutoIncrement placeholders are
    /// ignored); natural-key tables insert-if-absent and report a conflict when the key already exists. Modified:
    /// conditional update on the loaded <c>updated_at</c>; 0 rows = conflict. <paramref name="forceOverwrite"/>
    /// turns every check into an unconditional upsert.
    /// </summary>
    private static void ApplyVehicleTableDelta<T>(
        System.Data.IDbConnection conn, System.Data.IDbTransaction trans,
        VehicleTableSpec<T> spec, TableDelta<T> delta, bool forceOverwrite, List<ConcurrencyConflict> conflicts)
    {
        bool HasKey(T r) => !spec.UsesSequence || (spec.Key(r) is long id && id > 0);

        void AddConflict(T r, string what, DateTime? expected)
        {
            object key = spec.Key(r);
            string raw = Convert.ToString(key, CultureInfo.InvariantCulture) ?? string.Empty;
            string display = key is long ? $"{spec.Label} #{raw}" : $"{spec.Label} {raw}";
            string detail = spec.Describe(r);
            var serverTs = conn.ExecuteScalar<DateTime?>(
                $"SELECT updated_at FROM {spec.DbTable} WHERE {spec.KeyColumn} = @Key;", new { Key = key }, trans);
            conflicts.Add(new ConcurrencyConflict(spec.DbTable, display,
                string.IsNullOrWhiteSpace(detail) ? $"{display} {what}" : $"{display} ({detail}) {what}",
                expected, serverTs,
                ExcelTableName: spec.ExcelTable,
                KeyColumnHeader: spec.KeyHeader,
                RawKeyValue: raw,
                FieldDiffs: spec.Diff(conn, trans, r)));
        }

        // a. No baseline: batched unconditional upsert.
        var fallback = delta.FallbackUpsert.Where(spec.IsValid).ToList();
        var fallbackKeyed = fallback.Where(HasKey).Select(r => spec.ToParams(r, null)).ToList();
        var fallbackNew = fallback.Where(r => !HasKey(r)).Select(r => spec.ToParams(r, null)).ToList();
        if (fallbackKeyed.Count > 0) conn.Execute(spec.Sql.Upsert, fallbackKeyed, trans);
        if (fallbackNew.Count > 0) conn.Execute(spec.Sql.InsertNewId, fallbackNew, trans);

        // b. Added rows.
        var added = delta.Added.Where(spec.IsValid).ToList();
        if (spec.UsesSequence)
        {
            var rows = added.Select(r => spec.ToParams(r, null)).ToList();
            if (rows.Count > 0) conn.Execute(spec.Sql.InsertNewId, rows, trans);
        }
        else
        {
            foreach (var r in added)
            {
                if (forceOverwrite)
                    conn.Execute(spec.Sql.Upsert, spec.ToParams(r, null), trans);
                else if (conn.Execute(spec.Sql.InsertIfAbsent, spec.ToParams(r, null), trans) == 0)
                    AddConflict(r, "already exists on the server (added by another user).", null);
            }
        }

        // c. Modified rows (optimistic concurrency on updated_at).
        foreach (var item in delta.Modified.Where(m => spec.IsValid(m.Record)))
        {
            var r = item.Record;
            if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
            {
                if (conn.Execute(spec.Sql.ConditionalUpdate, spec.ToParams(r, item.ExpectedUpdatedAt), trans) == 0)
                    AddConflict(r, "was modified on the server.", item.ExpectedUpdatedAt);
            }
            else
            {
                conn.Execute(HasKey(r) ? spec.Sql.Upsert : spec.Sql.InsertNewId, spec.ToParams(r, null), trans);
            }
        }
    }

    private static string[] NormalizeFuelCompanies(IEnumerable<string>? names) =>
        names?
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray() ?? Array.Empty<string>();

    /// <summary>Append-only: fuel companies are never deleted (historical cards and fuel transactions keep their company).</summary>
    private static void InsertFuelCompanies(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, string[] names)
    {
        if (names.Length == 0) return;
        conn.Execute(@"
            INSERT INTO lookup_fuel_companies (fuel_company_name)
            VALUES (@Name)
            ON CONFLICT (fuel_company_name) DO NOTHING;", names.Select(n => new { Name = n }).ToList(), trans);
    }

    /// <summary>Child rows inherit the owning vehicle's plate when the sheet left it blank.</summary>
    private static List<TChild> CollectVehicleChildren<TChild>(IEnumerable<VehicleRecord> vehicles, Func<VehicleRecord, IEnumerable<TChild>?> selector)
        where TChild : VehicleChildRecord
    {
        var list = new List<TChild>();
        foreach (var v in vehicles)
        {
            foreach (var c in selector(v) ?? Enumerable.Empty<TChild>())
            {
                if (string.IsNullOrWhiteSpace(c.VehiclePlate)) c.VehiclePlate = v.VehiclePlate;
                list.Add(c);
            }
        }
        return list;
    }

    private static TableDelta<T> UpsertOnly<T>(IEnumerable<T> rows) => new() { HasBaseline = false, FallbackUpsert = rows.ToList() };

    // -------------------------------------------------------------------------
    // 2e. UPSERT-ONLY PERSISTENCE (offline pending upload / no-baseline fallback; never hard-deletes)
    // -------------------------------------------------------------------------
    /// <summary>
    /// Upsert-only vehicle save (no deletes, no concurrency checks): vehicles / child rows missing from the workbook
    /// are kept on the server. <paramref name="isFullBatch"/> is kept for call-site compatibility only.
    /// </summary>
    public static void PersistVehicles(IList<VehicleRecord> vehicles, bool isFullBatch)
    {
        _ = isFullBatch;
        var valid = vehicles.Where(IsValidVehicle).ToList();
        if (valid.Count == 0) return;

        var conflicts = new List<ConcurrencyConflict>();
        bool ensured = false;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            ensured = EnsureVehicleSchemaAndSequences(conn, trans);

            ApplyVehicleTableDelta(conn, trans, VehicleMasterSpec, UpsertOnly(valid), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, MaintenanceSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.Maintenances)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, RoadTaxSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.RoadTaxes)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, InspectionSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.Inspections)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, PermitSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.Permits)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.Insurances)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceClaimSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.InsuranceClaims)), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, OdometerSpec, UpsertOnly(CollectVehicleChildren(valid, v => v.OdometerLogs)), true, conflicts);
        });
        if (ensured) _vehicleSchemaEnsured = true;

        AppLogger.Info($"[ServerDispatcher] Vehicles upserted (no deletes): {valid.Count} vehicle(s).");
    }

    /// <summary>Upsert-only auxiliary save: workshops, insurance companies and fuel cards are never deleted; fuel companies are append-only.</summary>
    public static void SaveAuxiliaryTables(
        IEnumerable<Workshop> workshops,
        IEnumerable<InsuranceCompanyRecord> insuranceCompanies,
        IEnumerable<FuelCardRecord> fuelCards,
        IEnumerable<string> fuelCompanies)
    {
        var companyNames = NormalizeFuelCompanies(fuelCompanies);
        var workshopList = workshops?.ToList() ?? new List<Workshop>();
        var insuranceList = insuranceCompanies?.ToList() ?? new List<InsuranceCompanyRecord>();
        var cardList = fuelCards?.ToList() ?? new List<FuelCardRecord>();

        var conflicts = new List<ConcurrencyConflict>();
        bool ensured = false;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            ensured = EnsureVehicleSchemaAndSequences(conn, trans);

            InsertFuelCompanies(conn, trans, companyNames);
            ApplyVehicleTableDelta(conn, trans, WorkshopSpec, UpsertOnly(workshopList), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceCompanySpec, UpsertOnly(insuranceList), true, conflicts);
            ApplyVehicleTableDelta(conn, trans, FuelCardSpec, UpsertOnly(cardList), true, conflicts);
        });
        if (ensured) _vehicleSchemaEnsured = true;
    }

    // -------------------------------------------------------------------------
    // 2f. DELTA SAVE WITH OPTIMISTIC CONCURRENCY
    // -------------------------------------------------------------------------
    /// <summary>
    /// Writes only Added / Modified rows of every vehicle table in ONE transaction (lookup + auxiliary directories
    /// first, then vehicle_master, then the 7 child tables). Modified rows are updated only when the server
    /// <c>updated_at</c> still matches the loaded baseline; any mismatch (or a natural key added concurrently by
    /// another user) is collected and raised as <see cref="ConcurrencyConflictException"/>, rolling back the entire
    /// batch. New child / workshop rows always take a server sequence ID. <paramref name="forceOverwrite"/> skips
    /// the checks and upserts unconditionally. Never hard-deletes.
    /// </summary>
    public static void PersistVehicleDelta(VehicleDeltaPayload delta, bool forceOverwrite = false)
    {
        var fuelCompanies = NormalizeFuelCompanies(delta.FuelCompanies);

        bool ensured = false;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            ensured = EnsureVehicleSchemaAndSequences(conn, trans);
            var conflicts = new List<ConcurrencyConflict>();

            // a. Lookup + auxiliary directories (fuel companies before fuel cards).
            InsertFuelCompanies(conn, trans, fuelCompanies);
            ApplyVehicleTableDelta(conn, trans, WorkshopSpec, delta.Workshops, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceCompanySpec, delta.InsuranceCompanies, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, FuelCardSpec, delta.FuelCards, forceOverwrite, conflicts);

            // b. vehicle_master, then its child tables.
            ApplyVehicleTableDelta(conn, trans, VehicleMasterSpec, delta.Vehicles, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, MaintenanceSpec, delta.Maintenances, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, RoadTaxSpec, delta.RoadTaxes, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, InspectionSpec, delta.Inspections, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, PermitSpec, delta.Permits, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceSpec, delta.Insurances, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, InsuranceClaimSpec, delta.InsuranceClaims, forceOverwrite, conflicts);
            ApplyVehicleTableDelta(conn, trans, OdometerSpec, delta.OdometerLogs, forceOverwrite, conflicts);

            // c. Any conflict rolls back the whole batch (ExecuteInTransaction catches, rolls back, rethrows).
            if (conflicts.Count > 0) throw new ConcurrencyConflictException(conflicts);
        });
        if (ensured) _vehicleSchemaEnsured = true;

        static string Stats<T>(string name, TableDelta<T> d) =>
            $"{name} +{d.Added.Count}/~{d.Modified.Count}/fallback {d.FallbackUpsert.Count}/={d.UnchangedCount}";

        AppLogger.Info($"[ServerDispatcher] Vehicle delta saved (force={forceOverwrite}): " + string.Join("; ",
            Stats("vehicles", delta.Vehicles),
            Stats("maintenance", delta.Maintenances),
            Stats("road tax", delta.RoadTaxes),
            Stats("inspections", delta.Inspections),
            Stats("permits", delta.Permits),
            Stats("insurances", delta.Insurances),
            Stats("claims", delta.InsuranceClaims),
            Stats("odometers", delta.OdometerLogs),
            Stats("workshops", delta.Workshops),
            Stats("insurers", delta.InsuranceCompanies),
            Stats("fuel cards", delta.FuelCards)) + $"; fuel companies {fuelCompanies.Length}.");
    }

    // =========================================================================
    // 3. STAFF MANAGER PERSISTENCE
    // =========================================================================
    // History-preserving: staff, licenses and skill types are never hard-deleted. Staff are deactivated via
    // Resign Date / Contract Status; licenses accumulate as history rows (one row per renewal).

    /// <summary>Set once the legacy unique constraints are dropped and sequences synced in this process.</summary>
    private static volatile bool _staffHistoryConstraintsEnsured;

    /// <summary>
    /// Upsert-only staff save (no deletes, no concurrency checks). Used for the offline pending-upload commit where
    /// no server baseline exists. <paramref name="isFullBatch"/> is kept for call-site compatibility only.
    /// </summary>
    public static void PersistStaff(IList<StaffRecord> staffList, bool isFullBatch, IEnumerable<SkillLicenseType>? skillLicenseTypes = null)
    {
        _ = isFullBatch;
        var staffItems = staffList.Where(s => s.StaffID > 0).ToList();
        if (!staffItems.Any()) return;

        var typeItems = NormalizeSkillLicenseTypes(skillLicenseTypes);

        var allVLicenses = staffItems.SelectMany(s => s.VehicleLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList();
        var allSLicenses = staffItems.SelectMany(s => s.SkillLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList();
        var allSalaries = staffItems.SelectMany(s => (s.Salaries ?? new List<StaffSalary>()).Select(w => { if (w.StaffID == 0) w.StaffID = s.StaffID; return w; })).ToList();
        var allLeaves = staffItems.SelectMany(s => (s.Leaves ?? new List<StaffLeave>()).Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList();

        bool ensured = false;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            ensured = EnsureStaffHistoryConstraints(conn, trans);

            // a. Lookup types first so license FK references stay valid.
            if (typeItems.Count > 0) UpsertSkillLicenseTypes(conn, trans, typeItems);

            // b. Masters, then child licenses and salary history.
            UpsertStaffRows(conn, trans, staffItems);
            PersistStaffVehicleLicenses(conn, trans, allVLicenses);
            PersistStaffSkillLicenses(conn, trans, allSLicenses);
            PersistStaffSalaries(conn, trans, allSalaries);
            PersistStaffLeaves(conn, trans, allLeaves);
        });
        if (ensured) _staffHistoryConstraintsEnsured = true;
    }

    /// <summary>
    /// Idempotent migration: drops the legacy one-license-per-class/type unique constraints (licenses are now history
    /// rows) and moves the license sequences forward past existing PKs (never backwards, so in-flight
    /// <c>nextval</c> values from other sessions stay unique). Runs once per process; returns true when it ran.
    /// </summary>
    private static bool EnsureStaffHistoryConstraints(System.Data.IDbConnection conn, System.Data.IDbTransaction trans)
    {
        if (_staffHistoryConstraintsEnsured) return false;

        conn.Execute(@"
            ALTER TABLE staff_vehicle_licenses DROP CONSTRAINT IF EXISTS staff_vehicle_licenses_staff_id_vehicle_license_class_key;
            ALTER TABLE staff_vehicle_licenses DROP CONSTRAINT IF EXISTS uq_staff_vehicle_license;
            DROP INDEX IF EXISTS uq_staff_vehicle_license;
            ALTER TABLE staff_skill_licenses DROP CONSTRAINT IF EXISTS staff_skill_licenses_staff_id_skill_license_type_key;
            ALTER TABLE staff_skill_licenses DROP CONSTRAINT IF EXISTS uq_staff_skill_license;
            DROP INDEX IF EXISTS uq_staff_skill_license;", transaction: trans);

        conn.Execute(@"
            SELECT setval('staff_vehicle_licenses_vehicle_license_id_seq',
                GREATEST(COALESCE((SELECT MAX(vehicle_license_id) FROM staff_vehicle_licenses), 0),
                         (SELECT last_value FROM staff_vehicle_licenses_vehicle_license_id_seq), 1),
                (SELECT COUNT(*) > 0 FROM staff_vehicle_licenses) OR (SELECT is_called FROM staff_vehicle_licenses_vehicle_license_id_seq));
            SELECT setval('staff_skill_licenses_skill_license_id_seq',
                GREATEST(COALESCE((SELECT MAX(skill_license_id) FROM staff_skill_licenses), 0),
                         (SELECT last_value FROM staff_skill_licenses_skill_license_id_seq), 1),
                (SELECT COUNT(*) > 0 FROM staff_skill_licenses) OR (SELECT is_called FROM staff_skill_licenses_skill_license_id_seq));", transaction: trans);

        // Salary history synced from SQL Payroll (one row per effective salary revision).
        conn.Execute(@"
            CREATE TABLE IF NOT EXISTS staff_salaries (
                salary_id           BIGSERIAL PRIMARY KEY,
                staff_id            BIGINT NOT NULL REFERENCES staff_master(staff_id),
                effective_date      DATE NOT NULL,
                salary_before       NUMERIC(12,2) NULL,
                salary_before_basis VARCHAR(32) NOT NULL DEFAULT 'None',
                salary_basis        VARCHAR(32) NOT NULL DEFAULT 'Monthly',
                salary_rate         NUMERIC(12,2) NOT NULL DEFAULT 0,
                remarks             TEXT DEFAULT '',
                updated_at          TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE (staff_id, effective_date, salary_basis, salary_rate)
            );", transaction: trans);

        // One-time carry-over from the legacy staff_wages table (left in place, never dropped).
        conn.Execute(@"
            DO $$
            BEGIN
                IF to_regclass('public.staff_wages') IS NOT NULL AND NOT EXISTS (SELECT 1 FROM staff_salaries) THEN
                    INSERT INTO staff_salaries (salary_id, staff_id, effective_date, salary_basis, salary_rate, remarks, updated_at)
                    SELECT wage_id, staff_id, effective_date, wage_basis, wage_rate, COALESCE(remarks, ''), updated_at
                    FROM staff_wages
                    ON CONFLICT DO NOTHING;
                END IF;
            END $$;", transaction: trans);

        SyncStaffSalarySequence(conn, trans);

        // Leave applications mirrored from SQL Payroll LV_APP (same schema as the Schedule calendar cache).
        conn.Execute(@"
            CREATE TABLE IF NOT EXISTS staff_leaves (
                leave_id       BIGSERIAL PRIMARY KEY,
                payroll_dockey BIGINT UNIQUE,
                staff_id       BIGINT NOT NULL,
                leave_date     DATE NOT NULL,
                leave_type     VARCHAR(64) NOT NULL,
                leave_day      NUMERIC(5,2) NOT NULL DEFAULT 1.0,
                description    TEXT DEFAULT '',
                updated_at     TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            );
            CREATE INDEX IF NOT EXISTS ix_staff_leaves_leave_date ON staff_leaves (leave_date);", transaction: trans);
        return true;
    }

    /// <summary>Moves <c>staff_salaries_salary_id_seq</c> forward past MAX(salary_id) (never backwards).</summary>
    private static void SyncStaffSalarySequence(System.Data.IDbConnection conn, System.Data.IDbTransaction trans)
    {
        conn.Execute(@"
            SELECT setval('staff_salaries_salary_id_seq',
                GREATEST(COALESCE((SELECT MAX(salary_id) FROM staff_salaries), 0),
                         (SELECT last_value FROM staff_salaries_salary_id_seq), 1),
                (SELECT COUNT(*) > 0 FROM staff_salaries) OR (SELECT is_called FROM staff_salaries_salary_id_seq));", transaction: trans);
    }

    // -------------------------------------------------------------------------
    // 3a. ROW PARAMETER MODELS (DateOnly projected to DateTime for Dapper)
    // -------------------------------------------------------------------------
    private sealed class StaffRow
    {
        public long StaffId { get; init; }
        public string StaffName { get; init; } = string.Empty;
        public string ShortName { get; init; } = string.Empty;
        public string Department { get; init; } = "Worker";
        public long Nric { get; init; }
        public string Passport { get; init; } = string.Empty;
        public string Phone { get; init; } = string.Empty;
        public DateTime JoinDate { get; init; }
        public DateTime? ResignDate { get; init; }
        public string ContractStatus { get; init; } = string.Empty;
        public DateTime? ExpectedUpdatedAt { get; init; }
    }

    private sealed class VehicleLicenseRow
    {
        public long VehicleLicenseId { get; init; }
        public long StaffId { get; init; }
        public string VehicleLicenseClass { get; init; } = string.Empty;
        public DateTime RenewDate { get; init; }
        public DateTime ExpiryDate { get; init; }
        public string Status { get; init; } = string.Empty;
        public decimal Cost { get; init; }
        public DateTime? ExpectedUpdatedAt { get; init; }
    }

    private sealed class SkillLicenseRow
    {
        public long SkillLicenseId { get; init; }
        public long StaffId { get; init; }
        public string SkillLicenseType { get; init; } = string.Empty;
        public DateTime RenewDate { get; init; }
        public DateTime ExpiryDate { get; init; }
        public string Status { get; init; } = string.Empty;
        public decimal Cost { get; init; }
        public DateTime? ExpectedUpdatedAt { get; init; }
    }

    /// <summary>
    /// Concurrency token as sent back to PostgreSQL. Npgsql reads <c>timestamptz</c> as UTC and <c>timestamp</c> as
    /// Unspecified; keeping that Kind round-trips the token with the column's own type (only Local is normalized).
    /// </summary>
    private static DateTime? ToConcurrencyToken(DateTime? d) =>
        d.HasValue && d.Value.Kind == DateTimeKind.Local ? DateTime.SpecifyKind(d.Value.ToUniversalTime(), DateTimeKind.Utc) : d;

    private static StaffRow ToStaffRow(StaffRecord s, DateTime? expectedUpdatedAt = null) => new()
    {
        StaffId = s.StaffID,
        StaffName = s.StaffDetails.StaffName ?? string.Empty,
        ShortName = s.StaffDetails.StaffShortName ?? string.Empty,
        Department = s.StaffDetails.Department ?? "Worker",
        Nric = s.StaffDetails.NRIC,
        Passport = s.StaffDetails.Passport ?? string.Empty,
        Phone = s.StaffDetails.Phone ?? string.Empty,
        JoinDate = s.StaffDetails.JoinDate.ToDateTime(TimeOnly.MinValue),
        ResignDate = s.StaffDetails.ResignDate.HasValue ? s.StaffDetails.ResignDate.Value.ToDateTime(TimeOnly.MinValue) : null,
        ContractStatus = s.StaffDetails.ContractStatus.ToString(),
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static VehicleLicenseRow ToVehicleLicenseRow(VehicleLicense l, DateTime? expectedUpdatedAt = null) => new()
    {
        VehicleLicenseId = l.VehicleLicenseID,
        StaffId = l.StaffID,
        VehicleLicenseClass = l.VehicleLicenseClass.ToString(),
        RenewDate = l.RenewDate.ToDateTime(TimeOnly.MinValue),
        ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
        Status = l.VehicleLicenseStatus.ToString(),
        Cost = l.Cost,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static SkillLicenseRow ToSkillLicenseRow(SkillLicense l, DateTime? expectedUpdatedAt = null) => new()
    {
        SkillLicenseId = l.SkillLicenseID,
        StaffId = l.StaffID,
        SkillLicenseType = (l.SkillLicenseType ?? string.Empty).Trim(),
        RenewDate = l.RenewDate.ToDateTime(TimeOnly.MinValue),
        ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
        Status = l.SkillLicenseStatus.ToString(),
        Cost = l.Cost,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private sealed class StaffSalaryRow
    {
        public long SalaryId { get; init; }
        public long StaffId { get; init; }
        public DateTime EffectiveDate { get; init; }
        public decimal? SalaryBefore { get; init; }
        public string SalaryBeforeBasis { get; init; } = "None";
        public string SalaryBasis { get; init; } = "Monthly";
        public decimal SalaryRate { get; init; }
        public string Remarks { get; init; } = string.Empty;
    }

    private static StaffSalaryRow ToStaffSalaryRow(StaffSalary s, bool ignoreLocalId = false) => new()
    {
        SalaryId = ignoreLocalId ? 0 : s.SalaryID,
        StaffId = s.StaffID,
        EffectiveDate = s.EffectiveDate.ToDateTime(TimeOnly.MinValue),
        SalaryBefore = s.SalaryBefore.HasValue ? decimal.Round(s.SalaryBefore.Value, 2, MidpointRounding.AwayFromZero) : null,
        SalaryBeforeBasis = (s.SalaryBefore.HasValue ? s.SalaryBeforeBasis : SalaryBasis.None).ToString(),
        SalaryBasis = s.SalaryBasis.ToString(),
        SalaryRate = decimal.Round(s.SalaryRate, 2, MidpointRounding.AwayFromZero),
        Remarks = (s.Remarks ?? string.Empty).Trim()
    };

    private static bool IsValidStaffSalary(StaffSalary s) => s.StaffID > 0 && s.SalaryBasis != SalaryBasis.None && s.SalaryRate > 0m;

    private static bool IsValidVehicleLicense(VehicleLicense l) => l.StaffID > 0 && l.VehicleLicenseClass != VehicleLicenseClass.None;
    private static bool IsValidSkillLicense(SkillLicense l) => l.StaffID > 0 && !string.IsNullOrWhiteSpace(l.SkillLicenseType);

    // -------------------------------------------------------------------------
    // 3b. SQL (updated_at always strictly increases so concurrency tokens never repeat)
    // -------------------------------------------------------------------------
    private const string SqlStaffInsert = @"
        INSERT INTO staff_master (
            staff_id, staff_name, short_name, department, nric, passport, phone, join_date, resign_date, contract_status, updated_at
        )
        VALUES (
            @StaffId, @StaffName, @ShortName, @Department, @Nric, @Passport, @Phone, @JoinDate, @ResignDate, @ContractStatus, clock_timestamp()
        )";

    private const string SqlStaffUpsert = SqlStaffInsert + @"
        ON CONFLICT (staff_id) DO UPDATE SET
            staff_name      = EXCLUDED.staff_name,
            short_name      = EXCLUDED.short_name,
            department      = EXCLUDED.department,
            nric            = EXCLUDED.nric,
            passport        = EXCLUDED.passport,
            phone           = EXCLUDED.phone,
            join_date       = EXCLUDED.join_date,
            resign_date     = EXCLUDED.resign_date,
            contract_status = EXCLUDED.contract_status,
            updated_at      = GREATEST(clock_timestamp(), staff_master.updated_at + INTERVAL '1 millisecond');";

    private const string SqlStaffInsertIfAbsent = SqlStaffInsert + @"
        ON CONFLICT (staff_id) DO NOTHING;";

    private const string SqlStaffConditionalUpdate = @"
        UPDATE staff_master SET
            staff_name      = @StaffName,
            short_name      = @ShortName,
            department      = @Department,
            nric            = @Nric,
            passport        = @Passport,
            phone           = @Phone,
            join_date       = @JoinDate,
            resign_date     = @ResignDate,
            contract_status = @ContractStatus,
            updated_at      = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE staff_id = @StaffId
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    private const string SqlVehicleLicenseInsertNewId = @"
        INSERT INTO staff_vehicle_licenses (vehicle_license_id, staff_id, vehicle_license_class, renew_date, expiry_date, status, cost, updated_at)
        VALUES (nextval('staff_vehicle_licenses_vehicle_license_id_seq'), @StaffId, @VehicleLicenseClass, @RenewDate, @ExpiryDate, @Status, @Cost, clock_timestamp());";

    private const string SqlVehicleLicenseUpsert = @"
        INSERT INTO staff_vehicle_licenses (vehicle_license_id, staff_id, vehicle_license_class, renew_date, expiry_date, status, cost, updated_at)
        VALUES (@VehicleLicenseId, @StaffId, @VehicleLicenseClass, @RenewDate, @ExpiryDate, @Status, @Cost, clock_timestamp())
        ON CONFLICT (vehicle_license_id) DO UPDATE SET
            staff_id              = EXCLUDED.staff_id,
            vehicle_license_class = EXCLUDED.vehicle_license_class,
            renew_date            = EXCLUDED.renew_date,
            expiry_date           = EXCLUDED.expiry_date,
            status                = EXCLUDED.status,
            cost                  = EXCLUDED.cost,
            updated_at            = GREATEST(clock_timestamp(), staff_vehicle_licenses.updated_at + INTERVAL '1 millisecond');";

    private const string SqlVehicleLicenseConditionalUpdate = @"
        UPDATE staff_vehicle_licenses SET
            staff_id              = @StaffId,
            vehicle_license_class = @VehicleLicenseClass,
            renew_date            = @RenewDate,
            expiry_date           = @ExpiryDate,
            status                = @Status,
            cost                  = @Cost,
            updated_at            = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE vehicle_license_id = @VehicleLicenseId
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    private const string SqlSkillLicenseInsertNewId = @"
        INSERT INTO staff_skill_licenses (skill_license_id, staff_id, skill_license_type, renew_date, expiry_date, status, cost, updated_at)
        VALUES (nextval('staff_skill_licenses_skill_license_id_seq'), @StaffId, @SkillLicenseType, @RenewDate, @ExpiryDate, @Status, @Cost, clock_timestamp());";

    private const string SqlSkillLicenseUpsert = @"
        INSERT INTO staff_skill_licenses (skill_license_id, staff_id, skill_license_type, renew_date, expiry_date, status, cost, updated_at)
        VALUES (@SkillLicenseId, @StaffId, @SkillLicenseType, @RenewDate, @ExpiryDate, @Status, @Cost, clock_timestamp())
        ON CONFLICT (skill_license_id) DO UPDATE SET
            staff_id           = EXCLUDED.staff_id,
            skill_license_type = EXCLUDED.skill_license_type,
            renew_date         = EXCLUDED.renew_date,
            expiry_date        = EXCLUDED.expiry_date,
            status             = EXCLUDED.status,
            cost               = EXCLUDED.cost,
            updated_at         = GREATEST(clock_timestamp(), staff_skill_licenses.updated_at + INTERVAL '1 millisecond');";

    private const string SqlSkillLicenseConditionalUpdate = @"
        UPDATE staff_skill_licenses SET
            staff_id           = @StaffId,
            skill_license_type = @SkillLicenseType,
            renew_date         = @RenewDate,
            expiry_date        = @ExpiryDate,
            status             = @Status,
            cost               = @Cost,
            updated_at         = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE skill_license_id = @SkillLicenseId
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    private const string SqlStaffSalaryInsertNewId = @"
        INSERT INTO staff_salaries (salary_id, staff_id, effective_date, salary_before, salary_before_basis, salary_basis, salary_rate, remarks, updated_at)
        VALUES (nextval('staff_salaries_salary_id_seq'), @StaffId, @EffectiveDate, @SalaryBefore, @SalaryBeforeBasis, @SalaryBasis, @SalaryRate, @Remarks, clock_timestamp())
        ON CONFLICT (staff_id, effective_date, salary_basis, salary_rate) DO UPDATE SET
            salary_before       = EXCLUDED.salary_before,
            salary_before_basis = EXCLUDED.salary_before_basis,
            remarks             = EXCLUDED.remarks,
            updated_at          = GREATEST(clock_timestamp(), staff_salaries.updated_at + INTERVAL '1 millisecond');";

    private const string SqlStaffSalaryUpsertById = @"
        INSERT INTO staff_salaries (salary_id, staff_id, effective_date, salary_before, salary_before_basis, salary_basis, salary_rate, remarks, updated_at)
        VALUES (@SalaryId, @StaffId, @EffectiveDate, @SalaryBefore, @SalaryBeforeBasis, @SalaryBasis, @SalaryRate, @Remarks, clock_timestamp())
        ON CONFLICT (salary_id) DO UPDATE SET
            staff_id            = EXCLUDED.staff_id,
            effective_date      = EXCLUDED.effective_date,
            salary_before       = EXCLUDED.salary_before,
            salary_before_basis = EXCLUDED.salary_before_basis,
            salary_basis        = EXCLUDED.salary_basis,
            salary_rate         = EXCLUDED.salary_rate,
            remarks             = EXCLUDED.remarks,
            updated_at          = GREATEST(clock_timestamp(), staff_salaries.updated_at + INTERVAL '1 millisecond');";

    // -------------------------------------------------------------------------
    // 3c. UPSERT-ONLY HELPERS (offline pending upload / no-baseline fallback)
    // -------------------------------------------------------------------------
    /// <summary>
    /// Upserts salary history. Rows with a server <c>salary_id</c> upsert by ID; new rows take a sequence ID and merge
    /// on the natural key (staff, effective date, basis, rate), so re-syncing SQL Payroll never duplicates history.
    /// </summary>
    private static void PersistStaffSalaries(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, IEnumerable<StaffSalary> salaries, bool ignoreLocalIds = false)
    {
        var items = salaries.Where(IsValidStaffSalary)
            .Select(s => ToStaffSalaryRow(s, ignoreLocalIds))
            .GroupBy(r => (r.SalaryId > 0 ? r.SalaryId.ToString(CultureInfo.InvariantCulture) : string.Empty, r.StaffId, r.EffectiveDate, r.SalaryBasis, r.SalaryRate))
            .Select(g => g.Last())
            .ToList();
        var withId = items.Where(i => i.SalaryId > 0).ToList();
        var withoutId = items.Where(i => i.SalaryId <= 0).ToList();

        if (withId.Count > 0)
        {
            conn.Execute(SqlStaffSalaryUpsertById, withId, trans);
            SyncStaffSalarySequence(conn, trans); // explicit IDs must never be handed out again by nextval
        }
        if (withoutId.Count > 0) conn.Execute(SqlStaffSalaryInsertNewId, withoutId, trans);
    }

    private const string SqlStaffLeaveUpsertByDocKey = @"
        INSERT INTO staff_leaves (payroll_dockey, staff_id, leave_date, leave_type, leave_day, description, updated_at)
        VALUES (@DocKey, @StaffId, @LeaveDate, @LeaveType, @LeaveDay, @Description, clock_timestamp())
        ON CONFLICT (payroll_dockey) DO UPDATE SET
            staff_id    = EXCLUDED.staff_id,
            leave_date  = EXCLUDED.leave_date,
            leave_type  = EXCLUDED.leave_type,
            leave_day   = EXCLUDED.leave_day,
            description = EXCLUDED.description,
            updated_at  = GREATEST(clock_timestamp(), staff_leaves.updated_at + INTERVAL '1 millisecond');";

    /// <summary>
    /// Upserts leave rows mirrored from SQL Payroll <c>LV_APP</c> by <c>payroll_dockey</c> (= <see cref="StaffLeave.LeaveID"/>).
    /// SQL Payroll is authoritative, so no concurrency checks and no deletes.
    /// </summary>
    private static void PersistStaffLeaves(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, IEnumerable<StaffLeave> leaves)
    {
        var rows = leaves
            .Where(l => l != null && l.LeaveID > 0 && l.StaffID > 0)
            .GroupBy(l => l.LeaveID)
            .Select(g => g.Last())
            .Select(l => new
            {
                DocKey = l.LeaveID,
                StaffId = l.StaffID,
                LeaveDate = l.LeaveDate.ToDateTime(TimeOnly.MinValue),
                LeaveType = (l.LeaveType ?? string.Empty).Trim(),
                l.LeaveDay,
                Description = (l.Description ?? string.Empty).Trim()
            })
            .ToList();
        if (rows.Count > 0) conn.Execute(SqlStaffLeaveUpsertByDocKey, rows, trans);
    }

    private static void UpsertStaffRows(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, IEnumerable<StaffRecord> staff)
    {
        var rows = staff.Where(s => s.StaffID > 0).Select(s => ToStaffRow(s)).ToList();
        if (rows.Count > 0) conn.Execute(SqlStaffUpsert, rows, trans);
    }

    /// <summary>Upserts by <c>vehicle_license_id</c>; rows without an ID get a server sequence ID.</summary>
    private static void PersistStaffVehicleLicenses(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, IEnumerable<VehicleLicense> licenses)
    {
        var items = licenses.Where(IsValidVehicleLicense).Select(l => ToVehicleLicenseRow(l)).ToList();
        var withId = items.Where(i => i.VehicleLicenseId > 0).ToList();
        var withoutId = items.Where(i => i.VehicleLicenseId <= 0).ToList();

        if (withId.Count > 0) conn.Execute(SqlVehicleLicenseUpsert, withId, trans);
        if (withoutId.Count > 0) conn.Execute(SqlVehicleLicenseInsertNewId, withoutId, trans);
    }

    /// <summary>Upserts by <c>skill_license_id</c>; rows without an ID get a server sequence ID.</summary>
    private static void PersistStaffSkillLicenses(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, IEnumerable<SkillLicense> licenses)
    {
        var items = licenses.Where(IsValidSkillLicense).Select(l => ToSkillLicenseRow(l)).ToList();
        var withId = items.Where(i => i.SkillLicenseId > 0).ToList();
        var withoutId = items.Where(i => i.SkillLicenseId <= 0).ToList();

        if (withId.Count > 0) conn.Execute(SqlSkillLicenseUpsert, withId, trans);
        if (withoutId.Count > 0) conn.Execute(SqlSkillLicenseInsertNewId, withoutId, trans);
    }

    // -------------------------------------------------------------------------
    // 3d. DELTA SAVE WITH OPTIMISTIC CONCURRENCY
    // -------------------------------------------------------------------------
    /// <summary>
    /// Writes only Added / Modified rows. Modified rows are updated only when the server <c>updated_at</c> still
    /// matches the loaded baseline; any mismatch (or a staff ID added concurrently by another user) is collected
    /// and raised as <see cref="ConcurrencyConflictException"/>, rolling back the entire transaction.
    /// <paramref name="forceOverwrite"/> skips the checks and upserts unconditionally.
    /// </summary>
    public static void PersistStaffDelta(StaffDeltaPayload delta, bool forceOverwrite = false)
    {
        var typeItems = NormalizeSkillLicenseTypes(delta.SkillLicenseTypes);

        bool ensured = false;
        Db.ExecuteInTransaction((conn, trans) =>
        {
            ensured = EnsureStaffHistoryConstraints(conn, trans);
            if (typeItems.Count > 0) UpsertSkillLicenseTypes(conn, trans, typeItems);

            var conflicts = new List<ConcurrencyConflict>();

            DateTime? ServerUpdatedAt(string sql, object key) =>
                conn.ExecuteScalar<DateTime?>(sql, key, trans);

            // a. staff_master
            UpsertStaffRows(conn, trans, delta.Staff.FallbackUpsert);

            foreach (var s in delta.Staff.Added.Where(s => s.StaffID > 0))
            {
                var row = ToStaffRow(s);
                if (forceOverwrite)
                {
                    conn.Execute(SqlStaffUpsert, row, trans);
                    continue;
                }

                int affected = conn.Execute(SqlStaffInsertIfAbsent, row, trans);
                if (affected == 0)
                {
                    var serverTs = ServerUpdatedAt("SELECT updated_at FROM staff_master WHERE staff_id = @StaffId;", new { StaffId = s.StaffID });
                    conflicts.Add(new ConcurrencyConflict("staff_master", $"Staff #{s.StaffID}",
                        $"Staff #{s.StaffID} ({s.StaffDetails.StaffName}) already exists on the server (added by another user).",
                        null, serverTs,
                        ExcelTableName: "AllStaffsTable",
                        KeyColumnHeader: "Staff ID",
                        RawKeyValue: s.StaffID.ToString(CultureInfo.InvariantCulture),
                        FieldDiffs: ConflictDiff.ForStaff(conn, trans, s)));
                }
            }

            foreach (var item in delta.Staff.Modified.Where(m => m.Record.StaffID > 0))
            {
                var s = item.Record;
                if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
                {
                    int affected = conn.Execute(SqlStaffConditionalUpdate, ToStaffRow(s, item.ExpectedUpdatedAt), trans);
                    if (affected == 0)
                    {
                        var serverTs = ServerUpdatedAt("SELECT updated_at FROM staff_master WHERE staff_id = @StaffId;", new { StaffId = s.StaffID });
                        conflicts.Add(new ConcurrencyConflict("staff_master", $"Staff #{s.StaffID}",
                            $"Staff #{s.StaffID} ({s.StaffDetails.StaffName}) was modified on the server.",
                            item.ExpectedUpdatedAt, serverTs,
                            ExcelTableName: "AllStaffsTable",
                            KeyColumnHeader: "Staff ID",
                            RawKeyValue: s.StaffID.ToString(CultureInfo.InvariantCulture),
                            FieldDiffs: ConflictDiff.ForStaff(conn, trans, s)));
                    }
                }
                else
                {
                    conn.Execute(SqlStaffUpsert, ToStaffRow(s), trans);
                }
            }

            // b. staff_vehicle_licenses (new rows always take a server sequence ID, so concurrent users never collide)
            PersistStaffVehicleLicenses(conn, trans, delta.VehicleLicenses.FallbackUpsert);

            var newVehicle = delta.VehicleLicenses.Added.Where(IsValidVehicleLicense).Select(l => ToVehicleLicenseRow(l)).ToList();
            if (newVehicle.Count > 0) conn.Execute(SqlVehicleLicenseInsertNewId, newVehicle, trans);

            foreach (var item in delta.VehicleLicenses.Modified.Where(m => IsValidVehicleLicense(m.Record)))
            {
                var l = item.Record;
                if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
                {
                    int affected = conn.Execute(SqlVehicleLicenseConditionalUpdate, ToVehicleLicenseRow(l, item.ExpectedUpdatedAt), trans);
                    if (affected == 0)
                    {
                        var serverTs = ServerUpdatedAt("SELECT updated_at FROM staff_vehicle_licenses WHERE vehicle_license_id = @Id;", new { Id = l.VehicleLicenseID });
                        conflicts.Add(new ConcurrencyConflict("staff_vehicle_licenses", $"Vehicle License #{l.VehicleLicenseID}",
                            $"Vehicle License #{l.VehicleLicenseID} (Staff #{l.StaffID}, {l.VehicleLicenseClass}) was modified on the server.",
                            item.ExpectedUpdatedAt, serverTs,
                            ExcelTableName: "VehicleLicensesTable",
                            KeyColumnHeader: "Vehicle License ID",
                            RawKeyValue: l.VehicleLicenseID.ToString(CultureInfo.InvariantCulture),
                            FieldDiffs: ConflictDiff.ForVehicleLicense(conn, trans, l)));
                    }
                }
                else
                {
                    conn.Execute(SqlVehicleLicenseUpsert, ToVehicleLicenseRow(l), trans);
                }
            }

            // c. staff_skill_licenses
            PersistStaffSkillLicenses(conn, trans, delta.SkillLicenses.FallbackUpsert);

            var newSkill = delta.SkillLicenses.Added.Where(IsValidSkillLicense).Select(l => ToSkillLicenseRow(l)).ToList();
            if (newSkill.Count > 0) conn.Execute(SqlSkillLicenseInsertNewId, newSkill, trans);

            foreach (var item in delta.SkillLicenses.Modified.Where(m => IsValidSkillLicense(m.Record)))
            {
                var l = item.Record;
                if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
                {
                    int affected = conn.Execute(SqlSkillLicenseConditionalUpdate, ToSkillLicenseRow(l, item.ExpectedUpdatedAt), trans);
                    if (affected == 0)
                    {
                        var serverTs = ServerUpdatedAt("SELECT updated_at FROM staff_skill_licenses WHERE skill_license_id = @Id;", new { Id = l.SkillLicenseID });
                        conflicts.Add(new ConcurrencyConflict("staff_skill_licenses", $"Skill License #{l.SkillLicenseID}",
                            $"Skill License #{l.SkillLicenseID} (Staff #{l.StaffID}, {l.SkillLicenseType}) was modified on the server.",
                            item.ExpectedUpdatedAt, serverTs,
                            ExcelTableName: "SkillLicensesTable",
                            KeyColumnHeader: "Skill License ID",
                            RawKeyValue: l.SkillLicenseID.ToString(CultureInfo.InvariantCulture),
                            FieldDiffs: ConflictDiff.ForSkillLicense(conn, trans, l)));
                    }
                }
                else
                {
                    conn.Execute(SqlSkillLicenseUpsert, ToSkillLicenseRow(l), trans);
                }
            }

            // d. staff_salaries: read-only history mirrored from SQL Payroll (the payroll system is authoritative, so
            //    no concurrency checks). New rows ignore any local Excel ID and merge on the natural key.
            PersistStaffSalaries(conn, trans, delta.Salaries.FallbackUpsert);
            PersistStaffSalaries(conn, trans, delta.Salaries.Added, ignoreLocalIds: true);
            PersistStaffSalaries(conn, trans, delta.Salaries.Modified.Select(m => m.Record));

            // e. staff_leaves: read-only LV_APP mirror (SQL Payroll authoritative; upsert by payroll_dockey).
            PersistStaffLeaves(conn, trans, delta.Leaves.FallbackUpsert);
            PersistStaffLeaves(conn, trans, delta.Leaves.Added);
            PersistStaffLeaves(conn, trans, delta.Leaves.Modified.Select(m => m.Record));

            // f. Any conflict rolls back the whole batch (ExecuteInTransaction catches, rolls back, rethrows).
            if (conflicts.Count > 0) throw new ConcurrencyConflictException(conflicts);
        });
        if (ensured) _staffHistoryConstraintsEnsured = true;

        AppLogger.Info($"[ServerDispatcher] Staff delta saved (force={forceOverwrite}): " +
            $"staff +{delta.Staff.Added.Count}/~{delta.Staff.Modified.Count}/fallback {delta.Staff.FallbackUpsert.Count}/={delta.Staff.UnchangedCount}; " +
            $"vehicle lic +{delta.VehicleLicenses.Added.Count}/~{delta.VehicleLicenses.Modified.Count}/fallback {delta.VehicleLicenses.FallbackUpsert.Count}/={delta.VehicleLicenses.UnchangedCount}; " +
            $"skill lic +{delta.SkillLicenses.Added.Count}/~{delta.SkillLicenses.Modified.Count}/fallback {delta.SkillLicenses.FallbackUpsert.Count}/={delta.SkillLicenses.UnchangedCount}; " +
            $"salaries +{delta.Salaries.Added.Count}/~{delta.Salaries.Modified.Count}/fallback {delta.Salaries.FallbackUpsert.Count}/={delta.Salaries.UnchangedCount}; " +
            $"leaves +{delta.Leaves.Added.Count}/~{delta.Leaves.Modified.Count}/fallback {delta.Leaves.FallbackUpsert.Count}/={delta.Leaves.UnchangedCount}.");
    }

    // -------------------------------------------------------------------------
    // 3e. FIELD-LEVEL CONFLICT DIFFS (server row vs. the local row this save would write)
    // -------------------------------------------------------------------------
    /// <summary>
    /// Re-reads a conflicting row inside the (soon rolled back) save transaction and lists every Excel column whose
    /// server value differs from the local one. The server row is mapped through the same record types as the sheet
    /// so computed columns (Contract Status, license Status, derived Short Name) compare like-for-like.
    /// </summary>
    private static class ConflictDiff
    {
        private const string Empty = "(empty)";

        public static List<FieldConflict> ForStaff(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, StaffRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM staff_master WHERE staff_id = @StaffId;", new { StaffId = local.StaffID });
            if (row == null) return diffs;

            string dept = Text(row, "department");
            var server = new StaffDetails
            {
                StaffName = Text(row, "staff_name"),
                StaffShortName = Text(row, "short_name"),
                Department = dept.Length > 0 ? dept : "Worker",
                NRIC = Long(row, "nric"),
                Passport = Text(row, "passport"),
                Phone = Text(row, "phone"),
                JoinDate = Date(row, "join_date") ?? DateOnly.FromDateTime(DateTime.Today),
                ResignDate = Date(row, "resign_date"),
                ContractStatus = Enum.TryParse<ContractStatus>(Text(row, "contract_status"), ignoreCase: true, out var cs) ? cs : ContractStatus.Active
            };
            var d = local.StaffDetails ?? new StaffDetails();

            AddText(diffs, "Staff Name", server.StaffName, d.StaffName);
            AddText(diffs, "Short Name", server.StaffShortName, d.StaffShortName);
            AddText(diffs, "Department", server.Department, d.Department);
            Add(diffs, "IC Number", ShowId(server.NRIC), ShowId(d.NRIC), server.NRIC == 0 ? null : (object)server.NRIC);
            AddText(diffs, "Passport Number", server.Passport, d.Passport);
            AddText(diffs, "Phone Number", server.Phone, d.Phone);
            Add(diffs, "Join Date", ShowDate(server.JoinDate), ShowDate(d.JoinDate), ExcelDate(server.JoinDate));
            Add(diffs, "Resign Date", ShowDate(server.ResignDate), ShowDate(d.ResignDate), ExcelDate(server.ResignDate));
            AddEnum(diffs, "Contract Status", server.ContractStatus, d.ContractStatus);
            return diffs;
        }

        public static List<FieldConflict> ForVehicleLicense(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, VehicleLicense local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM staff_vehicle_licenses WHERE vehicle_license_id = @VehicleLicenseId;",
                new { VehicleLicenseId = local.VehicleLicenseID });
            if (row == null) return diffs;

            string cls = FirstText(row, "vehicle_license_class", "license_type", "license_class");
            var renew = Date(row, "renew_date") ?? DateOnly.FromDateTime(DateTime.Today);
            var server = new VehicleLicense
            {
                VehicleLicenseID = local.VehicleLicenseID,
                StaffID = Long(row, "staff_id"),
                VehicleLicenseClass = Enum.TryParse<VehicleLicenseClass>(cls, ignoreCase: true, out var lc) ? lc : VehicleLicenseClass.None,
                RenewDate = renew,
                ExpiryDate = Date(row, "expiry_date") ?? renew.AddYears(1),
                VehicleLicenseStatus = Enum.TryParse<VehicleLicenseStatus>(Text(row, "status"), ignoreCase: true, out var ls) ? ls : VehicleLicenseStatus.Valid,
                Cost = Decimal(row, "cost")
            };

            Add(diffs, "Staff ID", ShowId(server.StaffID), ShowId(local.StaffID), server.StaffID == 0 ? null : (object)server.StaffID);
            AddEnum(diffs, "Vehicle License Class", server.VehicleLicenseClass, local.VehicleLicenseClass);
            AddLicenseCommon(diffs, server, local);
            AddEnum(diffs, "Status", server.VehicleLicenseStatus, local.VehicleLicenseStatus);
            Add(diffs, "Cost", ShowCost(server.Cost), ShowCost(local.Cost), server.Cost);
            return diffs;
        }

        public static List<FieldConflict> ForSkillLicense(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, SkillLicense local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM staff_skill_licenses WHERE skill_license_id = @SkillLicenseId;",
                new { SkillLicenseId = local.SkillLicenseID });
            if (row == null) return diffs;

            var renew = Date(row, "renew_date") ?? DateOnly.FromDateTime(DateTime.Today);
            var server = new SkillLicense
            {
                SkillLicenseID = local.SkillLicenseID,
                StaffID = Long(row, "staff_id"),
                SkillLicenseType = FirstText(row, "skill_license_type", "license_type"),
                RenewDate = renew,
                ExpiryDate = Date(row, "expiry_date") ?? renew.AddYears(1),
                SkillLicenseStatus = Enum.TryParse<SkillLicenseStatus>(Text(row, "status"), ignoreCase: true, out var ss) ? ss : SkillLicenseStatus.Valid,
                Cost = Decimal(row, "cost")
            };

            Add(diffs, "Staff ID", ShowId(server.StaffID), ShowId(local.StaffID), server.StaffID == 0 ? null : (object)server.StaffID);
            AddText(diffs, "Skill License Type", server.SkillLicenseType, local.SkillLicenseType);
            AddLicenseCommon(diffs, server, local);
            AddEnum(diffs, "Status", server.SkillLicenseStatus, local.SkillLicenseStatus);
            Add(diffs, "Cost", ShowCost(server.Cost), ShowCost(local.Cost), server.Cost);
            return diffs;
        }

        public static List<FieldConflict> ForCompany(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, CompanyRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM company_master WHERE company_id = @CompanyId;", new { CompanyId = local.CompanyID });
            if (row == null) return diffs;

            AddText(diffs, "Legal Company Name", Text(row, "legal_company_name"), local.LegalCompanyName);
            AddText(diffs, "Short Company Name", Text(row, "short_company_name"), local.ShortCompanyName);
            AddText(diffs, "iJob SQL Account Project", Text(row, "default_sql_account_project"), local.DefaultSqlAccountProject);
            return diffs;
        }

        /// <summary>User-editable project columns only; work / cost summary and formula columns are never diffed.</summary>
        public static List<FieldConflict> ForProject(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, ProjectRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM project_master WHERE project_id = @ProjectId;",
                new { ProjectId = NormalizeProjectCode(local.ProjectCode) });
            if (row == null) return diffs;

            var d = local.ProjectDetails ?? new ProjectDetails();
            var start = Date(row, "start_date") ?? DateOnly.FromDateTime(DateTime.Today);
            var server = new ProjectDetails
            {
                ReferenceNo = Text(row, "reference_no"),
                ProjectType = ParseProjectTypes(Text(row, "project_type")),
                Priority = Enum.TryParse<ProjectPriority>(Text(row, "priority"), ignoreCase: true, out var pr) ? pr : ProjectPriority.None,
                Status = Enum.TryParse<ProjectStatus>(Text(row, "status"), ignoreCase: true, out var st) ? st : ProjectStatus.None,
                CompanyID = Long(row, "company_id"),
                CompanyName = Text(row, "company_name"),
                SqlAccountProjectCode = Text(row, "sql_account_project_code"),
                StaffID = Long(row, "staff_id") is long sid && sid > 0 ? sid : null,
                PIC = Text(row, "pic"),
                District = Enum.TryParse<ProjectDistrict>(Text(row, "district"), ignoreCase: true, out var di) ? di : ProjectDistrict.None,
                Address = Text(row, "address"),
                StartDate = start,
                TargetEndDate = Date(row, "target_end_date") ?? start.AddDays(365),
                ActualEndDate = Date(row, "actual_end_date"),
                EstimatedCost = Decimal(row, "estimated_cost"),
                PromisedReturn = Decimal(row, "promised_return"),
                ActualReturn = Decimal(row, "actual_return"),
                Remarks = Text(row, "remarks")
            };

            string serverTypes = JoinProjectTypes(server.ProjectType);
            long serverStaff = server.StaffID ?? 0;

            AddText(diffs, "Reference No", server.ReferenceNo, d.ReferenceNo);
            Add(diffs, "Type", ShowText(serverTypes), ShowText(JoinProjectTypes(d.ProjectType)), serverTypes.Length == 0 ? null : serverTypes);
            AddEnum(diffs, "Priority", server.Priority, d.Priority);
            AddEnum(diffs, "Status", server.Status, d.Status);
            Add(diffs, "Company ID", ShowId(server.CompanyID), ShowId(d.CompanyID), server.CompanyID == 0 ? null : (object)server.CompanyID);
            AddText(diffs, "Company Name", server.CompanyName, d.CompanyName);
            AddText(diffs, "SQL Account Project", server.SqlAccountProjectCode, d.SqlAccountProjectCode);
            Add(diffs, "Staff ID", ShowId(serverStaff), ShowId(d.StaffID ?? 0), serverStaff == 0 ? null : (object)serverStaff);
            AddText(diffs, "PIC", server.PIC, d.PIC);
            AddEnum(diffs, "District", server.District, d.District);
            AddText(diffs, "Address", server.Address, d.Address);
            Add(diffs, "Start Date", ShowDate(server.StartDate), ShowDate(d.StartDate), ExcelDate(server.StartDate));
            Add(diffs, "Target End Date", ShowDate(server.TargetEndDate), ShowDate(d.TargetEndDate), ExcelDate(server.TargetEndDate));
            Add(diffs, "Actual End Date", ShowDate(server.ActualEndDate), ShowDate(d.ActualEndDate), ExcelDate(server.ActualEndDate));
            Add(diffs, "Estimated Cost", ShowCost(server.EstimatedCost), ShowCost(d.EstimatedCost), server.EstimatedCost);
            Add(diffs, "Promised Return", ShowCost(server.PromisedReturn), ShowCost(d.PromisedReturn), server.PromisedReturn);
            Add(diffs, "Actual Return", ShowCost(server.ActualReturn), ShowCost(d.ActualReturn), server.ActualReturn);
            AddText(diffs, "Remarks", server.Remarks, d.Remarks);
            return diffs;
        }

        // --- Vehicle Manager (headers match Types.VehicleRecord.cs; rollup / calculated columns are never diffed) ---
        public static List<FieldConflict> ForVehicle(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, VehicleRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_master WHERE vehicle_plate = @Key;", new { Key = NormalizePlate(local.VehiclePlate) });
            if (row == null) return diffs;

            var d = local.VehicleDetail ?? new VehicleDetail();
            AddText(diffs, "VIN", Text(row, "vin"), d.VIN);
            AddParsedEnum(diffs, "Vehicle Type", Text(row, "vehicle_type"), d.Type);
            AddText(diffs, "Brand", Text(row, "brand"), d.Brand);
            AddText(diffs, "Model", Text(row, "model"), d.Model);
            AddText(diffs, "Driven By", Text(row, "driven_by"), d.DrivenBy);
            AddText(diffs, "Used At", Text(row, "used_at"), d.UsedAt);
            AddNum(diffs, "Current Odometer", DecN(row, "current_odometer"), d.CurrentOdometer, 1, money: false);
            AddDate(diffs, "Bought Date", Date(row, "bought_date"), d.BoughtDate);
            AddNum(diffs, "Bought Price", DecN(row, "bought_price"), d.BoughtPrice, 2, money: true);
            AddDate(diffs, "Sold Date", Date(row, "sold_date"), d.SoldDate);
            AddNum(diffs, "Sold Price", DecN(row, "sold_price"), d.SoldPrice, 2, money: true);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.VehicleStatus);
            return diffs;
        }

        public static List<FieldConflict> ForMaintenance(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, Maintenance local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_maintenance WHERE maintenance_id = @Key;", new { Key = local.MaintenanceID });
            if (row == null) return diffs;

            AddText(diffs, "Vehicle Plate", Text(row, "vehicle_plate"), NormalizePlate(local.VehiclePlate));
            AddParsedEnum(diffs, "Service Type", Text(row, "service_type"), local.ServiceType);
            AddDate(diffs, "Date In", Date(row, "date_in"), local.DateIn);
            AddDate(diffs, "Estimated Return", Date(row, "estimated_return"), local.EstimatedReturn);
            AddDate(diffs, "Date Completed", Date(row, "date_completed"), local.DateCompleted);
            AddNum(diffs, "Odometer In", DecN(row, "odometer_in"), local.OdometerIn, 1, money: false);
            AddNum(diffs, "Odometer Out", DecN(row, "odometer_out"), local.OdometerOut, 1, money: false);
            AddText(diffs, "Workshop Name", Text(row, "workshop_name"), local.WorkshopName);
            AddText(diffs, "Workshop Address", Text(row, "workshop_address"), local.WorkshopAddress);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.MaintenanceStatus);
            AddNum(diffs, "Tyre", DecN(row, "tyre_cost"), local.TyreCost, 2, money: true);
            AddNum(diffs, "Brake", DecN(row, "brake_cost"), local.BrakeCost, 2, money: true);
            AddNum(diffs, "Clutch", DecN(row, "clutch_cost"), local.ClutchCost, 2, money: true);
            AddNum(diffs, "Service", DecN(row, "service_cost"), local.ServiceCost, 2, money: true);
            AddNum(diffs, "Fuel Pump", DecN(row, "fuel_pump_cost"), local.FuelPumpCost, 2, money: true);
            AddNum(diffs, "Aircond", DecN(row, "aircond_cost"), local.AircondCost, 2, money: true);
            AddNum(diffs, "Wiring", DecN(row, "wiring_cost"), local.WiringCost, 2, money: true);
            AddNum(diffs, "Battery", DecN(row, "battery_cost"), local.BatteryCost, 2, money: true);
            AddNum(diffs, "Towing", DecN(row, "towing_cost"), local.TowingCost, 2, money: true);
            AddNum(diffs, "Cushion", DecN(row, "cushion_cost"), local.CushionCost, 2, money: true);
            AddNum(diffs, "Parts", DecN(row, "parts_cost"), local.PartsCost, 2, money: true);
            AddText(diffs, "Remarks", Text(row, "remarks"), local.Remarks);
            return diffs;
        }

        public static List<FieldConflict> ForRoadTax(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, RoadTax local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_road_taxes WHERE road_tax_id = @Key;", new { Key = local.RoadTaxID });
            if (row == null) return diffs;

            AddRenewalCommon(diffs, row, local);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.Status);
            return diffs;
        }

        public static List<FieldConflict> ForPermit(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, VehiclePermit local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_permits WHERE permit_id = @Key;", new { Key = local.PermitID });
            if (row == null) return diffs;

            AddParsedEnum(diffs, "Permit Type", Text(row, "permit_type"), local.PermitType);
            AddRenewalCommon(diffs, row, local);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.Status);
            return diffs;
        }

        public static List<FieldConflict> ForInsurance(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, VehicleInsurance local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_insurances WHERE insurance_id = @Key;", new { Key = local.InsuranceID });
            if (row == null) return diffs;

            AddText(diffs, "Insurance Number", Text(row, "insurance_number"), local.InsuranceNumber);
            AddText(diffs, "Insurance Company", Text(row, "insurance_company"), local.InsuranceCompany);
            AddParsedEnum(diffs, "Insurance Type", Text(row, "insurance_type"), local.InsuranceType);
            AddRenewalCommon(diffs, row, local);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.Status);
            return diffs;
        }

        public static List<FieldConflict> ForInspection(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, Inspection local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_inspections WHERE inspection_id = @Key;", new { Key = local.InspectionID });
            if (row == null) return diffs;

            AddText(diffs, "Vehicle Plate", Text(row, "vehicle_plate"), NormalizePlate(local.VehiclePlate));
            AddDate(diffs, "Schedule Date In", Date(row, "schedule_date_in"), local.ScheduleDateIn);
            AddDate(diffs, "Date Out", Date(row, "date_out"), local.DateOut);
            AddDate(diffs, "Expiry Date", Date(row, "expiry_date"), local.ExpiryDate);
            AddNum(diffs, "Cost", DecN(row, "amount"), local.Amount, 2, money: true);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.Status);
            AddText(diffs, "Address", Text(row, "address"), local.Address);
            AddText(diffs, "Remarks", Text(row, "remarks"), local.Remarks);
            return diffs;
        }

        public static List<FieldConflict> ForInsuranceClaim(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, InsuranceClaim local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_insurance_claims WHERE claim_id = @Key;", new { Key = local.ClaimID });
            if (row == null) return diffs;

            long staff = Long(row, "staff_id");
            long maint = Long(row, "maintenance_id");
            AddText(diffs, "Vehicle Plate", Text(row, "vehicle_plate"), NormalizePlate(local.VehiclePlate));
            AddDate(diffs, "Claim Date", Date(row, "claim_date"), local.ClaimDate);
            AddDate(diffs, "Settlement Date", Date(row, "settlement_date"), local.SettlementDate);
            Add(diffs, "Staff ID", ShowId(staff), ShowId(local.StaffID), staff == 0 ? null : (object)staff);
            Add(diffs, "Maintenance ID", ShowId(maint), ShowId(local.MaintenanceID), maint == 0 ? null : (object)maint);
            AddText(diffs, "Workshop Name", Text(row, "workshop_name"), local.WorkshopName);
            AddParsedEnum(diffs, "Status", Text(row, "status"), local.Status);
            AddNum(diffs, "Claim Amount", DecN(row, "claim_amount"), local.ClaimAmount, 2, money: true);
            AddText(diffs, "Remarks", Text(row, "remarks"), local.Remarks);
            return diffs;
        }

        public static List<FieldConflict> ForOdometer(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, OdometerLog local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_odometers WHERE odometer_id = @Key;", new { Key = local.OdometerID });
            if (row == null) return diffs;

            long staff = Long(row, "staff_id");
            AddText(diffs, "Vehicle Plate", Text(row, "vehicle_plate"), NormalizePlate(local.VehiclePlate));
            AddDate(diffs, "Date", Date(row, "date"), local.Date);
            Add(diffs, "Staff ID", ShowId(staff), ShowId(local.StaffID), staff == 0 ? null : (object)staff);
            AddText(diffs, "Project", Text(row, "project"), local.Project);
            AddNum(diffs, "Current Odometer", DecN(row, "current_odometer"), local.CurrentOdometer, 1, money: false);
            AddText(diffs, "Fuel Card Number", Text(row, "fuel_card_number"), local.FuelCardNumber);
            AddText(diffs, "Fuel Company", Text(row, "fuel_company"), local.FuelCompany);
            AddText(diffs, "Station Name", Text(row, "station_name"), local.StationName);
            AddText(diffs, "Receipt Number", Text(row, "receipt_number"), local.ReceiptNumber);
            AddText(diffs, "Fuel Type", Text(row, "fuel_type"), local.FuelType);
            AddNum(diffs, "Price at Fueling", DecN(row, "price_at_fueling"), local.PriceAtFueling, 4, money: true);
            AddNum(diffs, "Fuel Amount", DecN(row, "fuel_amount"), local.FuelAmount, 4, money: false);
            AddNum(diffs, "Transaction Amount", DecN(row, "amount"), local.Amount, 2, money: true);
            AddText(diffs, "Time", Text(row, "time"), local.Time);
            AddText(diffs, "Remarks", Text(row, "remarks"), local.Remarks);
            return diffs;
        }

        public static List<FieldConflict> ForWorkshop(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, Workshop local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_workshops WHERE workshop_id = @Key;", new { Key = local.WorkshopID });
            if (row == null) return diffs;

            AddText(diffs, "Workshop Name", Text(row, "workshop_name"), local.WorkshopName);
            AddText(diffs, "Workshop Address", Text(row, "workshop_address"), local.WorkshopAddress);
            AddText(diffs, "Workshop Owner", Text(row, "workshop_owner"), local.WorkshopOwner);
            AddText(diffs, "Workshop Primary Contact", Text(row, "primary_contact"), local.PrimaryContact);
            AddText(diffs, "Workshop Secondary Contact", Text(row, "secondary_contact"), local.SecondaryContact);
            return diffs;
        }

        public static List<FieldConflict> ForInsuranceCompany(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, InsuranceCompanyRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_insurance_companies WHERE insurance_company = @Key;",
                new { Key = (local.InsuranceCompany ?? string.Empty).Trim() });
            if (row == null) return diffs;

            AddText(diffs, "Company Contact Number", Text(row, "company_contact_number"), local.CompanyContactNumber);
            AddText(diffs, "Company Address", Text(row, "company_address"), local.CompanyAddress);
            AddText(diffs, "Insurance Agent", Text(row, "insurance_agent"), local.InsuranceAgent);
            AddText(diffs, "Agent Contact Number", Text(row, "agent_contact_number"), local.AgentContactNumber);
            return diffs;
        }

        public static List<FieldConflict> ForFuelCard(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, FuelCardRecord local)
        {
            var diffs = new List<FieldConflict>();
            var row = QueryRow(conn, trans, "SELECT * FROM vehicle_fuel_cards WHERE fuel_card_number = @Key;",
                new { Key = (local.FuelCardNumber ?? string.Empty).Trim() });
            if (row == null) return diffs;

            AddText(diffs, "Fuel Company Names", Text(row, "fuel_company_name"), local.FuelCompanyName);
            AddText(diffs, "Card Owner ID", Text(row, "card_owner_id"), local.CardOwnerID);
            return diffs;
        }

        /// <summary>Vehicle Plate / Renewal Date / Expiry Date / Cost / Remarks shared by insurance, permit and road tax rows.</summary>
        private static void AddRenewalCommon(List<FieldConflict> diffs, IDictionary<string, object> row, VehicleRenewalRecord local)
        {
            AddText(diffs, "Vehicle Plate", Text(row, "vehicle_plate"), NormalizePlate(local.VehiclePlate));
            AddDate(diffs, "Renewal Date", Date(row, "renew_date"), local.RenewDate);
            AddDate(diffs, "Expiry Date", Date(row, "expiry_date"), local.ExpiryDate);
            AddNum(diffs, "Cost", DecN(row, "amount"), local.Amount, 2, money: true);
            AddText(diffs, "Remarks", Text(row, "remarks"), local.Remarks);
        }

        private static void AddDate(List<FieldConflict> diffs, string header, DateOnly? server, DateOnly? local) =>
            Add(diffs, header, ShowDate(server), ShowDate(local), ExcelDate(server));

        private static void AddNum(List<FieldConflict> diffs, string header, decimal? server, decimal? local, int decimals, bool money)
        {
            decimal? s = server.HasValue ? decimal.Round(server.Value, decimals) : null;
            Add(diffs, header, ShowNum(s, decimals, money), ShowNum(local, decimals, money), s);
        }

        /// <summary>Unparseable / NULL server text falls back to the enum's <c>None</c> member (shown as empty).</summary>
        private static void AddParsedEnum<TEnum>(List<FieldConflict> diffs, string header, string serverRaw, TEnum local) where TEnum : struct, Enum
        {
            TEnum server = Enum.TryParse<TEnum>(serverRaw, ignoreCase: true, out var parsed) ? parsed
                : Enum.TryParse<TEnum>("None", ignoreCase: true, out var none) ? none : default;
            AddEnum(diffs, header, server, local);
        }

        private static string ShowNum(decimal? v, int decimals, bool money) =>
            v.HasValue
                ? (money ? "RM " : string.Empty) + decimal.Round(v.Value, decimals).ToString("F" + decimals.ToString(CultureInfo.InvariantCulture), CultureInfo.InvariantCulture)
                : Empty;

        private static decimal? DecN(IDictionary<string, object> row, string col) => Raw(row, col) == null ? null : Decimal(row, col);

        private static List<ProjectType> ParseProjectTypes(string raw) => raw
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => Enum.TryParse<ProjectType>(t, ignoreCase: true, out var pt) ? (ProjectType?)pt : null)
            .Where(t => t.HasValue && t.Value != ProjectType.None)
            .Select(t => t!.Value)
            .Distinct()
            .ToList();

        private static string JoinProjectTypes(IEnumerable<ProjectType>? types) =>
            string.Join(", ", (types ?? Enumerable.Empty<ProjectType>()).Where(t => t != ProjectType.None).Distinct().Select(t => t.ToString()));

        private static void AddLicenseCommon(List<FieldConflict> diffs, StaffLicenseBase server, StaffLicenseBase local)
        {
            Add(diffs, "Renew Date", ShowDate(server.RenewDate), ShowDate(local.RenewDate), ExcelDate(server.RenewDate));
            Add(diffs, "Expiry Date", ShowDate(server.ExpiryDate), ShowDate(local.ExpiryDate), ExcelDate(server.ExpiryDate));
        }

        // --- Diff collection ---
        private static void Add(List<FieldConflict> diffs, string header, string serverDisplay, string localDisplay, object? serverExcelValue)
        {
            if (!string.Equals(serverDisplay, localDisplay, StringComparison.Ordinal))
                diffs.Add(new FieldConflict(header, serverDisplay, localDisplay, serverExcelValue));
        }

        private static void AddText(List<FieldConflict> diffs, string header, string? server, string? local)
        {
            string s = (server ?? string.Empty).Trim();
            Add(diffs, header, ShowText(s), ShowText(local), s.Length == 0 ? null : s);
        }

        private static void AddEnum<TEnum>(List<FieldConflict> diffs, string header, TEnum server, TEnum local) where TEnum : struct, Enum
        {
            string s = server.ToString();
            bool isNone = string.Equals(s, "None", StringComparison.OrdinalIgnoreCase);
            Add(diffs, header, isNone ? Empty : s, ShowEnum(local), isNone ? null : s);
        }

        // --- Display formatting ---
        private static string ShowText(string? s) => string.IsNullOrWhiteSpace(s) ? Empty : s.Trim();
        private static string ShowId(long v) => v == 0 ? Empty : v.ToString(CultureInfo.InvariantCulture);
        private static string ShowDate(DateOnly? d) => d.HasValue ? d.Value.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture) : Empty;
        private static string ShowCost(decimal m) => "RM " + decimal.Round(m, 2).ToString("0.00", CultureInfo.InvariantCulture);
        private static string ShowEnum<TEnum>(TEnum v) where TEnum : struct, Enum
        {
            string s = v.ToString();
            return string.Equals(s, "None", StringComparison.OrdinalIgnoreCase) ? Empty : s;
        }
        private static object? ExcelDate(DateOnly? d) => d.HasValue ? d.Value.ToDateTime(TimeOnly.MinValue) : null;

        // --- Server row access ---
        private static IDictionary<string, object>? QueryRow(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, string sql, object param)
        {
            try
            {
                return (object?)conn.QueryFirstOrDefault(sql, param, trans) as IDictionary<string, object>;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[ServerDispatcher] Conflict diff query failed: {ex.Message}");
                return null;
            }
        }

        private static object? Raw(IDictionary<string, object> row, string col) =>
            row.TryGetValue(col, out var v) && v is not null && v is not DBNull ? v : null;

        private static string Text(IDictionary<string, object> row, string col) =>
            (Convert.ToString(Raw(row, col), CultureInfo.InvariantCulture) ?? string.Empty).Trim();

        private static string FirstText(IDictionary<string, object> row, params string[] cols)
        {
            foreach (var c in cols)
            {
                string v = Text(row, c);
                if (v.Length > 0) return v;
            }
            return string.Empty;
        }

        private static long Long(IDictionary<string, object> row, string col) =>
            long.TryParse(Text(row, col), NumberStyles.Integer, CultureInfo.InvariantCulture, out long v) ? v : 0;

        private static decimal Decimal(IDictionary<string, object> row, string col)
        {
            var v = Raw(row, col);
            if (v == null) return 0m;
            try { return Convert.ToDecimal(v, CultureInfo.InvariantCulture); } catch { return 0m; }
        }

        private static DateOnly? Date(IDictionary<string, object> row, string col) => Raw(row, col) switch
        {
            DateTime dt => DateOnly.FromDateTime(dt),
            DateOnly d => d,
            DateTimeOffset dto => DateOnly.FromDateTime(dto.DateTime),
            _ => null
        };
    }

    private static List<SkillTypeRow> NormalizeSkillLicenseTypes(IEnumerable<SkillLicenseType>? types)
    {
        if (types == null) return new();
        return types
            .Where(t => !string.IsNullOrWhiteSpace(t.SkillLicenseName))
            .GroupBy(t => t.SkillLicenseName.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => new SkillTypeRow
            {
                SkillLicenseName = g.Key,
                DisplayIcon = g.Select(x => (x.DisplayIcon ?? string.Empty).Trim()).FirstOrDefault(i => i.Length > 0) ?? string.Empty
            })
            .ToList();
    }

    private sealed class SkillTypeRow
    {
        public string SkillLicenseName { get; init; } = string.Empty;
        public string DisplayIcon { get; init; } = string.Empty;
    }

    private static void UpsertSkillLicenseTypes(System.Data.IDbConnection conn, System.Data.IDbTransaction trans, List<SkillTypeRow> items)
    {
        conn.Execute("ALTER TABLE lookup_skill_license_types ADD COLUMN IF NOT EXISTS display_icon VARCHAR(32) NOT NULL DEFAULT '';", transaction: trans);
        conn.Execute(@"
            INSERT INTO lookup_skill_license_types (skill_license_type, description, display_icon)
            VALUES (@SkillLicenseName, '', @DisplayIcon)
            ON CONFLICT (skill_license_type) DO UPDATE SET
                display_icon = EXCLUDED.display_icon;", items, trans);
    }

    /// <summary>Upsert-only (skill license types are never pruned, so historical licenses keep their type).</summary>
    private static void PersistSkillLicenseTypes(IEnumerable<SkillLicenseType> types)
    {
        var items = NormalizeSkillLicenseTypes(types);

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute("ALTER TABLE lookup_skill_license_types ADD COLUMN IF NOT EXISTS display_icon VARCHAR(32) NOT NULL DEFAULT '';", transaction: trans);
            if (items.Count == 0) return;

            UpsertSkillLicenseTypes(conn, trans, items);
        });
    }

    // =========================================================================
    // 4. PROJECT & COMPANY PERSISTENCE
    // =========================================================================
    // -------------------------------------------------------------------------
    // 4a. ROW PARAMETER MODELS (DateOnly projected to DateTime for Dapper)
    // -------------------------------------------------------------------------
    private static string NormalizeProjectCode(string? code) => (code ?? string.Empty).Trim().ToUpperInvariant();

    private sealed class ProjectRow
    {
        public string ProjectId { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string ReferenceNo { get; init; } = string.Empty;
        public string ProjectType { get; init; } = string.Empty;
        public string Priority { get; init; } = string.Empty;
        public string Status { get; init; } = string.Empty;
        public long? CompanyId { get; init; }
        public string CompanyName { get; init; } = string.Empty;
        public string SqlAccountProjectCode { get; init; } = string.Empty;
        public long? StaffId { get; init; }
        public string Pic { get; init; } = string.Empty;
        public string District { get; init; } = string.Empty;
        public string Address { get; init; } = string.Empty;
        public DateTime StartDate { get; init; }
        public DateTime TargetEndDate { get; init; }
        public DateTime? ActualEndDate { get; init; }
        public decimal EstimatedCost { get; init; }
        public decimal PromisedReturn { get; init; }
        public decimal ActualReturn { get; init; }
        public string Remarks { get; init; } = string.Empty;
        public int? DaysWorked { get; init; }
        public decimal? TotalRegularShifts { get; init; }
        public decimal? TotalHours { get; init; }
        public int? OvertimeShifts { get; init; }
        public decimal? TotalOvertimeHours { get; init; }
        public decimal? TotalOdometer { get; init; }
        public decimal? FuelCost { get; init; }
        public decimal? LabourCost { get; init; }
        public decimal? ClaimCost { get; init; }
        public decimal? ActualProjectCost { get; init; }
        public decimal? ProfitMargin { get; init; }
        public DateTime? ExpectedUpdatedAt { get; init; }
    }

    private sealed class CompanyRow
    {
        public long CompanyId { get; init; }
        public string LegalCompanyName { get; init; } = string.Empty;
        public string ShortCompanyName { get; init; } = string.Empty;
        public string DefaultSqlAccountProject { get; init; } = string.Empty;
        public DateTime? ExpectedUpdatedAt { get; init; }
    }

    private static ProjectRow ToProjectRow(ProjectRecord p, DateTime? expectedUpdatedAt = null)
    {
        var d = p.ProjectDetails ?? new ProjectDetails();
        return new ProjectRow
        {
            ProjectId = NormalizeProjectCode(p.ProjectCode),
            Category = p.Category.ToString(),
            ReferenceNo = d.ReferenceNo ?? string.Empty,
            ProjectType = string.Join(", ", (d.ProjectType ?? new List<ProjectType>()).Select(t => t.ToString())),
            Priority = d.Priority.ToString(),
            Status = d.Status.ToString(),
            CompanyId = d.CompanyID > 0 ? d.CompanyID : null,
            CompanyName = d.CompanyName ?? string.Empty,
            SqlAccountProjectCode = d.SqlAccountProjectCode?.Trim() ?? string.Empty,
            StaffId = d.StaffID,
            Pic = d.PIC ?? string.Empty,
            District = d.District.ToString(),
            Address = d.Address ?? string.Empty,
            StartDate = d.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = d.TargetEndDate.ToDateTime(TimeOnly.MinValue),
            ActualEndDate = d.ActualEndDate.HasValue ? d.ActualEndDate.Value.ToDateTime(TimeOnly.MinValue) : null,
            EstimatedCost = d.EstimatedCost,
            PromisedReturn = d.PromisedReturn,
            ActualReturn = d.ActualReturn,
            Remarks = d.Remarks ?? string.Empty,
            DaysWorked = d.DaysWorked,
            TotalRegularShifts = d.TotalRegularShifts,
            TotalHours = d.TotalHours,
            OvertimeShifts = d.OvertimeShifts,
            TotalOvertimeHours = d.TotalOvertimeHours,
            TotalOdometer = d.TotalOdometer,
            FuelCost = d.FuelCost,
            LabourCost = d.LabourCost,
            ClaimCost = d.ClaimCost,
            ActualProjectCost = d.ActualProjectCost,
            ProfitMargin = d.ProfitMargin,
            ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
        };
    }

    private static CompanyRow ToCompanyRow(CompanyRecord c, DateTime? expectedUpdatedAt = null) => new()
    {
        CompanyId = c.CompanyID,
        LegalCompanyName = c.LegalCompanyName ?? string.Empty,
        ShortCompanyName = c.ShortCompanyName ?? string.Empty,
        DefaultSqlAccountProject = c.DefaultSqlAccountProject?.Trim() ?? string.Empty,
        ExpectedUpdatedAt = ToConcurrencyToken(expectedUpdatedAt)
    };

    private static bool IsValidProject(ProjectRecord p) => !string.IsNullOrWhiteSpace(p.ProjectCode);
    private static bool IsValidCompany(CompanyRecord c) => c.CompanyID > 0;

    private static string ProjectExcelTable(ProjectCategory category) => category switch
    {
        ProjectCategory.iJob => "iJobProjectsTable",
        ProjectCategory.MWO => "MWOProjectsTable",
        ProjectCategory.PVT => "PVTProjectsTable",
        ProjectCategory.Other => "OtherProjectsTable",
        _ => "AllProjectsTable"
    };

    // -------------------------------------------------------------------------
    // 4b. SQL (updated_at always strictly increases so concurrency tokens never repeat)
    // -------------------------------------------------------------------------
    private const string SqlProjectInsert = @"
INSERT INTO project_master (
    project_id, category, reference_no, project_type, priority, status, company_id, company_name, staff_id, pic,
    district, address, start_date, target_end_date, actual_end_date, estimated_cost, promised_return, actual_return,
    remarks, days_worked, total_regular_shifts, total_hours, overtime_shifts, total_overtime_hours, total_odometer, fuel_cost, labour_cost, claim_cost,
    actual_project_cost, profit_margin, sql_account_project_code, updated_at
) VALUES (
    @ProjectId, @Category, @ReferenceNo, @ProjectType, @Priority, @Status, @CompanyId, @CompanyName, @StaffId, @Pic,
    @District, @Address, @StartDate, @TargetEndDate, @ActualEndDate, @EstimatedCost, @PromisedReturn, @ActualReturn,
    @Remarks, @DaysWorked, @TotalRegularShifts, @TotalHours, @OvertimeShifts, @TotalOvertimeHours, @TotalOdometer, @FuelCost, @LabourCost, @ClaimCost,
    @ActualProjectCost, @ProfitMargin, @SqlAccountProjectCode, clock_timestamp()
)";

    private const string SqlProjectUpsert = SqlProjectInsert + @"
ON CONFLICT (project_id) DO UPDATE SET
    category            = EXCLUDED.category,
    reference_no        = EXCLUDED.reference_no,
    project_type        = EXCLUDED.project_type,
    priority            = EXCLUDED.priority,
    status              = EXCLUDED.status,
    company_id          = EXCLUDED.company_id,
    company_name        = EXCLUDED.company_name,
    staff_id            = EXCLUDED.staff_id,
    pic                 = EXCLUDED.pic,
    district            = EXCLUDED.district,
    address             = EXCLUDED.address,
    start_date          = EXCLUDED.start_date,
    target_end_date     = EXCLUDED.target_end_date,
    actual_end_date     = EXCLUDED.actual_end_date,
    estimated_cost      = EXCLUDED.estimated_cost,
    promised_return     = EXCLUDED.promised_return,
    actual_return       = EXCLUDED.actual_return,
    remarks             = EXCLUDED.remarks,
    days_worked         = EXCLUDED.days_worked,
    total_regular_shifts = EXCLUDED.total_regular_shifts,
    total_hours         = EXCLUDED.total_hours,
    overtime_shifts     = EXCLUDED.overtime_shifts,
    total_overtime_hours = EXCLUDED.total_overtime_hours,
    total_odometer      = EXCLUDED.total_odometer,
    fuel_cost           = EXCLUDED.fuel_cost,
    labour_cost         = EXCLUDED.labour_cost,
    claim_cost          = EXCLUDED.claim_cost,
    actual_project_cost = EXCLUDED.actual_project_cost,
    profit_margin       = EXCLUDED.profit_margin,
    sql_account_project_code = EXCLUDED.sql_account_project_code,
    updated_at         = GREATEST(clock_timestamp(), project_master.updated_at + INTERVAL '1 millisecond');";

    private const string SqlProjectInsertIfAbsent = SqlProjectInsert + @"
ON CONFLICT (project_id) DO NOTHING;";

    private const string SqlProjectConditionalUpdate = @"
UPDATE project_master SET
    category            = @Category,
    reference_no        = @ReferenceNo,
    project_type        = @ProjectType,
    priority            = @Priority,
    status              = @Status,
    company_id          = @CompanyId,
    company_name        = @CompanyName,
    staff_id            = @StaffId,
    pic                 = @Pic,
    district            = @District,
    address             = @Address,
    start_date          = @StartDate,
    target_end_date     = @TargetEndDate,
    actual_end_date     = @ActualEndDate,
    estimated_cost      = @EstimatedCost,
    promised_return     = @PromisedReturn,
    actual_return       = @ActualReturn,
    remarks             = @Remarks,
    days_worked         = @DaysWorked,
    total_regular_shifts = @TotalRegularShifts,
    total_hours         = @TotalHours,
    overtime_shifts     = @OvertimeShifts,
    total_overtime_hours = @TotalOvertimeHours,
    total_odometer      = @TotalOdometer,
    fuel_cost           = @FuelCost,
    labour_cost         = @LabourCost,
    claim_cost          = @ClaimCost,
    actual_project_cost = @ActualProjectCost,
    profit_margin       = @ProfitMargin,
    sql_account_project_code = @SqlAccountProjectCode,
    updated_at         = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
WHERE project_id = @ProjectId
  AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    private const string SqlCompanyInsert = @"
        INSERT INTO company_master (company_id, legal_company_name, short_company_name, default_sql_account_project, updated_at)
        VALUES (@CompanyId, @LegalCompanyName, @ShortCompanyName, @DefaultSqlAccountProject, clock_timestamp())";

    private const string SqlCompanyUpsert = SqlCompanyInsert + @"
        ON CONFLICT (company_id) DO UPDATE SET
            legal_company_name = EXCLUDED.legal_company_name,
            short_company_name = EXCLUDED.short_company_name,
            default_sql_account_project = EXCLUDED.default_sql_account_project,
            updated_at         = GREATEST(clock_timestamp(), company_master.updated_at + INTERVAL '1 millisecond');";

    private const string SqlCompanyInsertIfAbsent = SqlCompanyInsert + @"
        ON CONFLICT (company_id) DO NOTHING;";

    private const string SqlCompanyConditionalUpdate = @"
        UPDATE company_master SET
            legal_company_name = @LegalCompanyName,
            short_company_name = @ShortCompanyName,
            default_sql_account_project = @DefaultSqlAccountProject,
            updated_at        = GREATEST(clock_timestamp(), updated_at + INTERVAL '1 millisecond')
        WHERE company_id = @CompanyId
          AND ABS(EXTRACT(EPOCH FROM (updated_at - @ExpectedUpdatedAt))) < 0.001;";

    // -------------------------------------------------------------------------
    // 4c. UPSERT-ONLY PERSISTENCE (offline pending upload / no-baseline fallback; never hard-deletes)
    // -------------------------------------------------------------------------
    /// <summary>Upsert-only: rows missing from the workbook are kept on the server (history preserved).</summary>
    public static void PersistProjects(IList<ProjectRecord> projects, bool isFullBatch)
    {
        var rows = projects.Where(IsValidProject).Select(p => ToProjectRow(p)).ToList();
        if (rows.Count == 0 && !isFullBatch) return;

        Db.ExecuteInTransaction((conn, trans) =>
        {
            EnsureProjectSchema(conn, trans);
            if (rows.Count > 0) conn.Execute(SqlProjectUpsert, rows, trans);
        });
    }

    /// <summary>Upsert-only: rows missing from the workbook are kept on the server (projects may still reference them).</summary>
    public static void PersistCompanies(IList<CompanyRecord> companies, bool isFullBatch)
    {
        var rows = companies.Where(IsValidCompany).Select(c => ToCompanyRow(c)).ToList();
        if (rows.Count == 0 && !isFullBatch) return;

        Db.ExecuteInTransaction((conn, trans) =>
        {
            EnsureCompanySchema(conn, trans);
            if (rows.Count > 0) conn.Execute(SqlCompanyUpsert, rows, trans);
        });
    }

    // -------------------------------------------------------------------------
    // 4d. DELTA SAVE WITH OPTIMISTIC CONCURRENCY
    // -------------------------------------------------------------------------
    /// <summary>
    /// Writes only Added / Modified companies and projects in one transaction (companies first so project rows can
    /// resolve them). Modified rows are updated only when the server <c>updated_at</c> still matches the loaded
    /// baseline; any mismatch (or a key added concurrently by another user) is collected and raised as
    /// <see cref="ConcurrencyConflictException"/>, rolling back the entire batch. <paramref name="forceOverwrite"/>
    /// skips the checks and upserts unconditionally.
    /// </summary>
    public static void PersistProjectDelta(ProjectDeltaPayload delta, bool forceOverwrite = false)
    {
        Db.ExecuteInTransaction((conn, trans) =>
        {
            EnsureCompanySchema(conn, trans);
            EnsureProjectSchema(conn, trans);

            var conflicts = new List<ConcurrencyConflict>();

            DateTime? ServerUpdatedAt(string sql, object key) =>
                conn.ExecuteScalar<DateTime?>(sql, key, trans);

            // a. company_master
            var fallbackCompanies = delta.Companies.FallbackUpsert.Where(IsValidCompany).Select(c => ToCompanyRow(c)).ToList();
            if (fallbackCompanies.Count > 0) conn.Execute(SqlCompanyUpsert, fallbackCompanies, trans);

            void AddCompanyConflict(CompanyRecord c, string what, DateTime? expected)
            {
                var serverTs = ServerUpdatedAt("SELECT updated_at FROM company_master WHERE company_id = @CompanyId;", new { CompanyId = c.CompanyID });
                conflicts.Add(new ConcurrencyConflict("company_master", $"Company #{c.CompanyID}",
                    $"Company #{c.CompanyID} ({c.ShortCompanyName}) {what}",
                    expected, serverTs,
                    ExcelTableName: "CompaniesTable",
                    KeyColumnHeader: "Company ID",
                    RawKeyValue: c.CompanyID.ToString(CultureInfo.InvariantCulture),
                    FieldDiffs: ConflictDiff.ForCompany(conn, trans, c)));
            }

            foreach (var c in delta.Companies.Added.Where(IsValidCompany))
            {
                if (forceOverwrite)
                {
                    conn.Execute(SqlCompanyUpsert, ToCompanyRow(c), trans);
                    continue;
                }
                if (conn.Execute(SqlCompanyInsertIfAbsent, ToCompanyRow(c), trans) == 0)
                    AddCompanyConflict(c, "already exists on the server (added by another user).", null);
            }

            foreach (var item in delta.Companies.Modified.Where(m => IsValidCompany(m.Record)))
            {
                var c = item.Record;
                if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
                {
                    if (conn.Execute(SqlCompanyConditionalUpdate, ToCompanyRow(c, item.ExpectedUpdatedAt), trans) == 0)
                        AddCompanyConflict(c, "was modified on the server.", item.ExpectedUpdatedAt);
                }
                else
                {
                    conn.Execute(SqlCompanyUpsert, ToCompanyRow(c), trans);
                }
            }

            // b. project_master
            var fallbackProjects = delta.Projects.FallbackUpsert.Where(IsValidProject).Select(p => ToProjectRow(p)).ToList();
            if (fallbackProjects.Count > 0) conn.Execute(SqlProjectUpsert, fallbackProjects, trans);

            void AddProjectConflict(ProjectRecord p, string what, DateTime? expected)
            {
                string code = NormalizeProjectCode(p.ProjectCode);
                var serverTs = ServerUpdatedAt("SELECT updated_at FROM project_master WHERE project_id = @ProjectId;", new { ProjectId = code });
                conflicts.Add(new ConcurrencyConflict("project_master", $"Project {code}",
                    $"Project {code} {what}",
                    expected, serverTs,
                    ExcelTableName: ProjectExcelTable(p.Category),
                    KeyColumnHeader: "Project Code",
                    RawKeyValue: code,
                    FieldDiffs: ConflictDiff.ForProject(conn, trans, p)));
            }

            foreach (var p in delta.Projects.Added.Where(IsValidProject))
            {
                if (forceOverwrite)
                {
                    conn.Execute(SqlProjectUpsert, ToProjectRow(p), trans);
                    continue;
                }
                if (conn.Execute(SqlProjectInsertIfAbsent, ToProjectRow(p), trans) == 0)
                    AddProjectConflict(p, "already exists on the server (added by another user).", null);
            }

            foreach (var item in delta.Projects.Modified.Where(m => IsValidProject(m.Record)))
            {
                var p = item.Record;
                if (!forceOverwrite && item.ExpectedUpdatedAt.HasValue)
                {
                    if (conn.Execute(SqlProjectConditionalUpdate, ToProjectRow(p, item.ExpectedUpdatedAt), trans) == 0)
                        AddProjectConflict(p, "was modified on the server.", item.ExpectedUpdatedAt);
                }
                else
                {
                    conn.Execute(SqlProjectUpsert, ToProjectRow(p), trans);
                }
            }

            // c. Any conflict rolls back the whole batch (ExecuteInTransaction catches, rolls back, rethrows).
            if (conflicts.Count > 0) throw new ConcurrencyConflictException(conflicts);
        });

        AppLogger.Info($"[ServerDispatcher] Project delta saved (force={forceOverwrite}): " +
            $"companies +{delta.Companies.Added.Count}/~{delta.Companies.Modified.Count}/fallback {delta.Companies.FallbackUpsert.Count}/={delta.Companies.UnchangedCount}; " +
            $"projects +{delta.Projects.Added.Count}/~{delta.Projects.Modified.Count}/fallback {delta.Projects.FallbackUpsert.Count}/={delta.Projects.UnchangedCount}.");
    }

    // -------------------------------------------------------------------------
    // 4e. IDEMPOTENT SCHEMA MIGRATIONS
    // -------------------------------------------------------------------------
    /// <summary>Safe on a fresh database and on the legacy project_master layout.</summary>
    private static void EnsureProjectSchema(System.Data.IDbConnection conn, System.Data.IDbTransaction trans)
    {
        conn.Execute(@"
CREATE TABLE IF NOT EXISTS project_master (
    project_id      VARCHAR(64) PRIMARY KEY,
    project_type    VARCHAR(64) DEFAULT 'None',
    company_name    VARCHAR(256) DEFAULT '',
    start_date      DATE NOT NULL DEFAULT CURRENT_DATE,
    target_end_date DATE NOT NULL DEFAULT (CURRENT_DATE + INTERVAL '365 days'),
    updated_at      TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
);
CREATE UNIQUE INDEX IF NOT EXISTS ux_project_master_project_id ON project_master (project_id);
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS category            VARCHAR(16)   DEFAULT 'iJob';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS reference_no        VARCHAR(128)  DEFAULT '';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS priority            VARCHAR(32)   DEFAULT 'Standard';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS status              VARCHAR(32)   DEFAULT 'Active';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS company_id          BIGINT;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS staff_id            BIGINT;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS pic                 VARCHAR(128)  DEFAULT '';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS district            VARCHAR(128)  DEFAULT 'None';
ALTER TABLE project_master ALTER COLUMN project_type TYPE VARCHAR(512);
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS address             TEXT          DEFAULT '';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS actual_end_date     DATE;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS estimated_cost      NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS promised_return     NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS actual_return       NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS remarks             TEXT          DEFAULT '';
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS days_worked         INT           DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS total_regular_shifts NUMERIC(14,1) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS total_hours         NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS overtime_shifts     INT           DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS total_overtime_hours NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS total_odometer      NUMERIC(14,1) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS fuel_cost           NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS labour_cost         NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS claim_cost          NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS actual_project_cost NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS profit_margin       NUMERIC(14,2) DEFAULT 0;
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS sql_account_project_code VARCHAR(80);
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS updated_at          TIMESTAMPTZ   DEFAULT CURRENT_TIMESTAMP;", transaction: trans);
    }

    private static void EnsureCompanySchema(System.Data.IDbConnection conn, System.Data.IDbTransaction trans)
    {
        conn.Execute(@"
CREATE TABLE IF NOT EXISTS company_master (
    company_id         BIGINT PRIMARY KEY,
    legal_company_name VARCHAR(256) NOT NULL DEFAULT '',
    short_company_name VARCHAR(128) NOT NULL DEFAULT '',
    created_at         TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at         TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
);
ALTER TABLE company_master ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP;
ALTER TABLE company_master ADD COLUMN IF NOT EXISTS default_sql_account_project VARCHAR(80);", transaction: trans);
    }

    // =========================================================================
    // 5. SQL ACCOUNTING PROJECT LOOKUP (lookup_sql_account_projects)
    // =========================================================================
    private const string SqlEnsureSqlAccountProjects = @"
CREATE TABLE IF NOT EXISTS lookup_sql_account_projects (
    code        VARCHAR(80)  PRIMARY KEY,
    description VARCHAR(320),
    is_active   BOOLEAN      NOT NULL DEFAULT TRUE,
    updated_at  TIMESTAMPTZ  NOT NULL DEFAULT CURRENT_TIMESTAMP
);";

    private const string SqlSqlAccountProjectUpsert = @"
INSERT INTO lookup_sql_account_projects (code, description, is_active, updated_at)
VALUES (@Code, @Description, @IsActive, CURRENT_TIMESTAMP)
ON CONFLICT (code) DO UPDATE SET
    description = EXCLUDED.description,
    is_active   = EXCLUDED.is_active,
    updated_at  = CURRENT_TIMESTAMP;";

    /// <summary>
    /// Mirrors the SQL Accounting PROJECT list into PostgreSQL: upserts every row, then marks codes no longer
    /// returned by SQL Accounting as inactive (never hard-deleted).
    /// </summary>
    public static void SyncSqlAccountProjects(IReadOnlyList<SKSB_App.Core.Integration.Erp.AccountingProjectRow> rows)
    {
        var clean = (rows ?? Array.Empty<SKSB_App.Core.Integration.Erp.AccountingProjectRow>())
            .Where(r => !string.IsNullOrWhiteSpace(r.Code))
            .GroupBy(r => r.Code.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .Select(r => new
            {
                Code = r.Code.Trim(),
                Description = (r.Description ?? string.Empty).Trim(),
                r.IsActive
            })
            .ToList();
        string[] codes = clean.Select(r => r.Code).ToArray();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute(SqlEnsureSqlAccountProjects, transaction: trans);
            if (clean.Count > 0) conn.Execute(SqlSqlAccountProjectUpsert, clean, trans);
            if (codes.Length > 0)
            {
                conn.Execute(
                    "UPDATE lookup_sql_account_projects SET is_active = FALSE, updated_at = CURRENT_TIMESTAMP WHERE NOT (code = ANY(@Codes));",
                    new { Codes = codes }, trans);
            }
        });

        AppLogger.Info($"[ServerDispatcher] lookup_sql_account_projects synced ({clean.Count} row(s)).");
    }

    /// <summary>Active SQL Accounting project codes cached in PostgreSQL (used when SQL Accounting itself is unreachable).</summary>
    public static List<string> FetchActiveSqlAccountProjectCodes()
    {
        var codes = new List<string>();
        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute(SqlEnsureSqlAccountProjects, transaction: trans);
            codes = conn.Query<string>(
                "SELECT code FROM lookup_sql_account_projects WHERE is_active = TRUE ORDER BY code;",
                transaction: trans).ToList();
        });
        return codes;
    }
}
