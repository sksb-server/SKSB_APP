namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
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
    /// <summary>Expiry dates of 1900 or earlier (0001-01-01 sentinels) are saved as NULL.</summary>
    private static DateOnly? ValidExpiry(DateOnly? d) => d.HasValue && d.Value.Year > 1900 ? d : null;

    public static void PersistVehicles(IList<VehicleRecord> vehicles, bool isFullBatch)
    {
        var validVehicles = vehicles.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)).ToList();
        if (!validVehicles.Any()) return;

        var activePlates = validVehicles.Select(v => v.VehiclePlate.Trim().ToUpperInvariant()).ToArray();

        var masterRecords = validVehicles.Select(v =>
        {
            var latestRoadTax = v.RoadTaxes.OrderByDescending(r => r.ExpiryDate).FirstOrDefault();
            DateOnly? roadTaxExpiry = ValidExpiry(v.VehicleDetail.RoadTaxExpiry ?? latestRoadTax?.ExpiryDate);
            var latestInspection = v.Inspections.OrderByDescending(i => i.ExpiryDate).FirstOrDefault();
            DateOnly? inspectionExpiry = ValidExpiry(v.VehicleDetail.InspectionExpiry ?? latestInspection?.ExpiryDate);
            return new
            {
                VehiclePlate = v.VehiclePlate.Trim().ToUpperInvariant(),
                Vin = v.VehicleDetail.VIN ?? string.Empty,
                VehicleType = v.VehicleDetail.Type.ToString(),
                Brand = v.VehicleDetail.Brand ?? string.Empty,
                Model = v.VehicleDetail.Model ?? string.Empty,
                RoadTaxExpiry = roadTaxExpiry.HasValue ? (DateTime?)roadTaxExpiry.Value.ToDateTime(TimeOnly.MinValue) : null,
                InspectionExpiry = inspectionExpiry.HasValue ? (DateTime?)inspectionExpiry.Value.ToDateTime(TimeOnly.MinValue) : null,
                DrivenBy = v.VehicleDetail.DrivenBy ?? string.Empty,
                UsedAt = v.VehicleDetail.UsedAt ?? string.Empty,
                CurrentOdometer = v.VehicleDetail.CurrentOdometer,
                BoughtDate = v.VehicleDetail.BoughtDate.HasValue
                    ? (DateTime?)v.VehicleDetail.BoughtDate.Value.ToDateTime(TimeOnly.MinValue)
                    : null,
                BoughtPrice = v.VehicleDetail.BoughtPrice,
                SoldDate = v.VehicleDetail.SoldDate.HasValue
                    ? (DateTime?)v.VehicleDetail.SoldDate.Value.ToDateTime(TimeOnly.MinValue)
                    : null,
                SoldPrice = v.VehicleDetail.SoldPrice,
                Status = v.VehicleStatus.ToString()
            };
        }).ToList();

        var maintenanceRecords = validVehicles.SelectMany(v => v.Maintenances.Select(m =>
        {
            if (string.IsNullOrWhiteSpace(m.VehiclePlate)) m.VehiclePlate = v.VehiclePlate;
            return new
            {
                MaintenanceId = m.MaintenanceID,
                VehiclePlate = m.VehiclePlate.Trim().ToUpperInvariant(),
                ServiceType = m.ServiceType.ToString(),
                DateIn = m.DateIn.ToDateTime(TimeOnly.MinValue),
                EstimatedReturn = m.EstimatedReturn.ToDateTime(TimeOnly.MinValue),
                DateCompleted = m.DateCompleted.HasValue ? (DateTime?)m.DateCompleted.Value.ToDateTime(TimeOnly.MinValue) : null,
                OdometerIn = m.OdometerIn,
                OdometerOut = m.OdometerOut,
                RepairCost = m.RepairCost,
                WorkshopName = m.WorkshopName,
                WorkshopAddress = m.WorkshopAddress ?? string.Empty,
                Status = m.MaintenanceStatus.ToString(),
                Remarks = m.Remarks,
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
                PartsCost = m.PartsCost
            };
        })).ToList();

        var inspectionRecords = validVehicles.SelectMany(v => v.Inspections.Select(i =>
        {
            if (string.IsNullOrWhiteSpace(i.VehiclePlate)) i.VehiclePlate = v.VehiclePlate;
            return new
            {
                InspectionId = i.InspectionID,
                VehiclePlate = i.VehiclePlate.Trim().ToUpperInvariant(),
                ScheduleDateIn = i.ScheduleDateIn.ToDateTime(TimeOnly.MinValue),
                DateOut = i.DateOut.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = i.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Amount = i.Amount,
                Status = i.Status.ToString(),
                Address = i.Address ?? string.Empty,
                Remarks = i.Remarks ?? string.Empty
            };
        })).ToList();

        var roadTaxRecords = validVehicles.SelectMany(v => v.RoadTaxes.Select(rt =>
        {
            if (string.IsNullOrWhiteSpace(rt.VehiclePlate)) rt.VehiclePlate = v.VehiclePlate;
            return new
            {
                RoadTaxId = rt.RoadTaxID,
                VehiclePlate = rt.VehiclePlate.Trim().ToUpperInvariant(),
                VehicleType = rt.VehicleType.ToString(),
                RenewDate = rt.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = rt.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Amount = rt.Amount,
                Status = rt.Status.ToString(),
                Remarks = rt.Remarks ?? string.Empty
            };
        })).ToList();

        var permitRecords = validVehicles.SelectMany(v => v.Permits.Select(p =>
        {
            if (string.IsNullOrWhiteSpace(p.VehiclePlate)) p.VehiclePlate = v.VehiclePlate;
            return new
            {
                PermitId = p.PermitID,
                VehiclePlate = p.VehiclePlate.Trim().ToUpperInvariant(),
                VehicleType = p.VehicleType.ToString(),
                PermitType = p.PermitType.ToString(),
                RenewDate = p.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = p.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Amount = p.Amount,
                Status = p.Status.ToString(),
                Remarks = p.Remarks ?? string.Empty
            };
        })).ToList();

        var insuranceRecords = validVehicles.SelectMany(v => v.Insurances.Select(ins =>
        {
            if (string.IsNullOrWhiteSpace(ins.VehiclePlate)) ins.VehiclePlate = v.VehiclePlate;
            return new
            {
                InsuranceId = ins.InsuranceID,
                VehiclePlate = ins.VehiclePlate.Trim().ToUpperInvariant(),
                VehicleType = ins.VehicleType.ToString(),
                InsuranceCompany = ins.InsuranceCompany ?? string.Empty,
                InsuranceType = ins.InsuranceType.ToString(),
                RenewDate = ins.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = ins.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Amount = ins.Amount,
                Status = ins.Status.ToString(),
                Remarks = ins.Remarks ?? string.Empty,
                InsuranceNumber = ins.InsuranceNumber ?? string.Empty // <-- Added
            };
        })).ToList();

        var odoRecords = validVehicles.SelectMany(v => v.OdometerLogs.Select(o =>
        {
            if (string.IsNullOrWhiteSpace(o.VehiclePlate)) o.VehiclePlate = v.VehiclePlate;
            return new
            {
                OdometerId = o.OdometerID,
                Date = o.Date.ToDateTime(TimeOnly.MinValue),
                VehiclePlate = o.VehiclePlate.Trim().ToUpperInvariant(),
                VehicleType = o.VehicleType.ToString(),
                StaffId = o.StaffID,                               // <-- Updated from DrivenBy
                Project = o.Project ?? string.Empty,
                OdometerBefore = o.OdometerBefore,
                CurrentOdometer = o.CurrentOdometer,
                FuelCardNumber = o.FuelCardNumber ?? string.Empty,
                FuelCompany = o.FuelCompany ?? string.Empty,
                StationName = o.StationName ?? string.Empty,
                ReceiptNumber = o.ReceiptNumber ?? string.Empty,
                Time = o.Time ?? string.Empty,
                FuelType = o.FuelType ?? string.Empty,
                PriceAtFueling = o.PriceAtFueling,
                FuelAmount = o.FuelAmount,
                Amount = o.Amount,
                Remarks = o.Remarks ?? string.Empty
            };
        })).ToList();

        var insuranceClaimRecords = validVehicles.SelectMany(v => (v.InsuranceClaims ?? new List<InsuranceClaim>()).Select(c => new
        {
            ClaimId = c.ClaimID,
            ClaimDate = c.ClaimDate.ToDateTime(TimeOnly.MinValue),
            SettlementDate = c.SettlementDate.HasValue
                ? (DateTime?)c.SettlementDate.Value.ToDateTime(TimeOnly.MinValue)
                : null,
            VehiclePlate = (string.IsNullOrWhiteSpace(c.VehiclePlate) ? v.VehiclePlate : c.VehiclePlate).Trim().ToUpperInvariant(),
            VehicleType = c.VehicleType.ToString(),
            StaffId = c.StaffID,
            MaintenanceId = c.MaintenanceID,
            WorkshopName = c.WorkshopName ?? string.Empty,
            Status = c.Status.ToString(),
            ClaimAmount = c.ClaimAmount,
            Remarks = c.Remarks ?? string.Empty
        })).ToList();

        // 2. Define UPSERT queries
        const string sqlMasterUpsert = @"
INSERT INTO vehicle_master (
    vehicle_plate, vin, vehicle_type, brand, model,
    road_tax_expiry, inspection_expiry, current_odometer,
    bought_date, bought_price, sold_date, sold_price,
    driven_by, used_at, status
)
VALUES (
    @VehiclePlate, @Vin, @VehicleType, @Brand, @Model,
    @RoadTaxExpiry, @InspectionExpiry, @CurrentOdometer,
    @BoughtDate, @BoughtPrice, @SoldDate, @SoldPrice,
    @DrivenBy, @UsedAt, @Status
)
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
    updated_at        = CURRENT_TIMESTAMP;";

        const string sqlMaintenanceUpsert = @"
        INSERT INTO vehicle_maintenance (
            maintenance_id, vehicle_plate, service_type, date_in, estimated_return,
            date_completed, odometer_in, odometer_out, repair_cost, workshop_name, workshop_address,
            status, remarks,
            tyre_cost, brake_cost, clutch_cost, service_cost, fuel_pump_cost,
            aircond_cost, wiring_cost, battery_cost, towing_cost, cushion_cost, parts_cost
        )
        VALUES (
            @MaintenanceId, @VehiclePlate, @ServiceType, @DateIn, @EstimatedReturn,
            @DateCompleted, @OdometerIn, @OdometerOut, @RepairCost, @WorkshopName, @WorkshopAddress,
            @Status, @Remarks,
            @TyreCost, @BrakeCost, @ClutchCost, @ServiceCost, @FuelPumpCost,
            @AircondCost, @WiringCost, @BatteryCost, @TowingCost, @CushionCost, @PartsCost
        )
        ON CONFLICT (maintenance_id) DO UPDATE SET
            vehicle_plate    = EXCLUDED.vehicle_plate,
            service_type     = EXCLUDED.service_type,
            date_in          = EXCLUDED.date_in,
            estimated_return = EXCLUDED.estimated_return,
            date_completed   = EXCLUDED.date_completed,
            odometer_in      = EXCLUDED.odometer_in,
            odometer_out     = EXCLUDED.odometer_out,
            repair_cost      = EXCLUDED.repair_cost,
            workshop_name    = EXCLUDED.workshop_name,
            workshop_address = EXCLUDED.workshop_address,
            status           = EXCLUDED.status,
            remarks          = EXCLUDED.remarks,
            tyre_cost        = EXCLUDED.tyre_cost,
            brake_cost       = EXCLUDED.brake_cost,
            clutch_cost      = EXCLUDED.clutch_cost,
            service_cost     = EXCLUDED.service_cost,
            fuel_pump_cost   = EXCLUDED.fuel_pump_cost,
            aircond_cost     = EXCLUDED.aircond_cost,
            wiring_cost      = EXCLUDED.wiring_cost,
            battery_cost     = EXCLUDED.battery_cost,
            towing_cost      = EXCLUDED.towing_cost,
            cushion_cost     = EXCLUDED.cushion_cost,
            parts_cost       = EXCLUDED.parts_cost,
            updated_at       = CURRENT_TIMESTAMP;";

        const string sqlInspectionUpsert = @"
            INSERT INTO vehicle_inspections (
                inspection_id, vehicle_plate, schedule_date_in, date_out,
                expiry_date, amount, status, address, remarks, updated_at
            ) VALUES (
                @InspectionId, @VehiclePlate, @ScheduleDateIn, @DateOut,
                @ExpiryDate, @Amount, @Status, @Address, @Remarks, CURRENT_TIMESTAMP
            )
            ON CONFLICT (inspection_id) DO UPDATE SET
                vehicle_plate    = EXCLUDED.vehicle_plate,
                schedule_date_in = EXCLUDED.schedule_date_in,
                date_out         = EXCLUDED.date_out,
                expiry_date      = EXCLUDED.expiry_date,
                amount           = EXCLUDED.amount,
                status           = EXCLUDED.status,
                address          = EXCLUDED.address,
                remarks          = EXCLUDED.remarks,
                updated_at       = CURRENT_TIMESTAMP;";

        const string sqlRoadTaxUpsert = @"
INSERT INTO vehicle_road_taxes (
    road_tax_id, vehicle_plate, vehicle_type, renew_date, expiry_date, amount, status, remarks, updated_at
) VALUES (
    @RoadTaxId, @VehiclePlate, @VehicleType, @RenewDate, @ExpiryDate, @Amount, @Status, @Remarks, CURRENT_TIMESTAMP
)
ON CONFLICT (road_tax_id) DO UPDATE SET
    vehicle_plate = EXCLUDED.vehicle_plate,
    vehicle_type  = EXCLUDED.vehicle_type,
    renew_date    = EXCLUDED.renew_date,
    expiry_date   = EXCLUDED.expiry_date,
    amount        = EXCLUDED.amount,
    status        = EXCLUDED.status,
    remarks       = EXCLUDED.remarks,
    updated_at    = CURRENT_TIMESTAMP;";

        const string sqlPermitUpsert = @"
INSERT INTO vehicle_permits (
    permit_id, vehicle_plate, vehicle_type, permit_type, renew_date, expiry_date, amount, status, remarks, updated_at
) VALUES (
    @PermitId, @VehiclePlate, @VehicleType, @PermitType, @RenewDate, @ExpiryDate, @Amount, @Status, @Remarks, CURRENT_TIMESTAMP
)
ON CONFLICT (permit_id) DO UPDATE SET
    vehicle_plate = EXCLUDED.vehicle_plate,
    vehicle_type  = EXCLUDED.vehicle_type,
    permit_type   = EXCLUDED.permit_type,
    renew_date    = EXCLUDED.renew_date,
    expiry_date   = EXCLUDED.expiry_date,
    amount        = EXCLUDED.amount,
    status        = EXCLUDED.status,
    remarks       = EXCLUDED.remarks,
    updated_at    = CURRENT_TIMESTAMP;";

        const string sqlInsuranceUpsert = @"
INSERT INTO vehicle_insurances (
    insurance_id, vehicle_plate, vehicle_type, insurance_company, insurance_type, 
    renew_date, expiry_date, amount, status, remarks, updated_at, insurance_number
) VALUES (
    @InsuranceId, @VehiclePlate, @VehicleType, @InsuranceCompany, @InsuranceType, 
    @RenewDate, @ExpiryDate, @Amount, @Status, @Remarks, CURRENT_TIMESTAMP, @InsuranceNumber
)
ON CONFLICT (insurance_id) DO UPDATE SET
    vehicle_plate     = EXCLUDED.vehicle_plate,
    vehicle_type      = EXCLUDED.vehicle_type,
    insurance_company = EXCLUDED.insurance_company,
    insurance_type    = EXCLUDED.insurance_type,
    renew_date        = EXCLUDED.renew_date,
    expiry_date       = EXCLUDED.expiry_date,
    amount            = EXCLUDED.amount,
    status            = EXCLUDED.status,
    remarks           = EXCLUDED.remarks,
    insurance_number  = EXCLUDED.insurance_number, -- <-- Added
    updated_at        = CURRENT_TIMESTAMP;";

        const string sqlOdometerUpsert = @"
INSERT INTO vehicle_odometers (
    odometer_id, date, vehicle_plate, vehicle_type, staff_id, project,
    odometer_before, current_odometer, fuel_card_number, fuel_company,
    station_name, receipt_number, amount, remarks, time, fuel_type, price_at_fueling, fuel_amount, updated_at
) VALUES (
    @OdometerId, @Date, @VehiclePlate, @VehicleType, @StaffId, @Project,
    @OdometerBefore, @CurrentOdometer, @FuelCardNumber, @FuelCompany,
    @StationName, @ReceiptNumber, @Amount, @Remarks, @Time, @FuelType, @PriceAtFueling, @FuelAmount, CURRENT_TIMESTAMP
)
ON CONFLICT (odometer_id) DO UPDATE SET
    date             = EXCLUDED.date,
    vehicle_plate    = EXCLUDED.vehicle_plate,
    vehicle_type     = EXCLUDED.vehicle_type,
    staff_id         = EXCLUDED.staff_id,
    project          = EXCLUDED.project,
    odometer_before  = EXCLUDED.odometer_before,
    current_odometer = EXCLUDED.current_odometer,
    fuel_card_number = EXCLUDED.fuel_card_number,
    fuel_company     = EXCLUDED.fuel_company,
    station_name     = EXCLUDED.station_name,
    receipt_number   = EXCLUDED.receipt_number,
    amount           = EXCLUDED.amount,
    remarks          = EXCLUDED.remarks,
    time             = EXCLUDED.time,
    fuel_type        = EXCLUDED.fuel_type,
    price_at_fueling = EXCLUDED.price_at_fueling,
    fuel_amount      = EXCLUDED.fuel_amount,
    updated_at       = CURRENT_TIMESTAMP;";

        const string sqlInsuranceClaimUpsert = @"
INSERT INTO vehicle_insurance_claims (
    claim_id, claim_date, settlement_date, vehicle_plate, vehicle_type,
    staff_id, maintenance_id, workshop_name, status, claim_amount, remarks, updated_at
) VALUES (
    @ClaimId, @ClaimDate, @SettlementDate, @VehiclePlate, @VehicleType,
    @StaffId, @MaintenanceId, @WorkshopName, @Status, @ClaimAmount, @Remarks, CURRENT_TIMESTAMP
)
ON CONFLICT (claim_id) DO UPDATE SET
    claim_date      = EXCLUDED.claim_date,
    settlement_date = EXCLUDED.settlement_date,
    vehicle_plate   = EXCLUDED.vehicle_plate,
    vehicle_type    = EXCLUDED.vehicle_type,
    staff_id        = EXCLUDED.staff_id,
    maintenance_id  = EXCLUDED.maintenance_id,
    workshop_name   = EXCLUDED.workshop_name,
    status          = EXCLUDED.status,
    claim_amount    = EXCLUDED.claim_amount,
    remarks         = EXCLUDED.remarks,
    updated_at      = CURRENT_TIMESTAMP;";

        // 3. Execute in Transaction
        Db.ExecuteInTransaction((conn, trans) =>
        {
            // Defensive migration: must precede the full-batch DELETE below on a server that predates this table.
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

            // Defensive idempotent migration: odometers / engine-hour meters carry 1 decimal place (NUMERIC(12,1)).
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

            if (isFullBatch)
            {
                conn.Execute("DELETE FROM vehicle_insurance_claims WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_inspections WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_road_taxes WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_permits WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_insurances WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_odometers WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_maintenance WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_master WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
            }

            // Idempotent migrations: no 10-year placeholder expiry defaults, plus the Driven By / Used At columns.
            conn.Execute(@"
ALTER TABLE vehicle_master ALTER COLUMN road_tax_expiry DROP DEFAULT;
ALTER TABLE vehicle_master ALTER COLUMN inspection_expiry DROP DEFAULT;
ALTER TABLE vehicle_master ADD COLUMN IF NOT EXISTS driven_by VARCHAR(128) DEFAULT '';
ALTER TABLE vehicle_master ADD COLUMN IF NOT EXISTS used_at   VARCHAR(256) DEFAULT '';
ALTER TABLE vehicle_maintenance ADD COLUMN IF NOT EXISTS workshop_address TEXT DEFAULT '';", transaction: trans);

            conn.Execute(sqlMasterUpsert, masterRecords, trans);

            if (maintenanceRecords.Any()) conn.Execute(sqlMaintenanceUpsert, maintenanceRecords, trans);
            if (inspectionRecords.Any()) conn.Execute(sqlInspectionUpsert, inspectionRecords, trans);
            if (roadTaxRecords.Any()) conn.Execute(sqlRoadTaxUpsert, roadTaxRecords, trans);
            if (permitRecords.Any()) conn.Execute(sqlPermitUpsert, permitRecords, trans);
            if (insuranceRecords.Any()) conn.Execute(sqlInsuranceUpsert, insuranceRecords, trans);
            if (insuranceClaimRecords.Any()) conn.Execute(sqlInsuranceClaimUpsert, insuranceClaimRecords, trans);
            if (odoRecords.Any())
            {
                conn.Execute("ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS time VARCHAR(16) DEFAULT '';", transaction: trans);
                conn.Execute("ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS fuel_type VARCHAR(64) DEFAULT '';", transaction: trans);
                conn.Execute("ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS price_at_fueling NUMERIC(12,4) DEFAULT 0;", transaction: trans);
                conn.Execute("ALTER TABLE vehicle_odometers ADD COLUMN IF NOT EXISTS fuel_amount NUMERIC(12,4) DEFAULT 0;", transaction: trans);
                conn.Execute(sqlOdometerUpsert, odoRecords, trans);
            }
        });

#if DEBUG
        var sb = new StringBuilder();
        sb.AppendLine("=== POSTGRESQL VEHICLE UPLOAD COMMITTED ===");
        sb.AppendLine($"Committed Vehicle Profiles: {validVehicles.Count}");
        sb.AppendLine($"InsuranceClaims: {insuranceClaimRecords.Count}");
        sb.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        AppLogger.Info(sb.ToString());
#endif
    }

    public static void SaveAuxiliaryTables(
        IEnumerable<Workshop> workshops,
        IEnumerable<InsuranceCompanyRecord> insuranceCompanies,
        IEnumerable<FuelCardRecord> fuelCards,
        IEnumerable<string> fuelCompanies)
    {
        var companyNames = fuelCompanies?
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToArray() ?? Array.Empty<string>();

        var workshopList = workshops?.ToList() ?? new List<Workshop>();
        var insuranceList = insuranceCompanies?.ToList() ?? new List<InsuranceCompanyRecord>();
        var cardList = fuelCards?.ToList() ?? new List<FuelCardRecord>();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            // Workshops
            if (workshopList.Count > 0)
            {
                conn.Execute(@"
                    INSERT INTO vehicle_workshops (workshop_id, workshop_name, workshop_address, workshop_owner, primary_contact, secondary_contact)
                    VALUES (@WorkshopID, @WorkshopName, @WorkshopAddress, @WorkshopOwner, @PrimaryContact, @SecondaryContact)
                    ON CONFLICT (workshop_id) DO UPDATE SET
                        workshop_name = EXCLUDED.workshop_name,
                        workshop_address = EXCLUDED.workshop_address,
                        workshop_owner = EXCLUDED.workshop_owner,
                        primary_contact = EXCLUDED.primary_contact,
                        secondary_contact = EXCLUDED.secondary_contact,
                        updated_at = CURRENT_TIMESTAMP;", workshopList, trans);
            }

            // Insurance Companies
            if (insuranceList.Count > 0)
            {
                conn.Execute(@"
                    INSERT INTO vehicle_insurance_companies (insurance_company, company_contact_number, company_address, insurance_agent, agent_contact_number)
                    VALUES (@InsuranceCompany, @CompanyContactNumber, @CompanyAddress, @InsuranceAgent, @AgentContactNumber)
                    ON CONFLICT (insurance_company) DO UPDATE SET
                        company_contact_number = EXCLUDED.company_contact_number,
                        company_address = EXCLUDED.company_address,
                        insurance_agent = EXCLUDED.insurance_agent,
                        agent_contact_number = EXCLUDED.agent_contact_number,
                        updated_at = CURRENT_TIMESTAMP;", insuranceList, trans);
            }

            // Fuel Companies Lookup - Sync
            if (companyNames.Length > 0)
            {
                // A. Delete only companies removed from Excel that are NO LONGER referenced by any fuel card
                const string sqlReconcile = @"
                    DELETE FROM lookup_fuel_companies 
                    WHERE fuel_company_name != ALL(@ActiveNames);";

                conn.Execute(sqlReconcile, new { ActiveNames = companyNames }, transaction: trans);

                // B. Insert newly added companies
                const string sqlUpsert = @"
                    INSERT INTO lookup_fuel_companies (fuel_company_name)
                    VALUES (@Name)
                    ON CONFLICT (fuel_company_name) DO NOTHING;";

                var payload = companyNames.Select(n => new { Name = n });
                conn.Execute(sqlUpsert, payload, transaction: trans);
            }
            else
            {
                // If the table was deliberately wiped completely
                conn.Execute("DELETE FROM lookup_fuel_companies;", transaction: trans);
            }

            // Fuel Cards
            if (cardList.Count > 0)
            {
                conn.Execute(@"
                    INSERT INTO vehicle_fuel_cards (fuel_card_number, fuel_company_name, card_owner_id, card_owner_name)
                    VALUES (@FuelCardNumber, @FuelCompanyName, @CardOwnerID, @CardOwnerName)
                    ON CONFLICT (fuel_card_number) DO UPDATE SET
                        fuel_company_name = EXCLUDED.fuel_company_name,
                        card_owner_id = EXCLUDED.card_owner_id,
                        card_owner_name = EXCLUDED.card_owner_name,
                        updated_at = CURRENT_TIMESTAMP;", cardList, trans);
            }
        });
    }

    // =========================================================================
    // 3. STAFF MANAGER PERSISTENCE
    // =========================================================================
    public static void PersistStaff(IList<StaffRecord> staffList, bool isFullBatch)
    {
        var staffItems = staffList.Where(s => s.StaffID > 0).ToList();
        if (!staffItems.Any()) return;

        var activeIds = staffItems.Select(s => s.StaffID).ToArray();

        var masterRecords = staffItems.Select(s => new
        {
            StaffId = s.StaffID,
            StaffName = s.StaffDetails.StaffName ?? string.Empty,
            ShortName = s.StaffDetails.StaffShortName ?? string.Empty,
            Department = s.StaffDetails.Department ?? "Worker",
            Nric = s.StaffDetails.NRIC,
            Passport = s.StaffDetails.Passport ?? string.Empty,
            Phone = s.StaffDetails.Phone ?? string.Empty,
            JoinDate = s.StaffDetails.JoinDate.ToDateTime(TimeOnly.MinValue),
            ResignDate = s.StaffDetails.ResignDate.HasValue ? (DateTime?)s.StaffDetails.ResignDate.Value.ToDateTime(TimeOnly.MinValue) : null,
            ContractStatus = s.StaffDetails.ContractStatus.ToString()
        }).ToList();

        const string sqlMasterUpsert = @"
        INSERT INTO staff_master (
            staff_id, staff_name, short_name, department, nric, passport, phone, join_date, resign_date, contract_status
        )
        VALUES (
            @StaffId, @StaffName, @ShortName, @Department, @Nric, @Passport, @Phone, @JoinDate, @ResignDate, @ContractStatus
        )
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
            updated_at      = CURRENT_TIMESTAMP;";

        Db.ExecuteInTransaction((conn, trans) =>
        {
            if (isFullBatch)
            {
                conn.Execute("DELETE FROM staff_vehicle_licenses WHERE staff_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
                conn.Execute("DELETE FROM staff_skill_licenses WHERE staff_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
                conn.Execute("DELETE FROM staff_master WHERE staff_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
            }

            conn.Execute(sqlMasterUpsert, masterRecords, trans);
        });

        // Child Licenses
        var allVLicenses = staffItems.SelectMany(s => s.VehicleLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList();
        PersistStaffVehicleLicenses(allVLicenses);

        var allSLicenses = staffItems.SelectMany(s => s.SkillLicenses.Select(l => { if (l.StaffID == 0) l.StaffID = s.StaffID; return l; })).ToList();
        PersistStaffSkillLicenses(allSLicenses);
    }

    private static void PersistStaffVehicleLicenses(IEnumerable<VehicleLicense> licenses)
    {
        var items = licenses.Where(l => l.StaffID > 0 && l.VehicleLicenseClass != VehicleLicenseClass.None).Select(l => new
        {
            VehicleLicenseId = l.VehicleLicenseID,
            StaffId = l.StaffID,
            VehicleLicenseClass = l.VehicleLicenseClass.ToString(),
            RenewDate = l.RenewDate.ToDateTime(TimeOnly.MinValue),
            ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
            Status = l.VehicleLicenseStatus.ToString(),
            Cost = l.Cost
        }).ToList();

        if (!items.Any()) return;

        Db.ExecuteTransaction(@"
            INSERT INTO staff_vehicle_licenses (vehicle_license_id, staff_id, vehicle_license_class, renew_date, expiry_date, status, cost)
            VALUES (
                CASE WHEN @VehicleLicenseId > 0 THEN @VehicleLicenseId ELSE nextval('staff_vehicle_licenses_vehicle_license_id_seq') END,
                @StaffId, @VehicleLicenseClass, @RenewDate, @ExpiryDate, @Status, @Cost
            )
            ON CONFLICT (staff_id, vehicle_license_class) DO UPDATE SET
                renew_date  = EXCLUDED.renew_date,
                expiry_date = EXCLUDED.expiry_date,
                status      = EXCLUDED.status,
                cost        = EXCLUDED.cost,
                updated_at  = CURRENT_TIMESTAMP;", items);
    }

    private static void PersistStaffSkillLicenses(IEnumerable<SkillLicense> licenses)
    {
        var items = licenses.Where(s => s.StaffID > 0 && !string.IsNullOrWhiteSpace(s.SkillLicenseType)).Select(s => new
        {
            SkillLicenseId = s.SkillLicenseID,
            StaffId = s.StaffID,
            SkillLicenseType = s.SkillLicenseType.Trim(),
            RenewDate = s.RenewDate.ToDateTime(TimeOnly.MinValue),
            ExpiryDate = s.ExpiryDate.ToDateTime(TimeOnly.MinValue),
            Status = s.SkillLicenseStatus.ToString(),
            Cost = s.Cost
        }).ToList();

        if (!items.Any()) return;

        Db.ExecuteTransaction(@"
            INSERT INTO staff_skill_licenses (skill_license_id, staff_id, skill_license_type, renew_date, expiry_date, status, cost)
            VALUES (
                CASE WHEN @SkillLicenseId > 0 THEN @SkillLicenseId ELSE nextval('staff_skill_licenses_skill_license_id_seq') END,
                @StaffId, @SkillLicenseType, @RenewDate, @ExpiryDate, @Status, @Cost
            )
            ON CONFLICT (staff_id, skill_license_type) DO UPDATE SET
                renew_date  = EXCLUDED.renew_date,
                expiry_date = EXCLUDED.expiry_date,
                status      = EXCLUDED.status,
                cost        = EXCLUDED.cost,
                updated_at  = CURRENT_TIMESTAMP;", items);
    }

    private static void PersistSkillLicenseTypes(IEnumerable<SkillLicenseType> types)
    {
        var items = types
            .Where(t => !string.IsNullOrWhiteSpace(t.SkillLicenseName))
            .GroupBy(t => t.SkillLicenseName.Trim(), StringComparer.OrdinalIgnoreCase)
            .Select(g => g.First())
            .Select(t => new
            {
                SkillLicenseName = t.SkillLicenseName.Trim(),
                DisplayIcon = (t.DisplayIcon ?? string.Empty).Trim()
            }).ToList();
        string[] names = items.Select(i => i.SkillLicenseName).ToArray();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute("ALTER TABLE lookup_skill_license_types ADD COLUMN IF NOT EXISTS display_icon VARCHAR(32) NOT NULL DEFAULT '';", transaction: trans);
            if (items.Count == 0) return;

            conn.Execute("DELETE FROM lookup_skill_license_types WHERE skill_license_type != ALL(@Names);", new { Names = names }, trans);
            conn.Execute(@"
                INSERT INTO lookup_skill_license_types (skill_license_type, description, display_icon)
                VALUES (@SkillLicenseName, '', @DisplayIcon)
                ON CONFLICT (skill_license_type) DO UPDATE SET
                    display_icon = EXCLUDED.display_icon;", items, trans);
        });
    }

    // =========================================================================
    // 4. PROJECT & COMPANY PERSISTENCE
    // =========================================================================
    public static void PersistProjects(IList<ProjectRecord> projects, bool isFullBatch)
    {
        var items = projects.Where(p => !string.IsNullOrWhiteSpace(p.ProjectCode)).ToList();
        if (!items.Any() && !isFullBatch) return;

        static DateTime? ToDt(DateOnly? d) => d.HasValue ? d.Value.ToDateTime(TimeOnly.MinValue) : null;

        var activeIds = items.Select(p => p.ProjectCode.Trim().ToUpperInvariant()).ToArray();
        var masterRecords = items.Select(p => new
        {
            ProjectId = p.ProjectCode.Trim().ToUpperInvariant(),
            Category = p.Category.ToString(),
            ReferenceNo = p.ProjectDetails.ReferenceNo ?? string.Empty,
            ProjectType = string.Join(", ", (p.ProjectDetails.ProjectType ?? new List<ProjectType>()).Select(t => t.ToString())),
            Priority = p.ProjectDetails.Priority.ToString(),
            Status = p.ProjectDetails.Status.ToString(),
            CompanyId = p.ProjectDetails.CompanyID > 0 ? p.ProjectDetails.CompanyID : (long?)null,
            CompanyName = p.ProjectDetails.CompanyName ?? string.Empty,
            StaffId = p.ProjectDetails.StaffID,
            Pic = p.ProjectDetails.PIC ?? string.Empty,
            District = p.ProjectDetails.District.ToString(),
            Address = p.ProjectDetails.Address ?? string.Empty,
            StartDate = p.ProjectDetails.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = p.ProjectDetails.TargetEndDate.ToDateTime(TimeOnly.MinValue),
            ActualEndDate = ToDt(p.ProjectDetails.ActualEndDate),
            EstimatedCost = p.ProjectDetails.EstimatedCost,
            PromisedReturn = p.ProjectDetails.PromisedReturn,
            ActualReturn = p.ProjectDetails.ActualReturn,
            Remarks = p.ProjectDetails.Remarks ?? string.Empty,
            DaysWorked = p.ProjectDetails.DaysWorked,
            TotalRegularShifts = p.ProjectDetails.TotalRegularShifts,
            TotalHours = p.ProjectDetails.TotalHours,
            OvertimeShifts = p.ProjectDetails.OvertimeShifts,
            TotalOvertimeHours = p.ProjectDetails.TotalOvertimeHours,
            TotalOdometer = p.ProjectDetails.TotalOdometer,
            FuelCost = p.ProjectDetails.FuelCost,
            LabourCost = p.ProjectDetails.LabourCost,
            ClaimCost = p.ProjectDetails.ClaimCost,
            ActualProjectCost = p.ProjectDetails.ActualProjectCost,
            ProfitMargin = p.ProjectDetails.ProfitMargin
        }).ToList();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            // Idempotent migrations: safe on a fresh database and on the legacy project_master layout.
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
ALTER TABLE project_master ADD COLUMN IF NOT EXISTS profit_margin       NUMERIC(14,2) DEFAULT 0;", transaction: trans);

            if (isFullBatch)
            {
                // Reconcile the master table with the workbook.
                conn.Execute("DELETE FROM project_master WHERE project_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
            }

            if (masterRecords.Count == 0) return;

            conn.Execute(@"
INSERT INTO project_master (
    project_id, category, reference_no, project_type, priority, status, company_id, company_name, staff_id, pic,
    district, address, start_date, target_end_date, actual_end_date, estimated_cost, promised_return, actual_return,
    remarks, days_worked, total_regular_shifts, total_hours, overtime_shifts, total_overtime_hours, total_odometer, fuel_cost, labour_cost, claim_cost,
    actual_project_cost, profit_margin, updated_at
) VALUES (
    @ProjectId, @Category, @ReferenceNo, @ProjectType, @Priority, @Status, @CompanyId, @CompanyName, @StaffId, @Pic,
    @District, @Address, @StartDate, @TargetEndDate, @ActualEndDate, @EstimatedCost, @PromisedReturn, @ActualReturn,
    @Remarks, @DaysWorked, @TotalRegularShifts, @TotalHours, @OvertimeShifts, @TotalOvertimeHours, @TotalOdometer, @FuelCost, @LabourCost, @ClaimCost,
    @ActualProjectCost, @ProfitMargin, CURRENT_TIMESTAMP
)
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
    updated_at          = CURRENT_TIMESTAMP;", masterRecords, trans);
        });
    }

    public static void PersistCompanies(IList<CompanyRecord> companies, bool isFullBatch)
    {
        var items = companies.Where(c => c.CompanyID > 0).ToList();
        if (!items.Any() && !isFullBatch) return;

        var activeIds = items.Select(c => c.CompanyID).ToArray();
        var records = items.Select(c => new
        {
            CompanyId = c.CompanyID,
            LegalCompanyName = c.LegalCompanyName ?? string.Empty,
            ShortCompanyName = c.ShortCompanyName ?? string.Empty
        }).ToList();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute(@"
CREATE TABLE IF NOT EXISTS company_master (
    company_id         BIGINT PRIMARY KEY,
    legal_company_name VARCHAR(256) NOT NULL DEFAULT '',
    short_company_name VARCHAR(128) NOT NULL DEFAULT '',
    created_at         TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP,
    updated_at         TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP
);
ALTER TABLE company_master ADD COLUMN IF NOT EXISTS updated_at TIMESTAMPTZ DEFAULT CURRENT_TIMESTAMP;", transaction: trans);

            if (isFullBatch)
            {
                conn.Execute("DELETE FROM company_master WHERE company_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
            }

            if (records.Count == 0) return;

            conn.Execute(@"
                INSERT INTO company_master (company_id, legal_company_name, short_company_name)
                VALUES (@CompanyId, @LegalCompanyName, @ShortCompanyName)
                ON CONFLICT (company_id) DO UPDATE SET
                    legal_company_name = EXCLUDED.legal_company_name,
                    short_company_name = EXCLUDED.short_company_name,
                    updated_at         = CURRENT_TIMESTAMP;", records, trans);
        });
    }
}