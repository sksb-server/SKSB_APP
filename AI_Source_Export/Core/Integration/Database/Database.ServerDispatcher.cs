namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Dapper;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types;

public static class ServerDispatcher
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
            case StaffRecord singleStaff:
                PersistStaff(new[] { singleStaff }, isFullBatch: false);
                break;

            // Project Manager
            case IEnumerable<ProjectRecord> projectBatch:
                PersistProjects(projectBatch.ToList(), isFullBatch: true);
                break;
            case ProjectRecord singleProject:
                PersistProjects(new[] { singleProject }, isFullBatch: false);
                break;

            // Company Master
            case IEnumerable<CompanyRecord> companyBatch:
                PersistCompanies(companyBatch.ToList(), isFullBatch: true);
                break;
            case CompanyRecord singleCompany:
                PersistCompanies(new[] { singleCompany }, isFullBatch: false);
                break;

            // Vehicle Manager
            case IEnumerable<VehicleRecord> vehicleBatch:
                PersistVehicles(vehicleBatch.ToList(), isFullBatch: true);
                break;
            case VehicleRecord singleVehicle:
                PersistVehicles(new[] { singleVehicle }, isFullBatch: false);
                break;

            // Standalone Child Entity Collections
            case IEnumerable<VehicleLicense> vLicenses:
                PersistStaffVehicleLicenses(vLicenses);
                break;
            case IEnumerable<SkillLicense> sLicenses:
                PersistStaffSkillLicenses(sLicenses);
                break;
            case IEnumerable<SkillLicenseType> sTypes:
                PersistSkillLicenseTypes(sTypes);
                break;
        }
    }

    // =========================================================================
    // 2. VEHICLE MANAGER PERSISTENCE
    // =========================================================================
    public static void PersistVehicles(IList<VehicleRecord> vehicles, bool isFullBatch)
    {
        var validVehicles = vehicles.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)).ToList();
        if (!validVehicles.Any()) return;

        var activePlates = validVehicles.Select(v => v.VehiclePlate.Trim().ToUpperInvariant()).ToArray();

        var masterRecords = validVehicles.Select(v =>
        {
            var latestRoadTax = v.RoadTaxes.OrderByDescending(r => r.ExpiryDate).FirstOrDefault();
            var roadTaxExpiry = v.VehicleDetail.RoadTaxExpiry ?? latestRoadTax?.ExpiryDate;

            var latestInspection = v.Inspections.OrderByDescending(i => i.ExpiryDate).FirstOrDefault();
            var inspectionExpiry = v.VehicleDetail.InspectionExpiry ?? latestInspection?.ExpiryDate;

            return new
            {
                VehiclePlate = v.VehiclePlate.Trim().ToUpperInvariant(),
                Vin = v.VehicleDetail.VIN ?? string.Empty,
                VehicleType = v.VehicleDetail.Type.ToString(),
                Brand = v.VehicleDetail.Brand ?? string.Empty,
                Model = v.VehicleDetail.Model ?? string.Empty,
                RoadTaxExpiry = roadTaxExpiry?.ToDateTime(TimeOnly.MinValue),
                InspectionExpiry = inspectionExpiry?.ToDateTime(TimeOnly.MinValue),
                CurrentOdometer = v.VehicleDetail.CurrentOdometer,
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
                WorkshopAddress = m.WorkshopAddress,
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

        const string sqlMasterUpsert = @"
        INSERT INTO vehicle_master (
            vehicle_plate, vin, vehicle_type, brand, model,
            road_tax_expiry, inspection_expiry, current_odometer, status
        )
        VALUES (
            @VehiclePlate, @Vin, @VehicleType, @Brand, @Model,
            @RoadTaxExpiry, @InspectionExpiry, @CurrentOdometer, @Status
        )
        ON CONFLICT (vehicle_plate) DO UPDATE SET
            vin               = COALESCE(NULLIF(EXCLUDED.vin, ''), vehicle_master.vin),
            vehicle_type      = EXCLUDED.vehicle_type,
            brand             = COALESCE(NULLIF(EXCLUDED.brand, ''), vehicle_master.brand),
            model             = COALESCE(NULLIF(EXCLUDED.model, ''), vehicle_master.model),
            road_tax_expiry   = EXCLUDED.road_tax_expiry,
            inspection_expiry = EXCLUDED.inspection_expiry,
            current_odometer  = CASE WHEN EXCLUDED.current_odometer > 0 THEN EXCLUDED.current_odometer ELSE vehicle_master.current_odometer END,
            status            = EXCLUDED.status,
            updated_at        = CURRENT_TIMESTAMP;";

        const string sqlMaintenanceUpsert = @"
        INSERT INTO vehicle_maintenance (
            maintenance_id, vehicle_plate, service_type, date_in, estimated_return,
            date_completed, odometer_in, odometer_out, repair_cost, workshop_name,
            workshop_address, status, remarks,
            tyre_cost, brake_cost, clutch_cost, service_cost, fuel_pump_cost,
            aircond_cost, wiring_cost, battery_cost, towing_cost, cushion_cost, parts_cost
        )
        VALUES (
            @MaintenanceId, @VehiclePlate, @ServiceType, @DateIn, @EstimatedReturn,
            @DateCompleted, @OdometerIn, @OdometerOut, @RepairCost, @WorkshopName,
            @WorkshopAddress, @Status, @Remarks,
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

        Db.ExecuteInTransaction((conn, trans) =>
        {
            if (isFullBatch)
            {
                conn.Execute("DELETE FROM vehicle_maintenance WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
                conn.Execute("DELETE FROM vehicle_master WHERE vehicle_plate != ALL(@ActivePlates);", new { ActivePlates = activePlates }, trans);
            }

            conn.Execute(sqlMasterUpsert, masterRecords, trans);

            if (maintenanceRecords.Any())
            {
                conn.Execute(sqlMaintenanceUpsert, maintenanceRecords, trans);
            }
        });

#if DEBUG
        var sb = new StringBuilder();
        sb.AppendLine("=== POSTGRESQL VEHICLE UPLOAD COMMITTED ===");
        sb.AppendLine($"Committed Vehicle Profiles: {validVehicles.Count}");
        sb.AppendLine($"Timestamp: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        DebugDialog.ShowLog("[DEBUG] PostgreSQL Upload - Vehicle Summary", sb.ToString());
#endif
    }

    public static void SaveAuxiliaryTables(
        IEnumerable<Workshop> workshops,
        IEnumerable<InsuranceCompanyRecord> insuranceCompanies,
        IEnumerable<FuelCardRecord> fuelCards,
        IEnumerable<string> fuelCompanies)
    {
        var compList = fuelCompanies?
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList() ?? new List<string>();

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

            // Fuel Companies Lookup - Sync
            if (compList.Count > 0)
            {
                conn.Execute("DELETE FROM lookup_fuel_companies WHERE fuel_company_name != ALL(@Names);", new { Names = compList.ToArray() }, trans);
                var compPayload = compList.Select(n => new { Name = n });
                conn.Execute(@"
                    INSERT INTO lookup_fuel_companies (fuel_company_name) 
                    VALUES (@Name) 
                    ON CONFLICT (fuel_company_name) DO NOTHING;", compPayload, trans);
            }
            else
            {
                conn.Execute("DELETE FROM lookup_fuel_companies;", transaction: trans);
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
        var names = types.Where(t => !string.IsNullOrWhiteSpace(t.SkillLicenseName)).Select(t => t.SkillLicenseName.Trim()).Distinct().ToList();
        Db.ExecuteInTransaction((conn, trans) =>
        {
            conn.Execute("DELETE FROM lookup_skill_license_types WHERE skill_license_type != ALL(@Names);", new { Names = names.ToArray() }, trans);
            conn.Execute(@"
                INSERT INTO lookup_skill_license_types (skill_license_type, description)
                VALUES (@SkillLicenseType, '')
                ON CONFLICT (skill_license_type) DO NOTHING;", names.Select(n => new { SkillLicenseType = n }), trans);
        });
    }

    // =========================================================================
    // 4. PROJECT & COMPANY PERSISTENCE
    // =========================================================================
    public static void PersistProjects(IList<ProjectRecord> projects, bool isFullBatch)
    {
        var items = projects.Where(p => !string.IsNullOrWhiteSpace(p.ProjectID)).ToList();
        if (!items.Any()) return;

        var activeIds = items.Select(p => p.ProjectID.Trim()).ToArray();
        var records = items.Select(p => new
        {
            ProjectId = p.ProjectID.Trim(),
            ProjectName = p.ProjectDetails.ProjectName ?? string.Empty,
            ProjectType = p.ProjectDetails.ProjectType.ToString(),
            CompanyName = p.ProjectDetails.CompanyName ?? string.Empty,
            StartDate = p.ProjectDetails.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = p.ProjectDetails.TargetEndDate.ToDateTime(TimeOnly.MinValue)
        }).ToList();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            if (isFullBatch)
            {
                conn.Execute("DELETE FROM project_master WHERE project_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
            }

            conn.Execute(@"
                INSERT INTO project_master (project_id, project_name, project_type, company_name, start_date, target_end_date)
                VALUES (@ProjectId, @ProjectName, @ProjectType, @CompanyName, @StartDate, @TargetEndDate)
                ON CONFLICT (project_id) DO UPDATE SET
                    project_name    = EXCLUDED.project_name,
                    project_type    = CASE WHEN EXCLUDED.project_type != 'None' THEN EXCLUDED.project_type ELSE project_master.project_type END,
                    company_name    = EXCLUDED.company_name,
                    start_date      = EXCLUDED.start_date,
                    target_end_date = EXCLUDED.target_end_date,
                    updated_at      = CURRENT_TIMESTAMP;", records, trans);
        });
    }

    public static void PersistCompanies(IList<CompanyRecord> companies, bool isFullBatch)
    {
        var items = companies.Where(c => c.CompanyID > 0).ToList();
        if (!items.Any()) return;

        var activeIds = items.Select(c => c.CompanyID).ToArray();
        var records = items.Select(c => new
        {
            CompanyId = c.CompanyID,
            LegalCompanyName = c.LegalCompanyName ?? string.Empty,
            ShortCompanyName = c.ShortCompanyName ?? string.Empty
        }).ToList();

        Db.ExecuteInTransaction((conn, trans) =>
        {
            if (isFullBatch)
            {
                conn.Execute("DELETE FROM company_master WHERE company_id != ALL(@ActiveIds);", new { ActiveIds = activeIds }, trans);
            }

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