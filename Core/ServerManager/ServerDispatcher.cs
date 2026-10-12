namespace SKSB_App.Core.ServerManager;

using Dapper;
using SKSB_App.Core.Types;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ServerDispatcher
{
    private static readonly PostgreSqlService DbService = new();

    // -------------------------------------------------------------
    // Main Routing Dispatcher:
    // -------------------------------------------------------------
    public static void Dispatch<T>(T data)
    {
        if (data is IEnumerable<StaffRecord> staffList)
        {
            var list = staffList.ToList();
            SaveStaffToDb(list);

            var allVehicleLicenses = list.SelectMany(s => s.VehicleLicenses.Select(lic =>
            {
                if (lic.StaffID == 0) lic.StaffID = s.StaffID;
                return lic;
            })).ToList();
            SaveVehicleLicensesToDb(allVehicleLicenses);

            var allSkillLicenses = list.SelectMany(s => s.SkillLicenses.Select(skill =>
            {
                if (skill.StaffID == 0) skill.StaffID = s.StaffID;
                return skill;
            })).ToList();
            SaveSkillLicensesToDb(allSkillLicenses);
        }
        else if (data is StaffRecord singleStaff)
        {
            SaveStaffToDb(new[] { singleStaff });
            if (singleStaff.VehicleLicenses.Any())
            {
                singleStaff.VehicleLicenses.ForEach(lic => { if (lic.StaffID == 0) lic.StaffID = singleStaff.StaffID; });
                SaveVehicleLicensesToDb(singleStaff.VehicleLicenses);
            }
            if (singleStaff.SkillLicenses.Any())
            {
                singleStaff.SkillLicenses.ForEach(skill => { if (skill.StaffID == 0) skill.StaffID = singleStaff.StaffID; });
                SaveSkillLicensesToDb(singleStaff.SkillLicenses);
            }
        }
        else if (data is IEnumerable<ProjectRecord> projectList)
        {
            SaveProjectsToDb(projectList);
        }
        else if (data is ProjectRecord singleProject)
        {
            SaveProjectsToDb(new[] { singleProject });
        }
        else if (data is IEnumerable<VehicleRecord> vehicleList)
        {
            var list = vehicleList.ToList();
            SaveVehiclesToDb(list);

            var allMaintenance = list.SelectMany(v => v.Maintenances.Select(m =>
            {
                // Inject parent plate if missing
                if (string.IsNullOrWhiteSpace(m.VehiclePlate))
                    m.VehiclePlate = v.VehiclePlate;
                return m;
            })).ToList();
            SaveVehicleMaintenanceToDb(allMaintenance);
        }
        else if (data is VehicleRecord singleVehicle)
        {
            SaveVehiclesToDb(new[] { singleVehicle });
            if (singleVehicle.Maintenances.Any())
            {
                SaveVehicleMaintenanceToDb(singleVehicle.Maintenances);
            }
        }
        else if (data is IEnumerable<VehicleLicense> vehicleLicenses)
        {
            SaveVehicleLicensesToDb(vehicleLicenses);
        }
        else if (data is IEnumerable<SkillLicense> skillLicenses)
        {
            SaveSkillLicensesToDb(skillLicenses);
        }
    }

    #region Staff Database Operations

    private static void SaveStaffToDb(IEnumerable<StaffRecord> staffList)
    {
        var staffItems = staffList.Where(s => s.StaffID > 0).ToList();
        if (!staffItems.Any()) return;

        var activeIds = staffItems.Select(s => s.StaffID).ToArray();

        var masterRecords = staffItems.Select(s => new
        {
            StaffId = s.StaffID,
            FamilyName = s.StaffDetails.FamilyName,
            MiddleName = s.StaffDetails.MiddleName,
            GivenName = s.StaffDetails.GivenName,
            DisplayName = s.StaffDetails.Name,
            Nric = s.StaffDetails.NRIC,
            Address = s.StaffDetails.Address,
            Email = s.StaffDetails.Email,
            Phone = s.StaffDetails.Phone,
            Department = s.StaffDetails.Department.ToString(),
            PayType = s.StaffDetails.PayType.ToString(),
            BaseRate = s.StaffDetails.BaseRate,
            ContractStatus = s.StaffDetails.ContractStatus.ToString(),
            BankName = s.StaffDetails.BankName,
            BankNumber = s.StaffDetails.BankNumber,
            EpfNumber = s.StaffDetails.EPFNumber,
            SocsoNumber = s.StaffDetails.SocsoNumber,
            TinNumber = s.StaffDetails.TINNumber,
            JoinDate = s.StaffDetails.JoinDate.ToDateTime(TimeOnly.MinValue),
            ConfirmDate = s.StaffDetails.ConfirmDate.HasValue
                ? (DateTime?)s.StaffDetails.ConfirmDate.Value.ToDateTime(TimeOnly.MinValue)
                : null,
            ResignDate = s.StaffDetails.ResignDate.HasValue
                ? (DateTime?)s.StaffDetails.ResignDate.Value.ToDateTime(TimeOnly.MinValue)
                : null
        }).ToList();

        const string sqlReconcileVehicleLicenses = "DELETE FROM staff_licenses WHERE staff_id != ALL(@ActiveIds);";
        const string sqlReconcileSkillLicenses = "DELETE FROM staff_skill_licenses WHERE staff_id != ALL(@ActiveIds);";
        const string sqlReconcileMaster = "DELETE FROM staff_master WHERE staff_id != ALL(@ActiveIds);";

        const string sqlMasterUpsert = @"
        INSERT INTO staff_master (
            staff_id, family_name, middle_name, given_name, display_name,
            nric, address, email, phone, department, pay_type, base_rate,
            contract_status, bank_name, bank_number, epf_number, socso_number,
            tin_number, join_date, confirm_date, resign_date
        )
        VALUES (
            @StaffId, @FamilyName, @MiddleName, @GivenName, @DisplayName,
            @Nric, @Address, @Email, @Phone, @Department, @PayType, @BaseRate,
            @ContractStatus, @BankName, @BankNumber, @EpfNumber, @SocsoNumber,
            @TinNumber, @JoinDate, @ConfirmDate, @ResignDate
        )
        ON CONFLICT (staff_id) DO UPDATE SET
            family_name     = EXCLUDED.family_name,
            middle_name     = EXCLUDED.middle_name,
            given_name      = EXCLUDED.given_name,
            display_name    = EXCLUDED.display_name,
            nric            = EXCLUDED.nric,
            address         = EXCLUDED.address,
            email           = EXCLUDED.email,
            phone           = EXCLUDED.phone,
            department      = EXCLUDED.department,
            pay_type        = EXCLUDED.pay_type,
            base_rate       = EXCLUDED.base_rate,
            contract_status = EXCLUDED.contract_status,
            bank_name       = EXCLUDED.bank_name,
            bank_number     = EXCLUDED.bank_number,
            epf_number      = EXCLUDED.epf_number,
            socso_number    = EXCLUDED.socso_number,
            tin_number      = EXCLUDED.tin_number,
            join_date       = EXCLUDED.join_date,
            confirm_date    = EXCLUDED.confirm_date,
            resign_date     = EXCLUDED.resign_date,
            updated_at      = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileVehicleLicenses, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcileSkillLicenses, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcileMaster, new { ActiveIds = activeIds }, transaction: trans);

            db.Execute(sqlMasterUpsert, masterRecords, transaction: trans);
        });
    }

    private static void SaveVehicleLicensesToDb(IEnumerable<VehicleLicense> licenses)
    {
        var licenseItems = licenses
            .Where(l => l.StaffID > 0 && l.LicenseClass != LicenseClass.None)
            .Select(l => new
            {
                StaffId = l.StaffID,
                LicenseType = l.LicenseClass.ToString(),
                RenewDate = l.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Status = l.LicenseStatus.ToString()
            }).ToList();

        if (!licenseItems.Any()) return;

        const string sqlLicenses = @"
        INSERT INTO staff_licenses (staff_id, license_type, renew_date, expiry_date, status)
        VALUES (@StaffId, @LicenseType, @RenewDate, @ExpiryDate, @Status)
        ON CONFLICT (staff_id, license_type) DO UPDATE SET
            renew_date  = EXCLUDED.renew_date,
            expiry_date = EXCLUDED.expiry_date,
            status      = EXCLUDED.status,
            updated_at  = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlLicenses, licenseItems);
    }

    private static void SaveSkillLicensesToDb(IEnumerable<SkillLicense> skillLicenses)
    {
        var skillItems = skillLicenses
            .Where(s => s.StaffID > 0 && !string.IsNullOrWhiteSpace(s.LicenseType))
            .Select(s => new
            {
                LicenseId = s.LicenseID,
                StaffId = s.StaffID,
                LicenseType = s.LicenseType,
                RenewDate = s.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = s.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Status = s.LicenseStatus.ToString()
            }).ToList();

        if (!skillItems.Any()) return;

        const string sqlSkillLicense = @"
        INSERT INTO staff_skill_licenses (
            license_id, staff_id, license_type, renew_date, expiry_date, status
        )
        VALUES (
            @LicenseId, @StaffId, @LicenseType, @RenewDate, @ExpiryDate, @Status
        )
        ON CONFLICT (staff_id, license_type) DO UPDATE SET
            renew_date  = EXCLUDED.renew_date,
            expiry_date = EXCLUDED.expiry_date,
            status      = EXCLUDED.status,
            updated_at  = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlSkillLicense, skillItems);
    }

    #endregion

    #region Project Database Operations

    private static void SaveProjectsToDb(IEnumerable<ProjectRecord> projectList)
    {
        var projectItems = projectList.Where(p => !string.IsNullOrWhiteSpace(p.ProjectID)).ToList();
        if (!projectItems.Any()) return;

        var activeIds = projectItems.Select(p => p.ProjectID).ToArray();

        var projectRecords = projectItems.Select(p => new
        {
            ProjectId = p.ProjectID,
            ProjectName = p.ProjectDetails.ProjectName,
            ProjectType = p.ProjectDetails.ProjectType.ToString(),
            CompanyId = p.ProjectDetails.CompanyID,
            CompanyName = p.ProjectDetails.CompanyName,
            Priority = p.ProjectDetails.Priority.ToString(),
            Status = p.ProjectDetails.Status.ToString(),
            District = p.ProjectDetails.District.ToString(),
            Address = p.ProjectDetails.Address,
            Remarks = p.ProjectDetails.Remarks,
            StartDate = p.ProjectDetails.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = p.ProjectDetails.TargetEndDate.ToDateTime(TimeOnly.MinValue),
            ActualEndDate = p.ProjectDetails.ActualEndDate.HasValue
                ? (DateTime?)p.ProjectDetails.ActualEndDate.Value.ToDateTime(TimeOnly.MinValue)
                : null,
            EstProjectCost = p.ProjectDetails.EstProjectCost,
            ProjectTotalCost = p.ProjectDetails.ProjectTotalCost,
            PromisedReturn = p.ProjectDetails.PromisedReturn,
            ActualReturn = p.ProjectDetails.ActualReturn,
            Profit = p.ProjectDetails.Profit
        }).ToList();

        const string sqlReconcileProjects = "DELETE FROM project_master WHERE project_id != ALL(@ActiveIds);";

        const string sqlProjectUpsert = @"
        INSERT INTO project_master (
            project_id, project_name, project_type, company_id, company_name,
            priority, status, district, address, remarks, start_date,
            target_end_date, actual_end_date, est_project_cost, project_total_cost,
            promised_return, actual_return, profit
        )
        VALUES (
            @ProjectId, @ProjectName, @ProjectType, @CompanyId, @CompanyName,
            @Priority, @Status, @District, @Address, @Remarks, @StartDate,
            @TargetEndDate, @ActualEndDate, @EstProjectCost, @ProjectTotalCost,
            @PromisedReturn, @ActualReturn, @Profit
        )
        ON CONFLICT (project_id) DO UPDATE SET
            project_name       = EXCLUDED.project_name,
            project_type       = EXCLUDED.project_type,
            company_id         = EXCLUDED.company_id,
            company_name       = EXCLUDED.company_name,
            priority           = EXCLUDED.priority,
            status             = EXCLUDED.status,
            district           = EXCLUDED.district,
            address            = EXCLUDED.address,
            remarks            = EXCLUDED.remarks,
            start_date         = EXCLUDED.start_date,
            target_end_date    = EXCLUDED.target_end_date,
            actual_end_date    = EXCLUDED.actual_end_date,
            est_project_cost   = EXCLUDED.est_project_cost,
            project_total_cost = EXCLUDED.project_total_cost,
            promised_return    = EXCLUDED.promised_return,
            actual_return      = EXCLUDED.actual_return,
            profit             = EXCLUDED.profit,
            updated_at         = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileProjects, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlProjectUpsert, projectRecords, transaction: trans);
        });
    }

    #endregion

    #region Vehicle Database Operations

    private static void SaveVehiclesToDb(IEnumerable<VehicleRecord> vehicleList)
    {
        var vehicleItems = vehicleList.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)).ToList();
        if (!vehicleItems.Any()) return;

        var activePlates = vehicleItems.Select(v => v.VehiclePlate).ToArray();

        var masterRecords = vehicleItems.Select(v => new
        {
            VehiclePlate = v.VehiclePlate.Trim().ToUpperInvariant(),
            Vin = v.VehicleDetail.VIN ?? string.Empty,
            VehicleType = v.VehicleDetail.Type.ToString(),
            Brand = v.VehicleDetail.Brand ?? string.Empty,
            Model = v.VehicleDetail.Model ?? string.Empty,
            RoadTaxExpiry = v.VehicleDetail.RoadTaxExpiry.ToDateTime(TimeOnly.MinValue),
            InspectionExpiry = v.VehicleDetail.InspectionExpiry.ToDateTime(TimeOnly.MinValue),
            CurrentOdometer = v.VehicleDetail.CurrentOdometer,
            Status = v.VehicleStatus.ToString()
        }).ToList();

        const string sqlReconcileVehicles = @"
        DELETE FROM vehicle_master 
        WHERE vehicle_plate != ALL(@ActivePlates);";

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
            vin               = EXCLUDED.vin,
            vehicle_type      = EXCLUDED.vehicle_type,
            brand             = EXCLUDED.brand,
            model             = EXCLUDED.model,
            road_tax_expiry   = EXCLUDED.road_tax_expiry,
            inspection_expiry = EXCLUDED.inspection_expiry,
            current_odometer  = EXCLUDED.current_odometer,
            status            = EXCLUDED.status,
            updated_at        = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileVehicles, new { ActivePlates = activePlates }, transaction: trans);
            db.Execute(sqlMasterUpsert, masterRecords, transaction: trans);
        });
    }

    private static void SaveVehicleMaintenanceToDb(IEnumerable<Maintenance> records)
    {
        var validRecords = records
            .Where(m => !string.IsNullOrWhiteSpace(m.VehiclePlate))
            .Select(m => new
            {
                MaintenanceId = m.MaintenanceID,
                VehiclePlate = m.VehiclePlate.Trim().ToUpperInvariant(),
                ServiceType = m.ServiceType.ToString(),
                DateIn = m.DateIn.ToDateTime(TimeOnly.MinValue),
                EstimatedReturn = m.EstimatedReturn.ToDateTime(TimeOnly.MinValue),
                DateCompleted = m.DateCompleted.HasValue
                    ? (DateTime?)m.DateCompleted.Value.ToDateTime(TimeOnly.MinValue)
                    : null,
                OdometerIn = m.OdometerIn,
                OdometerOut = m.OdometerOut,
                RepairCost = m.RepairCost,
                WorkshopName = m.WorkshopName,
                WorkshopAddress = m.WorkshopAddress,
                Status = m.MaintenanceStatus.ToString(),
                Remarks = m.Remarks
            }).ToList();

        if (!validRecords.Any()) return;

        const string sqlMaintenanceUpsert = @"
        INSERT INTO vehicle_maintenance (
            maintenance_id, vehicle_plate, service_type, date_in, estimated_return,
            date_completed, odometer_in, odometer_out, repair_cost, workshop_name,
            workshop_address, status, remarks
        )
        VALUES (
            @MaintenanceId, @VehiclePlate, @ServiceType, @DateIn, @EstimatedReturn,
            @DateCompleted, @OdometerIn, @OdometerOut, @RepairCost, @WorkshopName,
            @WorkshopAddress, @Status, @Remarks
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
            updated_at       = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlMaintenanceUpsert, validRecords);
    }

    #endregion
}