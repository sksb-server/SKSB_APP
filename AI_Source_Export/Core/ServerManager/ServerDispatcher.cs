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
        // 1. Worker Manager
        if (data is IEnumerable<WorkerRecord> workerList)
        {
            var list = workerList.ToList();
            SaveWorkerToDb(list, isFullBatch: true);

            var allVehicleLicenses = list.SelectMany(s => s.VehicleLicenses.Select(lic =>
            {
                if (lic.WorkerID == 0) lic.WorkerID = s.WorkerID;
                return lic;
            })).ToList();
            SaveVehicleLicensesToDb(allVehicleLicenses);

            var allSkillLicenses = list.SelectMany(s => s.SkillLicenses.Select(skill =>
            {
                if (skill.WorkerID == 0) skill.WorkerID = s.WorkerID;
                return skill;
            })).ToList();
            SaveSkillLicensesToDb(allSkillLicenses);
        }
        else if (data is WorkerRecord singleWorker)
        {
            SaveWorkerToDb(new[] { singleWorker }, isFullBatch: false);
            if (singleWorker.VehicleLicenses.Any())
            {
                singleWorker.VehicleLicenses.ForEach(lic => { if (lic.WorkerID == 0) lic.WorkerID = singleWorker.WorkerID; });
                SaveVehicleLicensesToDb(singleWorker.VehicleLicenses);
            }
            if (singleWorker.SkillLicenses.Any())
            {
                singleWorker.SkillLicenses.ForEach(skill => { if (skill.WorkerID == 0) skill.WorkerID = singleWorker.WorkerID; });
                SaveSkillLicensesToDb(singleWorker.SkillLicenses);
            }
        }

        // 2. Project Manager
        else if (data is IEnumerable<ProjectRecord> projectList)
        {
            SaveProjectsToDb(projectList, isFullBatch: true);
        }
        else if (data is ProjectRecord singleProject)
        {
            SaveProjectsToDb(new[] { singleProject }, isFullBatch: false);
        }

        // 3. Company Records
        else if (data is IEnumerable<CompanyRecord> companyList)
        {
            SaveCompaniesToDb(companyList, isFullBatch: true);
        }
        else if (data is CompanyRecord singleCompany)
        {
            SaveCompaniesToDb(new[] { singleCompany }, isFullBatch: false);
        }

        // 4. Vehicle Manager
        else if (data is IEnumerable<VehicleRecord> vehicleList)
        {
            var list = vehicleList.ToList();
            SaveVehiclesToDb(list, isFullBatch: true);

            var allMaintenance = list.SelectMany(v => v.Maintenances.Select(m =>
            {
                if (string.IsNullOrWhiteSpace(m.VehiclePlate))
                    m.VehiclePlate = v.VehiclePlate;
                return m;
            })).ToList();
            SaveVehicleMaintenanceToDb(allMaintenance);
        }
        else if (data is VehicleRecord singleVehicle)
        {
            SaveVehiclesToDb(new[] { singleVehicle }, isFullBatch: false);
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
        else if (data is IEnumerable<SkillLicenseType> skillTypes)
        {
            SaveSkillLicenseTypesToDb(skillTypes);
        }
        else if (data is SkillLicenseType singleSkillType)
        {
            SaveSkillLicenseTypesToDb(new[] { singleSkillType });
        }
    }

    #region Worker Database Operations

    private static void SaveWorkerToDb(IEnumerable<WorkerRecord> workerList, bool isFullBatch)
    {
        var workerItems = workerList.Where(w => w.WorkerID > 0).ToList();
        if (!workerItems.Any()) return;

        var activeIds = workerItems.Select(w => w.WorkerID).ToArray();

        var masterRecords = workerItems.Select(w => new
        {
            WorkerId = w.WorkerID,
            WorkerName = w.WorkerDetails.WorkerName ?? string.Empty,
            Nric = w.WorkerDetails.NRIC,
            Passport = w.WorkerDetails.Passport ?? string.Empty,
            Phone = w.WorkerDetails.Phone ?? string.Empty,
            JoinDate = w.WorkerDetails.JoinDate.ToDateTime(TimeOnly.MinValue),
            ResignDate = w.WorkerDetails.ResignDate.HasValue
                ? (DateTime?)w.WorkerDetails.ResignDate.Value.ToDateTime(TimeOnly.MinValue)
                : null,
            ContractStatus = w.WorkerDetails.ContractStatus.ToString()
        }).ToList();

        const string sqlReconcileVehicleLicenses = "DELETE FROM worker_licenses WHERE worker_id != ALL(@ActiveIds);";
        const string sqlReconcileSkillLicenses = "DELETE FROM worker_skill_licenses WHERE worker_id != ALL(@ActiveIds);";
        const string sqlReconcileMaster = "DELETE FROM worker_master WHERE worker_id != ALL(@ActiveIds);";

        const string sqlMasterUpsert = @"
    INSERT INTO worker_master (
        worker_id, worker_name, nric, passport, phone, join_date, resign_date, contract_status
    )
    VALUES (
        @WorkerId, @WorkerName, @Nric, @Passport, @Phone, @JoinDate, @ResignDate, @ContractStatus
    )
    ON CONFLICT (worker_id) DO UPDATE SET
        worker_name     = EXCLUDED.worker_name,
        nric            = EXCLUDED.nric,
        passport        = EXCLUDED.passport,
        phone           = EXCLUDED.phone,
        join_date       = EXCLUDED.join_date,
        resign_date     = EXCLUDED.resign_date,
        contract_status = EXCLUDED.contract_status,
        updated_at      = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            if (isFullBatch)
            {
                db.Execute(sqlReconcileVehicleLicenses, new { ActiveIds = activeIds }, transaction: trans);
                db.Execute(sqlReconcileSkillLicenses, new { ActiveIds = activeIds }, transaction: trans);
                db.Execute(sqlReconcileMaster, new { ActiveIds = activeIds }, transaction: trans);
            }

            db.Execute(sqlMasterUpsert, masterRecords, transaction: trans);
        });
    }

    private static void SaveVehicleLicensesToDb(IEnumerable<VehicleLicense> licenses)
    {
        var licenseItems = licenses
            .Where(l => l.WorkerID > 0 && l.LicenseClass != LicenseClass.None)
            .Select(l => new
            {
                WorkerId = l.WorkerID,
                LicenseType = l.LicenseClass.ToString(),
                RenewDate = l.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Status = l.LicenseStatus.ToString()
            }).ToList();

        if (!licenseItems.Any()) return;

        const string sqlLicenses = @"
        INSERT INTO worker_licenses (worker_id, license_type, renew_date, expiry_date, status)
        VALUES (@WorkerId, @LicenseType, @RenewDate, @ExpiryDate, @Status)
        ON CONFLICT (worker_id, license_type) DO UPDATE SET
            renew_date  = EXCLUDED.renew_date,
            expiry_date = EXCLUDED.expiry_date,
            status      = EXCLUDED.status,
            updated_at  = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlLicenses, licenseItems);
    }

    private static void SaveSkillLicensesToDb(IEnumerable<SkillLicense> skillLicenses)
    {
        var skillItems = skillLicenses
            .Where(s => s.WorkerID > 0 && !string.IsNullOrWhiteSpace(s.LicenseType))
            .Select(s => new
            {
                WorkerId = s.WorkerID,
                LicenseType = s.LicenseType.Trim(),
                RenewDate = s.RenewDate.ToDateTime(TimeOnly.MinValue),
                ExpiryDate = s.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Status = s.LicenseStatus.ToString()
            }).ToList();

        if (!skillItems.Any()) return;

        const string sqlSkillLicense = @"
        INSERT INTO worker_skill_licenses (worker_id, license_type, renew_date, expiry_date, status)
        VALUES (@WorkerId, @LicenseType, @RenewDate, @ExpiryDate, @Status)
        ON CONFLICT (worker_id, license_type) DO UPDATE SET
            renew_date  = EXCLUDED.renew_date,
            expiry_date = EXCLUDED.expiry_date,
            status      = EXCLUDED.status,
            updated_at  = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlSkillLicense, skillItems);
    }

    private static void SaveSkillLicenseTypesToDb(IEnumerable<SkillLicenseType> skillTypes)
    {
        var payload = skillTypes
            .Where(t => !string.IsNullOrWhiteSpace(t.LicenseType))
            .Select(t => new
            {
                LicenseType = t.LicenseType.Trim(),
                Description = t.Description ?? string.Empty
            }).ToList();

        if (!payload.Any()) return;

        const string sql = @"
            INSERT INTO lookup_skill_license_types (license_type, description)
            VALUES (@LicenseType, @Description)
            ON CONFLICT (license_type) DO UPDATE 
            SET description = EXCLUDED.description;";

        DbService.ExecuteTransaction(sql, payload);
    }

    #endregion

    #region Project Database Operations

    private static void SaveProjectsToDb(IEnumerable<ProjectRecord> projectList, bool isFullBatch)
    {
        var projectItems = projectList.Where(p => !string.IsNullOrWhiteSpace(p.ProjectID)).ToList();
        if (!projectItems.Any()) return;

        var activeIds = projectItems.Select(p => p.ProjectID).ToArray();

        // Extracts strictly the 6 columns displayed on your template
        var projectRecords = projectItems.Select(p => new
        {
            ProjectId = p.ProjectID.Trim(),
            ProjectName = p.ProjectDetails.ProjectName ?? string.Empty,
            ProjectType = p.ProjectDetails.ProjectType.ToString(),
            CompanyName = p.ProjectDetails.CompanyName ?? string.Empty,
            StartDate = p.ProjectDetails.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = p.ProjectDetails.TargetEndDate.ToDateTime(TimeOnly.MinValue)
        }).ToList();

        const string sqlReconcileProjects = "DELETE FROM project_master WHERE project_id != ALL(@ActiveIds);";

        // Preserves address, district, remarks, actual_end_date, and costs if already present
        const string sqlProjectUpsert = @"
        INSERT INTO project_master (
            project_id, project_name, project_type, company_name, start_date, target_end_date
        )
        VALUES (
            @ProjectId, @ProjectName, @ProjectType, @CompanyName, @StartDate, @TargetEndDate
        )
        ON CONFLICT (project_id) DO UPDATE SET
            project_name    = EXCLUDED.project_name,
            project_type    = CASE WHEN EXCLUDED.project_type != 'None' THEN EXCLUDED.project_type ELSE project_master.project_type END,
            company_name    = EXCLUDED.company_name,
            start_date      = EXCLUDED.start_date,
            target_end_date = EXCLUDED.target_end_date,
            updated_at      = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            if (isFullBatch)
            {
                db.Execute(sqlReconcileProjects, new { ActiveIds = activeIds }, transaction: trans);
            }
            db.Execute(sqlProjectUpsert, projectRecords, transaction: trans);
        });
    }

    #endregion

    #region Company Database Operations

    private static void SaveCompaniesToDb(IEnumerable<CompanyRecord> companyList, bool isFullBatch)
    {
        var companyItems = companyList.Where(c => c.CompanyID > 0).ToList();
        if (!companyItems.Any()) return;

        var activeIds = companyItems.Select(c => c.CompanyID).ToArray();

        // Extracts strictly CompanyID, LegalCompanyName, and ShortCompanyName
        var companyRecords = companyItems.Select(c => new
        {
            CompanyId = c.CompanyID,
            LegalCompanyName = c.LegalCompanyName ?? string.Empty,
            ShortCompanyName = c.ShortCompanyName ?? string.Empty
        }).ToList();

        const string sqlReconcileCompanies = "DELETE FROM company_master WHERE company_id != ALL(@ActiveIds);";

        // Preserves SSM, VendorCode, SPAN, and CIDB records if already populated
        const string sqlCompanyUpsert = @"
        INSERT INTO company_master (
            company_id, legal_company_name, short_company_name
        )
        VALUES (
            @CompanyId, @LegalCompanyName, @ShortCompanyName
        )
        ON CONFLICT (company_id) DO UPDATE SET
            legal_company_name = EXCLUDED.legal_company_name,
            short_company_name = EXCLUDED.short_company_name,
            updated_at         = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            if (isFullBatch)
            {
                db.Execute(sqlReconcileCompanies, new { ActiveIds = activeIds }, transaction: trans);
            }
            db.Execute(sqlCompanyUpsert, companyRecords, transaction: trans);
        });
    }

    #endregion

    #region Vehicle Database Operations

    private static void SaveVehiclesToDb(IEnumerable<VehicleRecord> vehicleList, bool isFullBatch)
    {
        var vehicleItems = vehicleList.Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)).ToList();
        if (!vehicleItems.Any()) return;

        var activePlates = vehicleItems.Select(v => v.VehiclePlate.Trim().ToUpperInvariant()).ToArray();

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

        const string sqlReconcileMaintenance = "DELETE FROM vehicle_maintenance WHERE vehicle_plate != ALL(@ActivePlates);";
        const string sqlReconcileVehicles = "DELETE FROM vehicle_master WHERE vehicle_plate != ALL(@ActivePlates);";

        // Preserves expiry dates if the user only supplied plate and vehicle type
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
            current_odometer  = CASE WHEN EXCLUDED.current_odometer > 0 THEN EXCLUDED.current_odometer ELSE vehicle_master.current_odometer END,
            status            = EXCLUDED.status,
            updated_at        = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            if (isFullBatch)
            {
                db.Execute(sqlReconcileMaintenance, new { ActivePlates = activePlates }, transaction: trans);
                db.Execute(sqlReconcileVehicles, new { ActivePlates = activePlates }, transaction: trans);
            }
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