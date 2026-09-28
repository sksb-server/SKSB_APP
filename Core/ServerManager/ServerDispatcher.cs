namespace SKSB_App.Core.ServerManager;

using Dapper;
using SKSB_App.Core.Types;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ServerDispatcher
{
    private static readonly PostgreSqlService DbService = new();

    public static void Dispatch<T>(T data)
    {
        if (data is IEnumerable<Staff> staffList)
        {
            var list = staffList.ToList();
            SaveStaffToDb(list);

            var allLicenses = list.SelectMany(s => s.Licenses.Select(lic =>
            {
                if (lic.StaffID == 0) lic.StaffID = s.StaffID;
                return lic;
            })).ToList();
            SaveLicensesToDb(allLicenses);

            var allLeaves = list.SelectMany(s => s.Leaves.Select(l =>
            {
                if (l.StaffID == 0) l.StaffID = s.StaffID;
                return l;
            })).ToList();
            SaveLeavesToDb(allLeaves);
        }
        else if (data is Staff singleStaff)
        {
            SaveStaffToDb(new[] { singleStaff });
            if (singleStaff.Licenses.Any())
            {
                singleStaff.Licenses.ForEach(lic => { if (lic.StaffID == 0) lic.StaffID = singleStaff.StaffID; });
                SaveLicensesToDb(singleStaff.Licenses);
            }
            if (singleStaff.Leaves.Any())
            {
                singleStaff.Leaves.ForEach(l => { if (l.StaffID == 0) l.StaffID = singleStaff.StaffID; });
                SaveLeavesToDb(singleStaff.Leaves);
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
        else if (data is IEnumerable<Vehicle> vehicleList)
        {
            var list = vehicleList.ToList();
            SaveVehiclesToDb(list);

            var allMaintenance = list.SelectMany(v => v.MaintenanceLogs.Select(m =>
            {
                if (m.VehicleID == 0) m.VehicleID = v.VehicleID;
                return m;
            })).ToList();
            SaveVehicleMaintenanceToDb(allMaintenance);
        }
        else if (data is Vehicle singleVehicle)
        {
            SaveVehiclesToDb(new[] { singleVehicle });
            if (singleVehicle.MaintenanceLogs.Any())
            {
                singleVehicle.MaintenanceLogs.ForEach(m => { if (m.VehicleID == 0) m.VehicleID = singleVehicle.VehicleID; });
                SaveVehicleMaintenanceToDb(singleVehicle.MaintenanceLogs);
            }
        }
        else if (data is IEnumerable<License> licenses)
        {
            SaveLicensesToDb(licenses);
        }
        else if (data is IEnumerable<Leave> leaves)
        {
            SaveLeavesToDb(leaves);
        }
    }

    #region Staff Database Operations

    private static void SaveStaffToDb(IEnumerable<Staff> staffList)
    {
        var staffItems = staffList.Where(s => s.StaffID > 0).ToList();
        if (!staffItems.Any()) return;

        var activeIds = staffItems.Select(s => s.StaffID).ToArray();

        var masterRecords = staffItems.Select(s => new
        {
            StaffId = s.StaffID,
            FamilyName = s.Master.FamilyName,
            MiddleName = s.Master.MiddleName,
            GivenName = s.Master.GivenName,
            Name = s.Master.Name,
            FullName = s.Master.Name,
            Nric = s.Master.NRIC,
            Department = s.Master.Department.ToString(),
            PayType = s.Master.PayType.ToString(),
            BaseRate = s.Master.BaseRate,
            OvertimeRate = s.Master.OvertimeRate,
            ActiveStatus = s.Master.ActiveStatus.ToString()
        }).ToList();

        var prefRecords = staffItems.Select(s => new
        {
            StaffId = s.StaffID,
            PrimaryTeam = s.Preference.PrimaryTeam,
            SecondaryTeam = s.Preference.SecondaryTeam,
            PrimarySkill = s.Preference.PrimarySkill,
            SecondarySkill = s.Preference.SecondarySkill,
            BlacklistTask = s.Preference.BlackListTask,
            PreferredShift = s.Preference.PreferredShift.ToString(),
            MaxDaysPerWeek = s.Preference.MaxDaysPerWeek,
            PreferredRestDay = s.Preference.PrimaryRestDay.ToString(),
            PreferredVehicleType = s.Preference.PreferredVehicleType.ToString(),
            DedicatedVehicleId = s.Preference.DedicatedVehicleID
        }).ToList();

        const string sqlReconcileLicenses = "DELETE FROM staff_licenses WHERE staff_id != ALL(@ActiveIds);";
        const string sqlReconcileLeaves = "DELETE FROM staff_leave WHERE staff_id != ALL(@ActiveIds);";
        const string sqlReconcilePrefs = "DELETE FROM staff_preference WHERE staff_id != ALL(@ActiveIds);";
        const string sqlReconcileMaster = "DELETE FROM staff_master WHERE staff_id != ALL(@ActiveIds);";

        const string sqlMasterUpsert = @"
        INSERT INTO staff_master (
            staff_id, family_name, middle_name, given_name, display_name, full_name,
            nric, department, pay_type, base_rate, overtime_rate, active_status
        )
        VALUES (
            @StaffId, @FamilyName, @MiddleName, @GivenName, @Name, @FullName,
            @Nric, @Department, @PayType, @BaseRate, @OvertimeRate, @ActiveStatus
        )
        ON CONFLICT (staff_id) DO UPDATE SET
            family_name   = EXCLUDED.family_name,
            middle_name   = EXCLUDED.middle_name,
            given_name    = EXCLUDED.given_name,
            display_name  = EXCLUDED.display_name,
            full_name     = EXCLUDED.full_name,
            nric          = EXCLUDED.nric,
            department    = EXCLUDED.department,
            pay_type      = EXCLUDED.pay_type,
            base_rate     = EXCLUDED.base_rate,
            overtime_rate = EXCLUDED.overtime_rate,
            active_status = EXCLUDED.active_status,
            updated_at    = CURRENT_TIMESTAMP;";

        const string sqlPrefUpsert = @"
        INSERT INTO staff_preference (
            staff_id, primary_team, secondary_team, primary_skill, 
            secondary_skill, blacklist_task, preferred_shift, 
            max_days_per_week, preferred_rest_day, preferred_vehicle_type, dedicated_vehicle_id
        )
        VALUES (
            @StaffId, @PrimaryTeam, @SecondaryTeam, @PrimarySkill, 
            @SecondarySkill, @BlacklistTask, 
            @PreferredShift, 
            @MaxDaysPerWeek, @PreferredRestDay, @PreferredVehicleType, @DedicatedVehicleId
        )
        ON CONFLICT (staff_id) DO UPDATE SET
            primary_team           = EXCLUDED.primary_team,
            secondary_team         = EXCLUDED.secondary_team,
            primary_skill          = EXCLUDED.primary_skill,
            secondary_skill        = EXCLUDED.secondary_skill,
            blacklist_task         = EXCLUDED.blacklist_task,
            preferred_shift        = EXCLUDED.preferred_shift,
            max_days_per_week      = EXCLUDED.max_days_per_week,
            preferred_rest_day     = EXCLUDED.preferred_rest_day,
            preferred_vehicle_type = EXCLUDED.preferred_vehicle_type,
            dedicated_vehicle_id   = EXCLUDED.dedicated_vehicle_id,
            updated_at             = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileLicenses, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcileLeaves, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcilePrefs, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcileMaster, new { ActiveIds = activeIds }, transaction: trans);

            db.Execute(sqlMasterUpsert, masterRecords, transaction: trans);
            db.Execute(sqlPrefUpsert, prefRecords, transaction: trans);
        });
    }

    private static void SaveLicensesToDb(IEnumerable<License> licenses)
    {
        var licenseItems = licenses
            .Where(l => l.StaffID > 0 && l.LicenseType != LicenseClass.None)
            .Select(l => new
            {
                StaffId = l.StaffID,
                LicenseType = l.LicenseType.ToString(),
                ExpiryDate = l.ExpiryDate.ToDateTime(TimeOnly.MinValue),
                Status = l.LicenseStatus.ToString()
            }).ToList();

        if (!licenseItems.Any()) return;

        const string sqlLicenses = @"
        INSERT INTO staff_licenses (staff_id, license_type, expiry_date, status)
        VALUES (@StaffId, @LicenseType, @ExpiryDate, @Status)
        ON CONFLICT (staff_id, license_type) DO UPDATE SET
            expiry_date = EXCLUDED.expiry_date,
            status      = EXCLUDED.status,
            updated_at  = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlLicenses, licenseItems);
    }

    private static void SaveLeavesToDb(IEnumerable<Leave> leaves)
    {
        var leaveItems = leaves.Where(l => l.StaffID > 0).Select(l => new
        {
            LeaveId = l.LeaveID,
            StaffId = l.StaffID,
            LeaveType = l.LeaveType.ToString(),
            StartDate = l.DateStart.ToDateTime(TimeOnly.MinValue),
            EndDate = l.DateEnd.ToDateTime(TimeOnly.MinValue),
            SessionType = l.SessionType.ToString(),
            DaysTaken = l.TotalDays,
            Status = l.ApprovalStatus.ToString()
        }).ToList();

        if (!leaveItems.Any()) return;

        const string sqlLeave = @"
        INSERT INTO staff_leave (
            leave_id, staff_id, leave_type, start_date, end_date, 
            days_taken, status, session_type
        )
        VALUES (
            @LeaveId, @StaffId, @LeaveType, @StartDate, @EndDate, 
            @DaysTaken, @Status, @SessionType
        )
        ON CONFLICT (leave_id) DO UPDATE SET
            staff_id     = EXCLUDED.staff_id,
            leave_type   = EXCLUDED.leave_type,
            start_date   = EXCLUDED.start_date,
            end_date     = EXCLUDED.end_date,
            days_taken   = EXCLUDED.days_taken,
            status       = EXCLUDED.status,
            session_type = EXCLUDED.session_type,
            updated_at   = CURRENT_TIMESTAMP;";

        const string sqlSyncSeq = @"
        SELECT setval(
            pg_get_serial_sequence('staff_leave', 'leave_id'),
            COALESCE(MAX(leave_id), 1)
        ) FROM staff_leave;";

        DbService.ExecuteTransaction(sqlLeave, leaveItems);
        DbService.Execute(sqlSyncSeq);
    }

    #endregion

    #region Project Database Operations

    private static void SaveProjectsToDb(IEnumerable<ProjectRecord> projectList)
    {
        var projectItems = projectList.Where(p => p.ProjectID > 0).ToList();
        if (!projectItems.Any()) return;

        var activeIds = projectItems.Select(p => p.ProjectID).ToArray();

        var projectRecords = projectItems.Select(p => new
        {
            ProjectId = p.ProjectID,
            ProjectCode = string.IsNullOrWhiteSpace(p.Master.ProjectID) ? $"PRJ-{p.ProjectID}" : p.Master.ProjectID,
            ProjectName = p.Master.ProjectName,
            ProjectClass = p.Master.ProjectClass.ToString(),
            Priority = p.Master.Priority.ToString(),
            Status = p.Master.Status.ToString(),
            StartDate = p.Master.StartDate.ToDateTime(TimeOnly.MinValue),
            TargetEndDate = p.Master.TargetEndDate.ToDateTime(TimeOnly.MinValue),
            ActualEndDate = p.Master.ActualEndDate == default
                ? (DateTime?)null
                : p.Master.ActualEndDate.ToDateTime(TimeOnly.MinValue),
            RequiredHeadcount = p.Master.RequiredHeadcount,
            RequiredVehicleType = p.Master.RequiredVehicleType.ToString(),
            RequiredSkills = p.Master.RequiredSkills,
            Address = p.Master.Address,
            Remarks = p.Master.Remarks
        }).ToList();

        const string sqlReconcileProjects = "DELETE FROM project_master WHERE project_id != ALL(@ActiveIds);";

        const string sqlProjectUpsert = @"
        INSERT INTO project_master (
            project_id, project_code, project_name, project_class, priority, status,
            start_date, target_end_date, actual_end_date, required_headcount,
            required_vehicle_type, required_skills, address, remarks
        )
        VALUES (
            @ProjectId, @ProjectCode, @ProjectName, @ProjectClass, @Priority, @Status,
            @StartDate, @TargetEndDate, @ActualEndDate, @RequiredHeadcount,
            @RequiredVehicleType, @RequiredSkills, @Address, @Remarks
        )
        ON CONFLICT (project_id) DO UPDATE SET
            project_code          = EXCLUDED.project_code,
            project_name          = EXCLUDED.project_name,
            project_class         = EXCLUDED.project_class,
            priority              = EXCLUDED.priority,
            status                = EXCLUDED.status,
            start_date            = EXCLUDED.start_date,
            target_end_date       = EXCLUDED.target_end_date,
            actual_end_date       = EXCLUDED.actual_end_date,
            required_headcount    = EXCLUDED.required_headcount,
            required_vehicle_type = EXCLUDED.required_vehicle_type,
            required_skills       = EXCLUDED.required_skills,
            address               = EXCLUDED.address,
            remarks               = EXCLUDED.remarks,
            updated_at            = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileProjects, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlProjectUpsert, projectRecords, transaction: trans);
        });
    }

    #endregion

    #region Vehicle Database Operations

    private static void SaveVehiclesToDb(IEnumerable<Vehicle> vehicleList)
    {
        var vehicleItems = vehicleList.Where(v => v.VehicleID > 0).ToList();
        if (!vehicleItems.Any()) return;

        var activeIds = vehicleItems.Select(v => v.VehicleID).ToArray();

        var masterRecords = vehicleItems.Select(v => new
        {
            VehicleId = v.VehicleID,
            RegistrationNumber = v.Master.RegistrationNumber.Trim().ToUpperInvariant(),
            VehicleType = v.Master.Type.ToString(),
            Brand = v.Master.Brand,
            Model = v.Master.Model,
            CapacityTonnage = v.Master.CapacityTonnage,
            RoadTaxExpiry = v.Master.RoadTaxExpiry.ToDateTime(TimeOnly.MinValue),
            PuspakomExpiry = v.Master.PuspakomExpiry.ToDateTime(TimeOnly.MinValue),
            CurrentOdometer = v.Master.CurrentOdometer,
            Status = v.Master.Status.ToString()
        }).ToList();

        const string sqlReconcileMaintenance = "DELETE FROM vehicle_maintenance WHERE vehicle_id != ALL(@ActiveIds);";
        const string sqlReconcileVehicles = "DELETE FROM vehicle_master WHERE vehicle_id != ALL(@ActiveIds);";

        const string sqlMasterUpsert = @"
        INSERT INTO vehicle_master (
            vehicle_id, registration_number, vehicle_type, brand, model,
            capacity_tonnage, road_tax_expiry, puspakom_expiry, current_odometer, status
        )
        VALUES (
            @VehicleId, @RegistrationNumber, @VehicleType, @Brand, @Model,
            @CapacityTonnage, @RoadTaxExpiry, @PuspakomExpiry, @CurrentOdometer, @Status
        )
        ON CONFLICT (vehicle_id) DO UPDATE SET
            registration_number = EXCLUDED.registration_number,
            vehicle_type        = EXCLUDED.vehicle_type,
            brand               = EXCLUDED.brand,
            model               = EXCLUDED.model,
            capacity_tonnage    = EXCLUDED.capacity_tonnage,
            road_tax_expiry     = EXCLUDED.road_tax_expiry,
            puspakom_expiry     = EXCLUDED.puspakom_expiry,
            current_odometer    = EXCLUDED.current_odometer,
            status              = EXCLUDED.status,
            updated_at          = CURRENT_TIMESTAMP;";

        DbService.ExecuteInTransaction((db, trans) =>
        {
            db.Execute(sqlReconcileMaintenance, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlReconcileVehicles, new { ActiveIds = activeIds }, transaction: trans);
            db.Execute(sqlMasterUpsert, masterRecords, transaction: trans);
        });
    }

    private static void SaveVehicleMaintenanceToDb(IEnumerable<MaintenanceRecord> records)
    {
        var validRecords = records.Where(m => m.VehicleID > 0).Select(m => new
        {
            MaintenanceId = m.MaintenanceID,
            VehicleId = m.VehicleID,
            ServiceType = m.ServiceType.ToString(),
            DateOut = m.DateOut.ToDateTime(TimeOnly.MinValue),
            EstimatedReturn = m.EstimatedReturn.ToDateTime(TimeOnly.MinValue),
            DateCompleted = m.DateCompleted.HasValue
                ? (DateTime?)m.DateCompleted.Value.ToDateTime(TimeOnly.MinValue)
                : null,
            OdometerAtService = m.OdometerAtService,
            CostAmount = m.CostAmount,
            WorkshopName = m.WorkshopName,
            Status = m.Status.ToString()
        }).ToList();

        if (!validRecords.Any()) return;

        const string sqlMaintenanceUpsert = @"
        INSERT INTO vehicle_maintenance (
            maintenance_id, vehicle_id, service_type, date_out, estimated_return,
            date_completed, odometer_at_service, cost_amount, workshop_name, status
        )
        VALUES (
            @MaintenanceId, @VehicleId, @ServiceType, @DateOut, @EstimatedReturn,
            @DateCompleted, @OdometerAtService, @CostAmount, @WorkshopName, @Status
        )
        ON CONFLICT (maintenance_id) DO UPDATE SET
            vehicle_id          = EXCLUDED.vehicle_id,
            service_type        = EXCLUDED.service_type,
            date_out            = EXCLUDED.date_out,
            estimated_return    = EXCLUDED.estimated_return,
            date_completed      = EXCLUDED.date_completed,
            odometer_at_service = EXCLUDED.odometer_at_service,
            cost_amount         = EXCLUDED.cost_amount,
            workshop_name       = EXCLUDED.workshop_name,
            status              = EXCLUDED.status,
            updated_at          = CURRENT_TIMESTAMP;";

        DbService.ExecuteTransaction(sqlMaintenanceUpsert, validRecords);
    }

    #endregion
}