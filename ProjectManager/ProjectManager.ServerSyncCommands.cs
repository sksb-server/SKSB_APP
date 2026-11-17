namespace SKSB_App.ProjectManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;
using System.Linq;

internal static class ServerSyncCommands
{
    private const string Module = "ProjectManager";

    // =========================================================================
    // 1. SAVE
    // =========================================================================
    internal static bool SaveProjectsSilent(List<ProjectRecord> projects, out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            // 1. Companies live on their own sheet; persist them with the projects.
            var companies = new List<CompanyRecord>();
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic? wb = xlApp.ActiveWorkbook;
                if (wb != null) companies = SheetMapper.ExtractCompanies((object)wb);
            }
            catch (Exception exCompanies)
            {
                AppLogger.Warn($"[{Module}] Could not read CompaniesTable for save: {exCompanies.Message}");
            }

            AppLogger.Info(
                $"[ProjectManager.SaveProjectsSilent] Projects={projects.Count}, " +
                $"Companies={companies.Count}");

            // 2. PostgreSQL (companies first so project rows can resolve them)
            ServerDispatcher.PersistCompanies(companies, isFullBatch: true);
            ServerDispatcher.PersistProjects(projects, isFullBatch: true);

            // 3. Local offline snapshots
            OfflineSnapshot.Save(projects, writeToDatabase: false);
            OfflineSnapshot.Save(companies, writeToDatabase: false);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error($"[{Module}] SaveProjectsSilent failed", ex);
            return false;
        }
    }

    // =========================================================================
    // 2. FETCH PROJECTS WITH OFFLINE FALLBACK
    // =========================================================================
    internal static List<ProjectRecord> FetchProjectRecords()
    {
        using var _ = AppLogger.Scope("ProjectManager.FetchProjectRecords");

        try
        {
            var db = new PostgreSqlService();

            var rawProjects = db.Query<dynamic>("SELECT * FROM project_master ORDER BY project_id;").ToList();
            AppLogger.Info($"[PostgreSQL Query] project_master returned {rawProjects.Count} record(s).");

            var projects = new List<ProjectRecord>();
            foreach (var p in rawProjects)
            {
                projects.Add(BuildProject(p));
            }

            if (projects.Count > 0)
            {
                try { OfflineSnapshot.Save(projects, writeToDatabase: false); } catch { }
                OfflineSnapshot.LogDataSource(Module, "ProjectRecord (project_master)", "SERVER", $"{projects.Count} project(s)");
                return projects;
            }

            // Server reachable but empty: the server is authoritative, so reset the local cache too.
            try { OfflineSnapshot.Save(new List<ProjectRecord>(), writeToDatabase: false); } catch { }
            OfflineSnapshot.LogDataSource(Module, "ProjectRecord", "SERVER", "server reachable but project_master is empty; local cache reset");
            return new List<ProjectRecord>();
        }
        catch (Exception ex)
        {
            AppLogger.Error("FetchProjectRecords server query failed", ex);
            OfflineSnapshot.LogDataSource(Module, "ProjectRecord", "OFFLINE CACHE", $"server query failed: {ex.GetType().Name}: {ex.Message}");
        }

        var fallback = OfflineSnapshot.Load<List<ProjectRecord>>() ?? new List<ProjectRecord>();
        AppLogger.Warn($"[Server Offline Fallback] Loaded {fallback.Count} project records from local offline snapshot.");
        return fallback;
    }

    private static ProjectRecord BuildProject(dynamic p)
    {
        static DateOnly Date(object? v, DateOnly fallback) => v != null ? DateOnly.FromDateTime(Convert.ToDateTime(v)) : fallback;
        static decimal Dec(object? v) => v != null ? Convert.ToDecimal(v) : 0m;
        static decimal? DecN(object? v) => v != null ? Convert.ToDecimal(v) : null;
        // DapperRow is an IDictionary: reading a column that does not exist yet (database not migrated) yields null.
        static object? Col(object row, string name) =>
            row is IDictionary<string, object?> d && d.TryGetValue(name, out var v) ? v : null;
        static string Str(object? v) => v?.ToString() ?? string.Empty;

        Enum.TryParse<ProjectCategory>(Str(p.category), true, out var category);
        Enum.TryParse<ProjectDistrict>(Str(p.district), true, out var district);
        var types = Str(p.project_type)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => Enum.TryParse<ProjectType>(t, true, out var pt) ? (ProjectType?)pt : null)
            .Where(t => t.HasValue && t.Value != ProjectType.None)
            .Select(t => t!.Value)
            .Distinct()
            .ToList();
        Enum.TryParse<ProjectPriority>(Str(p.priority), true, out var priority);
        Enum.TryParse<ProjectStatus>(Str(p.status), true, out var status);

        object? staffId = p.staff_id;
        object? companyId = p.company_id;
        object? actualEnd = p.actual_end_date;

        return new ProjectRecord
        {
            ProjectCode = Str(p.project_id).Trim().ToUpperInvariant(),
            Category = category,
            ProjectDetails = new ProjectDetails
            {
                ReferenceNo = Str(p.reference_no),
                ProjectType = types,
                Priority = priority,
                Status = status,
                CompanyID = companyId != null ? Convert.ToInt64(companyId) : 0,
                CompanyName = Str(p.company_name),
                StaffID = staffId != null ? Convert.ToInt64(staffId) : null,
                PIC = Str(p.pic),
                District = district,
                Address = Str(p.address),
                StartDate = Date(p.start_date, DateOnly.FromDateTime(DateTime.Today)),
                TargetEndDate = Date(p.target_end_date, DateOnly.FromDateTime(DateTime.Today.AddDays(365))),
                ActualEndDate = actualEnd != null ? DateOnly.FromDateTime(Convert.ToDateTime(actualEnd)) : null,
                EstimatedCost = Dec(p.estimated_cost),
                PromisedReturn = Dec(p.promised_return),
                ActualReturn = Dec(p.actual_return),
                Remarks = Str(p.remarks),
                DaysWorked = p.days_worked != null ? Convert.ToInt32(p.days_worked) : 0,
                TotalRegularShifts = DecN(Col((object)p, "total_regular_shifts") ?? Col((object)p, "total_shifts")),
                TotalHours = DecN(Col((object)p, "total_hours") ?? Col((object)p, "flattened_hours")),
                OvertimeShifts = Col((object)p, "overtime_shifts") is object os ? Convert.ToInt32(os) : null,
                TotalOvertimeHours = DecN(Col((object)p, "total_overtime_hours")),
                TotalOdometer = DecN(p.total_odometer),
                FuelCost = DecN(p.fuel_cost),
                LabourCost = DecN(p.labour_cost),
                ClaimCost = DecN(p.claim_cost),
                ActualProjectCost = DecN(p.actual_project_cost),
                ProfitMargin = DecN(p.profit_margin)
            }
        };
    }

    // =========================================================================
    // 2b. LOOKUP LISTS FOR THE 'Data' SHEET (server first, offline cache second)
    // =========================================================================

    /// <summary>
    /// Writes StaffIDList / StaffNameList / StaffShortNameList / VehiclePlateList / CompanyIDList /
    /// CompanyNameList / ProjectIDList straight from PostgreSQL, so the Data sheet is populated even when
    /// Staff Manager / Vehicle Manager have never been opened on this machine. Falls back to the local
    /// snapshots when the server is unreachable.
    /// </summary>
    internal static void SyncLookupLists()
    {
        SyncStaffLists();
        SyncVehicleList();
        SyncCompanyLists();
        SyncProjectList();
    }

    private static void SyncStaffLists()
    {
        List<StaffRecord> staff;
        try
        {
            var db = new PostgreSqlService();
            staff = db.Query<dynamic>("SELECT staff_id, staff_name, short_name FROM staff_master ORDER BY staff_id;")
                .Select(r => new StaffRecord
                {
                    StaffID = Convert.ToInt64(r.staff_id),
                    StaffDetails = new ProjectStaffStub((string?)r.staff_name, (string?)r.short_name).ToDetails()
                })
                .ToList();

            // Staff Manager owns the full staff snapshot (licences etc.): only seed the cache when it has none.
            try
            {
                var existing = OfflineSnapshot.Load<List<StaffRecord>>();
                if ((existing == null || existing.Count == 0) && staff.Count > 0)
                    OfflineSnapshot.Save(staff, writeToDatabase: false);
            }
            catch { }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] staff_master not reachable, using local snapshot: {ex.Message}");
            staff = OfflineSnapshot.Load<List<StaffRecord>>() ?? new List<StaffRecord>();
        }

        // StaffIDList and StaffShortNameList are index-aligned (Staff ID <-> PIC sync), so both skip the same rows.
        var valid = staff
            .Where(s => s.StaffID > 0)
            .Select(s => (Id: s.StaffID,
                          Short: !string.IsNullOrWhiteSpace(s.StaffDetails.StaffShortName) ? s.StaffDetails.StaffShortName.Trim() : s.StaffDetails.StaffName?.Trim() ?? string.Empty))
            .Where(s => s.Short.Length > 0)
            .OrderBy(s => s.Id)
            .ToList();

        if (valid.Count == 0) return;
        DataList.SyncValues("StaffIDList", valid.Select(s => s.Id.ToString()).ToList());
        DataList.SyncValues("StaffNameList", valid.Select(s => s.Short).ToList());
        DataList.SyncValues("StaffShortNameList", valid.Select(s => s.Short).ToList());
    }

    private static void SyncVehicleList()
    {
        List<string> plates;
        try
        {
            var db = new PostgreSqlService();
            plates = db.Query<string>("SELECT vehicle_plate FROM vehicle_master ORDER BY vehicle_plate;").ToList();

            try
            {
                var existing = OfflineSnapshot.Load<List<VehicleRecord>>();
                if ((existing == null || existing.Count == 0) && plates.Count > 0)
                    OfflineSnapshot.Save(plates.Select(pl => new VehicleRecord { VehiclePlate = pl }).ToList(), writeToDatabase: false);
            }
            catch { }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] vehicle_master not reachable, using local snapshot: {ex.Message}");
            plates = (OfflineSnapshot.Load<List<VehicleRecord>>() ?? new List<VehicleRecord>())
                .Select(v => v.VehiclePlate).ToList();
        }

        var clean = plates.Where(pl => !string.IsNullOrWhiteSpace(pl)).Select(pl => pl.Trim().ToUpperInvariant()).Distinct().OrderBy(pl => pl).ToList();
        if (clean.Count > 0) DataList.SyncValues("VehiclePlateList", clean);
    }

    private static void SyncCompanyLists()
    {
        var companies = FetchCompanies().Where(c => c.CompanyID > 0).OrderBy(c => c.CompanyID).ToList();
        if (companies.Count == 0) return;
        DataList.SyncValues("CompanyIDList", companies.Select(c => c.CompanyID.ToString()).ToList());
        DataList.SyncValues("CompanyNameList", companies
            .Select(c => !string.IsNullOrWhiteSpace(c.ShortCompanyName) ? c.ShortCompanyName : c.LegalCompanyName)
            .ToList());
    }

    private static void SyncProjectList()
    {
        List<string> codes;
        try
        {
            codes = new PostgreSqlService().Query<string>("SELECT project_id FROM project_master ORDER BY project_id;").ToList();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{Module}] project_master not reachable, using local snapshot: {ex.Message}");
            codes = (OfflineSnapshot.Load<List<ProjectRecord>>() ?? new List<ProjectRecord>()).Select(p => p.ProjectCode).ToList();
        }

        var clean = codes.Where(c => !string.IsNullOrWhiteSpace(c)).Select(c => c.Trim().ToUpperInvariant()).Distinct().ToList();
        if (clean.Count > 0) DataList.SyncValues("ProjectIDList", clean);
    }

    private readonly record struct ProjectStaffStub(string? Name, string? Short)
    {
        public StaffDetails ToDetails() => new() { StaffName = Name ?? string.Empty, StaffShortName = Short ?? string.Empty };
    }

    // =========================================================================
    // 3. COMPANIES
    // =========================================================================
    internal static List<CompanyRecord> FetchCompanies()
    {
        try
        {
            var db = new PostgreSqlService();
            var list = db.Query<CompanyRecord>(@"
                SELECT company_id AS CompanyID,
                       legal_company_name AS LegalCompanyName,
                       short_company_name AS ShortCompanyName
                FROM company_master
                ORDER BY company_id;").ToList();

            try { OfflineSnapshot.Save(list, writeToDatabase: false); } catch { }
            OfflineSnapshot.LogDataSource(Module, "CompaniesTable", "SERVER", $"{list.Count} row(s)");
            return list;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{Module}] company_master server query failed", ex);
            OfflineSnapshot.LogDataSource(Module, "CompaniesTable", "OFFLINE CACHE", $"server query failed: {ex.GetType().Name}: {ex.Message}");
            return OfflineSnapshot.Load<List<CompanyRecord>>() ?? new List<CompanyRecord>();
        }
    }
}
