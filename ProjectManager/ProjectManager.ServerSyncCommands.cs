namespace SKSB_App.ProjectManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

internal static class ServerSyncCommands
{
    private const string Module = "ProjectManager";

    // =========================================================================
    // 1. SAVE
    // =========================================================================
    private static List<CompanyRecord> ExtractCompaniesFromWorkbook()
    {
        var companies = new List<CompanyRecord>();
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb != null) companies = SheetMapper.ExtractCompanies((object)wb);
        }
        catch (Exception exCompanies)
        {
            AppLogger.Warn($"[{Module}] Could not read CompaniesTable: {exCompanies.Message}");
        }
        return companies;
    }

    /// <summary>Offline save: snapshots the projects AND the companies table, and queues the pending upload.</summary>
    internal static void SaveProjectsOffline(List<ProjectRecord> projects)
    {
        OfflineSnapshot.Save(ExtractCompaniesFromWorkbook(), writeToDatabase: false);
        OfflineSnapshot.SaveOfflineSnapshot(projects);
    }

    /// <summary>
    /// Non-interactive save (startup pending upload / contracts without conflict checks). Upsert-only: with no
    /// baseline every row is upserted; with a baseline only changed rows are written, overwriting the server.
    /// </summary>
    internal static bool SaveProjectsSilent(List<ProjectRecord> projects, out string errorMessage) =>
        SaveProjectsWithDelta(projects, forceOverwrite: true, out errorMessage, out _);

    /// <summary>
    /// Delta save: only companies / projects that differ from the last server load are written (no hard deletes).
    /// Modified rows are checked against their server <c>updated_at</c>; on conflict nothing is saved and
    /// <paramref name="conflicts"/> lists the affected rows.
    /// </summary>
    internal static bool SaveProjectsWithDelta(List<ProjectRecord> projects, bool forceOverwrite, out string errorMessage, out IReadOnlyList<ConcurrencyConflict>? conflicts)
    {
        errorMessage = string.Empty;
        conflicts = null;
        try
        {
            projects ??= new List<ProjectRecord>();

            // 1. Companies live on their own sheet; persist them with the projects.
            var companies = ExtractCompaniesFromWorkbook();

            // Startup pending-upload commit runs before PopulateWorksheet: an unpopulated CompaniesTable must not
            // drop the companies saved offline.
            if (companies.Count == 0 && OfflineSnapshot.HasPendingSnapshot<List<ProjectRecord>>())
                companies = OfflineSnapshot.Load<List<CompanyRecord>>() ?? companies;

            foreach (var p in projects) p.ProjectCode = NormalizeCode(p.ProjectCode);

            var payload = new ProjectDeltaPayload
            {
                Companies = CompanyBaseline.Diff(companies.Where(c => c.CompanyID > 0), c => c.CompanyID, CompanyFingerprint),
                Projects = ProjectBaseline.Diff(projects.Where(p => p.ProjectCode.Length > 0), p => p.ProjectCode, ProjectFingerprint)
            };

            AppLogger.Info(
                $"[ProjectManager.SaveProjectsWithDelta] force={forceOverwrite}, Projects={projects.Count}, Companies={companies.Count}");

            // 2. PostgreSQL: single transaction, companies first so project rows can resolve them (rolled back on any conflict).
            ServerDispatcher.PersistProjectDelta(payload, forceOverwrite);

            // 3. Local offline snapshots (only after the server commit succeeded)
            OfflineSnapshot.Save(projects, writeToDatabase: false);
            OfflineSnapshot.Save(companies, writeToDatabase: false);
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
            AppLogger.Error($"[{Module}] SaveProjectsWithDelta failed", ex);
            return false;
        }
    }

    // =========================================================================
    // 1b. DELTA BASELINES & FINGERPRINTS
    // =========================================================================
    // Rows as last loaded from the server; diffing the sheet against them yields the delta to save.
    private static readonly TableBaseline<string> ProjectBaseline = new();
    private static readonly TableBaseline<long> CompanyBaseline = new();

    private static string NormalizeCode(string? code) => (code ?? string.Empty).Trim().ToUpperInvariant();
    private static string Norm(string? s) => (s ?? string.Empty).Trim();
    private static string Norm(DateOnly d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static string Norm(DateOnly? d) => d.HasValue ? Norm(d.Value) : string.Empty;
    private static string Norm(decimal m) => decimal.Round(m, 2).ToString("F2", CultureInfo.InvariantCulture);

    private static string ProjectFingerprint(ProjectRecord p)
    {
        var d = p.ProjectDetails ?? new ProjectDetails();
        return string.Join("\u001F",
            NormalizeCode(p.ProjectCode),
            p.Category.ToString(),
            Norm(d.ReferenceNo),
            string.Join(",", (d.ProjectType ?? new List<ProjectType>()).Where(t => t != ProjectType.None).Distinct().Select(t => t.ToString())),
            d.Priority.ToString(),
            d.Status.ToString(),
            d.CompanyID.ToString(CultureInfo.InvariantCulture),
            Norm(d.CompanyName),
            (d.StaffID ?? 0).ToString(CultureInfo.InvariantCulture),
            Norm(d.PIC),
            d.District.ToString(),
            Norm(d.Address),
            Norm(d.StartDate),
            Norm(d.TargetEndDate),
            Norm(d.ActualEndDate),
            Norm(d.EstimatedCost),
            Norm(d.PromisedReturn),
            Norm(d.ActualReturn),
            Norm(d.Remarks));
    }

    private static string CompanyFingerprint(CompanyRecord c) => string.Join("\u001F",
        c.CompanyID.ToString(CultureInfo.InvariantCulture),
        Norm(c.LegalCompanyName),
        Norm(c.ShortCompanyName));

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

    // =========================================================================
    // 1c. CONCURRENCY CONFLICT RESOLUTION ("Keep Server Values")
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
            .Where(k => k.Length > 0);

        string[] projectIds = KeysFor("project_master").Select(NormalizeCode).Distinct().ToArray();
        long[] companyIds = KeysFor("company_master")
            .Select(k => long.TryParse(k, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) ? id : 0)
            .Where(id => id > 0)
            .Distinct()
            .ToArray();
        int refreshed = 0;

        try
        {
            var db = new PostgreSqlService();

            if (companyIds.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT * FROM company_master WHERE company_id = ANY(@Ids);", new { Ids = companyIds }))
                {
                    CompanyRecord c = MapCompanyRow((object)r);
                    CompanyBaseline.UpsertEntry(c.CompanyID, CompanyFingerprint(c), c.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (projectIds.Length > 0)
            {
                foreach (var r in db.Query<dynamic>("SELECT * FROM project_master WHERE project_id = ANY(@Ids);", new { Ids = projectIds }))
                {
                    ProjectRecord p = BuildProject(r);
                    ProjectBaseline.UpsertEntry(p.ProjectCode, ProjectFingerprint(p), p.ServerUpdatedAt);
                    refreshed++;
                }
            }

            AppLogger.Info($"[{Module}] Keep Server Values: refreshed {refreshed} conflict baseline(s) " +
                $"(companies {companyIds.Length}, projects {projectIds.Length}).");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{Module}] AcceptServerConflictBaselines failed", ex);
        }
    }

    // =========================================================================
    // 2. FETCH PROJECTS WITH OFFLINE FALLBACK
    // =========================================================================
    internal static List<ProjectRecord> FetchProjectRecords() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "ProjectRecord (project_master)", db =>
        {
            var rawProjects = db.Query<dynamic>("SELECT * FROM project_master ORDER BY project_id;").ToList();
            AppLogger.Info($"[PostgreSQL Query] project_master returned {rawProjects.Count} record(s).");
            var projects = new List<ProjectRecord>();
            foreach (var p in rawProjects) projects.Add(BuildProject(p));

            // Only reached when the server answered: an empty result is an authoritative (empty) baseline.
            ProjectBaseline.Capture(projects, p => p.ProjectCode, ProjectFingerprint, p => p.ServerUpdatedAt);
            return projects;
        });

    private static ProjectRecord BuildProject(dynamic p)
    {
        static DateOnly Date(object? v, DateOnly fallback) => ServerSyncHelper.DateOrDefault(v, fallback);
        static decimal Dec(object? v) => ServerSyncHelper.Dec(v);
        static decimal? DecN(object? v) => ServerSyncHelper.DecN(v);
        static object? Col(object row, string name) => ServerSyncHelper.Col(row, name);
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
            },
            ServerUpdatedAt = ToUtcDateTime(Col((object)p, "updated_at"))
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

    /// <summary>
    /// Online Mode: PostgreSQL only (a SQL error yields an empty list, never snapshot data).
    /// Offline Mode, or a connection failure that flips the app offline: the local snapshot.
    /// </summary>
    private static List<T> LookupFromServerOrSnapshot<T>(string label, Func<PostgreSqlService, List<T>> query, Func<List<T>> snapshot)
    {
        if (ServerEnvironment.IsOnline)
        {
            try
            {
                return query(new PostgreSqlService());
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[{Module}] {label} query failed: {ex.Message}");
                if (!ServerEnvironment.IsOffline) return new List<T>();
            }
        }
        return snapshot();
    }

    private static void SyncStaffLists()
    {
        var staff = LookupFromServerOrSnapshot(
            "staff_master",
            db => db.Query<dynamic>("SELECT staff_id, staff_name, short_name FROM staff_master ORDER BY staff_id;")
                .Select(r => new StaffRecord
                {
                    StaffID = Convert.ToInt64(r.staff_id),
                    StaffDetails = new ProjectStaffStub((string?)r.staff_name, (string?)r.short_name).ToDetails()
                })
                .ToList(),
            () => OfflineSnapshot.Load<List<StaffRecord>>() ?? new List<StaffRecord>());

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
        var plates = LookupFromServerOrSnapshot(
            "vehicle_master",
            db => db.Query<dynamic>("SELECT vehicle_plate, COALESCE(vehicle_type, '') AS vehicle_type FROM vehicle_master ORDER BY vehicle_plate;")
                .Select(r => (Plate: (string?)r.vehicle_plate ?? string.Empty, Type: (string?)r.vehicle_type ?? string.Empty))
                .ToList(),
            () => (OfflineSnapshot.Load<List<VehicleRecord>>() ?? new List<VehicleRecord>())
                .Select(v => (Plate: v.VehiclePlate ?? string.Empty,
                              Type: v.VehicleDetail.Type == VehicleType.None ? string.Empty : v.VehicleDetail.Type.ToString()))
                .ToList());

        // VehiclePlateList and VehiclePlateTypeList stay index-aligned ("PLATE — Type" dropdown labels).
        var clean = plates
            .Where(p => !string.IsNullOrWhiteSpace(p.Plate))
            .Select(p => (Plate: p.Plate.Trim().ToUpperInvariant(), Type: p.Type.Trim()))
            .GroupBy(p => p.Plate, StringComparer.OrdinalIgnoreCase).Select(g => g.First())
            .OrderBy(p => p.Plate, StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (clean.Count == 0) return;
        DataList.SyncAlignedLists(new Dictionary<string, IEnumerable<string>>
        {
            ["VehiclePlateList"] = clean.Select(p => p.Plate).ToList(),
            ["VehiclePlateTypeList"] = clean.Select(p => p.Type).ToList(),
        });
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
        var codes = LookupFromServerOrSnapshot(
            "project_master",
            db => db.Query<string>("SELECT project_id FROM project_master ORDER BY project_id;").ToList(),
            () => (OfflineSnapshot.Load<List<ProjectRecord>>() ?? new List<ProjectRecord>()).Select(p => p.ProjectCode).ToList());

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
    internal static List<CompanyRecord> FetchCompanies() =>
        ServerSyncHelper.FetchAuthoritativeList(Module, "CompaniesTable", db =>
        {
            var companies = new List<CompanyRecord>();
            foreach (var r in db.Query<dynamic>("SELECT * FROM company_master ORDER BY company_id;"))
                companies.Add(MapCompanyRow((object)r));

            // Only reached when the server answered (offline-cache rows must never become a baseline).
            CompanyBaseline.Capture(companies, c => c.CompanyID, CompanyFingerprint, c => c.ServerUpdatedAt);
            return companies;
        });

    /// <summary>Shared by the full load and the conflict-baseline refresh so fingerprints always agree.</summary>
    private static CompanyRecord MapCompanyRow(object row)
    {
        object? id = ServerSyncHelper.Col(row, "company_id");
        return new CompanyRecord
        {
            CompanyID = id != null ? Convert.ToInt64(id, CultureInfo.InvariantCulture) : 0,
            LegalCompanyName = ServerSyncHelper.ColumnText(row, "legal_company_name"),
            ShortCompanyName = ServerSyncHelper.ColumnText(row, "short_company_name"),
            ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col(row, "updated_at"))
        };
    }
}
