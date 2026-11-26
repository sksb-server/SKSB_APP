namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.Storage;

public static class ServerSyncCommands
{
    /// <summary>
    /// Fills a blank <see cref="StaffDetails.StaffShortName"/> from the staff name (first + last word), so saving never
    /// overwrites staff_master.short_name with an empty string when the sheet has no Short Name value.
    /// </summary>
    internal static void DeriveShortNames(IEnumerable<StaffRecord>? staffList)
    {
        if (staffList == null) return;
        foreach (var staff in staffList)
        {
            var details = staff.StaffDetails;
            if (details == null || !string.IsNullOrWhiteSpace(details.StaffShortName)) continue;
            details.StaffShortName = RuleEvaluator.ApplyTextTransform(details.StaffName ?? string.Empty, TransformKind.ShortNameFirstLast);
        }
    }

    // =========================================================================
    // 1. DELTA BASELINES & FINGERPRINTS
    // =========================================================================
    // Rows as last loaded from the server; diffing the sheet against them yields the delta to save.
    private static readonly TableBaseline<long> StaffBaseline = new();
    private static readonly TableBaseline<long> VehicleLicenseBaseline = new();
    private static readonly TableBaseline<long> SkillLicenseBaseline = new();
    private static readonly TableBaseline<long> StaffSalaryBaseline = new();
    private static readonly TableBaseline<long> StaffLeaveBaseline = new();

    private static string Norm(string? s) => (s ?? string.Empty).Trim();
    private static string Norm(DateOnly d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
    private static string Norm(DateOnly? d) => d.HasValue ? Norm(d.Value) : string.Empty;
    private static string Norm(decimal m) => decimal.Round(m, 2).ToString("F2", CultureInfo.InvariantCulture);

    private static string StaffFingerprint(StaffRecord s)
    {
        var d = s.StaffDetails ?? new StaffDetails();
        return string.Join("\u001F",
            s.StaffID.ToString(CultureInfo.InvariantCulture),
            Norm(d.StaffName),
            Norm(d.StaffShortName),
            Norm(d.Department),
            d.NRIC.ToString(CultureInfo.InvariantCulture),
            Norm(d.Passport),
            Norm(d.Phone),
            Norm(d.JoinDate),
            Norm(d.ResignDate),
            d.ContractStatus.ToString());
    }

    private static string VehicleLicenseFingerprint(VehicleLicense l) => string.Join("\u001F",
        l.VehicleLicenseID.ToString(CultureInfo.InvariantCulture),
        l.StaffID.ToString(CultureInfo.InvariantCulture),
        l.VehicleLicenseClass.ToString(),
        Norm(l.RenewDate),
        Norm(l.ExpiryDate),
        l.VehicleLicenseStatus.ToString(),
        Norm(l.Cost));

    private static string SkillLicenseFingerprint(SkillLicense l) => string.Join("\u001F",
        l.SkillLicenseID.ToString(CultureInfo.InvariantCulture),
        l.StaffID.ToString(CultureInfo.InvariantCulture),
        Norm(l.SkillLicenseType),
        Norm(l.RenewDate),
        Norm(l.ExpiryDate),
        l.SkillLicenseStatus.ToString(),
        Norm(l.Cost));

    private static string StaffSalaryFingerprint(StaffSalary s) => string.Join("\u001F",
        s.SalaryID.ToString(CultureInfo.InvariantCulture),
        s.StaffID.ToString(CultureInfo.InvariantCulture),
        Norm(s.EffectiveDate),
        s.SalaryBefore.HasValue ? Norm(s.SalaryBefore.Value) : string.Empty,
        s.SalaryBasis.ToString(),
        Norm(s.SalaryRate),
        Norm(s.Remarks));

    private static string StaffLeaveFingerprint(StaffLeave l) => string.Join("\u001F",
        l.LeaveID.ToString(CultureInfo.InvariantCulture),
        l.StaffID.ToString(CultureInfo.InvariantCulture),
        Norm(l.LeaveDate),
        Norm(l.LeaveType),
        Norm(l.LeaveDay),
        Norm(l.Description));

    private static void CaptureBaselines(List<StaffRecord> serverList)
    {
        StaffBaseline.Capture(serverList, s => s.StaffID, StaffFingerprint, s => s.ServerUpdatedAt);
        VehicleLicenseBaseline.Capture(serverList.SelectMany(s => s.VehicleLicenses), l => l.VehicleLicenseID, VehicleLicenseFingerprint, l => l.ServerUpdatedAt);
        SkillLicenseBaseline.Capture(serverList.SelectMany(s => s.SkillLicenses), l => l.SkillLicenseID, SkillLicenseFingerprint, l => l.ServerUpdatedAt);
        StaffSalaryBaseline.Capture(serverList.SelectMany(s => s.Salaries), w => w.SalaryID, StaffSalaryFingerprint, w => w.ServerUpdatedAt);
        StaffLeaveBaseline.Capture(serverList.SelectMany(s => s.Leaves), l => l.LeaveID, StaffLeaveFingerprint, l => l.ServerUpdatedAt);
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

    // =========================================================================
    // 2. SAVE COMMANDS
    // =========================================================================
    /// <summary>
    /// Non-interactive save (startup pending upload / contracts without conflict checks). Upsert-only: with no
    /// baseline every row is upserted; with a baseline only changed rows are written, overwriting the server.
    /// </summary>
    public static bool SaveStaffSilent(List<StaffRecord> staffList, out string errorMessage) =>
        SaveStaffWithDelta(staffList, forceOverwrite: true, out errorMessage, out _);

    /// <summary>
    /// Delta save: only rows that differ from the last server load are written (no hard deletes). Modified rows
    /// are checked against their server <c>updated_at</c>; on conflict nothing is saved and
    /// <paramref name="conflicts"/> lists the affected rows.
    /// </summary>
    public static bool SaveStaffWithDelta(List<StaffRecord> staffList, bool forceOverwrite, out string errorMessage, out IReadOnlyList<ConcurrencyConflict>? conflicts)
    {
        errorMessage = string.Empty;
        conflicts = null;
        try
        {
            DeriveShortNames(staffList);

            // 1. Auxiliary SkillLicenseTypesTable. Startup pending-upload commit runs before the sheets are populated:
            //    fall back to the offline snapshot so the lookup table is never wiped.
            var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
            if (skillTypes.Count == 0 && OfflineSnapshot.HasPendingSnapshot<List<StaffRecord>>())
                skillTypes = OfflineSnapshot.Load<List<SkillLicenseType>>() ?? skillTypes;

            // 2. PostgreSQL first; local snapshots only after the server accepted the data.
            if (staffList == null || staffList.Count == 0)
            {
                if (skillTypes.Count > 0)
                {
                    ServerDispatcher.Dispatch(skillTypes);
                    OfflineSnapshot.Save(skillTypes, writeToDatabase: false);
                }
                return true;
            }

            // Child rows inherit the owning staff ID when the sheet left it blank.
            foreach (var s in staffList)
            {
                foreach (var l in s.VehicleLicenses) if (l.StaffID == 0) l.StaffID = s.StaffID;
                foreach (var l in s.SkillLicenses) if (l.StaffID == 0) l.StaffID = s.StaffID;
                foreach (var w in s.Salaries) if (w.StaffID == 0) w.StaffID = s.StaffID;
                foreach (var l in s.Leaves) if (l.StaffID == 0) l.StaffID = s.StaffID;
                s.ResolveSalaryBeforeBases(); // display-only basis is not round-tripped through the sheet
            }

            var payload = new StaffDeltaPayload
            {
                Staff = StaffBaseline.Diff(staffList.Where(s => s.StaffID > 0), s => s.StaffID, StaffFingerprint),
                VehicleLicenses = VehicleLicenseBaseline.Diff(
                    staffList.SelectMany(s => s.VehicleLicenses).Where(l => l.StaffID > 0 && l.VehicleLicenseClass != VehicleLicenseClass.None),
                    l => l.VehicleLicenseID, VehicleLicenseFingerprint),
                SkillLicenses = SkillLicenseBaseline.Diff(
                    staffList.SelectMany(s => s.SkillLicenses).Where(l => l.StaffID > 0 && !string.IsNullOrWhiteSpace(l.SkillLicenseType)),
                    l => l.SkillLicenseID, SkillLicenseFingerprint),
                Salaries = StaffSalaryBaseline.Diff(
                    staffList.SelectMany(s => s.Salaries).Where(w => w.StaffID > 0 && w.SalaryRate > 0m && w.SalaryBasis != SalaryBasis.None),
                    w => w.SalaryID, StaffSalaryFingerprint),
                Leaves = StaffLeaveBaseline.Diff(
                    staffList.SelectMany(s => s.Leaves).Where(l => l.LeaveID > 0 && l.StaffID > 0),
                    l => l.LeaveID, StaffLeaveFingerprint),
                SkillLicenseTypes = skillTypes
            };

            // Single FK-safe transaction: lookup types + staff_master + child licenses (rolled back on any conflict).
            ServerDispatcher.PersistStaffDelta(payload, forceOverwrite);

            if (skillTypes.Count > 0) OfflineSnapshot.Save(skillTypes, writeToDatabase: false);
            OfflineSnapshot.Save(staffList, writeToDatabase: false);
            return true;
        }
        catch (ConcurrencyConflictException cex)
        {
            conflicts = cex.Conflicts;
            errorMessage = cex.Message;
            AppLogger.Warn($"[StaffManager] Delta save rolled back: {cex.Conflicts.Count} concurrency conflict(s).");
            return false;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error("[StaffManager] SaveStaffWithDelta failed", ex);
            return false;
        }
    }

    /// <summary>Offline save: snapshots the staff list AND the skill license types, and queues the pending upload.</summary>
    public static void SaveStaffOffline(List<StaffRecord> staffList)
    {
        DeriveShortNames(staffList);
        var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
        if (skillTypes.Count > 0) OfflineSnapshot.Save(skillTypes, writeToDatabase: false);
        OfflineSnapshot.SaveOfflineSnapshot(staffList);
    }

    public static List<SkillLicenseType> FetchSkillLicenseTypesFromDb() =>
        ServerSyncHelper.FetchAuthoritativeList("StaffManager", "SkillLicenseTypesTable", db =>
        {
            // Tolerate databases that have not been migrated to the newer column names.
            const string colSql = @"
                SELECT column_name
                FROM information_schema.columns
                WHERE table_name = 'lookup_skill_license_types';";
            var existingCols = db.Query<string>(colSql).ToList();

            string nameCol = existingCols.Any(c => c.Equals("skill_license_type", StringComparison.OrdinalIgnoreCase))
                ? "skill_license_type"
                : "license_type";

            string idCol = existingCols.Any(c => c.Equals("skill_license_type_id", StringComparison.OrdinalIgnoreCase))
                ? "skill_license_type_id"
                : "type_id";

            string iconExpr = existingCols.Any(c => c.Equals("display_icon", StringComparison.OrdinalIgnoreCase))
                ? "COALESCE(display_icon, '')"
                : "''";

            string querySql = $@"
                SELECT
                    {idCol}   AS SkillLicenseTypeID,
                    {nameCol} AS SkillLicenseName,
                    COALESCE(description, '') AS Description,
                    {iconExpr} AS DisplayIcon
                FROM lookup_skill_license_types
                ORDER BY {nameCol};";

            return db.Query<SkillLicenseType>(querySql)?.ToList() ?? new List<SkillLicenseType>();
        });

    // =========================================================================
    // 2b. LOAD (PostgreSQL baseline + optional SQL Payroll resync)
    // =========================================================================
    /// <summary>Advisory report of the last SQL Payroll sync, shown once after the sheets were populated.</summary>
    private static ImportSummaryReport? _pendingPayrollReport;
    private static readonly object _reportSync = new();

    /// <summary>Startup / "Load from System" entry point: PostgreSQL baseline merged with a live SQL Payroll resync.</summary>
    public static List<StaffRecord>? FetchAndSyncFromPayroll() => FetchStaffRecords(syncFromPayroll: true);

    /// <summary>
    /// Loads staff from PostgreSQL (or the offline snapshot). With <paramref name="syncFromPayroll"/> and Online Mode,
    /// profiles and salary history are then rebuilt from SQL Payroll (licenses / IDs / concurrency tokens kept from
    /// PostgreSQL); if SQL Payroll is unreachable the PostgreSQL copy is returned with an advisory queued.
    /// </summary>
    public static List<StaffRecord>? FetchStaffRecords(bool syncFromPayroll = false)
    {
        var pgRecords = ServerSyncHelper.FetchAuthoritativeList("StaffManager", "StaffRecord (all staff tables)", QueryStaffRecords);
        if (pgRecords != null)
        {
            foreach (var s in pgRecords) s.RefreshCurrentSalary();
        }

        if (!syncFromPayroll || !ServerEnvironment.IsOnline)
        {
            if (pgRecords != null) RegisterSalarySchedules(pgRecords);
            return pgRecords;
        }

        List<StaffRecord> merged;
        try
        {
            LoadingOverlay.Report("Reading SQL Payroll (HR_EMP & salary history)...");
            var sync = Importer.SyncFromSqlPayroll(pgRecords);
            merged = sync.ImportedStaffs;
            SetPendingReport(sync.ToGenericReport());
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StaffManager] SQL Payroll sync failed; falling back to PostgreSQL records", ex);
            merged = pgRecords ?? new List<StaffRecord>();
            SetPendingReport(new ImportSummaryReport
            {
                SourceSystemName = "SQL Payroll",
                EntityLabel = "staff profile(s) (from PostgreSQL)",
                SuccessfullyImportedCount = merged.Count,
                FilterRuleDescription = "SQL Payroll was unreachable; showing the last records saved in PostgreSQL.",
                AdvisoryBanner = StaffImportResult.AdvisoryBannerText,
                SkippedSectionTitle = "Advisories (Warnings)",
                Title = "SQL Payroll Unreachable",
                FooterHint = "Check that the SQL Payroll PC is on and reachable, then click 'Load from System' to resync.",
                SkippedErrors = new List<ImportSkippedRecord>
                {
                    new()
                    {
                        Category = "SQL Payroll",
                        Identifier = FirebirdPayrollService.PrimaryHost,
                        NameOrTitle = "SQL Payroll Database",
                        Reason = $"SQL Payroll could not be reached, so staff profiles and salaries were NOT resynced and may be out of date. ({ex.Message})"
                    }
                }
            });
        }

        RegisterSalarySchedules(merged);
        return merged;
    }

    private static void SetPendingReport(ImportSummaryReport? report)
    {
        lock (_reportSync) _pendingPayrollReport = report;
    }

    /// <summary>
    /// After-load hook (macro thread, Excel state already restored): consumes the pending SQL Payroll report and shows
    /// it when it carries any warnings / errors.
    /// </summary>
    public static void ShowPendingPayrollAdvisory()
    {
        ImportSummaryReport? report;
        lock (_reportSync)
        {
            report = _pendingPayrollReport;
            _pendingPayrollReport = null;
        }
        if (report == null || report.SkippedErrors.Count == 0) return;

        try
        {
            // Modal dialog exception (CLAUDE.md Rule 5.2): Excel must repaint & respond while the dialog is open.
            dynamic xlApp = ExcelDnaUtil.Application;
            try { xlApp.ScreenUpdating = true; } catch { }
            try { xlApp.EnableEvents = true; } catch { }
            try { Arrow.HideAll(); } catch { }

            using var dialog = new ImportResultDialog(report, "SQL Payroll (PAY-0003.FDB) + Thumbprint Controller");
            dialog.ShowDialog();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Payroll advisory dialog failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Registers every staff salary schedule (hourly rate by effective date) into the staged <see cref="WageProvider"/>
    /// for WAGE_RATE(...) lookups.
    /// </summary>
    private static void RegisterSalarySchedules(List<StaffRecord> staffList)
    {
        try
        {
            WageProvider.ClearSchedules();
            foreach (var s in staffList)
            {
                var schedule = new Dictionary<DateTime, double>();
                foreach (var w in s.Salaries.Where(w => w.SalaryRate > 0m).OrderBy(w => w.EffectiveDate).ThenBy(w => w.SalaryID))
                    schedule[w.EffectiveDate.ToDateTime(TimeOnly.MinValue)] = (double)decimal.Round(w.GetHourlyRateExact(), 4);
                if (schedule.Count == 0) continue;

                foreach (string key in new[] { s.StaffDetails.StaffName, s.StaffDetails.StaffShortName, s.StaffID.ToString(CultureInfo.InvariantCulture) }
                             .Where(k => !string.IsNullOrWhiteSpace(k)).Distinct(StringComparer.OrdinalIgnoreCase))
                    WageProvider.RegisterWorkerWages(key, schedule);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Registering salary schedules failed: {ex.Message}");
        }
    }

    /// <summary>Assembles every staff profile (master + licenses) from PostgreSQL. An empty server returns an empty list.</summary>
    private static List<StaffRecord> QueryStaffRecords(PostgreSqlService db)
    {
        const string sqlStaffs = "SELECT * FROM staff_master ORDER BY staff_id;";
        const string sqlVehicleLicenses = "SELECT * FROM staff_vehicle_licenses ORDER BY vehicle_license_id;";
        const string sqlSkillLicenses = "SELECT * FROM staff_skill_licenses ORDER BY skill_license_id;";
        const string sqlSalaries = "SELECT * FROM staff_salaries ORDER BY staff_id, effective_date, salary_id;";
        const string sqlLeaves = "SELECT * FROM staff_leaves WHERE payroll_dockey IS NOT NULL ORDER BY staff_id, leave_date DESC, payroll_dockey DESC;";

        var rawStaffs = db.Query<dynamic>(sqlStaffs)?.ToList();
        if (rawStaffs != null && rawStaffs.Count > 0)
        {
            var staffs = new Dictionary<long, StaffRecord>();
            foreach (var w in rawStaffs)
            {
                StaffRecord record = MapStaffRow((object)w);
                staffs[record.StaffID] = record;
            }

            var vehicleLicenses = db.Query<dynamic>(sqlVehicleLicenses)?.ToList() ?? new List<dynamic>();
            foreach (var vl in vehicleLicenses)
            {
                long wid = (long)vl.staff_id;
                if (staffs.TryGetValue(wid, out var staff))
                {
                    staff.VehicleLicenses.Add(MapVehicleLicenseRow((object)vl, staff.StaffDetails.StaffName));
                }
            }

            var skillLicenses = db.Query<dynamic>(sqlSkillLicenses)?.ToList() ?? new List<dynamic>();
            foreach (var sl in skillLicenses)
            {
                long wid = (long)sl.staff_id;
                if (staffs.TryGetValue(wid, out var staff))
                {
                    staff.SkillLicenses.Add(MapSkillLicenseRow((object)sl, staff.StaffDetails.StaffName));
                }
            }

            // staff_salaries is created lazily on the first staff save: tolerate its absence on first load.
            int salaryCount = 0;
            try
            {
                var salaries = db.Query<dynamic>(sqlSalaries)?.ToList() ?? new List<dynamic>();
                foreach (var sr in salaries)
                {
                    long sid = (long)sr.staff_id;
                    if (staffs.TryGetValue(sid, out var staff))
                    {
                        staff.Salaries.Add(MapStaffSalaryRow((object)sr, staff.StaffDetails.StaffName));
                        salaryCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[StaffManager] staff_salaries not loaded (table may not exist yet): {ex.Message}");
            }

            // staff_leaves is created lazily (Schedule or Staff save): tolerate its absence on first load.
            int leaveCount = 0;
            try
            {
                var leaves = db.Query<dynamic>(sqlLeaves)?.ToList() ?? new List<dynamic>();
                foreach (var lr in leaves)
                {
                    long sid = Convert.ToInt64(lr.staff_id);
                    if (staffs.TryGetValue(sid, out var staff))
                    {
                        var leave = MapStaffLeaveRow((object)lr, staff.StaffDetails.StaffName);
                        if (leave.LeaveID <= 0) continue;
                        staff.Leaves.Add(leave);
                        leaveCount++;
                    }
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[StaffManager] staff_leaves not loaded (table may not exist yet): {ex.Message}");
            }

            var assembledList = staffs.Values.ToList();
            foreach (var s in assembledList) s.RefreshCurrentSalary();

#if DEBUG
            var sb = new StringBuilder();
            sb.AppendLine("=== POSTGRESQL DOWNLOAD SUMMARY ===");
            sb.AppendLine($"Total Staff Profiles   : {assembledList.Count}");
            sb.AppendLine($"Total Vehicle Licenses : {vehicleLicenses.Count}");
            sb.AppendLine($"Total Skill Licenses   : {skillLicenses.Count}");
            sb.AppendLine($"Total Salary Rows      : {salaryCount}");
            sb.AppendLine($"Total Staff Leaves     : {leaveCount}");
            sb.AppendLine();
            foreach (var s in assembledList)
            {
                sb.AppendLine($"  [{s.StaffID}] {s.StaffDetails.StaffName} ({s.StaffDetails.StaffShortName}) | Dept: {s.StaffDetails.Department} | Status: {s.StaffDetails.ContractStatus}");
                if (s.VehicleLicenses.Count > 0)
                {
                    sb.AppendLine($"      Vehicle Lic: " + string.Join(", ", s.VehicleLicenses.Select(v => $"{v.VehicleLicenseClass} (RM{v.Cost:F2})")));
                }
                if (s.SkillLicenses.Count > 0)
                {
                    sb.AppendLine($"      Skill Lic: " + string.Join(", ", s.SkillLicenses.Select(sk => $"{sk.SkillLicenseType} (RM{sk.Cost:F2})")));
                }
            }
            AppLogger.Info(sb.ToString());
#endif
            CaptureBaselines(assembledList);
            return assembledList;
        }

        // Server reachable but staff_master is empty: the (empty) result is authoritative (empty baseline: all rows are new).
        var empty = new List<StaffRecord>();
        CaptureBaselines(empty);
        return empty;
    }

    // -------------------------------------------------------------------------
    // Row mapping (shared by the full load and conflict-baseline refresh so fingerprints always agree)
    // -------------------------------------------------------------------------
    private static StaffRecord MapStaffRow(dynamic w)
    {
        long staffId = (long)w.staff_id;
        string fullName = (string)(w.staff_name ?? string.Empty);
        string shortName = (string)(w.short_name ?? string.Empty);
        if (string.IsNullOrWhiteSpace(shortName) && !string.IsNullOrWhiteSpace(fullName))
        {
            shortName = RuleEvaluator.ApplyTextTransform(fullName, TransformKind.ShortNameFirstLast);
        }

        long parsedNric = 0;
        if (w.nric != null)
        {
            long.TryParse(w.nric.ToString(), out parsedNric);
        }

        return new StaffRecord
        {
            StaffID = staffId,
            StaffDetails = new StaffDetails
            {
                StaffName = fullName,
                StaffShortName = shortName,
                Department = (string)(w.department ?? "Worker"),
                NRIC = parsedNric,
                Passport = (string)(w.passport ?? string.Empty),
                Phone = (string)(w.phone ?? string.Empty),
                JoinDate = w.join_date != null ? DateOnly.FromDateTime((DateTime)w.join_date) : DateOnly.FromDateTime(DateTime.Today),
                ResignDate = w.resign_date != null ? DateOnly.FromDateTime((DateTime)w.resign_date) : null,
                ContractStatus = Enum.TryParse<ContractStatus>((string)(w.contract_status ?? "Active"), true, out var cs) ? cs : ContractStatus.Active
            },
            VehicleLicenses = new List<VehicleLicense>(),
            SkillLicenses = new List<SkillLicense>(),
            ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col((object)w, "updated_at"))
        };
    }

    private static VehicleLicense MapVehicleLicenseRow(dynamic vl, string staffName) => new VehicleLicense
    {
        VehicleLicenseID = vl.vehicle_license_id != null ? (long)vl.vehicle_license_id : 0,
        StaffID = (long)vl.staff_id,
        StaffName = staffName,
        VehicleLicenseClass = Enum.TryParse<VehicleLicenseClass>((string)(vl.vehicle_license_class ?? vl.license_type ?? vl.license_class ?? "None"), true, out var lc) ? lc : VehicleLicenseClass.None,
        RenewDate = vl.renew_date != null ? DateOnly.FromDateTime((DateTime)vl.renew_date) : DateOnly.FromDateTime(DateTime.Today),
        ExpiryDate = vl.expiry_date != null ? DateOnly.FromDateTime((DateTime)vl.expiry_date) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
        VehicleLicenseStatus = Enum.TryParse<VehicleLicenseStatus>((string)(vl.status ?? "Valid"), true, out var ls) ? ls : VehicleLicenseStatus.Valid,
        Cost = vl.cost != null ? Convert.ToDecimal(vl.cost) : 0.00m,
        ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col((object)vl, "updated_at"))
    };

    private static SkillLicense MapSkillLicenseRow(dynamic sl, string staffName) => new SkillLicense
    {
        SkillLicenseID = sl.skill_license_id != null ? (long)sl.skill_license_id : 0,
        StaffID = (long)sl.staff_id,
        StaffName = staffName,
        SkillLicenseType = (string)(sl.skill_license_type ?? sl.license_type ?? string.Empty),
        RenewDate = sl.renew_date != null ? DateOnly.FromDateTime((DateTime)sl.renew_date) : DateOnly.FromDateTime(DateTime.Today),
        ExpiryDate = sl.expiry_date != null ? DateOnly.FromDateTime((DateTime)sl.expiry_date) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
        SkillLicenseStatus = Enum.TryParse<SkillLicenseStatus>((string)(sl.status ?? "Valid"), true, out var ls) ? ls : SkillLicenseStatus.Valid,
        Cost = sl.cost != null ? Convert.ToDecimal(sl.cost) : 0.00m,
        ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col((object)sl, "updated_at"))
    };

    private static StaffSalary MapStaffSalaryRow(dynamic sr, string staffName) => new StaffSalary
    {
        SalaryID = sr.salary_id != null ? Convert.ToInt64(sr.salary_id) : 0,
        StaffID = Convert.ToInt64(sr.staff_id),
        StaffName = staffName,
        EffectiveDate = sr.effective_date != null ? DateOnly.FromDateTime((DateTime)sr.effective_date) : DateOnly.FromDateTime(DateTime.Today),
        SalaryBefore = sr.salary_before != null ? Convert.ToDecimal(sr.salary_before) : (decimal?)null,
        SalaryBeforeBasis = Enum.TryParse<SalaryBasis>((string)(sr.salary_before_basis ?? "None"), true, out SalaryBasis bb) ? bb : SalaryBasis.None,
        SalaryBasis = Enum.TryParse<SalaryBasis>((string)(sr.salary_basis ?? "Monthly"), true, out SalaryBasis sb) ? sb : SalaryBasis.Monthly,
        SalaryRate = sr.salary_rate != null ? Convert.ToDecimal(sr.salary_rate) : 0m,
        Remarks = (string)(sr.remarks ?? string.Empty),
        ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col((object)sr, "updated_at"))
    };

    /// <summary>Maps a <c>staff_leaves</c> row; <see cref="StaffLeave.LeaveID"/> is the SQL Payroll <c>payroll_dockey</c>.</summary>
    private static StaffLeave MapStaffLeaveRow(dynamic lr, string staffName) => new StaffLeave
    {
        LeaveID = lr.payroll_dockey != null ? Convert.ToInt64(lr.payroll_dockey) : 0,
        StaffID = Convert.ToInt64(lr.staff_id),
        StaffName = staffName,
        LeaveDate = lr.leave_date != null ? DateOnly.FromDateTime((DateTime)lr.leave_date) : DateOnly.FromDateTime(DateTime.Today),
        LeaveType = PayrollLeaveMapper.MapStatus((string)(lr.leave_type ?? string.Empty)).ToString(),
        LeaveDay = lr.leave_day != null ? Convert.ToDecimal(lr.leave_day) : 1m,
        Description = (string)(lr.description ?? string.Empty),
        ServerUpdatedAt = ToUtcDateTime(ServerSyncHelper.Col((object)lr, "updated_at"))
    };

    // =========================================================================
    // 3. CONCURRENCY CONFLICT RESOLUTION ("Keep Server Values")
    // =========================================================================
    /// <summary>
    /// After the pane wrote the server's values into the conflicting cells, re-reads ONLY those rows and moves their
    /// delta baselines to the server's latest fingerprint + <c>updated_at</c>. The next save then treats them as
    /// unchanged, while every other edited row keeps its local edits and its original baseline.
    /// </summary>
    public static void AcceptServerConflictBaselines(IReadOnlyList<ConcurrencyConflict> conflicts)
    {
        if (conflicts == null || conflicts.Count == 0) return;
        if (!ServerEnvironment.IsOnline)
        {
            AppLogger.Warn("[StaffManager] Keep Server Values: server offline, conflict baselines not refreshed.");
            return;
        }

        static long[] IdsFor(IReadOnlyList<ConcurrencyConflict> all, string table) => all
            .Where(c => string.Equals(c.TableName, table, StringComparison.OrdinalIgnoreCase))
            .Select(c => long.TryParse((c.RawKeyValue ?? string.Empty).Trim(), NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) ? id : 0)
            .Where(id => id > 0)
            .Distinct()
            .ToArray();

        long[] staffIds = IdsFor(conflicts, "staff_master");
        long[] vehicleIds = IdsFor(conflicts, "staff_vehicle_licenses");
        long[] skillIds = IdsFor(conflicts, "staff_skill_licenses");
        long[] salaryIds = IdsFor(conflicts, "staff_salaries");
        long[] leaveIds = IdsFor(conflicts, "staff_leaves");
        int refreshed = 0;

        try
        {
            var db = new PostgreSqlService();

            if (staffIds.Length > 0)
            {
                foreach (var w in db.Query<dynamic>("SELECT * FROM staff_master WHERE staff_id = ANY(@Ids);", new { Ids = staffIds }))
                {
                    StaffRecord s = MapStaffRow((object)w);
                    StaffBaseline.UpsertEntry(s.StaffID, StaffFingerprint(s), s.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (vehicleIds.Length > 0)
            {
                foreach (var vl in db.Query<dynamic>("SELECT * FROM staff_vehicle_licenses WHERE vehicle_license_id = ANY(@Ids);", new { Ids = vehicleIds }))
                {
                    VehicleLicense l = MapVehicleLicenseRow((object)vl, string.Empty);
                    VehicleLicenseBaseline.UpsertEntry(l.VehicleLicenseID, VehicleLicenseFingerprint(l), l.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (skillIds.Length > 0)
            {
                foreach (var sl in db.Query<dynamic>("SELECT * FROM staff_skill_licenses WHERE skill_license_id = ANY(@Ids);", new { Ids = skillIds }))
                {
                    SkillLicense l = MapSkillLicenseRow((object)sl, string.Empty);
                    SkillLicenseBaseline.UpsertEntry(l.SkillLicenseID, SkillLicenseFingerprint(l), l.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (salaryIds.Length > 0)
            {
                foreach (var sr in db.Query<dynamic>("SELECT * FROM staff_salaries WHERE salary_id = ANY(@Ids);", new { Ids = salaryIds }))
                {
                    StaffSalary s = MapStaffSalaryRow((object)sr, string.Empty);
                    StaffSalaryBaseline.UpsertEntry(s.SalaryID, StaffSalaryFingerprint(s), s.ServerUpdatedAt);
                    refreshed++;
                }
            }

            if (leaveIds.Length > 0)
            {
                foreach (var lr in db.Query<dynamic>("SELECT * FROM staff_leaves WHERE payroll_dockey = ANY(@Ids);", new { Ids = leaveIds }))
                {
                    StaffLeave l = MapStaffLeaveRow((object)lr, string.Empty);
                    StaffLeaveBaseline.UpsertEntry(l.LeaveID, StaffLeaveFingerprint(l), l.ServerUpdatedAt);
                    refreshed++;
                }
            }

            AppLogger.Info($"[StaffManager] Keep Server Values: refreshed {refreshed} conflict baseline(s) " +
                $"(staff {staffIds.Length}, vehicle lic {vehicleIds.Length}, skill lic {skillIds.Length}, salaries {salaryIds.Length}, leaves {leaveIds.Length}).");
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StaffManager] AcceptServerConflictBaselines failed", ex);
        }
    }

    // =========================================================================
    // STAFF SUMMARY (READ-ONLY FETCHES)
    // =========================================================================
    private static string FormatClock(TimeOnly? t) => t.HasValue ? t.Value.ToString("HH:mm") : "—";

    private static string FormatTimeRange(TimeOnly? inT, TimeOnly? outT) =>
        !inT.HasValue && !outT.HasValue ? "—" : $"{FormatClock(inT)} - {FormatClock(outT)}";

    private static TimeOnly? ToTimeOnly(object? v) => v switch
    {
        TimeOnly t => t,
        TimeSpan ts => TimeOnly.FromTimeSpan(ts),
        DateTime dt => TimeOnly.FromDateTime(dt),
        _ => null
    };

    private static DateOnly ToDateOnly(object? v) => v switch
    {
        DateOnly d => d,
        DateTime dt => DateOnly.FromDateTime(dt),
        _ => default
    };

    private static string JoinDistinct(IEnumerable<string> items) =>
        string.Join(", ", items.Where(i => !string.IsNullOrWhiteSpace(i)).Select(i => i.Trim()).Distinct(StringComparer.OrdinalIgnoreCase));

    /// <summary>
    /// Project fallback for present workers with no check-ins: planned projects that are still on the board
    /// (greyed-out removed-slot projects never leak), else the worker's on-board projects for that date.
    /// </summary>
    private static string OnBoardFallback(AttendanceStatus status, string plannedProjects, List<string>? onBoard)
    {
        if (status is not (AttendanceStatus.Presence or AttendanceStatus.Late or AttendanceStatus.NightShift)) return string.Empty;
        if (onBoard == null || onBoard.Count == 0) return string.Empty;
        var board = new HashSet<string>(onBoard.Select(p => p.Trim()), StringComparer.OrdinalIgnoreCase);
        var planned = plannedProjects.Split(',').Select(t => t.Trim()).Where(t => t.Length > 0 && board.Contains(t)).ToList();
        return JoinDistinct(planned.Count > 0 ? planned : onBoard);
    }

    public static List<StaffSummaryRow> FetchStaffSummaryRows(StaffRecord staff)
    {
        var rows = new List<StaffSummaryRow>();
        var names = new[] { staff.StaffDetails.StaffShortName, staff.StaffDetails.StaffName }
            .Where(n => !string.IsNullOrWhiteSpace(n)).Select(n => n.Trim().ToLowerInvariant()).Distinct().ToArray();

        if (ServerEnvironment.IsOnline)
        {
            try
            {
                try { ServerDispatcher.EnsureScheduleSchema(); }
                catch (Exception ex) { AppLogger.Warn($"[StaffManager] Schedule schema bootstrap skipped: {ex.Message}"); }
                var db = new PostgreSqlService();
                var att = db.Query<dynamic>(
                    "SELECT schedule_date, attendance_status, clock_in_time, clock_out_time, COALESCE(project_check_ins, '') AS project_check_ins, COALESCE(planned_projects, '') AS planned_projects, projects_done, regular_shift, additional_shifts, ot_hours, additional_ot_hours, total_flattened_hours FROM daily_attendances WHERE staff_id = @StaffId ORDER BY schedule_date DESC;",
                    new { StaffId = staff.StaffID }).ToList();

                var vehicles = new Dictionary<DateOnly, List<string>>();
                var scheduledProjects = new Dictionary<DateOnly, List<string>>();
                if (names.Length > 0)
                {
                    var onBoardRows = db.Query<dynamic>(
                        "SELECT schedule_date, project_id FROM daily_project_schedules WHERE EXISTS (SELECT 1 FROM unnest(workers) w WHERE lower(trim(w)) = ANY(@NamesLower)) ORDER BY schedule_date ASC, shift, slot_index;",
                        new { NamesLower = names });
                    foreach (var v in onBoardRows)
                    {
                        DateOnly d = ToDateOnly((object?)v.schedule_date);
                        string pid = ((string?)(v.project_id ?? string.Empty)).Trim();
                        if (pid.Length == 0) continue;
                        if (!scheduledProjects.TryGetValue(d, out var plist)) scheduledProjects[d] = plist = new List<string>();
                        plist.Add(pid);
                    }

                    var veh = db.Query<dynamic>(
                        "SELECT schedule_date, vehicles FROM daily_project_schedules WHERE COALESCE(array_length(vehicles, 1), 0) > 0 AND EXISTS (SELECT 1 FROM unnest(workers) w WHERE lower(trim(w)) = ANY(@NamesLower)) ORDER BY schedule_date ASC;",
                        new { NamesLower = names });
                    foreach (var v in veh)
                    {
                        DateOnly d = ToDateOnly((object?)v.schedule_date);
                        if (!vehicles.TryGetValue(d, out var list)) vehicles[d] = list = new List<string>();
                        if ((object?)v.vehicles is IEnumerable<string> plates) list.AddRange(plates);
                    }
                }

                foreach (var a in att)
                {
                    DateOnly date = ToDateOnly((object?)a.schedule_date);
                    string projects = ((string)(a.project_check_ins ?? string.Empty)).Trim();
                    if (projects.Length == 0 && (object?)a.projects_done is IEnumerable<string> done) projects = string.Join(", ", done);
                    var st = Enum.TryParse<AttendanceStatus>((string)(a.attendance_status ?? "Absence"), true, out AttendanceStatus parsedSt) ? parsedSt : AttendanceStatus.Absence;
                    if (projects.Length == 0)
                        projects = OnBoardFallback(st, (string)(a.planned_projects ?? string.Empty), scheduledProjects.TryGetValue(date, out var sp) ? sp : null);
                    rows.Add(new StaffSummaryRow
                    {
                        Date = date,
                        Time = FormatTimeRange(ToTimeOnly((object?)a.clock_in_time), ToTimeOnly((object?)a.clock_out_time)),
                        Status = st,
                        Projects = projects,
                        VehicleDriven = vehicles.TryGetValue(date, out var vl) ? JoinDistinct(vl) : string.Empty,
                        RegularShift = a.regular_shift != null ? Convert.ToDecimal(a.regular_shift) : 0m,
                        AdditionalShifts = a.additional_shifts != null ? Convert.ToDecimal(a.additional_shifts) : 0m,
                        OTHours = a.ot_hours != null ? Convert.ToDecimal(a.ot_hours) : 0m,
                        AdditionalOTHours = a.additional_ot_hours != null ? Convert.ToDecimal(a.additional_ot_hours) : 0m,
                        FlattenHours = a.total_flattened_hours != null ? Convert.ToDecimal(a.total_flattened_hours) : 0m
                    });
                }
                return ApplySalaryRates(staff, rows);
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[StaffManager] Staff summary query failed: {ex.Message}");
                if (!ServerEnvironment.IsOffline) return rows;
                rows.Clear();
            }
        }

        var snap = OfflineSnapshot.Load<ScheduleRecord>();
        if (snap == null) return rows;
        foreach (var w in snap.WorkerAttendance.Where(w => w.StaffID == staff.StaffID))
        {
            var plates = snap.Projects
                .Where(p => p.Workers.Any(x => names.Contains((x ?? string.Empty).Trim().ToLowerInvariant())))
                .SelectMany(p => p.Vehicles);
            string snapProjects = !string.IsNullOrWhiteSpace(w.ProjectCheckIns) ? w.ProjectCheckIns.Trim() : string.Join(", ", w.ProjectsDone);
            if (snapProjects.Length == 0)
            {
                var onBoardIds = snap.Projects
                    .Where(p => p.Workers.Any(x => names.Contains((x ?? string.Empty).Trim().ToLowerInvariant())))
                    .Select(p => p.ProjectID).ToList();
                snapProjects = OnBoardFallback(w.AttendanceStatus, w.PlannedProjects ?? string.Empty, onBoardIds);
            }
            rows.Add(new StaffSummaryRow
            {
                Date = w.DailyScheduleDate == default ? snap.DailyScheduleDate : w.DailyScheduleDate,
                Time = FormatTimeRange(w.ClockInTime, w.ClockOutTime),
                Status = w.AttendanceStatus,
                Projects = snapProjects,
                VehicleDriven = JoinDistinct(plates),
                RegularShift = w.RegularShift,
                AdditionalShifts = w.AdditionalShifts,
                OTHours = w.OTHours,
                AdditionalOTHours = w.AdditionalOTHours,
                FlattenHours = w.TotalFlattenedHours
            });
        }
        return ApplySalaryRates(staff, rows.OrderByDescending(r => r.Date).ToList());
    }

    /// <summary>Resolves the salary effective on each row's date: hourly Salary Rate and Flatten Hours x rate.</summary>
    private static List<StaffSummaryRow> ApplySalaryRates(StaffRecord staff, List<StaffSummaryRow> rows)
    {
        foreach (var row in rows)
        {
            var salary = staff.GetEffectiveSalary(row.Date);
            decimal hourly = salary?.GetHourlyRateExact() ?? 0m;
            row.SalaryRate = salary != null ? decimal.Round(hourly, 2) : null;
            row.TotalSalary = salary != null ? decimal.Round(row.FlattenHours * hourly, 2) : null;
        }
        return rows;
    }
}
