namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
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

    private static void CaptureBaselines(List<StaffRecord> serverList)
    {
        StaffBaseline.Capture(serverList, s => s.StaffID, StaffFingerprint, s => s.ServerUpdatedAt);
        VehicleLicenseBaseline.Capture(serverList.SelectMany(s => s.VehicleLicenses), l => l.VehicleLicenseID, VehicleLicenseFingerprint, l => l.ServerUpdatedAt);
        SkillLicenseBaseline.Capture(serverList.SelectMany(s => s.SkillLicenses), l => l.SkillLicenseID, SkillLicenseFingerprint, l => l.ServerUpdatedAt);
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

    public static List<StaffRecord>? FetchStaffRecords() =>
        ServerSyncHelper.FetchAuthoritativeList("StaffManager", "StaffRecord (all staff tables)", QueryStaffRecords);

    /// <summary>Assembles every staff profile (master + licenses) from PostgreSQL. An empty server returns an empty list.</summary>
    private static List<StaffRecord> QueryStaffRecords(PostgreSqlService db)
    {
        const string sqlStaffs = "SELECT * FROM staff_master ORDER BY staff_id;";
        const string sqlVehicleLicenses = "SELECT * FROM staff_vehicle_licenses ORDER BY vehicle_license_id;";
        const string sqlSkillLicenses = "SELECT * FROM staff_skill_licenses ORDER BY skill_license_id;";

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

            var assembledList = staffs.Values.ToList();

#if DEBUG
            var sb = new StringBuilder();
            sb.AppendLine("=== POSTGRESQL DOWNLOAD SUMMARY ===");
            sb.AppendLine($"Total Staff Profiles   : {assembledList.Count}");
            sb.AppendLine($"Total Vehicle Licenses : {vehicleLicenses.Count}");
            sb.AppendLine($"Total Skill Licenses   : {skillLicenses.Count}");
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

            AppLogger.Info($"[StaffManager] Keep Server Values: refreshed {refreshed} conflict baseline(s) " +
                $"(staff {staffIds.Length}, vehicle lic {vehicleIds.Length}, skill lic {skillIds.Length}).");
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StaffManager] AcceptServerConflictBaselines failed", ex);
        }
    }

    // =========================================================================
    // STAFF SUMMARY (READ-ONLY FETCHES)
    // =========================================================================
    public static Dictionary<long, string> FetchWorkerPrimarySkills()
    {
        var result = new Dictionary<long, string>();
        if (ServerEnvironment.IsOnline)
        {
            try
            {
                var rows = new PostgreSqlService().Query<dynamic>(
                    "SELECT staff_id, COALESCE(primary_skill, '') AS primary_skill FROM worker_preferences WHERE COALESCE(trim(primary_skill), '') <> '';");
                foreach (var r in rows)
                    result[Convert.ToInt64(r.staff_id)] = ((string)(r.primary_skill ?? string.Empty)).Trim();
                return result;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[StaffManager] worker_preferences query failed: {ex.Message}");
                if (!ServerEnvironment.IsOffline) return result;
            }
        }

        var prefs = OfflineSnapshot.Load<ScheduleRecord>()?.Preferences;
        if (prefs != null)
            foreach (var p in prefs)
                if (!string.IsNullOrWhiteSpace(p.PrimarySkill)) result[p.StaffID] = p.PrimarySkill.Trim();
        return result;
    }

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

    public static List<StaffSummaryRow> FetchStaffSummaryRows(StaffRecord staff)
    {
        var rows = new List<StaffSummaryRow>();
        var names = new[] { staff.StaffDetails.StaffShortName, staff.StaffDetails.StaffName }
            .Where(n => !string.IsNullOrWhiteSpace(n)).Select(n => n.Trim().ToLowerInvariant()).Distinct().ToArray();

        if (ServerEnvironment.IsOnline)
        {
            try
            {
                var db = new PostgreSqlService();
                var att = db.Query<dynamic>(
                    "SELECT schedule_date, attendance_status, clock_in_time, clock_out_time, COALESCE(project_check_ins, '') AS project_check_ins, projects_done, regular_shift, additional_shifts, ot_hours, additional_ot_hours, total_flattened_hours FROM daily_attendances WHERE staff_id = @StaffId ORDER BY schedule_date DESC;",
                    new { StaffId = staff.StaffID }).ToList();

                var vehicles = new Dictionary<DateOnly, List<string>>();
                if (names.Length > 0)
                {
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
                    rows.Add(new StaffSummaryRow
                    {
                        Date = date,
                        Time = FormatTimeRange(ToTimeOnly((object?)a.clock_in_time), ToTimeOnly((object?)a.clock_out_time)),
                        Status = Enum.TryParse<AttendanceStatus>((string)(a.attendance_status ?? "Absence"), true, out AttendanceStatus st) ? st : AttendanceStatus.Absence,
                        Projects = projects,
                        VehicleDriven = vehicles.TryGetValue(date, out var vl) ? JoinDistinct(vl) : string.Empty,
                        RegularShift = a.regular_shift != null ? Convert.ToDecimal(a.regular_shift) : 0m,
                        AdditionalShifts = a.additional_shifts != null ? Convert.ToDecimal(a.additional_shifts) : 0m,
                        OTHours = a.ot_hours != null ? Convert.ToDecimal(a.ot_hours) : 0m,
                        AdditionalOTHours = a.additional_ot_hours != null ? Convert.ToDecimal(a.additional_ot_hours) : 0m,
                        FlattenHours = a.total_flattened_hours != null ? Convert.ToDecimal(a.total_flattened_hours) : 0m
                    });
                }
                return rows;
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
            rows.Add(new StaffSummaryRow
            {
                Date = w.DailyScheduleDate == default ? snap.DailyScheduleDate : w.DailyScheduleDate,
                Time = FormatTimeRange(w.ClockInTime, w.ClockOutTime),
                Status = w.AttendanceStatus,
                Projects = !string.IsNullOrWhiteSpace(w.ProjectCheckIns) ? w.ProjectCheckIns.Trim() : string.Join(", ", w.ProjectsDone),
                VehicleDriven = JoinDistinct(plates),
                RegularShift = w.RegularShift,
                AdditionalShifts = w.AdditionalShifts,
                OTHours = w.OTHours,
                AdditionalOTHours = w.AdditionalOTHours,
                FlattenHours = w.TotalFlattenedHours
            });
        }
        return rows.OrderByDescending(r => r.Date).ToList();
    }
}
