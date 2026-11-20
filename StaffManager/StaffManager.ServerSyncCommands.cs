namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
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

    public static bool SaveStaffSilent(List<StaffRecord> staffList, out string errorMessage)
    {
        errorMessage = string.Empty;
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

            // Single FK-safe transaction: lookup types + staff_master + child licenses.
            ServerDispatcher.PersistStaff(staffList, isFullBatch: true, skillTypes);

            if (skillTypes.Count > 0) OfflineSnapshot.Save(skillTypes, writeToDatabase: false);
            OfflineSnapshot.Save(staffList, writeToDatabase: false);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error("[StaffManager] SaveStaffSilent failed", ex);
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

                staffs[staffId] = new StaffRecord
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
                    SkillLicenses = new List<SkillLicense>()
                };
            }

            var vehicleLicenses = db.Query<dynamic>(sqlVehicleLicenses)?.ToList() ?? new List<dynamic>();
            foreach (var vl in vehicleLicenses)
            {
                long wid = (long)vl.staff_id;
                if (staffs.TryGetValue(wid, out var staff))
                {
                    staff.VehicleLicenses.Add(new VehicleLicense
                    {
                        VehicleLicenseID = vl.vehicle_license_id != null ? (long)vl.vehicle_license_id : 0,
                        StaffID = wid,
                        StaffName = staff.StaffDetails.StaffName,
                        VehicleLicenseClass = Enum.TryParse<VehicleLicenseClass>((string)(vl.vehicle_license_class ?? vl.license_type ?? vl.license_class ?? "None"), true, out var lc) ? lc : VehicleLicenseClass.None,
                        RenewDate = vl.renew_date != null ? DateOnly.FromDateTime((DateTime)vl.renew_date) : DateOnly.FromDateTime(DateTime.Today),
                        ExpiryDate = vl.expiry_date != null ? DateOnly.FromDateTime((DateTime)vl.expiry_date) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                        VehicleLicenseStatus = Enum.TryParse<VehicleLicenseStatus>((string)(vl.status ?? "Valid"), true, out var ls) ? ls : VehicleLicenseStatus.Valid,
                        Cost = vl.cost != null ? Convert.ToDecimal(vl.cost) : 0.00m
                    });
                }
            }

            var skillLicenses = db.Query<dynamic>(sqlSkillLicenses)?.ToList() ?? new List<dynamic>();
            foreach (var sl in skillLicenses)
            {
                long wid = (long)sl.staff_id;
                if (staffs.TryGetValue(wid, out var staff))
                {
                    staff.SkillLicenses.Add(new SkillLicense
                    {
                        SkillLicenseID = sl.skill_license_id != null ? (long)sl.skill_license_id : 0,
                        StaffID = wid,
                        StaffName = staff.StaffDetails.StaffName,
                        SkillLicenseType = (string)(sl.skill_license_type ?? sl.license_type ?? string.Empty),
                        RenewDate = sl.renew_date != null ? DateOnly.FromDateTime((DateTime)sl.renew_date) : DateOnly.FromDateTime(DateTime.Today),
                        ExpiryDate = sl.expiry_date != null ? DateOnly.FromDateTime((DateTime)sl.expiry_date) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                        SkillLicenseStatus = Enum.TryParse<SkillLicenseStatus>((string)(sl.status ?? "Valid"), true, out var ls) ? ls : SkillLicenseStatus.Valid,
                        Cost = sl.cost != null ? Convert.ToDecimal(sl.cost) : 0.00m
                    });
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
            return assembledList;
        }

        // Server reachable but staff_master is empty: the (empty) result is authoritative.
        return new List<StaffRecord>();
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
                    "SELECT schedule_date, attendance_status, clock_in_time, clock_out_time, COALESCE(project_check_ins, '') AS project_check_ins, projects_done, regular_shift, additional_shifts, ot_hours, additional_ot_hours, total_flattened_hours FROM daily_attendances WHERE staff_id = @StaffId ORDER BY schedule_date ASC;",
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
        return rows.OrderBy(r => r.Date).ToList();
    }
}
