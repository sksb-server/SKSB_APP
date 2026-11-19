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
}
