namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookSave.Storage;

public static class ServerSyncCommands
{
    public static bool SaveStaffSilent(List<StaffRecord> staffList, out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            // 1. Persist auxiliary SkillLicenseTypesTable
            var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
            if (skillTypes.Count > 0)
            {
                OfflineSnapshot.Save(skillTypes, writeToDatabase: true);
            }

            if (staffList == null || staffList.Count == 0)
            {
                return true;
            }

            // 2. Persist validated staff records to JSON cache & PostgreSQL
            string snapshotPath = OfflineSnapshot.Save(staffList, writeToDatabase: true);
            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool SaveStaffSilent(out string errorMessage)
    {
        errorMessage = string.Empty;
        var (staffList, errors, _) = SheetMapper.ExtractStaffRecordsWithDiagnostics();
        if (errors.Count > 0)
        {
            errorMessage = $"{errors.Count} critical validation error(s) must be resolved.";
            return false;
        }
        return SaveStaffSilent(staffList ?? new List<StaffRecord>(), out errorMessage);
    }

    public static bool LoadStaffSilent()
    {
        try
        {
            var staffList = FetchStaffRecords();
            if (staffList != null && staffList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(staffList);
                return true;
            }
            return false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Load Error: {ex.Message}");
            return false;
        }
    }

    public static async Task<(bool Success, int Count, string Path)> SaveStaffAsync(List<StaffRecord> staffList)
    {
        if (staffList == null || staffList.Count == 0) return (false, 0, string.Empty);

        var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
        if (skillTypes.Count > 0)
        {
            await OfflineSnapshot.SaveAsync(skillTypes, writeToDatabase: true);
        }

        string path = await OfflineSnapshot.SaveAsync(staffList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), staffList.Count, path);
    }

    public static List<SkillLicenseType> FetchSkillLicenseTypesFromDb()
    {
        try
        {
            if (ServerEnvironment.SimulateOffline)
            {
                OfflineSnapshot.LogDataSource("StaffManager", "SkillLicenseTypesTable", "OFFLINE CACHE", "Simulate Server Offline is ON");
                return OfflineSnapshot.Load<List<SkillLicenseType>>() ?? new List<SkillLicenseType>();
            }

            var db = new PostgreSqlService();
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

            string querySql = $@"
                SELECT 
                    {idCol}   AS SkillLicenseTypeID, 
                    {nameCol} AS SkillLicenseName, 
                    COALESCE(description, '') AS Description 
                FROM lookup_skill_license_types 
                ORDER BY {nameCol};";

            var result = db.Query<SkillLicenseType>(querySql)?.ToList() ?? new List<SkillLicenseType>();
            // Server is authoritative: cache the result even when empty so wiped types are not resurrected.
            OfflineSnapshot.Save(result, writeToDatabase: false);
            OfflineSnapshot.LogDataSource("StaffManager", "SkillLicenseTypesTable", "SERVER", $"{result.Count} row(s)");
            return result;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FetchSkillLicenseTypesFromDb Error] {ex.Message}");
            OfflineSnapshot.LogDataSource("StaffManager", "SkillLicenseTypesTable", "OFFLINE CACHE", $"server query failed: {ex.GetType().Name}: {ex.Message}");
            return OfflineSnapshot.Load<List<SkillLicenseType>>() ?? new List<SkillLicenseType>();
        }
    }

    public static List<StaffRecord>? FetchStaffRecords()
    {
        using var _ = AppLogger.Scope("StaffManager.FetchStaffRecords");
        try
        {
            if (!ServerEnvironment.SimulateOffline)
            {
                // 1. Pre-fetch and cache SkillLicenseTypes locally
                FetchSkillLicenseTypesFromDb();

                var db = new PostgreSqlService();
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
                                StaffID = staffId,
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

                    // Cache to local JSON snapshot so VehicleManager and DailyScheduler have fresh staff lookups
                    OfflineSnapshot.Save(assembledList, writeToDatabase: false);
                    OfflineSnapshot.LogDataSource("StaffManager", "StaffRecord (all staff tables)", "SERVER", $"{assembledList.Count} staff");

#if DEBUG
                    if (DebugDialog.DebugEnabled)
                    {
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
                        DebugDialog.ShowLog("[DEBUG] PostgreSQL Download - Staff Summary", sb.ToString());
                    }
#endif
                    return assembledList;
                }

                // Server reachable but staff_master is empty: server is authoritative, so reset the local
                // cache instead of falling back to a stale snapshot.
                try { OfflineSnapshot.Save(new List<StaffRecord>(), writeToDatabase: false); } catch { }
                OfflineSnapshot.LogDataSource("StaffManager", "StaffRecord (all staff tables)", "SERVER", "server reachable but staff_master is empty; local cache reset (no offline fallback)");
                return new List<StaffRecord>();
            }

            OfflineSnapshot.LogDataSource("StaffManager", "StaffRecord (all staff tables)", "OFFLINE CACHE", "Simulate Server Offline is ON");
        }
        catch (Exception ex)
        {
            AppLogger.Error("[FetchStaffRecords DB Fallback]", ex);
            OfflineSnapshot.LogDataSource("StaffManager", "StaffRecord (all staff tables)", "OFFLINE CACHE", $"server query failed: {ex.GetType().Name}: {ex.Message}");
        }

        return OfflineSnapshot.Load<List<StaffRecord>>();
    }
}