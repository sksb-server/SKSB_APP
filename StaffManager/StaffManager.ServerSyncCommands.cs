namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using Dapper;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.Storage;

public static class ServerSyncCommands
{
    private static readonly PostgreSqlService Db = new();

    public static List<StaffRecord>? FetchStaffRecords()
    {
        try
        {
            if (!ServerEnvironment.IsSimulatedOffline())
            {
                const string sqlStaffs = @"
                    SELECT 
                        staff_id AS StaffID,
                        staff_name AS StaffName,
                        short_name AS ShortName,
                        department AS Department,
                        nric AS NRIC,
                        passport AS Passport,
                        phone AS Phone,
                        join_date AS JoinDateRaw,
                        resign_date AS ResignDateRaw,
                        contract_status AS StatusStr
                    FROM staff_master
                    ORDER BY staff_id;";

                const string sqlVLicenses = @"
                    SELECT 
                        vehicle_license_id AS VehicleLicenseID,
                        staff_id AS StaffID,
                        vehicle_license_class AS VehicleLicenseClassStr,
                        renew_date AS RenewDateRaw,
                        expiry_date AS ExpiryDateRaw,
                        status AS StatusStr,
                        cost AS Cost
                    FROM staff_vehicle_licenses;";

                const string sqlSLicenses = @"
                    SELECT 
                        skill_license_id AS SkillLicenseID,
                        staff_id AS StaffID,
                        skill_license_type AS SkillLicenseType,
                        renew_date AS RenewDateRaw,
                        expiry_date AS ExpiryDateRaw,
                        status AS StatusStr,
                        cost AS Cost
                    FROM staff_skill_licenses;";

                var staffRows = Db.Query<dynamic>(sqlStaffs).ToList();
                if (staffRows.Count > 0)
                {
                    var staffDict = new Dictionary<long, StaffRecord>();

                    foreach (var r in staffRows)
                    {
                        long staffId = Convert.ToInt64(r.staffid);
                        string statusStr = r.statusstr?.ToString() ?? string.Empty;

                        ContractStatus cStatus = ContractStatus.Active;
                        if (Enum.TryParse<ContractStatus>(statusStr, true, out var parsedStatus))
                            cStatus = parsedStatus;

                        var staff = new StaffRecord
                        {
                            StaffID = staffId,
                            StaffDetails = new StaffDetails
                            {
                                StaffID = staffId,
                                StaffName = r.staffname?.ToString() ?? string.Empty,
                                StaffShortName = r.shortname?.ToString() ?? string.Empty,
                                Department = r.department?.ToString() ?? "Worker",
                                NRIC = r.nric != null ? Convert.ToInt64(r.nric) : 0,
                                Passport = r.passport?.ToString() ?? string.Empty,
                                Phone = r.phone?.ToString() ?? string.Empty,
                                JoinDate = r.joindateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(r.joindateraw)) : DateOnly.FromDateTime(DateTime.Today),
                                ResignDate = r.resigndateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(r.resigndateraw)) : null,
                                ContractStatus = cStatus
                            }
                        };

                        staffDict[staffId] = staff;
                    }

                    // Attach Vehicle Licenses
                    var vLicRows = Db.Query<dynamic>(sqlVLicenses).ToList();
                    foreach (var vl in vLicRows)
                    {
                        long sid = Convert.ToInt64(vl.staffid);
                        if (staffDict.TryGetValue(sid, out var s))
                        {
                            string cStr = vl.vehiclelicenseclassstr?.ToString() ?? string.Empty;
                            VehicleLicenseClass vClass = Enum.TryParse<VehicleLicenseClass>(cStr, true, out var parsedVc) ? parsedVc : VehicleLicenseClass.None;

                            string sStr = vl.statusstr?.ToString() ?? string.Empty;
                            VehicleLicenseStatus vStatus = Enum.TryParse<VehicleLicenseStatus>(sStr, true, out var parsedVs) ? parsedVs : VehicleLicenseStatus.Valid;

                            s.VehicleLicenses.Add(new VehicleLicense
                            {
                                VehicleLicenseID = Convert.ToInt64(vl.vehiclelicenseid),
                                StaffID = sid,
                                StaffName = s.StaffDetails.StaffName,
                                VehicleLicenseClass = vClass,
                                RenewDate = vl.renewdateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(vl.renewdateraw)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = vl.expirydateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(vl.expirydateraw)) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                                VehicleLicenseStatus = vStatus,
                                Cost = vl.cost != null ? Convert.ToDecimal(vl.cost) : 0.00m
                            });
                        }
                    }

                    // Attach Skill Licenses
                    var sLicRows = Db.Query<dynamic>(sqlSLicenses).ToList();
                    foreach (var sl in sLicRows)
                    {
                        long sid = Convert.ToInt64(sl.staffid);
                        if (staffDict.TryGetValue(sid, out var s))
                        {
                            string sStr = sl.statusstr?.ToString() ?? string.Empty;
                            SkillLicenseStatus skStatus = Enum.TryParse<SkillLicenseStatus>(sStr, true, out var parsedSs) ? parsedSs : SkillLicenseStatus.Valid;

                            s.SkillLicenses.Add(new SkillLicense
                            {
                                SkillLicenseID = Convert.ToInt64(sl.skilllicenseid),
                                StaffID = sid,
                                StaffName = s.StaffDetails.StaffName,
                                SkillLicenseType = sl.skilllicensetype?.ToString() ?? string.Empty,
                                RenewDate = sl.renewdateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(sl.renewdateraw)) : DateOnly.FromDateTime(DateTime.Today),
                                ExpiryDate = sl.expirydateraw != null ? DateOnly.FromDateTime(Convert.ToDateTime(sl.expirydateraw)) : DateOnly.FromDateTime(DateTime.Today.AddYears(1)),
                                SkillLicenseStatus = skStatus,
                                Cost = sl.cost != null ? Convert.ToDecimal(sl.cost) : 0.00m
                            });
                        }
                    }

                    var list = staffDict.Values.ToList();
                    OfflineSnapshot.Save(list, writeToDatabase: false);
                    return list;
                }
            }
        }
        catch { }

        return OfflineSnapshot.Load<List<StaffRecord>>();
    }

    public static List<SkillLicenseType> FetchSkillLicenseTypesFromDb()
    {
        try
        {
            if (!ServerEnvironment.IsSimulatedOffline())
            {
                const string sql = @"
                    SELECT 
                        skill_license_type_id AS SkillLicenseTypeID, 
                        skill_license_type    AS SkillLicenseName, 
                        COALESCE(description, '') AS Description 
                    FROM lookup_skill_license_types 
                    ORDER BY skill_license_type;";

                var result = Db.Query<SkillLicenseType>(sql).ToList();
                if (result.Count > 0)
                {
                    OfflineSnapshot.Save(result, writeToDatabase: false);
                    return result;
                }
            }
        }
        catch { }

        return OfflineSnapshot.Load<List<SkillLicenseType>>() ?? new List<SkillLicenseType>();
    }

    public static bool SaveStaffSilent(List<StaffRecord> records, out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            OfflineSnapshot.Save(records, writeToDatabase: true);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            return false;
        }
    }
}