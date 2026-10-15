namespace SKSB_App.WorkerManager;

using SKSB_App.Core.FileManager;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Types;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

public static class ServerSyncCommands
{
    public static bool SaveWorkerSilent(out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            var (workerList, errors, _) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

            // 1. Validation errors block saving
            if (errors.Count > 0)
            {
                errorMessage = $"{errors.Count} critical validation error(s) must be resolved.";
                return false;
            }

            // 2. Persist skill license types after validation passes
            var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
            if (skillTypes.Count > 0)
            {
                Services.Save(skillTypes, writeToDatabase: true);
            }

            if (workerList == null || workerList.Count == 0)
            {
                return true;
            }

            // 3. Persist worker records
            string snapshotPath = Services.Save(workerList, writeToDatabase: true);
            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool LoadWorkerSilent()
    {
        try
        {
            var workerList = FetchWorkerRecords();
            if (workerList != null && workerList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(workerList);
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

    public static async Task<(bool Success, int Count, string Path)> SaveWorkerAsync(List<WorkerRecord> workerList)
    {
        if (workerList == null || workerList.Count == 0) return (false, 0, string.Empty);

        // Harvest and save skill license types
        var skillTypes = SheetMapper.ExtractSkillLicenseTypes();
        if (skillTypes.Count > 0)
        {
            await Services.SaveAsync(skillTypes, writeToDatabase: true);
        }

        string path = await Services.SaveAsync(workerList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), workerList.Count, path);
    }

    public static Task<(bool Success, int Count, string Path)> SaveWorkersAsync(List<WorkerRecord> workerList)
        => SaveWorkerAsync(workerList);


    public static List<SkillLicenseType> FetchSkillLicenseTypesFromDb()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sql = "SELECT type_id AS TypeID, license_type AS LicenseType, COALESCE(description, '') AS Description FROM lookup_skill_license_types ORDER BY license_type;";
            var result = db.Query<SkillLicenseType>(sql);
            return result?.ToList() ?? new List<SkillLicenseType>();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FetchSkillLicenseTypesFromDb Error] {ex.Message}");
            throw; // Let caller capture for diagnostics
        }
    }
    public static List<WorkerRecord>? FetchWorkerRecords()
    {
        try
        {
            var db = new PostgreSqlService();
            const string sqlWorkers = "SELECT * FROM worker_master ORDER BY worker_id;";
            const string sqlVehicleLicenses = "SELECT * FROM worker_licenses;";
            const string sqlSkillLicenses = "SELECT * FROM worker_skill_licenses;";

            var rawWorkers = db.Query<dynamic>(sqlWorkers);
            if (rawWorkers != null && rawWorkers.Any())
            {
                var workers = rawWorkers.Select(w => new WorkerRecord
                {
                    WorkerID = (long)w.worker_id,
                    WorkerDetails = new WorkerDetails
                    {
                        WorkerID = (long)w.worker_id,
                        WorkerName = (string)w.worker_name,
                        NRIC = w.nric != null ? (long)w.nric : 0,
                        Passport = (string)(w.passport ?? string.Empty),
                        Phone = (string)(w.phone ?? string.Empty),
                        JoinDate = DateOnly.FromDateTime((DateTime)w.join_date),
                        ResignDate = w.resign_date != null ? DateOnly.FromDateTime((DateTime)w.resign_date) : null,
                        ContractStatus = Enum.TryParse<ContractStatus>((string)w.contract_status, out var cs) ? cs : ContractStatus.Active
                    },
                    VehicleLicenses = new List<VehicleLicense>(),
                    SkillLicenses = new List<SkillLicense>()
                }).ToDictionary(w => w.WorkerID);

                var vehicleLicenses = db.Query<dynamic>(sqlVehicleLicenses);
                if (vehicleLicenses != null)
                {
                    foreach (var vl in vehicleLicenses)
                    {
                        long wid = (long)vl.worker_id;
                        if (workers.TryGetValue(wid, out var worker))
                        {
                            worker.VehicleLicenses.Add(new VehicleLicense
                            {
                                LicenseID = vl.license_id != null ? (long)vl.license_id : 0,
                                WorkerID = wid,
                                LicenseClass = Enum.TryParse<LicenseClass>((string)vl.license_type, out var lc) ? lc : LicenseClass.None,
                                RenewDate = DateOnly.FromDateTime((DateTime)vl.renew_date),
                                ExpiryDate = DateOnly.FromDateTime((DateTime)vl.expiry_date),
                                LicenseStatus = Enum.TryParse<LicenseStatus>((string)vl.status, out var ls) ? ls : LicenseStatus.Valid
                            });
                        }
                    }
                }

                var skillLicenses = db.Query<dynamic>(sqlSkillLicenses);
                if (skillLicenses != null)
                {
                    foreach (var sl in skillLicenses)
                    {
                        long wid = (long)sl.worker_id;
                        if (workers.TryGetValue(wid, out var worker))
                        {
                            worker.SkillLicenses.Add(new SkillLicense
                            {
                                LicenseID = sl.license_id != null ? (long)sl.license_id : 0,
                                WorkerID = wid,
                                LicenseType = (string)sl.license_type,
                                RenewDate = DateOnly.FromDateTime((DateTime)sl.renew_date),
                                ExpiryDate = DateOnly.FromDateTime((DateTime)sl.expiry_date),
                                LicenseStatus = Enum.TryParse<LicenseStatus>((string)sl.status, out var ls) ? ls : LicenseStatus.Valid
                            });
                        }
                    }
                }

                return workers.Values.ToList();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FetchWorkerRecords DB Fallback] {ex.Message}");
        }

        return Services.Load<List<WorkerRecord>>();
    }
}