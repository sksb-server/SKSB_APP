namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using SKSB_App.Core.ErpBridge;
using SKSB_App.Core.Types;

public static class Importer
{
    public static List<WorkerRecord> ImportWorkers(string filePath)
    {
        var doc = ERPImporter.LoadFromFile(filePath);
        var workers = new List<WorkerRecord>();

        foreach (var row in doc.Rows)
        {
            // 1. Department Filtering: Exclude OFFICE / ADMIN / MANAGEMENT worker
            string dept = row.GetString("Department");
            if (!string.IsNullOrWhiteSpace(dept))
            {
                if (dept.Equals("OFFICE", StringComparison.OrdinalIgnoreCase) ||
                    dept.Equals("ADMIN", StringComparison.OrdinalIgnoreCase) ||
                    dept.Equals("MANAGEMENT", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }
            }

            // 2. Active Status Filtering: Ignore inactive records if column present
            if (doc.HasColumn("Active"))
            {
                string activeVal = row.GetString("Active");
                if (activeVal.Equals("False", StringComparison.OrdinalIgnoreCase) || activeVal.Equals("0"))
                {
                    continue;
                }
            }

            // 3. Extract Worker ID / Code
            long workerId = row.GetLong("Code");
            if (workerId <= 0)
            {
                string rawCode = row.GetString("Code");
                if (!long.TryParse(rawCode, out workerId) || workerId <= 0)
                    continue;
            }

            // 4. Identity & Contact Details
            string name = row.GetString("Name");
            string phone = row.GetString("Handphone");
            if (string.IsNullOrWhiteSpace(phone))
                phone = row.GetString("Phone");

            // Extract IC (Local numeric identity)
            long nric = 0;
            string rawIc = row.GetString("New IC");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("ICNo");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("IC");

            if (!string.IsNullOrWhiteSpace(rawIc))
            {
                // Strip hyphens if present before parsing to long
                string cleanIc = rawIc.Replace("-", string.Empty).Trim();
                long.TryParse(cleanIc, out nric);
            }

            // Extract Passport (Foreign worker identity)
            string passport = row.GetString("Passport");
            if (string.IsNullOrWhiteSpace(passport)) passport = row.GetString("PassportNo");

            var worker = new WorkerRecord
            {
                WorkerID = workerId,
                WorkerDetails = new WorkerDetails
                {
                    WorkerID = workerId,
                    WorkerName = name,
                    NRIC = nric,
                    Passport = passport,
                    Phone = phone,
                    ContractStatus = ContractStatus.Active
                },
                VehicleLicenses = new List<VehicleLicense>(),
                SkillLicenses = new List<SkillLicense>()
            };

            workers.Add(worker);
        }

        return workers;
    }
}