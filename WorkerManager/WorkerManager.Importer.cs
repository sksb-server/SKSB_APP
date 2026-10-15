namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.ErpBridge;
using SKSB_App.Core.Types;

public class WorkerSkippedEntry
{
    public int ExcelRowNumber { get; set; }
    public string RawCode { get; set; } = string.Empty;
    public string Name { get; set; } = string.Empty;
    public string Reason { get; set; } = string.Empty;
}

public class WorkerImportResult
{
    public List<WorkerRecord> ImportedWorkers { get; set; } = new();
    public List<WorkerSkippedEntry> SkippedErrors { get; set; } = new();
    public int OfficeFilteredCount { get; set; } = 0;
    public int TotalRowsProcessed { get; set; } = 0;
}

public static class Importer
{
    public static WorkerImportResult ImportWorkers(string filePath)
    {
        var doc = ERPImporter.LoadFromFile(filePath);
        var result = new WorkerImportResult();

        for (int r = 0; r < doc.Rows.Count; r++)
        {
            var row = doc.Rows[r];
            int excelRowNum = r + 2; // Row 1 is header, data begins on Row 2
            result.TotalRowsProcessed++;

            string dept = row.GetString("Department");
            string rawCode = row.GetString("Code");
            string name = row.GetString("Name");

            // 1. Department Filtering: Exclude OFFICE / ADMIN / MANAGEMENT staff
            if (!string.IsNullOrWhiteSpace(dept))
            {
                if (dept.Equals("OFFICE", StringComparison.OrdinalIgnoreCase) ||
                    dept.Equals("ADMIN", StringComparison.OrdinalIgnoreCase) ||
                    dept.Equals("MANAGEMENT", StringComparison.OrdinalIgnoreCase))
                {
                    result.OfficeFilteredCount++;
                    continue; // Expected exclusion, not an error
                }
            }

            // 2. Validate Worker ID / Code
            if (string.IsNullOrWhiteSpace(rawCode))
            {
                result.SkippedErrors.Add(new WorkerSkippedEntry
                {
                    ExcelRowNumber = excelRowNum,
                    RawCode = "[Empty]",
                    Name = name,
                    Reason = "Missing Worker Code / ID."
                });
                continue;
            }

            long workerId = row.GetLong("Code");
            if (workerId <= 0)
            {
                if (!long.TryParse(rawCode, out workerId) || workerId <= 0)
                {
                    result.SkippedErrors.Add(new WorkerSkippedEntry
                    {
                        ExcelRowNumber = excelRowNum,
                        RawCode = rawCode,
                        Name = name,
                        Reason = $"Invalid non-numeric Worker Code '{rawCode}'."
                    });
                    continue;
                }
            }

            // 3. Validate Name
            if (string.IsNullOrWhiteSpace(name))
            {
                result.SkippedErrors.Add(new WorkerSkippedEntry
                {
                    ExcelRowNumber = excelRowNum,
                    RawCode = rawCode,
                    Name = "[Empty]",
                    Reason = $"Worker ID {workerId} has no name specified."
                });
                continue;
            }

            // 4. Extract Identity & Contact Info
            long nric = 0;
            string rawIc = row.GetString("New IC");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("NewIC");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("ICNo");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("IC");

            if (!string.IsNullOrWhiteSpace(rawIc))
            {
                string cleanIc = System.Text.RegularExpressions.Regex.Replace(rawIc, @"[^\d]", string.Empty);
                long.TryParse(cleanIc, out nric);
            }

            string passport = row.GetString("Passport");
            if (string.IsNullOrWhiteSpace(passport)) passport = row.GetString("PassportNo");

            string phone = row.GetString("Mobile");
            if (string.IsNullOrWhiteSpace(phone)) phone = row.GetString("Handphone");
            if (string.IsNullOrWhiteSpace(phone)) phone = row.GetString("Phone");

            // 5. Employment & Resignation Dates
            DateOnly joinDate = row.GetDateOnly("Join Date")
                ?? row.GetDateOnly("JoinDate")
                ?? DateOnly.FromDateTime(DateTime.Today);

            DateOnly? resignDate = row.GetDateOnly("Resign")
                ?? row.GetDateOnly("Resign Date")
                ?? row.GetDateOnly("ResignDate");

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
                    JoinDate = joinDate,
                    ResignDate = resignDate,
                    ContractStatus = ContractStatus.Active
                },
                VehicleLicenses = new List<VehicleLicense>(),
                SkillLicenses = new List<SkillLicense>()
            };

            result.ImportedWorkers.Add(worker);
        }

        return result;
    }
}