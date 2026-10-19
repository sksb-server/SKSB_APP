namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.ErpBridge;
using SKSB_App.Core.Types;

public class WorkerImportResult
{
    public List<WorkerRecord> ImportedWorkers { get; set; } = new();
    public List<ImportSkippedRecord> SkippedErrors { get; set; } = new();
    public int OfficeFilteredCount { get; set; } = 0;
    public int TotalRowsProcessed { get; set; } = 0;

    /// <summary>
    /// Projects this domain result into the Core Generic Report.
    /// </summary>
    public ImportSummaryReport ToGenericReport() => new()
    {
        SourceSystemName = "SQL Payroll",
        EntityLabel = "worker(s)",
        SuccessfullyImportedCount = ImportedWorkers.Count,
        FilteredExclusionCount = OfficeFilteredCount,
        FilterRuleDescription = "Only synced personnel where Department = 'PROJECT' (Office / Admin staff excluded).",
        SkippedErrors = SkippedErrors
    };
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
            int excelRowNum = r + 2;
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
                    continue;
                }
            }

            // 2. Validate Worker ID / Code
            if (string.IsNullOrWhiteSpace(rawCode))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = "[Empty]",
                    NameOrTitle = name,
                    Reason = "Missing Worker Code / ID."
                });
                continue;
            }

            long workerId = row.GetLong("Code");
            if (workerId <= 0 && (!long.TryParse(rawCode, out workerId) || workerId <= 0))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = rawCode,
                    NameOrTitle = name,
                    Reason = $"Invalid non-numeric Worker Code '{rawCode}'."
                });
                continue;
            }

            // 3. Validate Name
            if (string.IsNullOrWhiteSpace(name))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = rawCode,
                    NameOrTitle = "[Empty]",
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