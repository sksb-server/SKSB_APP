namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.Types.Attributes;

public static class StaffImportConfig
{
    public const string DefaultDepartment = "Worker";

    private static readonly Dictionary<string, string> DepartmentMappings = new(StringComparer.OrdinalIgnoreCase)
    {
        { "PROJECT", "Worker" },
        { "OFFICE", "Office" },
        { "ADMIN", "Office" },
        { "MANAGEMENT", "Office" }
    };

    public static string MapDepartment(string? rawDepartment)
    {
        if (string.IsNullOrWhiteSpace(rawDepartment))
            return DefaultDepartment;

        string clean = rawDepartment.Trim();
        if (DepartmentMappings.TryGetValue(clean, out string? targetDept) && !string.IsNullOrWhiteSpace(targetDept))
        {
            return targetDept;
        }

        return clean;
    }
}

public class StaffImportResult
{
    public List<StaffRecord> ImportedStaffs { get; set; } = new();
    public List<ImportSkippedRecord> SkippedErrors { get; set; } = new();
    public int TotalRowsProcessed { get; set; } = 0;

    public ImportSummaryReport ToGenericReport() => new()
    {
        SourceSystemName = "SQL Payroll",
        EntityLabel = "staff(s)",
        SuccessfullyImportedCount = ImportedStaffs.Count,
        FilteredExclusionCount = 0,
        FilterRuleDescription = "Mapped departments: PROJECT -> Worker, OFFICE/ADMIN -> Office.",
        SkippedErrors = SkippedErrors
    };
}

public static class Importer
{
    public static StaffImportResult ImportStaffs(string filePath)
    {
        var doc = ErpImporter.LoadFromFile(filePath);
        var result = new StaffImportResult();

        for (int r = 0; r < doc.Rows.Count; r++)
        {
            var row = doc.Rows[r];
            int excelRowNum = r + 2;
            result.TotalRowsProcessed++;

            string rawCode = row.GetString("Code");
            string name = row.GetString("Name")?.Trim() ?? string.Empty;
            string rawDept = row.GetString("Department");

            string mappedDept = StaffImportConfig.MapDepartment(rawDept);

            // 1. Validate Staff ID / Code
            if (string.IsNullOrWhiteSpace(rawCode))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = "[Empty]",
                    NameOrTitle = name,
                    Reason = "Missing Staff Code / ID."
                });
                continue;
            }

            long staffId = row.GetLong("Code");
            if (staffId <= 0 && (!long.TryParse(rawCode, out staffId) || staffId <= 0))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = rawCode,
                    NameOrTitle = name,
                    Reason = $"Invalid non-numeric Staff Code '{rawCode}'."
                });
                continue;
            }

            // 2. Validate Name
            if (string.IsNullOrWhiteSpace(name))
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = rawCode,
                    NameOrTitle = "[Empty]",
                    Reason = $"Staff ID {staffId} has no name specified."
                });
                continue;
            }

            string shortName = RuleEvaluator.ApplyTextTransform(name, TransformKind.ShortNameFirstLast);

            // 3. Validate Employment Dates
            DateOnly? rawJoinDate = row.GetDateOnly("Join Date") ?? row.GetDateOnly("JoinDate");
            DateOnly? rawResignDate = row.GetDateOnly("Resign") ?? row.GetDateOnly("Resign Date") ?? row.GetDateOnly("ResignDate");

            if (!rawJoinDate.HasValue && rawResignDate.HasValue)
            {
                result.SkippedErrors.Add(new ImportSkippedRecord
                {
                    ExcelRowNumber = excelRowNum,
                    Identifier = rawCode,
                    NameOrTitle = name,
                    Reason = $"Staff is marked resigned ({rawResignDate.Value:dd/MM/yyyy}) but has no Join Date recorded. Skipped from active operations."
                });
                continue;
            }

            DateOnly joinDate = rawJoinDate ?? DateOnly.FromDateTime(DateTime.Today);

            // 4. Contact & Identity
            long nric = 0;
            string rawIc = row.GetString("New IC");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("NewIC");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("ICNo");
            if (string.IsNullOrWhiteSpace(rawIc)) rawIc = row.GetString("IC");
            if (!string.IsNullOrWhiteSpace(rawIc))
            {
                string cleanIc = Regex.Replace(rawIc, @"[^\d]", string.Empty);
                long.TryParse(cleanIc, out nric);
            }

            string passport = row.GetString("Passport");
            if (string.IsNullOrWhiteSpace(passport)) passport = row.GetString("PassportNo");

            string phone = row.GetString("Mobile");
            if (string.IsNullOrWhiteSpace(phone)) phone = row.GetString("Handphone");
            if (string.IsNullOrWhiteSpace(phone)) phone = row.GetString("Phone");

            ContractStatus initialStatus = (rawResignDate.HasValue && rawResignDate.Value <= DateOnly.FromDateTime(DateTime.Today))
                ? ContractStatus.Resigned
                : ContractStatus.Active;

            var staff = new StaffRecord
            {
                StaffID = staffId,
                StaffDetails = new StaffDetails
                {
                    StaffID = staffId,
                    StaffName = name,
                    StaffShortName = shortName,
                    Department = mappedDept,
                    NRIC = nric,
                    Passport = passport,
                    Phone = phone,
                    JoinDate = joinDate,
                    ResignDate = rawResignDate,
                    ContractStatus = initialStatus
                },
                VehicleLicenses = new List<VehicleLicense>(),
                SkillLicenses = new List<SkillLicense>()
            };

            result.ImportedStaffs.Add(staff);
        }

        return result;
    }
}