namespace SKSB_App.ProjectManager;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;
using System.Linq;

internal static partial class SheetMapper
{
    // =========================================================================
    // 1. SHEET / TABLE IDENTITY
    // =========================================================================
    public const string MasterSheetName = "iJob Projects";
    public const string MasterTableName = "iJobProjectsTable";

    internal sealed record CategorySheet(ProjectCategory Category, string[] Aliases, string DefaultSheetName, string TableName);

    /// <summary>One master sheet / table per <see cref="ProjectCategory"/>; all share the same 27-column schema.</summary>
    internal static readonly CategorySheet[] CategorySheets =
    {
        new(ProjectCategory.iJob, new[] { "iJob Projects", "iJobProjects", "iJob" }, "iJob Projects", "iJobProjectsTable"),
        new(ProjectCategory.MWO,  new[] { "MWO Projects", "MWOProjects", "MWO" },     "MWO Projects",  "MWOProjectsTable"),
        new(ProjectCategory.PVT,  new[] { "PVT Projects", "PVTProjects", "PVT" },     "PVT Projects",  "PVTProjectsTable")
    };

    internal static readonly string[] MasterSheetAliases = CategorySheets[0].Aliases;
    internal static readonly string[] CompaniesAliases = { "Companies", "Company" };

    // =========================================================================
    // 2. DECLARATIVE WORKBOOK MAPPING (3 MASTER TABLES, NO CHILD TABLES)
    // =========================================================================
    private static readonly WorkbookMapping<ProjectRecord, string> Mapping = new()
    {
        MasterSheetAliases = MasterSheetAliases,
        MasterSheetName = MasterSheetName,
        MasterTableName = MasterTableName,
        EntityLabel = "project",
        MasterKeyLabel = "Project Code",
        MasterKeySelector = p => p.ProjectCode,
        KeyComparer = StringComparer.OrdinalIgnoreCase,
        IsValidKey = code => !string.IsNullOrWhiteSpace(code),

        // Three tables are read / written by the module; primary-key uniqueness is then checked across all of them.
        CustomMasterExtractor = ExtractAllMasterTables,
        CustomMasterPopulator = PopulateAllMasterTables,

        PreValidationHook = (wb, masters, errors, warnings) => ExtractCompaniesForValidation((object)wb, errors),
        CustomValidator = (wb, masters, errors, warnings) => ValidateProjects(masters, errors, warnings),

        OnPopulateAuxiliaryTables = (wb, masters) => PopulateCompanies((object)wb),

        RecalculateAfterPopulate = true,

        HealTableOrder = new[]
        {
            "CompaniesTable",
            "iJobProjectsTable",
            "MWOProjectsTable",
            "PVTProjectsTable"
        },

        AuditTitle = "Project Manager Populate Audit",
        MasterLabel = "Projects",
        InspectTables = new[]
        {
            "iJobProjectsTable", "MWOProjectsTable", "PVTProjectsTable", "CompaniesTable"
        }
    };

    // =========================================================================
    // 3. EXTRACT / POPULATE ENTRY POINTS
    // =========================================================================
    internal static (List<ProjectRecord>? Data, List<string> Errors, List<string> Warnings) ExtractProjectRecordsWithDiagnostics()
        => WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);

    internal static List<ProjectRecord> ExtractProjectRecords()
    {
        var (data, _, _) = ExtractProjectRecordsWithDiagnostics();
        return data ?? new List<ProjectRecord>();
    }

    internal static void PopulateWorksheet(List<ProjectRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 4. MULTI-SHEET MASTER EXTRACTION
    // =========================================================================
    private static List<ProjectRecord> ExtractAllMasterTables(object wbObj, List<string> errors)
    {
        var all = new List<ProjectRecord>();

        foreach (var sheet in CategorySheets)
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, sheet.Aliases, sheet.DefaultSheetName);
            if (ws == null)
            {
                errors.Add($"  Worksheet '{sheet.DefaultSheetName}' was not found in workbook.");
                continue;
            }

            try
            {
                var rows = TableMapper.ExtractFromTable<ProjectRecord>((object)ws, sheet.TableName);
                foreach (var record in rows)
                {
                    record.ProjectCode = (record.ProjectCode ?? string.Empty).Trim().ToUpperInvariant();
                    if (string.IsNullOrWhiteSpace(record.ProjectCode)) continue; // reported by the required-field rule

                    record.Category = sheet.Category;
                    RuleCalculator.ApplyRecordCalculations(record); // Actual Project Cost, Profit / Margin
                    all.Add(record);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[ProjectManager] Failed extracting '{sheet.TableName}'", ex);
                errors.Add($"  [{sheet.DefaultSheetName}]: Failed reading '{sheet.TableName}': {ex.Message}");
            }
        }

        return all;
    }

    // =========================================================================
    // 5. MULTI-SHEET MASTER POPULATION
    // =========================================================================
    private static void PopulateAllMasterTables(object wbObj, List<ProjectRecord> records)
    {
        foreach (var sheet in CategorySheets)
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, sheet.Aliases, sheet.DefaultSheetName);
            if (ws == null) continue;

            var forSheet = records.Where(r => r.Category == sheet.Category).ToList();
            foreach (var record in forSheet) RuleCalculator.ApplyRecordCalculations(record);

            try
            {
                // Actual Project Cost / Profit / Margin are written from the C# calculation above.
                TableMapper.PopulateTableSafe<ProjectRecord>((object)ws, sheet.TableName, forSheet);
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[ProjectManager] Failed populating '{sheet.TableName}'", ex);
            }
        }
    }

    // =========================================================================
    // 6. DOMAIN VALIDATION
    // =========================================================================
    private static void ValidateProjects(List<ProjectRecord> masters, List<string> errors, List<string> warnings)
    {
        // Every project must point at a company that exists on the Companies sheet.
        if (_extractedCompanies != null && _extractedCompanies.Count > 0)
        {
            var companyIds = new HashSet<long>(_extractedCompanies.Select(c => c.CompanyID));
            foreach (var p in masters)
            {
                long id = p.ProjectDetails.CompanyID;
                if (id > 0 && !companyIds.Contains(id))
                {
                    errors.Add($"  [{p.Category} Projects]: Project '{p.ProjectCode}' references Company ID '{id}' which does not exist in 'CompaniesTable'.");
                }
            }
        }
    }
}
