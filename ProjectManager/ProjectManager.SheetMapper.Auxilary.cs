namespace SKSB_App.ProjectManager;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;
using System.Linq;

internal static partial class SheetMapper
{
    private const string CompaniesTableName = "CompaniesTable";

    /// <summary>Companies read during the latest extract; used by the project validators.</summary>
    private static List<CompanyRecord>? _extractedCompanies;

    internal static bool HasExtractedCompanies => (_extractedCompanies?.Count ?? 0) > 0;

    // =========================================================================
    // 1. EXTRACTION & VALIDATION
    // =========================================================================

    /// <summary>Reads 'CompaniesTable' from the active workbook (empty list if the sheet is missing).</summary>
    internal static List<CompanyRecord> ExtractCompanies(object wbObj)
    {
        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, CompaniesAliases, CompaniesAliases[0]);
        if (ws == null) return new List<CompanyRecord>();

        return TableMapper.ExtractFromTable<CompanyRecord>((object)ws, CompaniesTableName)
            .Where(c => c.CompanyID > 0 || !string.IsNullOrWhiteSpace(c.ShortCompanyName))
            .ToList();
    }

    private static void ExtractCompaniesForValidation(object wbObj, List<string> errors)
    {
        _extractedCompanies = null;
        try
        {
            var companies = ExtractCompanies(wbObj);
            _extractedCompanies = companies;

            foreach (var dup in companies.Where(c => c.CompanyID > 0).GroupBy(c => c.CompanyID).Where(g => g.Count() > 1))
            {
                errors.Add($"  [{CompaniesAliases[0]}]: Duplicate Company ID '{dup.Key}' detected {dup.Count()} times in '{CompaniesTableName}'.");
            }

            foreach (var dup in companies
                .Where(c => !string.IsNullOrWhiteSpace(c.ShortCompanyName))
                .GroupBy(c => c.ShortCompanyName.Trim(), StringComparer.OrdinalIgnoreCase)
                .Where(g => g.Count() > 1))
            {
                errors.Add($"  [{CompaniesAliases[0]}]: Duplicate Short Company Name '{dup.Key}' detected {dup.Count()} times in '{CompaniesTableName}'.");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("[ProjectManager] Failed extracting CompaniesTable", ex);
            errors.Add($"Failed reading '{CompaniesTableName}': {ex.Message}");
        }
    }

    // =========================================================================
    // 2. POPULATION
    // =========================================================================
    private static void PopulateCompanies(object wbObj)
    {
        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, CompaniesAliases, CompaniesAliases[0]);
        if (ws == null) return;

        var companies = ServerSyncCommands.FetchCompanies();
        TableMapper.PopulateTableSafe<CompanyRecord>((object)ws, CompaniesTableName, companies);
    }
}
