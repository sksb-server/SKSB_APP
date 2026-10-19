namespace SKSB_App.ProjectManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;

public static class SheetMapper
{
    private const int ProjectColumnSpan = 6;  // Columns A to F
    private const int CompanyColumnSpan = 3;  // Columns A to C

    private static readonly string[] AllProjectSheetAliases = { "All Projects", "All_Projects", "Project_Master", "Projects" };
    private static readonly string[] CompanySheetAliases = { "Companies", "Company", "Company_Master" };

    public static (List<ProjectRecord> Projects, List<CompanyRecord> Companies) ExtractProjectRecords()
    {
        var projectResult = new List<ProjectRecord>();
        var companyResult = new List<CompanyRecord>();
        var validationErrors = new List<string>();

        // -------------------------------------------------------------
        // 1. EXTRACT COMPANIES (Sheet: Companies, Columns A to C)
        // -------------------------------------------------------------
        string compSheet = WorkbookSession.ResolveSheetName(CompanySheetAliases, "Companies");
        IntPtr compSheetId = WorkbookSession.GetSheetId(compSheet);

        var compRef = compSheetId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, CompanyColumnSpan - 1, compSheetId)
            : new ExcelReference(1, 500, 0, CompanyColumnSpan - 1);

        if (compRef.GetValue() is object[,] compGrid)
        {
            int rows = compGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = compGrid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long compId) || compId <= 0)
                    continue;

                string legalName = CleanCellString(compGrid[r, 1]);
                string shortName = CleanCellString(compGrid[r, 2]);

                if (string.IsNullOrWhiteSpace(legalName) && string.IsNullOrWhiteSpace(shortName))
                    continue;

                companyResult.Add(new CompanyRecord
                {
                    CompanyID = compId,
                    LegalCompanyName = legalName,
                    ShortCompanyName = shortName
                });
            }
        }

        // -------------------------------------------------------------
        // 2. EXTRACT PROJECTS (Sheet: All Projects, Columns A to F)
        // -------------------------------------------------------------
        string projSheet = WorkbookSession.ResolveSheetName(AllProjectSheetAliases, "All Projects");
        IntPtr projSheetId = WorkbookSession.GetSheetId(projSheet);

        var projRef = projSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, ProjectColumnSpan - 1, projSheetId)
            : new ExcelReference(1, 1000, 0, ProjectColumnSpan - 1);

        if (projRef.GetValue() is object[,] projGrid)
        {
            int rows = projGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                // Col A (0): Project Type
                string typeStr = CleanCellString(projGrid[r, 0]);
                // Col B (1): Project ID (e.g. PRJ-001 or IJ-102)
                string projectId = CleanCellString(projGrid[r, 1]);
                if (string.IsNullOrWhiteSpace(projectId))
                    continue;

                // Col C (2): Project Name
                string projectName = CleanCellString(projGrid[r, 2]);
                // Col D (3): Company
                string companyName = CleanCellString(projGrid[r, 3]);
                // Col E (4): Start Date
                DateOnly startDate = ParseExcelDate(projGrid[r, 4]);
                // Col F (5): End Date
                DateOnly endDate = ParseExcelDate(projGrid[r, 5]);

                if (string.IsNullOrWhiteSpace(projectName))
                {
                    validationErrors.Add($"• Row {r + 2} [All Projects]: Project '{projectId}' is missing a Project Name.");
                    continue;
                }

                ProjectType pType = Enum.TryParse(typeStr, true, out ProjectType parsedType)
                    ? parsedType
                    : ProjectType.None;

                if (startDate == default) startDate = DateOnly.FromDateTime(DateTime.Today);
                if (endDate == default) endDate = startDate.AddDays(30);

                var record = new ProjectRecord
                {
                    ProjectID = projectId,
                    ProjectDetails = new ProjectDetails
                    {
                        ProjectID = projectId,
                        ProjectType = pType,
                        ProjectName = projectName,
                        CompanyName = companyName,
                        StartDate = startDate,
                        TargetEndDate = endDate,
                        Status = ProjectStatus.Active
                    }
                };

                projectResult.Add(record);
            }
        }

        // -------------------------------------------------------------
        // PRE-SAVE VALIDATION GATEWAY: Prevent saving invalid rows
        // -------------------------------------------------------------
        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save to system due to validation errors:\n");
            foreach (var err in validationErrors) sb.AppendLine(err);
            sb.AppendLine("\nPlease correct the entries and try saving again.");

            MessageBox.Show(sb.ToString(), "Pre-Save Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return (new List<ProjectRecord>(), new List<CompanyRecord>());
        }

        return (projectResult, companyResult);
    }

    public static void PopulateWorksheet(List<ProjectRecord> projectList, List<CompanyRecord> companyList)
    {
        string targetProj = WorkbookSession.ResolveSheetName(AllProjectSheetAliases, "All Projects");
        string targetComp = WorkbookSession.ResolveSheetName(CompanySheetAliases, "Companies");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            // 1. POPULATE ALL PROJECTS (Columns A to F)
            if (projectList != null && projectList.Count > 0)
            {
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetProj });
                var clearProjRef = new ExcelReference(1, 1000, 0, ProjectColumnSpan - 1);
                clearProjRef.SetValue(new object[1000, ProjectColumnSpan]);

                int pCount = projectList.Count;
                object[,] projOutput = new object[pCount, ProjectColumnSpan];

                for (int i = 0; i < pCount; i++)
                {
                    var p = projectList[i];
                    projOutput[i, 0] = p.ProjectDetails.ProjectType == ProjectType.None ? string.Empty : p.ProjectDetails.ProjectType.ToString();
                    projOutput[i, 1] = p.ProjectID;
                    projOutput[i, 2] = p.ProjectDetails.ProjectName;
                    projOutput[i, 3] = p.ProjectDetails.CompanyName;
                    projOutput[i, 4] = p.ProjectDetails.StartDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    projOutput[i, 5] = p.ProjectDetails.TargetEndDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                }

                var writeProjRef = new ExcelReference(1, pCount, 0, ProjectColumnSpan - 1);
                writeProjRef.SetValue(projOutput);
            }

            // 2. POPULATE COMPANIES (Columns A to C)
            if (companyList != null && companyList.Count > 0)
            {
                XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetComp });
                var clearCompRef = new ExcelReference(1, 500, 0, CompanyColumnSpan - 1);
                clearCompRef.SetValue(new object[500, CompanyColumnSpan]);

                int cCount = companyList.Count;
                object[,] compOutput = new object[cCount, CompanyColumnSpan];

                for (int i = 0; i < cCount; i++)
                {
                    var c = companyList[i];
                    compOutput[i, 0] = c.CompanyID;
                    compOutput[i, 1] = c.LegalCompanyName;
                    compOutput[i, 2] = c.ShortCompanyName;
                }

                var writeCompRef = new ExcelReference(1, cCount, 0, CompanyColumnSpan - 1);
                writeCompRef.SetValue(compOutput);
            }
        }
        finally
        {
            if (initialActiveSheet != null)
            {
                string sheetStr = initialActiveSheet.ToString() ?? string.Empty;
                int bIdx = sheetStr.LastIndexOf(']');
                string restoreName = bIdx >= 0 ? sheetStr.Substring(bIdx + 1) : sheetStr;

                if (!string.IsNullOrEmpty(restoreName))
                {
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { restoreName }); } catch { }
                }
            }

            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }

    private static string CleanCellString(object? cell) =>
        (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            ? string.Empty
            : cell.ToString()?.Trim() ?? string.Empty;

    private static DateOnly ParseExcelDate(object? cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanCellString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return default;
    }
}