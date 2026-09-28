namespace SKSB_App.ProjectManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public static class ProjectSheetMapper
{
    private const int ColumnCount = 14;
    private static readonly string[] MasterSheetAliases = { "Project_Master", "Project Master", "Master" };

    public static List<ProjectRecord> ExtractProjectRecords()
    {
        var result = new List<ProjectRecord>();

        string masterSheet = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Project_Master");
        IntPtr masterSheetId = WorkbookSession.GetSheetId(masterSheet);

        var masterRef = masterSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, ColumnCount - 1, masterSheetId)
            : new ExcelReference(1, 1000, 0, ColumnCount - 1);

        if (masterRef.GetValue() is object[,] grid)
        {
            int rows = grid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = grid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long projectId) || projectId == 0)
                    continue;

                string projectCode = CleanCellString(grid[r, 1]);
                string classStr = CleanCellString(grid[r, 2]);
                string projectName = CleanCellString(grid[r, 3]);
                string priorityStr = CleanCellString(grid[r, 4]);
                string statusStr = CleanCellString(grid[r, 5]);

                DateOnly startDate = ParseExcelDate(grid[r, 6]);
                DateOnly targetEnd = ParseExcelDate(grid[r, 7]);
                DateOnly actualEnd = ParseExcelDate(grid[r, 8]);

                int.TryParse(CleanCellString(grid[r, 9]), out int headcount);
                string reqVehicleStr = CleanCellString(grid[r, 10]);
                string reqSkills = CleanCellString(grid[r, 11]);
                string address = CleanCellString(grid[r, 12]);
                string remarks = CleanCellString(grid[r, 13]);

                var project = new ProjectRecord
                {
                    ProjectID = projectId,
                    Master = new ProjectMaster
                    {
                        ProjectID = projectCode, // String ID/Code mapping
                        ProjectClass = Enum.TryParse(classStr, true, out ProjectClassification pClass) ? pClass : ProjectClassification.None,
                        ProjectName = projectName,
                        Priority = Enum.TryParse(priorityStr, true, out ProjectPriority priority) ? priority : ProjectPriority.Standard,
                        Status = Enum.TryParse(statusStr, true, out ProjectStatus status) ? status : ProjectStatus.Planned,
                        StartDate = startDate,
                        TargetEndDate = targetEnd,
                        ActualEndDate = actualEnd,
                        RequiredHeadcount = headcount > 0 ? headcount : 1,
                        RequiredVehicleType = Enum.TryParse(reqVehicleStr, true, out VehicleType vType) ? vType : VehicleType.None,
                        RequiredSkills = reqSkills,
                        Address = address,
                        Remarks = remarks
                    }
                };

                result.Add(project);
            }
        }
        return result;
    }

    public static void PopulateWorksheet(List<ProjectRecord> projectList)
    {
        if (projectList == null || projectList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Project_Master");
        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });

            var clearRef = new ExcelReference(1, 1000, 0, ColumnCount - 1);
            clearRef.SetValue(new object[1000, ColumnCount]);

            object[,] output = new object[projectList.Count, ColumnCount];
            for (int i = 0; i < projectList.Count; i++)
            {
                var p = projectList[i];
                output[i, 0] = p.ProjectID;
                output[i, 1] = p.Master.ProjectID;
                output[i, 2] = p.Master.ProjectClass == ProjectClassification.None ? string.Empty : p.Master.ProjectClass.ToString();
                output[i, 3] = p.Master.ProjectName;
                output[i, 4] = p.Master.Priority.ToString();
                output[i, 5] = p.Master.Status.ToString();
                output[i, 6] = p.Master.StartDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                output[i, 7] = p.Master.TargetEndDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                output[i, 8] = p.Master.ActualEndDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                output[i, 9] = p.Master.RequiredHeadcount;
                output[i, 10] = p.Master.RequiredVehicleType == VehicleType.None ? string.Empty : p.Master.RequiredVehicleType.ToString();
                output[i, 11] = p.Master.RequiredSkills;
                output[i, 12] = p.Master.Address;
                output[i, 13] = p.Master.Remarks;
            }

            var writeRef = new ExcelReference(1, projectList.Count, 0, ColumnCount - 1);
            writeRef.SetValue(output);
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

    private static string CleanCellString(object? cell)
    {
        if (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError) return string.Empty;
        return cell.ToString()?.Trim() ?? string.Empty;
    }

    private static DateOnly ParseExcelDate(object cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanCellString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return DateOnly.FromDateTime(DateTime.Today);
    }
}