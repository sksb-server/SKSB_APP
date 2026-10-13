namespace SKSB_App.ProjectManager;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    public static bool SaveProjectsSilent()
    {
        try
        {
            var (projectList, companyList) = SheetMapper.ExtractProjectRecords();
            if (projectList == null || projectList.Count == 0)
            {
                return false;
            }

            if (companyList != null && companyList.Count > 0)
            {
                Services.SaveAsync(companyList, writeToDatabase: true).GetAwaiter().GetResult();
            }

            string snapshotPath = Services.SaveAsync(projectList, writeToDatabase: true)
                                          .GetAwaiter()
                                          .GetResult();

            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool LoadProjectsSilent()
    {
        try
        {
            var projectList = Services.Load<List<ProjectRecord>>();
            var companyList = Services.Load<List<CompanyRecord>>() ?? new List<CompanyRecord>();

            if (projectList != null && projectList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(projectList, companyList);
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

    public static async Task<(bool Success, int Count, string Path)> SaveProjectsAsync(List<ProjectRecord> projectList)
    {
        if (projectList == null || projectList.Count == 0) return (false, 0, string.Empty);

        string path = await Services.SaveAsync(projectList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), projectList.Count, path);
    }

    public static async Task<bool> SaveCompaniesAsync(List<CompanyRecord> companyList)
    {
        if (companyList == null || companyList.Count == 0) return false;
        string path = await Services.SaveAsync(companyList, writeToDatabase: true);
        return !string.IsNullOrEmpty(path);
    }

    public static List<ProjectRecord>? FetchProjectRecords()
    {
        return Services.Load<List<ProjectRecord>>();
    }

    public static List<CompanyRecord>? FetchCompanyRecords()
    {
        return Services.Load<List<CompanyRecord>>();
    }
}