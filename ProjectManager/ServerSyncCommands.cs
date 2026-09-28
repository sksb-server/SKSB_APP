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
            var projectList = ProjectSheetMapper.ExtractProjectRecords();
            if (projectList == null || projectList.Count == 0) return true;

            string snapshotPath = Services.SaveAsync(projectList, writeToDatabase: true).GetAwaiter().GetResult();
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
            if (projectList != null && projectList.Count > 0)
            {
                ProjectSheetMapper.PopulateWorksheet(projectList);
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

    // Explicitly accepts List<ProjectRecord>
    public static async Task<(bool Success, int Count, string Path)> SaveProjectsAsync(List<ProjectRecord> projectList)
    {
        if (projectList == null || projectList.Count == 0) return (false, 0, string.Empty);

        string path = await Services.SaveAsync(projectList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), projectList.Count, path);
    }

    public static List<ProjectRecord>? FetchProjectRecords()
    {
        return Services.Load<List<ProjectRecord>>();
    }
}