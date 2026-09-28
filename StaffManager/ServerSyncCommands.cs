namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    /// <summary>
    /// Synchronously extracts sheet data and saves to backend/database without UI notifications.
    /// </summary>
    public static bool SaveStaffSilent()
    {
        try
        {
            var staffList = StaffSheetMapper.ExtractStaffRecords();
            // If validation failed, ExtractStaffRecords() returns an empty list
            if (staffList == null || staffList.Count == 0)
            {
                return false; // Prevents marking workbook as pristine when errors exist
            }

            string snapshotPath = Services.SaveAsync(staffList, writeToDatabase: true)
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

    /// <summary>
    /// Loads staff records from storage and writes them directly into the sheet.
    /// </summary>
    public static bool LoadStaffSilent()
    {
        try
        {
            var staffList = Services.Load<List<Staff>>();
            if (staffList != null && staffList.Count > 0)
            {
                StaffSheetMapper.PopulateWorksheet(staffList);
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

    /// <summary>
    /// Persists pre-extracted staff records to disk and PostgreSQL on a background worker thread.
    /// </summary>
    public static async Task<(bool Success, int Count, string Path)> SaveStaffAsync(List<Staff> staffList)
    {
        if (staffList == null || staffList.Count == 0)
        {
            return (false, 0, string.Empty);
        }

        string path = await Services.SaveAsync(staffList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), staffList.Count, path);
    }

    public static List<Staff>? FetchStaffRecords()
    {
        return Services.Load<List<Staff>>();
    }
}