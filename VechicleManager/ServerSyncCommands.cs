namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    public static bool SaveVehiclesSilent()
    {
        try
        {
            var vehicleList = VehicleSheetMapper.ExtractVehicleRecords();
            if (vehicleList == null || vehicleList.Count == 0) return true;

            string snapshotPath = Services.SaveAsync(vehicleList, writeToDatabase: true).GetAwaiter().GetResult();
            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool LoadVehiclesSilent()
    {
        try
        {
            var vehicleList = Services.Load<List<Vehicle>>();
            if (vehicleList != null && vehicleList.Count > 0)
            {
                VehicleSheetMapper.PopulateWorksheet(vehicleList);
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
    /// Persists pre-extracted vehicle records to disk and PostgreSQL on a background worker thread.
    /// </summary>
    public static async Task<(bool Success, int Count, string Path)> SaveVehiclesAsync(List<Vehicle> vehicleList)
    {
        if (vehicleList == null || vehicleList.Count == 0) return (false, 0, string.Empty);

        string path = await Services.SaveAsync(vehicleList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), vehicleList.Count, path);
    }

    public static List<Vehicle>? FetchVehicleRecords()
    {
        return Services.Load<List<Vehicle>>();
    }
}