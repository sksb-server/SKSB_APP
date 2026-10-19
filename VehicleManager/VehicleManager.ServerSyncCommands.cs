namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    public static bool SaveVehiclesSilent()
    {
        try
        {
            var vehicleList = SheetMapper.ExtractVehicleRecords();
            // Critical: If extraction failed or errors exist, do NOT report success
            if (vehicleList == null || vehicleList.Count == 0)
            {
                return false;
            }

            string snapshotPath = OfflineSaveServices.SaveAsync(vehicleList, writeToDatabase: true)
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

    public static bool LoadVehiclesSilent()
    {
        try
        {
            var vehicleList = OfflineSaveServices.Load<List<VehicleRecord>>();
            if (vehicleList != null && vehicleList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(vehicleList);
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

    public static async Task<(bool Success, int Count, string Path)> SaveVehiclesAsync(List<VehicleRecord> vehicleList)
    {
        if (vehicleList == null || vehicleList.Count == 0) return (false, 0, string.Empty);

        string path = await OfflineSaveServices.SaveAsync(vehicleList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), vehicleList.Count, path);
    }

    public static List<VehicleRecord>? FetchVehicleRecords()
    {
        return OfflineSaveServices.Load<List<VehicleRecord>>();
    }
}