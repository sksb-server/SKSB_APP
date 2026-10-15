namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    public static bool SaveWorkerSilent(out string errorMessage)
    {
        errorMessage = string.Empty;
        try
        {
            var (workerList, errors, _) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

            // Critical errors strictly block saving
            if (errors.Count > 0)
            {
                errorMessage = $"{errors.Count} critical error(s) must be resolved.";
                return false;
            }

            if (workerList == null || workerList.Count == 0)
            {
                return true;
            }

            string snapshotPath = Services.Save(workerList, writeToDatabase: true);
            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            System.Diagnostics.Debug.WriteLine($"[ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool LoadWorkerSilent()
    {
        try
        {
            var workerList = Services.Load<List<WorkerRecord>>();
            if (workerList != null && workerList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(workerList);
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

    public static async Task<(bool Success, int Count, string Path)> SaveWorkerAsync(List<WorkerRecord> workerList)
    {
        if (workerList == null || workerList.Count == 0) return (false, 0, string.Empty);

        string path = await Services.SaveAsync(workerList, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), workerList.Count, path);
    }

    public static Task<(bool Success, int Count, string Path)> SaveWorkersAsync(List<WorkerRecord> workerList)
        => SaveWorkerAsync(workerList);

    public static List<WorkerRecord>? FetchWorkerRecords()
    {
        return Services.Load<List<WorkerRecord>>();
    }
}