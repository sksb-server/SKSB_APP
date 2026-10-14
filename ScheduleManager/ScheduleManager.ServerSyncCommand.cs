namespace SKSB_App.ScheduleManager;

using System;
using System.Threading.Tasks;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;

public static class ServerSyncCommands
{
    public static bool SaveScheduleSilent(DateOnly scheduleDate)
    {
        try
        {
            var schedule = SheetMapper.ExtractScheduleRecord(scheduleDate);
            if (schedule == null)
            {
                return false;
            }

            string snapshotPath = Services.SaveAsync(schedule, writeToDatabase: true)
                                          .GetAwaiter()
                                          .GetResult();

            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ScheduleManager.ServerSyncCommands] Silent Save Error: {ex.Message}");
            return false;
        }
    }

    public static bool LoadScheduleSilent(DateOnly scheduleDate)
    {
        try
        {
            var schedule = Services.Load<ScheduleRecord>();
            if (schedule != null && schedule.DailyScheduleDate == scheduleDate)
            {
                SheetMapper.PopulateScheduleSheets(schedule);
                return true;
            }
            return false;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ScheduleManager.ServerSyncCommands] Silent Load Error: {ex.Message}");
            return false;
        }
    }

    public static async Task<(bool Success, string Path)> SaveScheduleAsync(ScheduleRecord schedule)
    {
        if (schedule == null) return (false, string.Empty);

        string path = await Services.SaveAsync(schedule, writeToDatabase: true);
        return (!string.IsNullOrEmpty(path), path);
    }

    public static ScheduleRecord? FetchSchedule(DateOnly scheduleDate)
    {
        var record = Services.Load<ScheduleRecord>();
        return record?.DailyScheduleDate == scheduleDate ? record : null;
    }
}