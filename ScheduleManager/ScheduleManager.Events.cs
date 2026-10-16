namespace SKSB_App.ScheduleManager;

using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Structure;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;

public class Events : IExcelAddIn
{
    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private static dynamic? _appEvents;
    private static WorkbookBeforeCloseDelegate? _closeHandler;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookApplicationClose();

            // 1. Sync Core static enums into 'Lookups' tab
            TypesToExcel.Sync(
                typeof(Shifts),
                typeof(Days),
                typeof(AttendanceStatus),
                typeof(BiometricStatus),
                typeof(VehicleType)
            );

            // 2. Sync dynamic master entities (Worker, Vehicle, Project) into 'Data' tab
            SyncScheduleLookupsData();

            // 3. Delayed startup macro for auto-load & offline reconciliation
            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadScheduleMacro");

            // 4. Hotkeys (Ctrl+S and Ctrl+Shift+S)
            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeyScheduleSave");
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeyScheduleSave");
        });
    }

    public void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            UnhookApplicationClose();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    public static void SyncScheduleLookupsData()
    {
        // Spawns 'Data' sheet and registers defined names
        DataToExcel.SyncData<WorkerRecord, long>("WorkerID", s => s.WorkerID);
        DataToExcel.SyncData<WorkerRecord, string>("WorkerName", s => s.WorkerDetails.WorkerName);
        DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate);
        DataToExcel.SyncData<ProjectRecord, string>("ProjectID", p => p.ProjectID);
    }

    private static void HookApplicationClose()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            _appEvents = xlApp;
            _closeHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            xlApp.WorkbookBeforeClose += _closeHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ScheduleManager.Events] Hook error: {ex.Message}");
        }
    }

    private static void UnhookApplicationClose()
    {
        try
        {
            if (_appEvents != null && _closeHandler != null)
            {
                _appEvents.WorkbookBeforeClose -= _closeHandler;
                _closeHandler = null;
                _appEvents = null;
            }
        }
        catch { }
    }

    private static void OnWorkbookBeforeClose(dynamic wb, ref bool cancel)
    {
        try
        {
            if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode) return;

            while (true)
            {
                // Attempt standard silent server save
                bool saved = ServerSyncCommands.SaveScheduleSilent(RibbonController.CurrentScheduleDate);

                if (saved)
                {
                    OfflineSaveServices.ClearPendingOfflineSnapshot<ScheduleRecord>();
                    wb.Saved = true;
                    return;
                }

                // Failed: Show 3-way conflict dialog
                using var dialog = new ExitSyncConflict(
                    "Schedule Manager",
                    $"Could not persist the schedule for {RibbonController.CurrentScheduleDate:yyyy-MM-dd} to the PostgreSQL server."
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        var currentSchedule = SheetMapper.ExtractScheduleRecord(RibbonController.CurrentScheduleDate);
                        if (currentSchedule != null)
                        {
                            OfflineSaveServices.SaveOfflineSnapshot(currentSchedule);
                            MessageBox.Show(
                                $"Schedule for {RibbonController.CurrentScheduleDate:yyyy-MM-dd} saved locally to offline staging.",
                                "Saved Offline",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                        }
                        wb.Saved = true;
                        return;

                    case ExitResolution.Cancel:
                    default:
                        cancel = true;
                        return;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ScheduleManager.Events] OnWorkbookBeforeClose error: {ex.Message}");
        }
    }

    [ExcelCommand(Name = "DelayedAutoLoadScheduleMacro")]
    public static void DelayedAutoLoadScheduleMacro()
    {
        try
        {
            if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
            {
                WorkbookSession.ShowStatusConfirmation("Design Mode: Auto-load bypassed.");
                return;
            }

            // -------------------------------------------------------------
            // STARTUP RECONCILIATION: Check for pending offline file
            // -------------------------------------------------------------
            if (OfflineSaveServices.HasPendingOfflineSnapshot<ScheduleRecord>())
            {
                var pendingSchedule = OfflineSaveServices.LoadPendingOfflineSnapshot<ScheduleRecord>();
                DateTime localOfflineTime = File.GetLastWriteTime(
                    Folder.Data.GetFilePath("ScheduleRecord_PendingUpload.json")
                );

                DateTime? serverLastUpdate = OfflineSaveServices.GetServerLatestTimestamp("daily_schedule");
                bool canCommit = true;

                if (serverLastUpdate.HasValue && serverLastUpdate.Value > localOfflineTime)
                {
                    var conflictChoice = MessageBox.Show(
                        $"[Synchronization Conflict Detected]\n\n" +
                        $"• Your offline saved schedule: {localOfflineTime:yyyy-MM-dd HH:mm:ss}\n" +
                        $"• Newer server version:         {serverLastUpdate.Value:yyyy-MM-dd HH:mm:ss}\n\n" +
                        "Do you still want to COMMIT your offline changes to the server?\n" +
                        "• Yes: Overwrite server schedule with your local version.\n" +
                        "• No: Discard local changes and load latest server copy.",
                        "Data Version Conflict",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );
                    canCommit = (conflictChoice == DialogResult.Yes);
                }

                if (canCommit && pendingSchedule != null)
                {
                    var (success, _) = ServerSyncCommands.SaveScheduleAsync(pendingSchedule)
                                                         .GetAwaiter()
                                                         .GetResult();
                    if (success)
                    {
                        OfflineSaveServices.ClearPendingOfflineSnapshot<ScheduleRecord>();
                        RibbonController.SetScheduleDate(pendingSchedule.DailyScheduleDate);
                        SheetMapper.PopulateScheduleSheets(pendingSchedule);
                        WorkbookSession.ShowStatusConfirmation($"Committed offline schedule for {pendingSchedule.DailyScheduleDate} to server.");
                        return;
                    }
                }
                else
                {
                    OfflineSaveServices.ClearPendingOfflineSnapshot<ScheduleRecord>();
                }
            }

            // Standard operational load for the current date
            if (ServerSyncCommands.LoadScheduleSilent(RibbonController.CurrentScheduleDate))
            {
                WorkbookSession.ShowStatusConfirmation($"Schedule for {RibbonController.CurrentScheduleDate:yyyy-MM-dd} loaded from server.");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ScheduleManager.Events] DelayedAutoLoad error: {ex.Message}");
        }
    }

    [ExcelCommand(Name = "OnHotKeyScheduleSave")]
    public static void OnHotKeyScheduleSave()
    {
        if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
        {
            XlCall.Excel(XlCall.xlcSave);
            return;
        }

        bool success = ServerSyncCommands.SaveScheduleSilent(RibbonController.CurrentScheduleDate);
        if (success)
        {
            OfflineSaveServices.ClearPendingOfflineSnapshot<ScheduleRecord>();
            MessageBox.Show(
                $"Schedule for {RibbonController.CurrentScheduleDate:yyyy-MM-dd} saved successfully.",
                "SKSB Schedule Manager",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
    }
}