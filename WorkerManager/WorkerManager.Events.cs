namespace SKSB_App.WorkerManager;

using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Constants;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;
    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private static WorkbookBeforeCloseDelegate? _beforeCloseHandler;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookApplicationEvents();

            SyncWorkerLookupsData();

            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadMacro");

            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeySave");
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeySave");
        });
    }

    public void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            UnhookApplicationEvents();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    public static void SyncWorkerLookupsData()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SyncWorkerLookupsData Error] {ex.Message}");
            }
        });
    }

    private static void HookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            _xlAppRoot = xlApp;

            _beforeCloseHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            _xlAppRoot.WorkbookBeforeClose += _beforeCloseHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Events] Failed to hook WorkbookBeforeClose: {ex.Message}");
        }
    }

    private static void UnhookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null && _beforeCloseHandler != null)
            {
                _xlAppRoot.WorkbookBeforeClose -= _beforeCloseHandler;
                _beforeCloseHandler = null;
                _xlAppRoot = null;
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
                var (workerList, validationErrors, validationWarnings) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

                // 1. If errors exist or warnings need review
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new ValidationSummaryDialog("Worker Manager", validationErrors, validationWarnings);
                    valDialog.ShowDialog();

                    // If blocking errors exist, or operator canceled: keep Excel open
                    if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        cancel = true;
                        return;
                    }
                }

                // Empty template: allow clean exit
                if (workerList == null || workerList.Count == 0)
                {
                    wb.Saved = true;
                    return;
                }

                // 2. Attempt Silent Save
                if (ServerSyncCommands.SaveWorkerSilent(out string dbError))
                {
                    Services.ClearPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();

                    // Populate sheet before closing so Excel's file buffer contains the generated values
                    if (workerList != null && workerList.Count > 0)
                    {
                        SheetMapper.PopulateWorksheet(workerList);
                    }
                    wb.Saved = true;

                    SaveNotificationDialog.ShowSuccess(
                        "Worker Manager",
                        $"Successfully saved {workerList?.Count ?? 0} worker record(s) to the server database before closing."
                    );
                    return;
                }

                // 3. Database Offline / Failed
                using var dialog = new ExitSyncConflictDialog(
                    "Worker Manager",
                    string.IsNullOrWhiteSpace(dbError) ? "Could not persist records to the PostgreSQL server." : dbError,
                    allowSaveOffline: true
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        Services.SaveOfflineSnapshot(workerList);
                        SaveNotificationDialog.ShowOffline(
                            "Worker Manager",
                            "Your worker records have been safely saved to local disk.\n\n" +
                            "The system will automatically verify and commit them to the server when you next start."
                        );

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
            System.Diagnostics.Debug.WriteLine($"[OnWorkbookBeforeClose] Error: {ex.Message}");
            cancel = true;
        }
    }

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave()
    {
        if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
        {
            XlCall.Excel(XlCall.xlcSave);
            return;
        }

        var (workerList, validationErrors, validationWarnings) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

        if (validationErrors.Count > 0 || validationWarnings.Count > 0)
        {
            using var valDialog = new ValidationSummaryDialog("Worker Manager", validationErrors, validationWarnings);
            valDialog.ShowDialog();

            if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
            {
                return;
            }
        }

        if (ServerSyncCommands.SaveWorkerSilent(out string dbErr))
        {
            Services.ClearPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();

            // Immediately reflect applied defaults and IDs onto Excel tables
            if (workerList != null && workerList.Count > 0)
            {
                SheetMapper.PopulateWorksheet(workerList);
            }

            int count = workerList?.Count ?? 0;
            SaveNotificationDialog.ShowSuccess(
                "Worker Manager",
                $"Save completed successfully. {count} record(s) persisted to the database.\n\nAll default dates, statuses, and auto-generated IDs have been refreshed on your sheets."
            );
        }
        else
        {
            SaveNotificationDialog.ShowError(
                "Worker Manager",
                $"Failed to persist records to the server database:\n\n{dbErr}"
            );
        }
    }

    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro()
    {
        try
        {
            if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
            {
                WorkbookSession.ShowStatusConfirmation("Design Mode: Auto-load bypassed.");
                return;
            }

            if (Services.HasPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>())
            {
                var pendingRecords = Services.LoadPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();

                bool isValidPayload = pendingRecords != null &&
                                      pendingRecords.Count > 0 &&
                                      pendingRecords.TrueForAll(w => w.WorkerID > 0 && !string.IsNullOrWhiteSpace(w.WorkerDetails.WorkerName));

                if (!isValidPayload)
                {
                    MessageBox.Show(
                        "Found a corrupt offline pending file. The upload has been aborted to protect the database.",
                        "Upload Rejected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    Services.ClearPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();
                }
                else
                {
                    DateTime localOfflineTime = File.GetLastWriteTime(
                        Folder.Data.GetFilePath("List`1_PendingUpload.json")
                    );

                    DateTime? serverLastUpdate = Services.GetServerLatestTimestamp("worker_master");

                    bool canCommit = true;

                    if (serverLastUpdate.HasValue && serverLastUpdate.Value > localOfflineTime)
                    {
                        var conflictChoice = MessageBox.Show(
                            $"[Synchronization Conflict Detected]\n\n" +
                            $"• Your offline saved file: {localOfflineTime:yyyy-MM-dd HH:mm:ss}\n" +
                            $"• Newer server version:  {serverLastUpdate.Value:yyyy-MM-dd HH:mm:ss}\n\n" +
                            "The server was updated after your offline save.\n\n" +
                            "Do you still want to COMMIT your offline changes to the server?\n" +
                            "• Yes: Overwrite server records with your local offline version.\n" +
                            "• No: Discard local offline changes and load the latest server copy.",
                            "Data Version Conflict",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        canCommit = (conflictChoice == DialogResult.Yes);
                    }

                    if (canCommit)
                    {
                        var (success, count, _) = ServerSyncCommands.SaveWorkerAsync(pendingRecords!)
                                                                    .GetAwaiter()
                                                                    .GetResult();
                        if (success)
                        {
                            Services.ClearPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();
                            SheetMapper.PopulateWorksheet(pendingRecords!);
                            WorkbookSession.ShowStatusConfirmation($"Committed {count} offline worker records to server.");
                            return;
                        }
                    }
                    else
                    {
                        Services.ClearPendingOfflineSnapshot<System.Collections.Generic.List<WorkerRecord>>();
                    }
                }
            }

            if (ServerSyncCommands.LoadWorkerSilent())
            {
                WorkbookSession.ShowStatusConfirmation("Worker records loaded from server.");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DelayedAutoLoad] Error: {ex.Message}");
        }
    }
}