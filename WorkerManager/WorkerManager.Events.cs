namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Constants;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

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

            // 1. Sync live database entities (e.g. Registered Vehicle Plates) to the 'Data' sheet
            SyncWorkerLookupsData();

            // 2. Startup reconciliation with conflict detection
            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadMacro");

            // 3. Hotkeys
            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeySave");
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeySave");
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

    public static void SyncWorkerLookupsData()
    {
        // Pushes Vehicle plates from database snapshot into 'Data' sheet -> Generates =VehiclePlateList in Name Manager
        DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate);
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
            System.Diagnostics.Debug.WriteLine($"[Events] Failed to hook WorkbookBeforeClose: {ex.Message}");
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
                bool saved = ServerSyncCommands.SaveWorkerSilent();

                if (saved)
                {
                    // Success: Clear any offline stash if it existed
                    Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                    wb.Saved = true;
                    return;
                }

                // Failed: Show 3-way conflict dialog
                using var dialog = new ExitSyncConflictDialog(
                    "Worker Manager",
                    "Could not persist records to the PostgreSQL server (connection timeout or schema conflict)."
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        // Loops back to top and retries SaveWorkerSilent()
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        // Extract current rows from Excel
                        var currentRecords = SheetMapper.ExtractWorkerRecords();
                        if (currentRecords != null && currentRecords.Count > 0)
                        {
                            Services.SaveOfflineSnapshot(currentRecords);
                            MessageBox.Show(
                                "Your worker records have been safely saved to local disk.\n\n" +
                                "The system will automatically detect this file and prompt to commit it to the server when you next start.",
                                "Saved Offline",
                                MessageBoxButtons.OK,
                                MessageBoxIcon.Information
                            );
                        }

                        // Mark clean so Excel closes without prompt
                        wb.Saved = true;
                        return;

                    case ExitResolution.Cancel:
                    default:
                        // Abort close sequence and keep workbook open
                        cancel = true;
                        return;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[OnWorkbookBeforeClose] Error: {ex.Message}");
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

            // -------------------------------------------------------------
            // STARTUP RECONCILIATION: Check for pending offline file
            // -------------------------------------------------------------
            if (Services.HasPendingOfflineSnapshot<List<WorkerRecord>>())
            {
                var pendingRecords = Services.LoadPendingOfflineSnapshot<List<WorkerRecord>>();
                DateTime localOfflineTime = File.GetLastWriteTime(
                    Folder.Data.GetFilePath("List`1_PendingUpload.json")
                );

                // Check server's latest timestamp
                DateTime? serverLastUpdate = Services.GetServerLatestTimestamp("worker_master");

                bool canCommit = true;

                if (serverLastUpdate.HasValue && serverLastUpdate.Value > localOfflineTime)
                {
                    // Conflict: Server has newer data than our offline exit file
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

                if (canCommit && pendingRecords != null && pendingRecords.Count > 0)
                {
                    // Commit offline records to the server
                    var (success, count, _) = ServerSyncCommands.SaveWorkerAsync(pendingRecords)
                                                                .GetAwaiter()
                                                                .GetResult();
                    if (success)
                    {
                        Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                        SheetMapper.PopulateWorksheet(pendingRecords);
                        WorkbookSession.ShowStatusConfirmation($"Committed {count} offline worker records to server.");
                        return;
                    }
                }
                else
                {
                    // User chose 'No': Clear pending offline file and fall through to regular server load
                    Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                }
            }

            // Standard operational load from PostgreSQL
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

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave()
    {
        if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
        {
            XlCall.Excel(XlCall.xlcSave);
            return;
        }

        bool success = ServerSyncCommands.SaveWorkerSilent();
        if (success)
        {
            Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
            MessageBox.Show("Save completed successfully.", "SKSB Worker Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}