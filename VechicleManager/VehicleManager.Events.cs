namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Structure;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;
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

            // 1. Live server data lookups
            SyncVehicleLookupsData();

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

    public static void SyncVehicleLookupsData()
    {
        DataToExcel.SyncData<WorkerRecord, string>("WorkerName", s => s.WorkerDetails.WorkerName);
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
        catch { }
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
                bool saved = ServerSyncCommands.SaveVehiclesSilent();

                if (saved)
                {
                    // Success: Clear any offline stash if it existed
                    OfflineSaveServices.ClearPendingOfflineSnapshot<List<VehicleRecord>>();
                    wb.Saved = true;
                    return;
                }

                // Failed: Show 3-way conflict dialog
                using var dialog = new ExitSyncConflict(
                    "Vehicle Manager",
                    "Could not persist records to the PostgreSQL server (connection timeout or schema conflict)."
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        // Loops back to top and retries SaveVehiclesSilent()
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        // Extract current rows from Excel
                        var currentRecords = SheetMapper.ExtractVehicleRecords();
                        if (currentRecords != null && currentRecords.Count > 0)
                        {
                            OfflineSaveServices.SaveOfflineSnapshot(currentRecords);
                            MessageBox.Show(
                                "Your records have been safely saved to local disk.\n\n" +
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
            if (OfflineSaveServices.HasPendingOfflineSnapshot<List<VehicleRecord>>())
            {
                var pendingRecords = OfflineSaveServices.LoadPendingOfflineSnapshot<List<VehicleRecord>>();
                DateTime localOfflineTime = File.GetLastWriteTime(
                    Folder.Data.GetFilePath("List`1_PendingUpload.json")
                );

                // Check server's latest timestamp
                DateTime? serverLastUpdate = OfflineSaveServices.GetServerLatestTimestamp("vehicle_master");

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
                    var (success, count, _) = ServerSyncCommands.SaveVehiclesAsync(pendingRecords)
                                                                 .GetAwaiter()
                                                                 .GetResult();
                    if (success)
                    {
                        OfflineSaveServices.ClearPendingOfflineSnapshot<List<VehicleRecord>>();
                        SheetMapper.PopulateWorksheet(pendingRecords);
                        WorkbookSession.ShowStatusConfirmation($"Committed {count} offline records to server.");
                        return;
                    }
                }
                else
                {
                    // User chose 'No': Clear pending offline file and fall through to regular server load
                    OfflineSaveServices.ClearPendingOfflineSnapshot<List<VehicleRecord>>();
                }
            }

            // Standard operational load from PostgreSQL
            if (ServerSyncCommands.LoadVehiclesSilent())
            {
                WorkbookSession.ShowStatusConfirmation("Vehicle records loaded from server.");
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

        bool success = ServerSyncCommands.SaveVehiclesSilent();
        if (success)
        {
            OfflineSaveServices.ClearPendingOfflineSnapshot<List<VehicleRecord>>();
            MessageBox.Show("Save completed successfully.", "SKSB Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}