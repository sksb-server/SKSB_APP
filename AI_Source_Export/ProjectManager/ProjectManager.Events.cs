namespace SKSB_App.ProjectManager;

using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Structure;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;
using SKSB_App.Core.DialogBoxes;
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

            // 1. Sync live database entities (Companies) to the 'Data' sheet
            SyncProjectLookupsData();

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

    public static void SyncProjectLookupsData()
    {
        // Pushes short/legal company names into 'Data' sheet -> Generates =CompanyNamesList in Name Manager
        DataToExcel.SyncData<CompanyRecord, string>("CompanyName", c =>
            !string.IsNullOrWhiteSpace(c.ShortCompanyName) ? c.ShortCompanyName : c.LegalCompanyName);
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
                bool saved = ServerSyncCommands.SaveProjectsSilent();

                if (saved)
                {
                    // Success: Clear any offline stash if it existed
                    OfflineSaveServices.ClearPendingOfflineSnapshot<List<ProjectRecord>>();
                    wb.Saved = true;
                    return;
                }

                // Failed: Show 3-way conflict dialog
                using var dialog = new ExitSyncConflict(
                    "Project Manager",
                    "Could not persist records to the PostgreSQL server (connection timeout or schema conflict)."
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        var (currentProjects, _) = SheetMapper.ExtractProjectRecords();
                        if (currentProjects != null && currentProjects.Count > 0)
                        {
                            OfflineSaveServices.SaveOfflineSnapshot(currentProjects);
                            MessageBox.Show(
                                "Your project records have been safely saved to local disk.\n\n" +
                                "The system will automatically detect this file and prompt to commit it to the server when you next start.",
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
            if (OfflineSaveServices.HasPendingOfflineSnapshot<List<ProjectRecord>>())
            {
                var pendingRecords = OfflineSaveServices.LoadPendingOfflineSnapshot<List<ProjectRecord>>();
                DateTime localOfflineTime = File.GetLastWriteTime(
                    Folder.Data.GetFilePath("List`1_PendingUpload.json")
                );

                DateTime? serverLastUpdate = OfflineSaveServices.GetServerLatestTimestamp("project_master");
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

                if (canCommit && pendingRecords != null && pendingRecords.Count > 0)
                {
                    var (success, count, _) = ServerSyncCommands.SaveProjectsAsync(pendingRecords)
                                                                .GetAwaiter()
                                                                .GetResult();
                    if (success)
                    {
                        OfflineSaveServices.ClearPendingOfflineSnapshot<List<ProjectRecord>>();
                        var companies = ServerSyncCommands.FetchCompanyRecords() ?? new List<CompanyRecord>();
                        SheetMapper.PopulateWorksheet(pendingRecords, companies);
                        WorkbookSession.ShowStatusConfirmation($"Committed {count} offline project records to server.");
                        return;
                    }
                }
                else
                {
                    OfflineSaveServices.ClearPendingOfflineSnapshot<List<ProjectRecord>>();
                }
            }

            // Standard operational load from PostgreSQL
            if (ServerSyncCommands.LoadProjectsSilent())
            {
                WorkbookSession.ShowStatusConfirmation("Project records loaded from server.");
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

        bool success = ServerSyncCommands.SaveProjectsSilent();
        if (success)
        {
            OfflineSaveServices.ClearPendingOfflineSnapshot<List<ProjectRecord>>();
            MessageBox.Show("Save completed successfully.", "SKSB Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}