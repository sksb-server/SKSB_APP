namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Constants;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.TemplateEditor;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;

    // Delegates for COM Events
    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private delegate void WorkbookBeforeSaveDelegate(dynamic wb, bool saveAsUi, ref bool cancel);

    private static WorkbookBeforeCloseDelegate? _beforeCloseHandler;
    private static WorkbookBeforeSaveDelegate? _beforeSaveHandler;

    private static bool _isExiting = false;
    private static bool _isInternalSaving = false;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookApplicationEvents();
            SyncWorkerLookupsData();

            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadMacro");

            // Keyboard Hijack for Ctrl+S and Ctrl+Shift+S
            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeySave");
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeySave");
        });
    }

    public void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SKSB_App.Core.DropDown.Arrow.Teardown();
            UnhookApplicationEvents();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    private static void HookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;
            dynamic xlApp = ExcelDnaUtil.Application;
            _xlAppRoot = xlApp;

            // 1. Hook Close
            _beforeCloseHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            _xlAppRoot.WorkbookBeforeClose += _beforeCloseHandler;

            // 2. Hook Save (Catches QAT Save Icon, File > Save, and in-cell editing Ctrl+S)
            _beforeSaveHandler = new WorkbookBeforeSaveDelegate(OnWorkbookBeforeSave);
            _xlAppRoot.WorkbookBeforeSave += _beforeSaveHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Events] Failed to hook application events: {ex.Message}");
        }
    }

    private static void UnhookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null)
            {
                if (_beforeCloseHandler != null)
                {
                    _xlAppRoot.WorkbookBeforeClose -= _beforeCloseHandler;
                    _beforeCloseHandler = null;
                }
                if (_beforeSaveHandler != null)
                {
                    _xlAppRoot.WorkbookBeforeSave -= _beforeSaveHandler;
                    _beforeSaveHandler = null;
                }
                _xlAppRoot = null;
            }
        }
        catch { }
    }

    public static void SyncWorkerLookupsData(bool isInteractive = false)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            var summary = new StringBuilder();

            try
            {
                // 1. Vehicle plates -> 'Data' sheet via DataToExcel
                DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate);
                summary.AppendLine("✔ Vehicle plates synchronized to 'Data' sheet.");
            }
            catch (Exception exV)
            {
                summary.AppendLine($"✖ Vehicle plates sync failed: {exV.Message}");
            }

            // 2. Skill License Types
            var (success, report) = SheetMapper.PopulateSkillLicenseTypesTable(showDialogIfFailed: isInteractive);
            summary.AppendLine(success ? "✔ Skill License Types populated into table." : "✖ Skill License Types failed to populate.");

            if (isInteractive)
            {
                MessageBox.Show(
                    $"{summary}\n\n{report}",
                    "Sync Lookups Diagnostic Results",
                    MessageBoxButtons.OK,
                    success ? MessageBoxIcon.Information : MessageBoxIcon.Warning
                );
            }
        });
    }

    private static void SyncSkillLicenseTypes()
    {
        var distinctTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // Fetch from PostgreSQL
        try
        {
            var db = new PostgreSqlService();
            const string sql = "SELECT DISTINCT license_type FROM worker_skill_licenses WHERE license_type IS NOT NULL AND TRIM(license_type) != '';";
            var dbTypes = db.Query<string>(sql);
            if (dbTypes != null)
            {
                foreach (var t in dbTypes)
                {
                    if (!string.IsNullOrWhiteSpace(t)) distinctTypes.Add(t.Trim());
                }
            }
        }
        catch { }

        // Fallback to local snapshot
        if (distinctTypes.Count == 0)
        {
            var workers = ServerSyncCommands.FetchWorkerRecords();
            if (workers != null)
            {
                foreach (var w in workers)
                {
                    if (w.SkillLicenses != null)
                    {
                        foreach (var sl in w.SkillLicenses)
                        {
                            if (!string.IsNullOrWhiteSpace(sl.LicenseType))
                                distinctTypes.Add(sl.LicenseType.Trim());
                        }
                    }
                }
            }
        }

        if (distinctTypes.Count == 0) return;

        string[] sortedTypes = distinctTypes.OrderBy(x => x).ToArray();

        // 1. Sync to hidden 'Data' sheet and update =LicenseTypesList
        DataToExcel.SyncValues("LicenseTypesList", sortedTypes);

        // 2. Populate 'SkillLicenseTypesTable' on the sheet
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            dynamic? targetTable = null;

            foreach (dynamic ws in wb.Sheets)
            {
                try
                {
                    targetTable = ws.ListObjects["SkillLicenseTypesTable"];
                    if (targetTable != null) break;
                }
                catch { }
            }

            if (targetTable != null)
            {
                try
                {
                    if (targetTable.DataBodyRange != null)
                        targetTable.DataBodyRange.Delete();
                }
                catch { }

                int count = sortedTypes.Length;
                object[,] matrix = new object[count, 1];
                for (int i = 0; i < count; i++) matrix[i, 0] = sortedTypes[i];

                dynamic headerCell = targetTable.HeaderRowRange.Cells[1, 1];
                dynamic firstDataCell = headerCell.Offset[1, 0];
                dynamic writeRange = firstDataCell.Resize[count, 1];
                writeRange.Value2 = matrix;
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[SyncToSkillLicenseTypesTable Error] {ex.Message}");
        }
    }

    /// <summary>
    /// Intercepts native Excel Save operations:
    /// - QAT floppy disk icon (top-left)
    /// - File > Save / Save As
    /// - Ctrl+S during cell edit mode
    /// </summary>
    private static void OnWorkbookBeforeSave(dynamic wb, bool saveAsUi, ref bool cancel)
    {
        // Allow native save if in Master Template Design Mode or internal operations
        if (_isInternalSaving || WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode)
        {
            return;
        }

        // 1. Tell Excel to CANCEL the native file-dialog save
        cancel = true;

        // 2. Execute custom save routine on the Excel macro queue
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            ExecuteUnifiedSave();
        });
    }

    private static void OnWorkbookBeforeClose(dynamic wb, ref bool cancel)
    {
        if (_isExiting)
        {
            cancel = false;
            return;
        }

        try
        {
            if (WorkbookSession.IsMasterTemplateFile() || WorkbookSession.IsTemplateEditMode) return;

            while (true)
            {
                var (workerList, validationErrors, validationWarnings) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

                // -------------------------------------------------------------
                // 1. DATA VALIDATION CHECK
                // -------------------------------------------------------------
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new ValidationSummaryDialog("Worker Manager", validationErrors, validationWarnings);
                    valDialog.ShowDialog();

                    if (validationErrors.Count > 0)
                    {
#if DEBUG
                        // DEBUG MODE OVERRIDE: Allow developer to bypass blocking errors and quit without saving
                        var debugChoice = MessageBox.Show(
                            $"[DEBUG MODE CLOSE OVERRIDE]\n\n" +
                            $"{validationErrors.Count} critical validation error(s) were found.\n\n" +
                            "Would you like to FORCE QUIT anyway without saving?\n" +
                            "• Yes: Discard all changes and close Excel immediately.\n" +
                            "• No: Cancel closing and stay in Excel to fix the errors.",
                            "Debug Force Quit",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Warning
                        );

                        if (debugChoice == DialogResult.Yes)
                        {
                            _isExiting = true;
                            UnhookApplicationEvents();
                            wb.Saved = true; // Tell Excel all changes are discarded/handled
                            cancel = false;  // Allow native close to exit cleanly
                            return;
                        }
#endif
                        cancel = true; // Stay in Excel in Release mode (or if developer clicks 'No')
                        return;
                    }
                    else if (valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        cancel = true;
                        return;
                    }
                }

                // If completely empty, allow clean exit
                if (workerList == null || workerList.Count == 0)
                {
                    _isExiting = true;
                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                // -------------------------------------------------------------
                // 2. ATTEMPT SILENT SERVER SAVE
                // -------------------------------------------------------------
                if (ServerSyncCommands.SaveWorkerSilent(out string dbError))
                {
                    Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();

                    _isExiting = true;
                    UnhookApplicationEvents();

                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                // -------------------------------------------------------------
                // 3. SERVER OFFLINE / UNREACHABLE
                // -------------------------------------------------------------
                using var dialog = new ExitSyncConflictDialog(
                    "Worker Manager",
                    string.IsNullOrWhiteSpace(dbError) ? "Could not persist records to the PostgreSQL server." : dbError,
                    allowSaveOffline: true,
                    isClosing: true
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        Services.SaveOfflineSnapshot(workerList);

                        _isExiting = true;
                        UnhookApplicationEvents();

                        wb.Saved = true;
                        cancel = false;
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
            _isInternalSaving = true;
            try
            {
                XlCall.Excel(XlCall.xlcSave);
            }
            finally
            {
                _isInternalSaving = false;
            }
            return;
        }

        ExecuteUnifiedSave();
    }

    /// <summary>
    /// Centralized save routine shared by Ctrl+S, QAT Save icon, and Ribbon Save button.
    /// </summary>
    public static void ExecuteUnifiedSave()
    {
        while (true)
        {
            // 1. Validation & Extraction
            var (workerList, validationErrors, validationWarnings) = SheetMapper.ExtractWorkerRecordsWithDiagnostics();

            if (validationErrors.Count > 0 || validationWarnings.Count > 0)
            {
                using var valDialog = new ValidationSummaryDialog("Worker Manager", validationErrors, validationWarnings);
                valDialog.ShowDialog();

                if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                {
                    return; // Blocking errors present or user cancelled
                }
            }

            // 2. Attempt Database Save
            if (ServerSyncCommands.SaveWorkerSilent(out string dbErr))
            {
                Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();

                if (workerList != null && workerList.Count > 0)
                {
                    SheetMapper.PopulateWorksheet(workerList);
                }

                try
                {
                    dynamic xlApp = ExcelDnaUtil.Application;
                    xlApp.ActiveWorkbook.Saved = true;
                }
                catch { }

                int count = workerList?.Count ?? 0;
                SaveNotificationDialog.ShowSuccess(
                    "Worker Manager",
                    $"Save completed successfully. {count} record(s) persisted to the database.\n\nAll default dates, statuses, and auto-generated IDs have been refreshed on your sheets."
                );
                return;
            }

            // 3. Server Unreachable: Display Offline Stash Dialog
            using var conflictDialog = new ExitSyncConflictDialog(
                "Worker Manager",
                string.IsNullOrWhiteSpace(dbErr) ? "Could not establish a connection to PostgreSQL." : dbErr,
                allowSaveOffline: true,
                isClosing: false // Indicates this is an in-session save, not an exit
            );
            conflictDialog.ShowDialog();

            switch (conflictDialog.Resolution)
            {
                case ExitResolution.Retry:
                    continue;

                case ExitResolution.SaveLocallyAndQuit:
                    Services.SaveOfflineSnapshot(workerList);

                    if (workerList != null && workerList.Count > 0)
                    {
                        SheetMapper.PopulateWorksheet(workerList);
                    }

                    try
                    {
                        dynamic xlApp = ExcelDnaUtil.Application;
                        xlApp.ActiveWorkbook.Saved = true;
                    }
                    catch { }

                    SaveNotificationDialog.ShowOffline(
                        "Worker Manager",
                        $"Saved {workerList?.Count ?? 0} record(s) to local offline storage.\n\nWhen the server is back online, your offline changes will automatically synchronize."
                    );
                    return;

                case ExitResolution.Cancel:
                default:
                    return;
            }
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

            // 1. Sync Lookups first so dropdowns and validation tables are fresh
            SyncWorkerLookupsData();

            // 2. Check for Pending Offline Snapshot
            string pendingFilePath = Folder.Data.GetFilePath("List`1_PendingUpload.json");
            if (Services.HasPendingOfflineSnapshot<List<WorkerRecord>>())
            {
                var pendingRecords = Services.LoadPendingOfflineSnapshot<List<WorkerRecord>>();

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
                    Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                }
                else
                {
                    DateTime localOfflineTime = File.GetLastWriteTime(pendingFilePath);
                    DateTime? serverLastUpdate = Services.GetServerLatestTimestamp("worker_master");

                    bool canCommit = true;

                    // -------------------------------------------------------------
                    // CONFLICT DETECTED: SERVER IS NEWER THAN LOCAL OFFLINE SAVE
                    // -------------------------------------------------------------
                    if (serverLastUpdate.HasValue && serverLastUpdate.Value > localOfflineTime)
                    {
                        var serverRecords = ServerSyncCommands.FetchWorkerRecords() ?? new List<WorkerRecord>();

                        // Build a visual comparison preview
                        var diffSb = new System.Text.StringBuilder();
                        diffSb.AppendLine("=== DATA VERSION CONFLICT DETECTED ===");
                        diffSb.AppendLine("The server was modified after your last offline save.\n");
                        diffSb.AppendLine($"• Your Local Offline Save: {localOfflineTime:yyyy-MM-dd HH:mm:ss} ({pendingRecords.Count} workers)");
                        diffSb.AppendLine($"• Current Server Copy:    {serverLastUpdate.Value:yyyy-MM-dd HH:mm:ss} ({serverRecords.Count} workers)\n");
                        diffSb.AppendLine("Summary of Differences:");

                        var serverIdMap = serverRecords.ToDictionary(s => s.WorkerID);
                        int newLocalCount = 0;
                        int modifiedCount = 0;

                        foreach (var local in pendingRecords)
                        {
                            if (!serverIdMap.TryGetValue(local.WorkerID, out var srv))
                            {
                                newLocalCount++;
                            }
                            else if (!local.WorkerDetails.WorkerName.Equals(srv.WorkerDetails.WorkerName, StringComparison.OrdinalIgnoreCase) ||
                                     local.SkillLicenses.Count != srv.SkillLicenses.Count)
                            {
                                modifiedCount++;
                            }
                        }

                        diffSb.AppendLine($"  - New workers in your offline file: {newLocalCount}");
                        diffSb.AppendLine($"  - Modified workers in your offline file: {modifiedCount}\n");
                        diffSb.AppendLine("Would you like to OVERWRITE the server with your local offline version?");
                        diffSb.AppendLine("• Yes: Overwrite server database with your local offline records.");
                        diffSb.AppendLine("• No:  Discard your offline file and load the latest server copy.");

                        var dialogChoice = MessageBox.Show(
                            diffSb.ToString(),
                            "Synchronization Conflict Preview",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        canCommit = (dialogChoice == DialogResult.Yes);
                    }

                    if (canCommit)
                    {
                        var (success, count, path) = ServerSyncCommands.SaveWorkerAsync(pendingRecords!)
                                                                        .GetAwaiter()
                                                                        .GetResult();
                        if (success)
                        {
                            Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                            SheetMapper.PopulateWorksheet(pendingRecords!);

                            SaveNotificationDialog.ShowSuccess(
                                "Worker Manager",
                                $"Offline changes uploaded successfully!\n\n• Committed {count} worker record(s) to PostgreSQL server.\n• Local pending upload stash has been cleared."
                            );
                            return;
                        }
                        else
                        {
                            SaveNotificationDialog.ShowError(
                                "Worker Manager",
                                "Failed to upload offline changes to the server database. Records remain in offline storage."
                            );
                        }
                    }
                    else
                    {
                        // User chose to discard local offline changes
                        Services.ClearPendingOfflineSnapshot<List<WorkerRecord>>();
                        SaveNotificationDialog.ShowSuccess(
                            "Worker Manager",
                            "Local offline stash discarded. Loading current version from server..."
                        );
                    }
                }
            }

            // 3. Load latest copy from server
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