namespace SKSB_App.Core.WorkbookHelper;

using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.FileManager.Structure;

public static class GenericHookEvents
{
    /// <summary>
    /// Unified save routine (Ctrl+S, QAT Save icon, Ribbon Save button).
    /// </summary>
    public static void Save<T>(SyncDescriptor<T> descriptor)
    {
        while (true)
        {
            var (data, validationErrors, validationWarnings) = descriptor.ExtractAndValidate();

            if (validationErrors.Count > 0 || validationWarnings.Count > 0)
            {
                using var valDialog = new ValidationSummary(descriptor.ModuleName, validationErrors, validationWarnings);
                valDialog.ShowDialog();

                if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                {
                    return; // Blocking errors present or user cancelled
                }
            }

            if (data == null || descriptor.IsDatasetEmpty(data))
            {
                SaveNotification.ShowSuccess(descriptor.ModuleName, "The worksheet contains no records to save.");
                return;
            }

            // Attempt Database Save
            var (success, dbErr) = descriptor.SaveToServer(data);
            if (success)
            {
                OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                descriptor.PopulateWorksheet?.Invoke(data);

                try
                {
                    dynamic xlApp = ExcelDnaUtil.Application;
                    xlApp.ActiveWorkbook.Saved = true;
                }
                catch { }

                SaveNotification.ShowSuccess(
                    descriptor.ModuleName,
                    "Save completed successfully. All records have been persisted to the database."
                );
                return;
            }

            // Server Unreachable -> Prompt for offline stash
            using var conflictDialog = new ExitSyncConflict(
                descriptor.ModuleName,
                string.IsNullOrWhiteSpace(dbErr) ? "Could not establish a connection to PostgreSQL." : dbErr,
                allowSaveOffline: true,
                isClosing: false
            );
            conflictDialog.ShowDialog();

            switch (conflictDialog.Resolution)
            {
                case ExitResolution.Retry:
                    continue;

                case ExitResolution.SaveLocallyAndQuit:
                    OfflineSaveServices.SaveOfflineSnapshot(data);
                    descriptor.PopulateWorksheet?.Invoke(data);

                    try
                    {
                        dynamic xlApp = ExcelDnaUtil.Application;
                        xlApp.ActiveWorkbook.Saved = true;
                    }
                    catch { }

                    SaveNotification.ShowOffline(
                        descriptor.ModuleName,
                        "Saved records to local offline storage.\n\nWhen the server is back online, changes will automatically synchronize."
                    );
                    return;

                case ExitResolution.Cancel:
                default:
                    return;
            }
        }
    }

    /// <summary>
    /// Intercepts Excel window close event to ensure changes are verified and persisted before quitting.
    /// </summary>
    public static void InterceptClose<T>(
        dynamic wb,
        ref bool cancel,
        SyncDescriptor<T> descriptor,
        ref bool isExiting,
        Action unhookEvents)
    {
        if (isExiting)
        {
            cancel = false;
            return;
        }

        try
        {
            if (SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode) return;

            while (true)
            {
                var (data, validationErrors, validationWarnings) = descriptor.ExtractAndValidate();

                // 1. Validation Check
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new ValidationSummary(descriptor.ModuleName, validationErrors, validationWarnings);
                    valDialog.ShowDialog();

#if DEBUG
                    // Quits right at validation without prompting with save dialog
                    if (valDialog.Action == ValidationDialogAction.DebugQuitAnyway)
                    {
                        isExiting = true;
                        unhookEvents();
                        wb.Saved = true;
                        cancel = false;
                        return;
                    }
#endif

                    if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        cancel = true; // Keep Excel open
                        return;
                    }
                }

                if (data == null || descriptor.IsDatasetEmpty(data))
                {
                    isExiting = true;
                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                // 2. Attempt Silent Save
                var (success, dbError) = descriptor.SaveToServer(data);
                if (success)
                {
                    OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                    isExiting = true;
                    unhookEvents();

                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                // 3. Database Offline Dialog
                using var dialog = new ExitSyncConflict(
                    descriptor.ModuleName,
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
                        OfflineSaveServices.SaveOfflineSnapshot(data);
                        isExiting = true;
                        unhookEvents();

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
            System.Diagnostics.Debug.WriteLine($"[GenericHookEvents.InterceptClose] Error: {ex.Message}");
            cancel = true;
        }
    }

    /// <summary>
    /// Delayed startup initialization: synchronizes lookups, pending offline uploads, and version conflicts.
    /// </summary>
    public static void HandleStartup<T>(SyncDescriptor<T> descriptor)
    {
        try
        {
            if (SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode)
            {
                SheetOps.ShowStatusConfirmation("Design Mode: Auto-load bypassed.");
                return;
            }

            descriptor.SyncLookups?.Invoke(false);

            string pendingFilePath = Folder.Data.GetFilePath("List`1_PendingUpload.json");

            if (OfflineSaveServices.HasPendingOfflineSnapshot<T>())
            {
                var pendingData = OfflineSaveServices.LoadPendingOfflineSnapshot<T>();

                bool isValid = pendingData != null && (descriptor.ValidateOfflinePayload?.Invoke(pendingData) ?? true);
                if (!isValid)
                {
                    MessageBox.Show(
                        "Found a corrupt offline pending file. The upload has been aborted to protect the database.",
                        "Upload Rejected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                }
                else
                {
                    DateTime localOfflineTime = File.GetLastWriteTime(pendingFilePath);
                    DateTime? serverLastUpdate = string.IsNullOrEmpty(descriptor.ServerTableName)
                        ? null
                        : OfflineSaveServices.GetServerLatestTimestamp(descriptor.ServerTableName);

                    bool canCommit = true;

                    if (serverLastUpdate.HasValue && serverLastUpdate.Value > localOfflineTime)
                    {
                        var serverData = descriptor.FetchFromServer != null ? descriptor.FetchFromServer() : default;

                        string diffMessage = descriptor.FormatConflictSummary != null && serverData != null
                            ? descriptor.FormatConflictSummary(pendingData!, serverData)
                            : $"• Your Offline Save: {localOfflineTime:yyyy-MM-dd HH:mm:ss}\n• Current Server:     {serverLastUpdate.Value:yyyy-MM-dd HH:mm:ss}\n";

                        var dialogChoice = MessageBox.Show(
                            $"=== DATA VERSION CONFLICT DETECTED ===\n" +
                            $"The server was modified after your last offline save.\n\n" +
                            $"{diffMessage}\n\n" +
                            $"Would you like to OVERWRITE the server with your local offline version?\n" +
                            $"• Yes: Overwrite server database with your local offline records.\n" +
                            $"• No:  Discard your offline file and load the latest server copy.",
                            "Synchronization Conflict Preview",
                            MessageBoxButtons.YesNo,
                            MessageBoxIcon.Question
                        );

                        canCommit = (dialogChoice == DialogResult.Yes);
                    }

                    if (canCommit)
                    {
                        var (success, dbError) = descriptor.SaveToServer(pendingData!);
                        if (success)
                        {
                            OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                            descriptor.PopulateWorksheet?.Invoke(pendingData!);

                            SaveNotification.ShowSuccess(
                                descriptor.ModuleName,
                                "Offline changes uploaded successfully!\n\n• Committed records to PostgreSQL server.\n• Local pending upload stash has been cleared."
                            );
                            return;
                        }
                    }
                    else
                    {
                        OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                    }
                }
            }

            if (descriptor.FetchFromServer != null)
            {
                var serverData = descriptor.FetchFromServer();
                if (serverData != null && !descriptor.IsDatasetEmpty(serverData))
                {
                    descriptor.PopulateWorksheet?.Invoke(serverData);
                    SheetOps.ShowStatusConfirmation($"{descriptor.ModuleName} records loaded from server.");
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[GenericHookEvents.HandleStartup] Error: {ex.Message}");
        }
    }

    /// <summary>
    /// Generic state check for Ribbon toggle buttons across any manager add-in.
    /// </summary>
    public static bool IsSimulatedOffline()
    {
#if DEBUG
        return SKSB_App.Core.ServerManager.ServerEnvironment.SimulateOffline;
#else
        return false;
#endif
    }

    /// <summary>
    /// Generic toggle handler for Ribbon buttons. Toggles ServerEnvironment.SimulateOffline,
    /// invalidates the calling ribbon control, and alerts the developer.
    /// </summary>
    public static void ToggleSimulatedOffline(ExcelDna.Integration.CustomUI.IRibbonUI? ribbonUi, string controlId, bool pressed)
    {
#if DEBUG
        SKSB_App.Core.ServerManager.ServerEnvironment.SimulateOffline = pressed;
        ribbonUi?.InvalidateControl(controlId);

        System.Windows.Forms.MessageBox.Show(
            pressed
                ? "Server is now SIMULATED OFFLINE.\nAll PostgreSQL database operations will fail immediately to test offline local stashing."
                : "Server is now ONLINE.\nNormal PostgreSQL database connectivity restored.",
            "Debug Server Mode",
            System.Windows.Forms.MessageBoxButtons.OK,
            pressed ? System.Windows.Forms.MessageBoxIcon.Warning : System.Windows.Forms.MessageBoxIcon.Information
        );
#endif
    }
}