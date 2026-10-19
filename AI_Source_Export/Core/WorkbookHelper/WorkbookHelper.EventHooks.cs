namespace SKSB_App.Core.WorkbookHelper;

using System;
using System.IO;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.FileManager;

public static class EventHooks
{
    /// <summary>
    /// Unified save routine (Ctrl+S, QAT Save icon, Ribbon Save button).
    /// </summary>
    public static void Save<T>(HookProperties<T> descriptor)
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

                if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                {
                    return; // Blocking errors present or user cancelled
                }
            }

#if DEBUG
            // In DEBUG: ALWAYS prompt. NEVER auto-save without asking.
            using var debugPrompt = new ExitSyncConflict(
                descriptor.ModuleName,
                string.Empty,
                allowSaveOffline: true,
                isClosing: false,
                isDebugPrompt: true
            );
            debugPrompt.ShowDialog();

            switch (debugPrompt.Resolution)
            {
                case ExitResolution.Retry: // Developer clicked "Save to Server"
                    var (dbSuccess, dbErr) = descriptor.SaveToServer(data);
                    if (dbSuccess)
                    {
                        OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                        descriptor.PopulateWorksheet?.Invoke(data);
                        try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                        SaveNotification.ShowSuccess(descriptor.ModuleName, "Debug: Data committed to PostgreSQL server successfully.");
                        return;
                    }
                    else
                    {
                        MessageBox.Show($"Server Save Failed:\n\n{dbErr}", "Debug Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }

                case ExitResolution.SaveLocallyAndQuit: // Developer clicked "Save Offline"
                    OfflineSaveServices.SaveOfflineSnapshot(data);
                    descriptor.PopulateWorksheet?.Invoke(data);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    SaveNotification.ShowOffline(descriptor.ModuleName, "Debug: Records stashed to offline JSON storage.");
                    return;

                case ExitResolution.QuitAnyway:
                case ExitResolution.Cancel:
                default:
                    return; // Return to sheet without saving
            }
#else
            // In RELEASE: Standard direct save to server
            if (data == null || descriptor.IsDatasetEmpty(data))
            {
                SaveNotification.ShowSuccess(descriptor.ModuleName, "The worksheet contains no records to save.");
                return;
            }

            var (success, dbErrRelease) = descriptor.SaveToServer(data);
            if (success)
            {
                OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                descriptor.PopulateWorksheet?.Invoke(data);
                try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                SaveNotification.ShowSuccess(descriptor.ModuleName, "Save completed successfully. All records persisted to database.");
                return;
            }

            // Fallback when server is offline
            using var conflictDialog = new ExitSyncConflict(
                descriptor.ModuleName,
                string.IsNullOrWhiteSpace(dbErrRelease) ? "Could not establish a connection to PostgreSQL." : dbErrRelease,
                allowSaveOffline: true,
                isClosing: false,
                isDebugPrompt: false
            );
            conflictDialog.ShowDialog();

            switch (conflictDialog.Resolution)
            {
                case ExitResolution.Retry:
                    continue;

                case ExitResolution.SaveLocallyAndQuit:
                    OfflineSaveServices.SaveOfflineSnapshot(data);
                    descriptor.PopulateWorksheet?.Invoke(data);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    SaveNotification.ShowOffline(descriptor.ModuleName, "Saved records to local offline storage.");
                    return;

                case ExitResolution.Cancel:
                default:
                    return;
            }
#endif
        }
    }

    /// <summary>
    /// Intercepts Excel window close event to verify, prompt in debug, or persist before quitting.
    /// </summary>
    public static void InterceptClose<T>(
        dynamic wb,
        ref bool cancel,
        HookProperties<T> descriptor,
        ref bool isExiting,
        Action unEventHooks)
    {
        if (isExiting)
        {
            cancel = false;
            return;
        }

        // Never intercept master template files
        if (SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode) return;

        try
        {
            while (true)
            {
                var (data, validationErrors, validationWarnings) = descriptor.ExtractAndValidate();

                // 1. Validation Check
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new ValidationSummary(descriptor.ModuleName, validationErrors, validationWarnings);
                    valDialog.ShowDialog();

#if DEBUG
                    if (valDialog.Action == ValidationDialogAction.DebugQuitAnyway)
                    {
                        isExiting = true;
                        unEventHooks();
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

#if DEBUG
                // In DEBUG: ALWAYS prompt on exit! Do NOT auto-save silently!
                using var closePrompt = new ExitSyncConflict(
                    descriptor.ModuleName,
                    string.Empty,
                    allowSaveOffline: true,
                    isClosing: true,
                    isDebugPrompt: true
                );
                closePrompt.ShowDialog();

                switch (closePrompt.Resolution)
                {
                    case ExitResolution.Retry: // Developer clicked "Save to Server"
                        var (dbSuccess, dbErr) = descriptor.SaveToServer(data);
                        if (dbSuccess)
                        {
                            OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                            SaveNotification.ShowSuccess(descriptor.ModuleName, "Debug: Changes committed to PostgreSQL server before exit.");
                            isExiting = true;
                            unEventHooks();
                            wb.Saved = true;
                            cancel = false;
                            return;
                        }
                        else
                        {
                            MessageBox.Show($"Server Save Failed:\n\n{dbErr}", "Debug Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }

                    case ExitResolution.SaveLocallyAndQuit: // Developer clicked "Save Offline & Quit"
                        OfflineSaveServices.SaveOfflineSnapshot(data);
                        SaveNotification.ShowOffline(descriptor.ModuleName, "Debug: Changes saved to offline storage before exit.");
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        return;

                    case ExitResolution.QuitAnyway: // Developer clicked "Quit Anyway (Discard)"
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;  // Mark true so Excel's native prompt is silenced
                        cancel = false;   // Release Excel close lock
                        return;

                    case ExitResolution.Cancel:
                    default:
                        cancel = true; // Stay in Excel
                        return;
                }
#else
                // In RELEASE on Close: Attempt silent save
                if (data == null || descriptor.IsDatasetEmpty(data))
                {
                    isExiting = true;
                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                var (success, dbError) = descriptor.SaveToServer(data);
                if (success)
                {
                    OfflineSaveServices.ClearPendingOfflineSnapshot<T>();
                    isExiting = true;
                    unEventHooks();
                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                // If DB is offline in production:
                using var dialog = new ExitSyncConflict(
                    descriptor.ModuleName,
                    string.IsNullOrWhiteSpace(dbError) ? "Could not persist records to the PostgreSQL server." : dbError,
                    allowSaveOffline: true,
                    isClosing: true,
                    isDebugPrompt: false
                );
                dialog.ShowDialog();

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;

                    case ExitResolution.SaveLocallyAndQuit:
                        OfflineSaveServices.SaveOfflineSnapshot(data);
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        return;

                    case ExitResolution.Cancel:
                    default:
                        cancel = true;
                        return;
                }
#endif
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[EventHooks.InterceptClose] Error: {ex.Message}");
            cancel = true;
        }
    }

    /// <summary>
    /// Delayed startup initialization: synchronizes lookups, pending offline uploads, and version conflicts.
    /// </summary>
    public static void HandleStartup<T>(HookProperties<T> descriptor)
    {
        try
        {
            if (SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode)
            {
                SheetOps.ShowStatusConfirmation("Design Mode: Auto-load bypassed.");
                return;
            }

            descriptor.SyncLookups?.Invoke(false);

            if (OfflineSaveServices.HasPendingOfflineSnapshot<T>())
            {
                string pendingFilePath = OfflineSaveServices.GetPendingUploadPath<T>();
                var pendingData = OfflineSaveServices.LoadPendingOfflineSnapshot<T>();
                bool isValid = pendingData != null && (descriptor.ValidateOfflinePayload?.Invoke(pendingData) ?? true);

                if (!isValid)
                {
                    MessageBox.Show(
                        "Found a corrupt or invalid offline pending file. The upload has been aborted to protect the database.",
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
                            : $"  Your Offline Save: {localOfflineTime:yyyy-MM-dd HH:mm:ss}\n  Current Server:     {serverLastUpdate.Value:yyyy-MM-dd HH:mm:ss}\n";

                        var dialogChoice = MessageBox.Show(
                            $"=== DATA VERSION CONFLICT DETECTED ===\n" +
                            $"The server was modified after your last offline save.\n\n" +
                            $"{diffMessage}\n\n" +
                            $"Would you like to OVERWRITE the server with your local offline version?\n" +
                            $"  Yes: Overwrite server database with your local offline records.\n" +
                            $"  No:  Discard your offline file and load the latest server copy.",
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
            System.Diagnostics.Debug.WriteLine($"[EventHooks.HandleStartup] Error: {ex.Message}");
        }
    }

    public static bool IsSimulatedOffline()
    {
#if DEBUG
        return SKSB_App.Core.ServerManager.ServerEnvironment.SimulateOffline;
#else
        return false;
#endif
    }

    public static void ToggleSimulatedOffline(ExcelDna.Integration.CustomUI.IRibbonUI? ribbonUi, string controlId, bool pressed)
    {
#if DEBUG
        SKSB_App.Core.ServerManager.ServerEnvironment.SimulateOffline = pressed;
        ribbonUi?.InvalidateControl(controlId);
        MessageBox.Show(
            pressed
                ? "Server is now SIMULATED OFFLINE.\nAll PostgreSQL database operations will fail immediately to test offline local stashing."
                : "Server is now ONLINE.\nNormal PostgreSQL database connectivity restored.",
            "Debug Server Mode",
            MessageBoxButtons.OK,
            pressed ? MessageBoxIcon.Warning : MessageBoxIcon.Information
        );
#endif
    }
}