namespace SKSB_App.Core.WorkbookSave.Entry;

using SKSB_App.Core.Diagnostics;
using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

public static class SaveEntry
{
    public static void Save<T>(SaveContract<T> contract) => TrySave(contract);

    public static bool TrySave<T>(SaveContract<T> contract, bool showSuccessNotification = true, bool repopulateAfterSave = true)
    {
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return false;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }

        try
        {
            using var progress = LoadingOverlay.BeginScope($"{contract.ModuleName} — Saving to System", "Extracting and validating workbook tables...");

            // 1. FREEZE ENTIRE SCREEN BEFORE ANY EXTRACTION OR POPULATION
            try { xlApp.ScreenUpdating = false; } catch { }

            while (true)
            {
                LoadingOverlay.Report("Extracting and validating workbook tables...");
                var (data, validationErrors, validationWarnings) = contract.ExtractAndValidate();
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    // Rule 2 modal exception: unfreeze + hide floating widgets before showing the dialog.
                    SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                    try { xlApp.EnableEvents = true; } catch { }
                    try { xlApp.ScreenUpdating = true; } catch { }

                    using var valDialog = new SummaryDialog(contract.ModuleName, validationErrors, validationWarnings);
                    using (LoadingOverlay.Suspend())
                    {
                        valDialog.ShowDialog();
                    }
                    if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        return false;
                    }
                    try { xlApp.ScreenUpdating = false; } catch { }
                }
#if DEBUG
                // Debug: empty datasets still reach the Debug Save Intercept; only a null dataset bails out.
                if (data == null)
#else
                if (data == null || contract.IsDatasetEmpty(data))
#endif
                {
                    if (data == null) return false;
                    if (showSuccessNotification)
                        SaveNotificationDialog.ShowSuccess(contract.ModuleName, "The worksheet contains no records to save.");
                    return true; // Nothing to save
                }

                bool confirmed;
                using (LoadingOverlay.Suspend())
                {
                    confirmed = contract.ConfirmBeforeSave == null || contract.ConfirmBeforeSave(data);
                }
                if (!confirmed)
                {
                    AppLogger.Info($"[{contract.ModuleName}] Save cancelled at the pre-save confirmation.");
                    return false;
                }

#if DEBUG
                // DEBUG ONLY: confirm before every commit. Rule 2 modal exception: unfreeze + hide floating widgets first.
                SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                try { xlApp.EnableEvents = true; } catch { }
                try { xlApp.ScreenUpdating = true; } catch { }

                using var debugPrompt = new ExitConflictDialog(
                    contract.ModuleName,
                    string.Empty,
                    isClosing: false,
                    isDebugPrompt: true
                );
                using (LoadingOverlay.Suspend())
                {
                    debugPrompt.ShowDialog();
                }
                switch (debugPrompt.Resolution)
                {
                    case ExitResolution.Retry:
                        try { xlApp.ScreenUpdating = false; } catch { }
                        break; // Proceed to contract.SaveToServer(data) below
                    case ExitResolution.SaveLocallyAndQuit:
                        LoadingOverlay.Report("Saving records to local offline storage...");
                        contract.SaveOffline(data);
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        contract.UpdateWindowCaption();
                        LoadingOverlay.Report("Refreshing worksheet tables...");
                        if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                        try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                        if (showSuccessNotification)
                            SaveNotificationDialog.ShowOffline(contract.ModuleName, "Debug: Records stashed to offline JSON storage.");
                        return true;
                    case ExitResolution.QuitAnyway:
                    case ExitResolution.Cancel:
                    default:
                        AppLogger.Info($"[{contract.ModuleName}] Debug save cancelled by user.");
                        return false;
                }
#endif

                LoadingOverlay.Report("Committing records to PostgreSQL server...", "Executing database transaction...");
                var (success, dbErrRelease) = contract.SaveToServer(data);
                if (success)
                {
                    AppLogger.Info($"[{contract.ModuleName}] Save committed to PostgreSQL.");
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    LoadingOverlay.Report("Refreshing worksheet tables...");
                    if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    if (showSuccessNotification)
                        SaveNotificationDialog.ShowSuccess(contract.ModuleName, "Save completed successfully. All records persisted to database.");
                    return true;
                }

                // Rule 2 modal exception: unfreeze + hide floating widgets before showing the dialog.
                SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                try { xlApp.EnableEvents = true; } catch { }
                try { xlApp.ScreenUpdating = true; } catch { }

                using var conflictDialog = new ExitConflictDialog(
                    contract.ModuleName,
                    string.IsNullOrWhiteSpace(dbErrRelease) ? "Could not establish a connection to PostgreSQL." : dbErrRelease,
                    isClosing: false
                );
                using (LoadingOverlay.Suspend())
                {
                    conflictDialog.ShowDialog();
                }
                switch (conflictDialog.Resolution)
                {
                    case ExitResolution.Retry:
                        try { xlApp.ScreenUpdating = false; } catch { }
                        if (ServerEnvironment.IsOffline && !ServerEnvironment.SimulateOffline) ServerEnvironment.PingServer(out _);
                        continue;
                    case ExitResolution.SaveLocallyAndQuit:
                        LoadingOverlay.Report("Saving records to local offline storage...");
                        contract.SaveOffline(data);
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        contract.UpdateWindowCaption();
                        LoadingOverlay.Report("Refreshing worksheet tables...");
                        if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                        try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                        if (showSuccessNotification)
                            SaveNotificationDialog.ShowOffline(contract.ModuleName, "Saved records to local offline storage.");
                        return true;
                    case ExitResolution.Cancel:
                    default:
                        return false;
                }
            }
        }
        finally
        {
            // 2. UNFREEZE SCREEN ONLY AT THE VERY END
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    public static void InterceptClose<T>(
        dynamic wb,
        ref bool cancel,
        SaveContract<T> contract,
        ref bool isExiting,
        Action unEventHooks)
    {
        if (isExiting)
        {
            cancel = false;
            return;
        }
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;

        // 1. Immediately hide any floating dropdown widgets to prevent Z-order fighting
        SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }
        dynamic? activeSheetBefore = null;
        try { activeSheetBefore = xlApp.ActiveSheet; } catch { }

        try
        {
            using var progress = LoadingOverlay.BeginScope($"{contract.ModuleName} — Saving to System", "Extracting and validating workbook tables...");

            while (true)
            {
                // 2. Freeze screen updating and event processing during multi-sheet validation
                try { xlApp.ScreenUpdating = false; } catch { }
                try { xlApp.EnableEvents = false; } catch { }

                LoadingOverlay.Report("Extracting and validating workbook tables...");
                var (data, validationErrors, validationWarnings) = contract.ExtractAndValidate();

                // 3. Unfreeze before showing any modal dialogs or exiting
                try { xlApp.ScreenUpdating = true; } catch { }
                try { xlApp.EnableEvents = true; } catch { }

                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new SummaryDialog(contract.ModuleName, validationErrors, validationWarnings);
                    using (LoadingOverlay.Suspend())
                    {
                        valDialog.ShowDialog();
                    }

#if DEBUG
                    if (valDialog.Action == ValidationDialogAction.DebugQuitAnyway)
                    {
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    }
#endif
                    if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        cancel = true;
                        return;
                    }
                }

#if DEBUG
                // Debug: empty datasets still reach the Debug Save Intercept; only a null dataset exits early.
                if (data == null)
#else
                if (data == null || contract.IsDatasetEmpty(data))
#endif
                {
                    isExiting = true;
                    unEventHooks();
                    wb.Saved = true;
                    cancel = false;
                    QuitExcelIfLastWorkbook(wb);
                    return;
                }

                // Pre-save failsafe (e.g. past-record confirmation). Declining keeps the workbook open.
                bool confirmed;
                using (LoadingOverlay.Suspend())
                {
                    confirmed = contract.ConfirmBeforeSave == null || contract.ConfirmBeforeSave(data);
                }
                if (!confirmed)
                {
                    AppLogger.Info($"[{contract.ModuleName}] Close cancelled at the pre-save confirmation.");
                    cancel = true;
                    return;
                }

#if DEBUG
                // DEBUG ONLY: confirm before every commit. ScreenUpdating / EnableEvents are already true here (step 3 above).
                using var closePrompt = new ExitConflictDialog(
                    contract.ModuleName,
                    string.Empty,
                    isClosing: true,
                    isDebugPrompt: true
                );
                using (LoadingOverlay.Suspend())
                {
                    closePrompt.ShowDialog();
                }
                switch (closePrompt.Resolution)
                {
                    case ExitResolution.Retry:
                        break; // Proceed to freeze + contract.SaveToServer(data) below
                    case ExitResolution.SaveLocallyAndQuit:
                        LoadingOverlay.Report("Saving records to local offline storage...");
                        contract.SaveOffline(data);
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    case ExitResolution.QuitAnyway:
                        AppLogger.Info($"[{contract.ModuleName}] Debug close: user chose Quit Anyway (discarded unsaved changes).");
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    case ExitResolution.Cancel:
                    default:
                        cancel = true;
                        return;
                }
#endif

                try { xlApp.ScreenUpdating = false; } catch { }
                try { xlApp.EnableEvents = false; } catch { }

                LoadingOverlay.Report("Committing records to PostgreSQL server...", "Executing database transaction...");
                var (success, dbError) = contract.SaveToServer(data);

                if (success)
                {
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    isExiting = true;
                    unEventHooks();
                    wb.Saved = true;
                    cancel = false;
                    QuitExcelIfLastWorkbook(wb);
                    return;
                }

                try { xlApp.ScreenUpdating = true; } catch { }
                try { xlApp.EnableEvents = true; } catch { }

                using var dialog = new ExitConflictDialog(
                    contract.ModuleName,
                    string.IsNullOrWhiteSpace(dbError) ? "Could not persist records to the PostgreSQL server." : dbError,
                    isClosing: true
                );
                using (LoadingOverlay.Suspend())
                {
                    dialog.ShowDialog();
                }

                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        if (ServerEnvironment.IsOffline && !ServerEnvironment.SimulateOffline) ServerEnvironment.PingServer(out _);
                        continue;
                    case ExitResolution.SaveLocallyAndQuit:
                        LoadingOverlay.Report("Saving records to local offline storage...");
                        contract.SaveOffline(data);
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
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
            AppLogger.Error("Error encountered during exit validation", ex);
            using (LoadingOverlay.Suspend())
            {
                MessageBox.Show($"Error encountered during exit validation:\n\n{ex.Message}\n\nDetails were written to the runtime log.",
                    "Exit Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            cancel = true;
        }
        finally
        {
            // Restore active sheet if user cancelled out of the exit prompt
            if (activeSheetBefore != null && !isExiting && cancel)
            {
                try { activeSheetBefore!.Activate(); } catch { }
            }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    private static void QuitExcelIfLastWorkbook(dynamic wb)
    {
        try
        {
            dynamic app = wb.Application;
            // Defer quit asynchronously so it executes after the WorkbookBeforeClose COM event returns cleanly
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    if (app != null && (int)app!.Workbooks.Count <= 1)
                    {
                        app!.Quit();
                    }
                }
                catch { }
            });
        }
        catch { }
    }
}