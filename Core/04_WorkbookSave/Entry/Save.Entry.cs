namespace SKSB_App.Core.WorkbookSave.Entry;

using SKSB_App.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.Undo;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

public static class SaveEntry
{
    // =========================================================================
    // 1. SAVE PIPELINE
    // =========================================================================
    public static void Save<T>(SaveContract<T> contract) => TrySave(contract);

    /// <param name="bypassWarnings">
    /// True when the user already reviewed the warnings in the Validation Task Pane and chose "Proceed with Save".
    /// Errors always block.
    /// </param>
    public static bool TrySave<T>(SaveContract<T> contract, bool showSuccessNotification = true, bool repopulateAfterSave = true, bool bypassWarnings = false)
        => TrySaveCore(contract, showSuccessNotification, repopulateAfterSave, bypassWarnings, showIssuesOverride: null);

    /// <param name="showIssuesOverride">
    /// Replaces the default task-pane presentation when validation blocks the save (used by the close flow so the
    /// pane stays in "closing" mode).
    /// </param>
    private static bool TrySaveCore<T>(
        SaveContract<T> contract,
        bool showSuccessNotification,
        bool repopulateAfterSave,
        bool bypassWarnings,
        Action<List<ValidationIssue>, List<ValidationIssue>>? showIssuesOverride)
    {
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return false;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }

        dynamic? activeWb = null;
        try { activeWb = xlApp.ActiveWorkbook; } catch { }

        try
        {
            using var progress = LoadingOverlay.BeginScope($"{contract.ModuleName} — Saving to System", "Extracting and validating workbook tables...");

            // 1. FREEZE ENTIRE SCREEN BEFORE ANY EXTRACTION OR POPULATION
            try { xlApp.ScreenUpdating = false; } catch { }

            while (true)
            {
                LoadingOverlay.Report("Extracting and validating workbook tables...");
                var (data, validationErrors, validationWarnings) = contract.ExtractAndValidate();
                validationErrors ??= new List<ValidationIssue>();
                validationWarnings ??= new List<ValidationIssue>();
                validationWarnings.AddRange(StackedEntryGuard.CollectStackedEditWarnings((object?)activeWb));
                if (validationErrors.Count > 0 || (!bypassWarnings && validationWarnings.Count > 0))
                {
                    // Unfreeze + hide floating widgets, then hand the issues to the non-blocking docked pane.
                    SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                    try { xlApp.EnableEvents = true; } catch { }
                    try { xlApp.ScreenUpdating = true; } catch { }

                    AppLogger.Info($"[{contract.ModuleName}] Save blocked by validation: {validationErrors.Count} error(s), {validationWarnings.Count} warning(s).");
                    if (showIssuesOverride != null)
                    {
                        showIssuesOverride(validationErrors, validationWarnings);
                    }
                    else
                    {
                        Action proceed = () => TrySave(contract, showSuccessNotification, repopulateAfterSave, bypassWarnings: true);
                        ValidationTaskPaneManager.ShowIssues(
                            contract.ModuleName,
                            validationErrors,
                            validationWarnings,
                            onRevalidate: () => ValidateAndShow(contract, proceed),
                            onProceedWithSave: proceed,
                            showBlockingNotice: true);
                    }
                    return false;
                }

                // Validation passed: any open issue list for this workbook is now stale.
                ValidationTaskPaneManager.HideForActiveWorkbook();
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
                        StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                        UndoManager.ClearWorkbook(activeWb);
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
                    ValidationTaskPaneManager.HideForActiveWorkbook();
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    LoadingOverlay.Report("Refreshing worksheet tables...");
                    if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                    StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                    UndoManager.ClearWorkbook(activeWb);
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
                        StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                        UndoManager.ClearWorkbook(activeWb);
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

    // =========================================================================
    // 2. VALIDATE-ONLY ("CHECK ISSUES")
    // =========================================================================
    /// <summary>
    /// Runs extraction + validation without saving and shows the outcome in the docked Validation Task Pane
    /// (including an "All checks passed" state). Must run on the macro thread.
    /// </summary>
    public static void Validate<T>(SaveContract<T> contract, bool allowProceedWithSave = true)
    {
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode)
        {
            SheetOperations.ShowStatusConfirmation("Check Issues is unavailable while editing a template.");
            return;
        }

        Action? proceed = allowProceedWithSave
            ? () => TrySave(contract, bypassWarnings: true)
            : null;
        ValidateAndShow(contract, proceed);
    }

    private static void ValidateAndShow<T>(SaveContract<T> contract, Action? onProceedWithSave)
    {
        var (errors, warnings) = RunValidation(contract);
        SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
        ValidationTaskPaneManager.ShowIssues(
            contract.ModuleName,
            errors,
            warnings,
            onRevalidate: () => ValidateAndShow(contract, onProceedWithSave),
            onProceedWithSave: onProceedWithSave);
    }

    /// <summary>Extract + validate under a frozen screen (restored in <c>finally</c>); nothing is persisted.</summary>
    private static (List<ValidationIssue> Errors, List<ValidationIssue> Warnings) RunValidation<T>(SaveContract<T> contract)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }

        try
        {
            using var progress = LoadingOverlay.BeginScope($"{contract.ModuleName} — Checking Issues", "Extracting and validating workbook tables...");
            try { xlApp.ScreenUpdating = false; } catch { }

            var (_, errors, warnings) = contract.ExtractAndValidate();
            warnings ??= new List<ValidationIssue>();
            dynamic? activeWb = null;
            try { activeWb = xlApp.ActiveWorkbook; } catch { }
            warnings.AddRange(StackedEntryGuard.CollectStackedEditWarnings((object?)activeWb));
            return (errors ?? new List<ValidationIssue>(), warnings);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{contract.ModuleName}] Validation run failed", ex);
            return (
                new List<ValidationIssue> { new(ValidationSeverity.Error, string.Empty, string.Empty, 0, 0, string.Empty, $"Validation failed: {ex.Message}") },
                new List<ValidationIssue>());
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    // =========================================================================
    // 3. CLOSE INTERCEPTION
    // =========================================================================
    /// <param name="onForceClose">
    /// Closes the workbook without re-entering this interception (e.g. after "Save &amp; Close" from the pane).
    /// When null, a default that unhooks events and closes the workbook is used.
    /// </param>
    public static void InterceptClose<T>(
        dynamic wb,
        ref bool cancel,
        SaveContract<T> contract,
        ref bool isExiting,
        Action unEventHooks,
        Action? onForceClose = null)
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
        bool showingIssuePane = false;

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
                validationErrors ??= new List<ValidationIssue>();
                validationWarnings ??= new List<ValidationIssue>();
                validationWarnings.AddRange(StackedEntryGuard.CollectStackedEditWarnings((object)wb));

                // 3. Unfreeze before showing any modal dialogs or exiting
                try { xlApp.ScreenUpdating = true; } catch { }
                try { xlApp.EnableEvents = true; } catch { }

                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    // Keep the workbook open (and the user on the sheet they are on) and list the issues in the
                    // docked pane once this COM event has returned.
                    cancel = true;
                    showingIssuePane = true;

                    object wbRef = (object)wb;
                    Action forceClose = onForceClose ?? (() => DefaultForceClose(wbRef, unEventHooks));
                    var closeErrors = validationErrors;
                    var closeWarnings = validationWarnings;
                    ExcelAsyncUtil.QueueAsMacro(() => ShowClosingIssues(contract, closeErrors, closeWarnings, forceClose, showBlockingNotice: true));
                    return;
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
                    QuitExcelIfLastWorkbook();
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
                        QuitExcelIfLastWorkbook();
                        return;
                    case ExitResolution.QuitAnyway:
                        AppLogger.Info($"[{contract.ModuleName}] Debug close: user chose Quit Anyway (discarded unsaved changes).");
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook();
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
                    QuitExcelIfLastWorkbook();
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
                        QuitExcelIfLastWorkbook();
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
            // Restore active sheet if user cancelled out of the exit prompt. Skipped while the issue pane is
            // shown so the user stays where they are and can navigate from the pane.
            if (activeSheetBefore != null && !isExiting && cancel && !showingIssuePane)
            {
                try { activeSheetBefore!.Activate(); } catch { }
            }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    /// <summary>
    /// Close-time issue pane: Re-validate stays in closing mode, "Save &amp; Close" saves (warnings bypassed) and then
    /// closes, and in DEBUG "Quit Anyway" closes without saving.
    /// </summary>
    private static void ShowClosingIssues<T>(
        SaveContract<T> contract,
        List<ValidationIssue> errors,
        List<ValidationIssue> warnings,
        Action forceClose,
        bool showBlockingNotice = false)
    {
        ValidationTaskPaneManager.ShowIssues(
            contract.ModuleName,
            errors,
            warnings,
            onRevalidate: () =>
            {
                var (e, w) = RunValidation(contract);
                SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                ShowClosingIssues(contract, e, w, forceClose, showBlockingNotice: false);
            },
            onProceedWithSave: () =>
            {
                bool saved = TrySaveCore(
                    contract,
                    showSuccessNotification: false,
                    repopulateAfterSave: false,
                    bypassWarnings: true,
                    showIssuesOverride: (e, w) => ShowClosingIssues(contract, e, w, forceClose, showBlockingNotice: true));
                if (saved) forceClose();
            },
            isClosing: true,
            showBlockingNotice: showBlockingNotice,
#if DEBUG
            onDebugQuitAnyway: () =>
            {
                AppLogger.Info($"[{contract.ModuleName}] Debug close: user chose Quit Anyway from the validation pane.");
                forceClose();
            });
#else
            onDebugQuitAnyway: null);
#endif
    }

    /// <summary>Fallback for <see cref="InterceptClose{T}"/> callers that do not supply their own force-close.</summary>
    private static void DefaultForceClose(object wbObj, Action unEventHooks)
    {
        try { unEventHooks(); } catch { }
        ForceCloseAndQuit(wbObj, "SaveEntry");
    }

    /// <summary>
    /// Discards and closes <paramref name="wbObj"/> (falling back to <c>ActiveWorkbook</c> if the reference is
    /// stale), then quits Excel when no workbook remains. Must run on the macro thread, outside
    /// <c>WorkbookBeforeClose</c>. Callers are responsible for unhooking their own events first.
    /// </summary>
    internal static void ForceCloseAndQuit(object? wbObj, string logTag)
    {
        ValidationTaskPaneManager.HideAndDisposeForWorkbook(wbObj);

        bool closed = TryDiscardAndClose(wbObj, logTag, "workbook reference");
        if (!closed)
        {
            object? active = null;
            try { active = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook; } catch { }
            if (active != null)
            {
                ValidationTaskPaneManager.HideAndDisposeForWorkbook(active);
                closed = TryDiscardAndClose(active, logTag, "ActiveWorkbook fallback");
            }
        }

        if (!closed)
        {
            AppLogger.Warn($"[{logTag}] Force close could not close the workbook; Excel will not be quit.");
            return;
        }

        QuitExcelIfLastWorkbook(workbookAlreadyClosed: true);
    }

    private static bool TryDiscardAndClose(object? wbObj, string logTag, string source)
    {
        if (wbObj == null) return false;
        try
        {
            dynamic wb = wbObj;
            string name = string.Empty;
            try { name = Convert.ToString((object)wb.Name) ?? string.Empty; } catch { }
            try { wb.Saved = true; } catch { }
            wb.Close(false);
            AppLogger.Info($"[{logTag}] Force-closed workbook '{name}' without saving ({source}).");
            return true;
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{logTag}] Force close via {source} failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>
    /// Quits Excel when no other workbook remains.
    /// From <c>WorkbookBeforeClose</c> (<paramref name="workbookAlreadyClosed"/> = false) the closing workbook still
    /// counts (threshold 1) and the quit is deferred until the COM event returns. After an explicit <c>wb.Close()</c>
    /// on the macro thread (<paramref name="workbookAlreadyClosed"/> = true) the quit runs immediately when 0 remain.
    /// </summary>
    internal static void QuitExcelIfLastWorkbook(bool workbookAlreadyClosed = false)
    {
        if (workbookAlreadyClosed)
        {
            try
            {
                dynamic app = ExcelDnaUtil.Application;
                int count = (int)app.Workbooks.Count;
                if (count == 0)
                {
                    AppLogger.Info("[SaveEntry] No workbooks remain after force close; quitting Excel.");
                    app.Quit();
                }
                else
                {
                    AppLogger.Info($"[SaveEntry] {count} workbook(s) still open after force close; Excel stays running.");
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[SaveEntry] Quit after force close failed: {ex.Message}");
            }
            return;
        }

        try
        {
            // Defer quit asynchronously so it executes after the WorkbookBeforeClose COM event returns cleanly
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    if (app == null) return;
                    int count = (int)app.Workbooks.Count;
                    if (count <= 1)
                    {
                        AppLogger.Info($"[SaveEntry] Last workbook closing ({count} open); quitting Excel.");
                        app.Quit();
                    }
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"[SaveEntry] Deferred quit failed: {ex.Message}");
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[SaveEntry] Could not queue deferred quit: {ex.Message}");
        }
    }
}