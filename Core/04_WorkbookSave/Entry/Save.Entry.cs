namespace SKSB_App.Core.WorkbookSave.Entry;

using SKSB_App.Core.Diagnostics;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.BugReport;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.RuntimeUI.Undo;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.SheetMapper;
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
    /// <param name="forceOverwrite">
    /// True when the user clicked "Overwrite Server &amp; Save" in the pane's Conflict Mode: concurrency checks are
    /// skipped (and, in DEBUG, so is the Debug Save Intercept, since the user already confirmed explicitly).
    /// </param>
    public static bool TrySave<T>(
        SaveContract<T> contract,
        bool showSuccessNotification = true,
        bool repopulateAfterSave = true,
        bool bypassWarnings = false,
        bool forceOverwrite = false)
        => TrySaveCore(contract, showSuccessNotification, repopulateAfterSave, bypassWarnings, forceOverwrite, showIssuesOverride: null);

    /// <param name="showIssuesOverride">
    /// Replaces the default task-pane presentation when validation blocks the save (used by the close flow so the
    /// pane stays in "closing" mode).
    /// </param>
    /// <param name="showConflictsOverride">
    /// Replaces the default Conflict Mode presentation when the commit is rolled back by concurrency conflicts (used
    /// by the close flow so "Overwrite &amp; Close" closes the workbook afterwards).
    /// </param>
    private static bool TrySaveCore<T>(
        SaveContract<T> contract,
        bool showSuccessNotification,
        bool repopulateAfterSave,
        bool bypassWarnings,
        bool forceOverwrite,
        Action<List<ValidationIssue>, List<ValidationIssue>>? showIssuesOverride,
        Action<IReadOnlyList<ConcurrencyConflict>>? showConflictsOverride = null)
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
            string displayModule = Loc.ModuleDisplayName(contract.ModuleName);
            using var progress = LoadingOverlay.BeginScope(
                Loc.T("save.overlay.title", displayModule),
                Loc.T("save.overlay.extracting"),
                operationKey: $"{contract.ModuleName} — Saving to System");

            // 1. FREEZE ENTIRE SCREEN BEFORE ANY EXTRACTION OR POPULATION
            try { xlApp.ScreenUpdating = false; } catch { }

            while (true)
            {
                LoadingOverlay.Report(Loc.T("save.overlay.extracting"));
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
                            showBlockingNotice: true,
                            onDiscardDraftAndReload: BuildDiscardDraftAction(contract));
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
                        SaveNotificationDialog.ShowSuccess(contract.ModuleName, Loc.T("save.msg.noRecords"));
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
                // DEBUG ONLY: confirm before every commit (skipped for "Overwrite Server & Save", already confirmed).
                // Rule 2 modal exception: unfreeze + hide floating widgets first.
                if (!forceOverwrite)
                {
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
                            break; // Proceed to the server commit below
                        case ExitResolution.SaveLocallyAndQuit:
                            LoadingOverlay.Report(Loc.T("save.overlay.savingOffline"));
                            contract.SaveOffline(data);
                            WorkbookDraft.Clear<T>();
                            ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                            contract.UpdateWindowCaption();
                            ManagerRibbonBase.InvalidateRibbon();
                            LoadingOverlay.Report(Loc.T("save.overlay.refreshing"));
                            if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                            StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                            UndoManager.ClearWorkbook(activeWb);
                            try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                            if (showSuccessNotification)
                                SaveNotificationDialog.ShowOffline(contract.ModuleName, Loc.T("save.msg.debugOffline"));
                            return true;
                        case ExitResolution.QuitAnyway:
                        case ExitResolution.Cancel:
                        default:
                            AppLogger.Info($"[{contract.ModuleName}] Debug save cancelled by user.");
                            return false;
                    }
                }
#endif

                LoadingOverlay.Report(Loc.T("save.overlay.committing"), Loc.T("save.overlay.committingDetail"));
                var (success, dbErrRelease, conflicts) = CommitToServer(contract, data, forceOverwrite);
                if (!success && conflicts is { Count: > 0 })
                {
                    // Nothing was saved. List the field-level diffs in the docked pane (Conflict Mode): no popup.
                    AppLogger.Warn($"[{contract.ModuleName}] Save rolled back: {conflicts.Count} concurrency conflict(s).");
                    SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                    try { xlApp.EnableEvents = true; } catch { }
                    try { xlApp.ScreenUpdating = true; } catch { }

                    if (showConflictsOverride != null)
                    {
                        showConflictsOverride(conflicts);
                        return false;
                    }

                    object? wbRef = (object?)activeWb;
                    var conflictIssues = BuildConflictIssues(wbRef, conflicts);
                    Action overwrite = () =>
                    {
                        AppLogger.Info($"[{contract.ModuleName}] Concurrency conflict: user chose Overwrite Server & Save.");
                        TrySave(contract, showSuccessNotification, repopulateAfterSave, bypassWarnings: true, forceOverwrite: true);
                    };
                    Action keepServer = () =>
                    {
                        AppLogger.Info($"[{contract.ModuleName}] Concurrency conflict: user chose Keep Server Values.");
                        ApplyServerValuesToWorkbook(wbRef, conflictIssues);
                        contract.OnAcceptServerValues?.Invoke(conflicts);
                        ValidateAndShow(contract, () => TrySave(contract, showSuccessNotification, repopulateAfterSave, bypassWarnings: true, forceOverwrite: false));
                    };

                    ValidationTaskPaneManager.ShowIssues(
                        contract.ModuleName,
                        new List<ValidationIssue>(),
                        conflictIssues,
                        onRevalidate: null,
                        onProceedWithSave: overwrite,
                        isClosing: false,
                        onKeepServerValues: keepServer,
                        showBlockingNotice: true);
                    SheetOperations.ShowStatusConfirmation(
                        Loc.T("save.status.pausedConflicts", displayModule, conflictIssues.Count));
                    return false;
                }

                if (success)
                {
                    AppLogger.Info($"[{contract.ModuleName}] Save committed to PostgreSQL.");
                    ValidationTaskPaneManager.HideForActiveWorkbook();
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    WorkbookDraft.Clear<T>();
                    contract.UpdateWindowCaption();
                    ManagerRibbonBase.InvalidateRibbon();
                    LoadingOverlay.Report(Loc.T("save.overlay.refreshing"));
                    if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                    StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                    UndoManager.ClearWorkbook(activeWb);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    if (showSuccessNotification)
                        SaveNotificationDialog.ShowSuccess(contract.ModuleName, Loc.T("save.msg.success"));
                    return true;
                }

                // Rule 2 modal exception: unfreeze + hide floating widgets before showing the dialog.
                SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                try { xlApp.EnableEvents = true; } catch { }
                try { xlApp.ScreenUpdating = true; } catch { }

                using var conflictDialog = new ExitConflictDialog(
                    contract.ModuleName,
                    string.IsNullOrWhiteSpace(dbErrRelease) ? Loc.T("save.err.noConnection") : dbErrRelease,
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
                        LoadingOverlay.Report(Loc.T("save.overlay.savingOffline"));
                        contract.SaveOffline(data);
                        WorkbookDraft.Clear<T>();
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        contract.UpdateWindowCaption();
                        ManagerRibbonBase.InvalidateRibbon();
                        LoadingOverlay.Report(Loc.T("save.overlay.refreshing"));
                        if (repopulateAfterSave) (contract.RefreshAfterSave ?? contract.PopulateWorksheet)?.Invoke(data);
                        StackedEntryGuard.RefreshWorkbookSnapshots(activeWb);
                        UndoManager.ClearWorkbook(activeWb);
                        try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                        if (showSuccessNotification)
                            SaveNotificationDialog.ShowOffline(contract.ModuleName, Loc.T("save.msg.savedOffline"));
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

    /// <summary>Uses the module's concurrency-aware save when wired, else the plain <see cref="SaveContract{T}.SaveToServer"/>.</summary>
    private static (bool Success, string ErrorMessage, IReadOnlyList<ConcurrencyConflict>? Conflicts) CommitToServer<T>(
        SaveContract<T> contract, T data, bool forceOverwrite)
    {
        if (contract.SaveToServerWithConflictCheck != null)
            return contract.SaveToServerWithConflictCheck(data, forceOverwrite);

        var (success, error) = contract.SaveToServer(data);
        return (success, error, null);
    }

    // =========================================================================
    // 1b. CONCURRENCY CONFLICTS (Validation Task Pane "Conflict Mode")
    // =========================================================================
    /// <summary>Bulk-read snapshot of one ListObject, used to locate conflicting rows / columns.</summary>
    private sealed class ConflictTableSnapshot
    {
        public string SheetName = string.Empty;
        public int FirstRow;
        public int FirstCol;
        public int RowCount;
        public string[] Headers = Array.Empty<string>();
        public object[,] Body = new object[2, 2];

        public int FindColumn(string header)
        {
            if (string.IsNullOrWhiteSpace(header)) return 0;
            for (int j = 0; j < Headers.Length; j++)
            {
                if (string.Equals(Headers[j], header.Trim(), StringComparison.OrdinalIgnoreCase)) return FirstCol + j;
            }
            return 0;
        }

        /// <summary>Excel row whose <paramref name="keyHeader"/> cell equals <paramref name="rawKey"/> (0 = not found).</summary>
        public int FindRow(string keyHeader, string rawKey)
        {
            int keyCol = FindColumn(keyHeader);
            if (keyCol == 0 || string.IsNullOrWhiteSpace(rawKey)) return 0;
            int j = keyCol - FirstCol + 1;
            for (int i = 1; i <= RowCount; i++)
            {
                if (KeysMatch(Body[i, j], rawKey)) return FirstRow + i - 1;
            }
            return 0;
        }

        /// <summary>Numeric keys compare as integers (so "0001" matches 1 / 1.0); anything else as trimmed text.</summary>
        private static bool KeysMatch(object? cell, string rawKey)
        {
            if (cell == null) return false;
            string cellText = (Convert.ToString(cell, CultureInfo.InvariantCulture) ?? string.Empty).Trim();
            string key = rawKey.Trim();
            if (TryParseKey(cell, cellText, out long a) && long.TryParse(key, NumberStyles.Integer, CultureInfo.InvariantCulture, out long b))
                return a == b;
            return string.Equals(cellText, key, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryParseKey(object cell, string text, out long value)
        {
            if (cell is double d && Math.Abs(d - Math.Round(d)) < 1e-9 && d >= long.MinValue && d <= long.MaxValue)
            {
                value = (long)Math.Round(d);
                return true;
            }
            return long.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out value);
        }
    }

    /// <summary>Scans the workbook's ListObjects once and snapshots the requested tables (header + body, bulk 2D reads).</summary>
    private static Dictionary<string, ConflictTableSnapshot> SnapshotConflictTables(object? wbObj, IEnumerable<string> tableNames)
    {
        var wanted = new HashSet<string>(tableNames.Where(n => !string.IsNullOrWhiteSpace(n)), StringComparer.OrdinalIgnoreCase);
        var result = new Dictionary<string, ConflictTableSnapshot>(StringComparer.OrdinalIgnoreCase);
        if (wbObj == null || wanted.Count == 0) return result;

        try
        {
            dynamic wb = wbObj;
            foreach (dynamic ws in wb.Worksheets)
            {
                string sheetName = string.Empty;
                try { sheetName = Convert.ToString((object)ws.Name) ?? string.Empty; } catch { }

                foreach (dynamic lo in ws.ListObjects)
                {
                    string name = string.Empty;
                    try { name = Convert.ToString((object)lo.Name) ?? string.Empty; } catch { }
                    if (!wanted.Contains(name) || result.ContainsKey(name)) continue;

                    var snap = new ConflictTableSnapshot { SheetName = sheetName };
                    try
                    {
                        dynamic header = lo.HeaderRowRange;
                        int colCount = (int)header.Columns.Count;
                        snap.FirstCol = (int)header.Column;
                        object[,] headerValues = TableMapper.GetRange2DValues(header, 1, colCount);
                        snap.Headers = new string[colCount];
                        for (int j = 1; j <= colCount; j++)
                            snap.Headers[j - 1] = (Convert.ToString(headerValues[1, j], CultureInfo.InvariantCulture) ?? string.Empty).Trim();

                        dynamic? body = null;
                        try { body = lo.DataBodyRange; } catch { }
                        if (body != null)
                        {
                            snap.FirstRow = (int)body.Row;
                            snap.RowCount = (int)body.Rows.Count;
                            snap.Body = TableMapper.GetRange2DValues(body, snap.RowCount, colCount);
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Warn($"[SaveEntry] Could not read table '{name}' for conflict mapping: {ex.Message}");
                    }
                    result[name] = snap;
                }
                if (result.Count == wanted.Count) break;
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[SaveEntry] Conflict table scan failed: {ex.Message}");
        }
        return result;
    }

    /// <summary>
    /// Turns server-side concurrency conflicts into Conflict-Mode pane issues: one per differing column, pointing at
    /// the exact sheet cell, with the server value carried in <see cref="ValidationIssue.OriginalValue"/>. Conflicts
    /// without field diffs (or whose row is not on the sheet) produce a single row-level issue.
    /// </summary>
    internal static List<ValidationIssue> BuildConflictIssues(object? wbObj, IReadOnlyList<ConcurrencyConflict> conflicts)
    {
        var issues = new List<ValidationIssue>();
        var tables = SnapshotConflictTables(wbObj, conflicts.Select(c => c.ExcelTableName));

        foreach (var conflict in conflicts)
        {
            ConflictTableSnapshot? table = null;
            if (!string.IsNullOrWhiteSpace(conflict.ExcelTableName)) tables.TryGetValue(conflict.ExcelTableName, out table);

            string sheetName = table?.SheetName ?? string.Empty;
            int excelRow = table?.FindRow(conflict.KeyColumnHeader, conflict.RawKeyValue) ?? 0;
            string? rowKey = string.IsNullOrWhiteSpace(conflict.RawKeyValue) ? null : conflict.RawKeyValue;

            if (conflict.FieldDiffs is { Count: > 0 } diffs)
            {
                foreach (var diff in diffs)
                {
                    int excelCol = excelRow > 0 ? table!.FindColumn(diff.ColumnHeader) : 0;
                    issues.Add(new ValidationIssue(
                        ValidationSeverity.Warning,
                        sheetName,
                        conflict.ExcelTableName,
                        excelRow,
                        excelCol,
                        diff.ColumnHeader,
                        Loc.T("save.conflict.fieldDiff", conflict.RecordKey, diff.ServerDisplay, diff.LocalDisplay),
                        IsStackedHistoryEdit: false,
                        OriginalValue: diff.ServerExcelValue,
                        RowKey: rowKey,
                        IsConcurrencyConflict: true));
                }
            }
            else
            {
                issues.Add(new ValidationIssue(
                    ValidationSeverity.Warning,
                    sheetName,
                    conflict.ExcelTableName,
                    excelRow,
                    0,
                    conflict.KeyColumnHeader,
                    conflict.Description,
                    IsStackedHistoryEdit: false,
                    OriginalValue: null,
                    RowKey: rowKey,
                    IsConcurrencyConflict: true));
            }
        }
        return issues;
    }

    /// <summary>
    /// "Keep Server Values": writes each conflict issue's server value into its exact cell (only cell-located issues;
    /// conflicts are a handful of scattered cells, so no bulk range applies), then resets the stacked-edit snapshots
    /// and undo history. Excel state is frozen and restored in <c>finally</c>. Must run on the macro thread.
    /// </summary>
    internal static void ApplyServerValuesToWorkbook(object? wbObj, List<ValidationIssue> conflictIssues)
    {
        if (wbObj == null) return;
        var targets = conflictIssues.Where(i => i.IsConcurrencyConflict && i.ExcelRow > 0 && i.ExcelCol > 0).ToList();

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        bool origEvents = true;
        object? origCalc = null;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }
        try { origCalc = (object)xlApp.Calculation; } catch { }

        int written = 0;
        try
        {
            try { xlApp.ScreenUpdating = false; } catch { }
            try { xlApp.EnableEvents = false; } catch { }
            try { xlApp.Calculation = -4135; } catch { } // xlCalculationManual

            dynamic wb = wbObj;
            foreach (var group in targets.GroupBy(i => i.SheetName, StringComparer.OrdinalIgnoreCase))
            {
                dynamic? ws = null;
                try { ws = wb.Worksheets[group.Key]; } catch { }
                if (ws == null)
                {
                    AppLogger.Warn($"[SaveEntry] Keep Server Values: sheet '{group.Key}' not found.");
                    continue;
                }

                foreach (var issue in group)
                {
                    try
                    {
                        ws.Cells[issue.ExcelRow, issue.ExcelCol].Value2 = ToExcelCellValue(issue.OriginalValue);
                        written++;
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Warn($"[SaveEntry] Keep Server Values: could not write {group.Key}!{issue.CellAddress}: {ex.Message}");
                    }
                }
            }

            StackedEntryGuard.RefreshWorkbookSnapshots(wbObj);
            UndoManager.ClearWorkbook(wbObj);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[SaveEntry] Keep Server Values failed", ex);
        }
        finally
        {
            try { if (origCalc != null) xlApp.Calculation = origCalc; else xlApp.Calculation = -4105; } catch { } // xlCalculationAutomatic
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }

        AppLogger.Info($"[SaveEntry] Keep Server Values: wrote {written}/{targets.Count} conflicting cell(s).");
        SheetOperations.ShowStatusConfirmation(Loc.T("save.status.restoredServerValues", written));
    }

    /// <summary>Server value -> COM-safe Value2 (OADate for dates, double for numbers; null clears the cell).</summary>
    private static object? ToExcelCellValue(object? value) => value switch
    {
        null => null,
        DateTime dt => dt.ToOADate(),
        DateOnly d => d.ToDateTime(TimeOnly.MinValue).ToOADate(),
        decimal m => (double)m,
        long l => (double)l,
        int i => (double)i,
        string s when s.Length == 0 => null,
        _ => value
    };

    /// <summary>
    /// Close-time Conflict Mode: "Overwrite &amp; Close" force-saves then closes; "Keep Server Values" restores the
    /// conflicting cells and re-validates in closing mode so the user can review before "Save &amp; Close".
    /// Must run on the macro thread (outside <c>WorkbookBeforeClose</c>).
    /// </summary>
    private static void ShowClosingConflicts<T>(
        SaveContract<T> contract,
        object? wbObj,
        IReadOnlyList<ConcurrencyConflict> conflicts,
        List<ValidationIssue> conflictIssues,
        Action forceClose)
    {
        Action overwrite = () =>
        {
            AppLogger.Info($"[{contract.ModuleName}] Close: concurrency conflict, user chose Overwrite & Close.");
            bool saved = TrySaveCore(
                contract,
                showSuccessNotification: false,
                repopulateAfterSave: false,
                bypassWarnings: true,
                forceOverwrite: true,
                showIssuesOverride: (e, w) => ShowClosingIssues(contract, e, w, forceClose, showBlockingNotice: true));
            if (saved) forceClose();
        };
        Action keepServer = () =>
        {
            AppLogger.Info($"[{contract.ModuleName}] Close: concurrency conflict, user chose Keep Server Values.");
            ApplyServerValuesToWorkbook(wbObj, conflictIssues);
            contract.OnAcceptServerValues?.Invoke(conflicts);
            var (e, w) = RunValidation(contract);
            SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
            ShowClosingIssues(contract, e, w, forceClose, showBlockingNotice: false);
        };

        ValidationTaskPaneManager.ShowIssues(
            contract.ModuleName,
            new List<ValidationIssue>(),
            conflictIssues,
            onRevalidate: null,
            onProceedWithSave: overwrite,
            isClosing: true,
#if DEBUG
            onDebugQuitAnyway: () =>
            {
                AppLogger.Info($"[{contract.ModuleName}] Debug close: user chose Quit Anyway from the conflict pane.");
                forceClose();
            },
#else
            onDebugQuitAnyway: null,
#endif
            onKeepServerValues: keepServer,
            showBlockingNotice: true);
        SheetOperations.ShowStatusConfirmation(
            Loc.T("save.status.closePausedConflicts", Loc.ModuleDisplayName(contract.ModuleName), conflictIssues.Count));
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
            SheetOperations.ShowStatusConfirmation(Loc.T("save.status.validateUnavailableTemplate"));
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
            onProceedWithSave: onProceedWithSave,
            onDiscardDraftAndReload: BuildDiscardDraftAction(contract));
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
            using var progress = LoadingOverlay.BeginScope(
                Loc.T("save.overlay.checkingTitle", Loc.ModuleDisplayName(contract.ModuleName)),
                Loc.T("save.overlay.extracting"),
                operationKey: $"{contract.ModuleName} — Checking Issues");
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
                new List<ValidationIssue> { new(ValidationSeverity.Error, string.Empty, string.Empty, 0, 0, string.Empty, Loc.T("save.err.validationFailed", ex.Message)) },
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
            using var progress = LoadingOverlay.BeginScope(
                Loc.T("save.overlay.title", Loc.ModuleDisplayName(contract.ModuleName)),
                Loc.T("save.overlay.extracting"),
                operationKey: $"{contract.ModuleName} — Saving to System");

            while (true)
            {
                // 2. Freeze screen updating and event processing during multi-sheet validation
                try { xlApp.ScreenUpdating = false; } catch { }
                try { xlApp.EnableEvents = false; } catch { }

                LoadingOverlay.Report(Loc.T("save.overlay.extracting"));
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
                // DEBUG ONLY: confirm before every commit. ("Overwrite & Close" runs through TrySaveCore instead.)
                // ScreenUpdating / EnableEvents are already true here (step 3 above).
                {
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
                            break; // Proceed to freeze + server commit below
                        case ExitResolution.SaveLocallyAndQuit:
                            LoadingOverlay.Report(Loc.T("save.overlay.savingOffline"));
                            contract.SaveOffline(data);
                            WorkbookDraft.Clear<T>();
                            ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                            ManagerRibbonBase.InvalidateRibbon();
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
                }
#endif

                try { xlApp.ScreenUpdating = false; } catch { }
                try { xlApp.EnableEvents = false; } catch { }

                LoadingOverlay.Report(Loc.T("save.overlay.committing"), Loc.T("save.overlay.committingDetail"));
                var (success, dbError, conflicts) = CommitToServer(contract, data, forceOverwrite: false);

                if (!success && conflicts is { Count: > 0 })
                {
                    // Nothing was saved: keep the workbook open and list the field-level diffs in the docked pane
                    // (Conflict Mode) once this COM event has returned. Bulk reads only here; writes happen later.
                    AppLogger.Warn($"[{contract.ModuleName}] Close: save rolled back by {conflicts.Count} concurrency conflict(s).");
                    var conflictIssues = BuildConflictIssues((object)wb, conflicts);

                    SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
                    try { xlApp.EnableEvents = true; } catch { }
                    try { xlApp.ScreenUpdating = true; } catch { }
                    cancel = true;
                    showingIssuePane = true;

                    object wbRef = (object)wb;
                    Action forceClose = onForceClose ?? (() => DefaultForceClose(wbRef, unEventHooks));
                    var capturedConflicts = conflicts;
                    ExcelAsyncUtil.QueueAsMacro(() => ShowClosingConflicts(contract, wbRef, capturedConflicts, conflictIssues, forceClose));
                    return;
                }

                if (success)
                {
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    WorkbookDraft.Clear<T>();
                    contract.UpdateWindowCaption();
                    ManagerRibbonBase.InvalidateRibbon();
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
                    string.IsNullOrWhiteSpace(dbError) ? Loc.T("save.err.couldNotPersist") : dbError,
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
                        LoadingOverlay.Report(Loc.T("save.overlay.savingOffline"));
                        contract.SaveOffline(data);
                        WorkbookDraft.Clear<T>();
                        ServerEnvironment.EnterOfflineMode("Saved records to local offline storage.");
                        ManagerRibbonBase.InvalidateRibbon();
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
                MessageBox.Show(Loc.T("save.err.exitValidationBody", ex.Message),
                    Loc.T("save.err.exitValidationTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        Action? saveDraftAndClose = errors.Count > 0
            ? () => SaveDraftAndClose(contract, errors.Count, warnings.Count, forceClose)
            : null;

        ValidationTaskPaneManager.ShowIssues(
            contract.ModuleName,
            errors,
            warnings,
            onSaveDraftAndClose: saveDraftAndClose,
            onDiscardDraftAndReload: BuildDiscardDraftAction(contract),
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
                    forceOverwrite: false,
                    showIssuesOverride: (e, w) => ShowClosingIssues(contract, e, w, forceClose, showBlockingNotice: true),
                    // A conflict on "Save & Close" stays in closing mode so "Overwrite & Close" still closes.
                    showConflictsOverride: conflicts =>
                    {
                        object? wbObj = null;
                        try { wbObj = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook; } catch { }
                        ShowClosingConflicts(contract, wbObj, conflicts, BuildConflictIssues(wbObj, conflicts), forceClose);
                    });
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

    // =========================================================================
    // 4. LOCAL DRAFT ("SAVE DRAFT & CLOSE")
    // =========================================================================
    /// <summary>
    /// "Save Draft &amp; Close": writes a lossless cell-level draft of the workbook (never uploaded to PostgreSQL),
    /// then closes. On a failed write the workbook stays open. Must run on the macro thread.
    /// </summary>
    private static void SaveDraftAndClose<T>(SaveContract<T> contract, int errorCount, int warningCount, Action forceClose)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? activeWb = null;
        try { activeWb = xlApp.ActiveWorkbook; } catch { }

        string displayModule = Loc.ModuleDisplayName(contract.ModuleName);
        try
        {
            using (LoadingOverlay.BeginScope(
                Loc.T("save.overlay.draftTitle", displayModule),
                Loc.T("save.overlay.draftStep"),
                operationKey: $"{contract.ModuleName} — Saving Local Draft"))
            {
                WorkbookDraft.Save<T>((object?)activeWb, contract.ModuleName, errorCount, warningCount);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{contract.ModuleName}] Save Draft & Close failed", ex);
            SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
            try { xlApp.EnableEvents = true; } catch { }
            try { xlApp.ScreenUpdating = true; } catch { }
            MessageBox.Show(
                Loc.T("save.draft.failedBody", ex.Message),
                Loc.T("save.draft.failedTitle", displayModule), MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        AppLogger.Info($"[{contract.ModuleName}] Close: saved local draft with {errorCount} error(s), {warningCount} warning(s).");
        contract.UpdateWindowCaption();
        ManagerRibbonBase.InvalidateRibbon();

        // Rule 2 modal exception: unfreeze + hide floating widgets before the notification.
        SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
        try { xlApp.EnableEvents = true; } catch { }
        try { xlApp.ScreenUpdating = true; } catch { }
        SaveNotificationDialog.ShowOffline(
            contract.ModuleName,
            Loc.T("save.draft.savedBody"));

        forceClose();
    }

    /// <summary>
    /// Startup after a draft was restored: validates and opens the Validation Issues pane (no blocking popup) with
    /// "Proceed with Save" (enabled once errors are fixed) and "Discard Draft &amp; Load Server". Macro thread only.
    /// </summary>
    public static void ShowDraftIssues<T>(SaveContract<T> contract)
    {
        Action proceed = () => TrySave(contract, bypassWarnings: true);
        var (errors, warnings) = RunValidation(contract);
        SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
        ValidationTaskPaneManager.ShowIssues(
            contract.ModuleName,
            errors,
            warnings,
            onRevalidate: () => ValidateAndShow(contract, proceed),
            onProceedWithSave: proceed,
            showBlockingNotice: false,
            onDiscardDraftAndReload: BuildDiscardDraftAction(contract));
    }

    /// <summary>"Discard Draft &amp; Load Server" action, or null when no local draft exists.</summary>
    private static Action? BuildDiscardDraftAction<T>(SaveContract<T> contract)
    {
        if (!contract.HasLocalDraft) return null;
        return () =>
        {
            SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
            try { ((dynamic)ExcelDnaUtil.Application).ScreenUpdating = true; } catch { }
            try { ((dynamic)ExcelDnaUtil.Application).EnableEvents = true; } catch { }

            var answer = MessageBox.Show(
                Loc.T("save.draft.discardConfirmBody"),
                Loc.T("save.draft.discardConfirmTitle", Loc.ModuleDisplayName(contract.ModuleName)),
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);
            if (answer != DialogResult.Yes)
            {
                // Re-enable the pane's draft buttons (disabled by the double-click guard) without leaving closing mode.
                ValidationTaskPaneManager.ReenableActionsForActiveWorkbook();
                return;
            }

            AppLogger.Info($"[{contract.ModuleName}] User discarded local draft; reloading from server.");
            OfflineSave.DiscardDraftAndReload(contract);
        };
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
                    AutoBugReporter.MarkShuttingDown();
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
                        AutoBugReporter.MarkShuttingDown();
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