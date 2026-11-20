namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.WorkbookOpen.Sheet;

/// <summary>
/// Owns the docked (bottom) "Validation Issues" Custom Task Pane, one per workbook window.
/// All public members must be called on Excel's macro thread. Every WinForms callback is re-dispatched
/// through <see cref="ExcelAsyncUtil.QueueAsMacro(ExcelAction)"/> so no COM work runs inside a click handler.
/// </summary>
public static class ValidationTaskPaneManager
{
    // =========================================================================
    // 1. PANE REGISTRY
    // =========================================================================
    private const int DefaultPaneHeight = 240;

    private sealed class PaneEntry
    {
        public CustomTaskPane Pane = null!;
        public ValidationTaskPaneControl Control = null!;
        public object? Workbook;
        public string WorkbookName = string.Empty;
    }

    private static readonly Dictionary<string, PaneEntry> _panes = new(StringComparer.OrdinalIgnoreCase);

    private static string BuildKey(object xlAppObj, object wbObj)
    {
        dynamic xlApp = xlAppObj;
        dynamic wb = wbObj;
        string fullName = string.Empty;
        try { fullName = Convert.ToString(wb.FullName) ?? string.Empty; } catch { }
        string hwnd = "0";
        try { hwnd = Convert.ToString(xlApp.ActiveWindow?.Hwnd) ?? "0"; } catch { }
        return $"{fullName}|{hwnd}";
    }

    private static bool IsAlive(PaneEntry entry)
    {
        try
        {
            if (entry.Control.IsDisposed) return false;
            _ = entry.Pane.Visible;
            _ = entry.Pane.Window; // throws once the hosting window has been closed
            return true;
        }
        catch
        {
            return false;
        }
    }

    private static void DisposeEntry(PaneEntry entry)
    {
        try { entry.Control.IssueSelected = null; } catch { }
        try { entry.Pane.Visible = false; } catch { }
        try { entry.Pane.Delete(); } catch { }
        entry.Workbook = null;
    }

    private static void PruneDeadPanes()
    {
        foreach (var key in _panes.Where(kv => !IsAlive(kv.Value)).Select(kv => kv.Key).ToList())
        {
            DisposeEntry(_panes[key]);
            _panes.Remove(key);
        }
    }

    private static PaneEntry GetOrCreate(string key, object wb, string moduleName)
    {
        PruneDeadPanes();
        if (_panes.TryGetValue(key, out var existing)) return existing;

        // Created without an explicit parent: Excel attaches it to the active window (the validated workbook).
        CustomTaskPane pane = CustomTaskPaneFactory.CreateCustomTaskPane(
            typeof(ValidationTaskPaneControl),
            $"{moduleName} — Validation Issues");
        pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionBottom;
        try { pane.Height = DefaultPaneHeight; } catch { }

        var entry = new PaneEntry
        {
            Pane = pane,
            Control = (ValidationTaskPaneControl)pane.ContentControl,
            Workbook = wb
        };
        try { entry.WorkbookName = Convert.ToString(((dynamic)wb).Name) ?? string.Empty; } catch { }
        _panes[key] = entry;
        return entry;
    }

    // =========================================================================
    // 2. SHOW / HIDE / CLOSE
    // =========================================================================
    /// <summary>
    /// Shows (creating if needed) the active workbook's pane populated with <paramref name="errors"/> /
    /// <paramref name="warnings"/>. With zero issues the pane reads "All checks passed". Call on the macro thread.
    /// </summary>
    public static void ShowIssues(
        string moduleName,
        List<ValidationIssue> errors,
        List<ValidationIssue> warnings,
        Action? onRevalidate,
        Action? onProceedWithSave,
        bool isClosing = false,
        Action? onDebugQuitAnyway = null,
        Action? onRevertHistoryEdits = null,
        Action? onCloneHistoryAsNewRows = null,
        bool showBlockingNotice = false)
    {
        errors ??= new List<ValidationIssue>();
        warnings ??= new List<ValidationIssue>();
        AppLogger.Info($"[{moduleName}] Validation: {errors.Count} error(s), {warnings.Count} warning(s){(isClosing ? " (closing)" : string.Empty)}.");

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            object wbObj = wb;

            var entry = GetOrCreate(BuildKey((object)xlApp, wbObj), wbObj, moduleName);
            entry.Workbook = wbObj;

            // Stacked-history edits: default one-click actions that fix the sheet, then re-validate (macro thread).
            Action? revertAction = onRevertHistoryEdits;
            Action? cloneAction = onCloneHistoryAsNewRows;
            var stackedWarnings = warnings.Where(w => w.IsStackedHistoryEdit).ToList();
            if (stackedWarnings.Count > 0)
            {
                revertAction ??= () =>
                {
                    StackedEntryGuard.RevertStackedEdits(wbObj, stackedWarnings);
                    onRevalidate?.Invoke();
                };
                cloneAction ??= () =>
                {
                    StackedEntryGuard.CloneStackedEditsAsNewRows(wbObj, stackedWarnings);
                    onRevalidate?.Invoke();
                };
            }

            entry.Control.IssueSelected = issue => QueueMacro(() => NavigateToIssue(issue, wbObj), "navigate");
            entry.Control.SetIssues(
                errors,
                warnings,
                WrapMacro(onRevalidate, "re-validate"),
                WrapMacro(onProceedWithSave, "proceed with save"),
                isClosing,
                WrapMacro(onDebugQuitAnyway, "debug quit"),
                WrapMacro(revertAction, "revert history edits"),
                WrapMacro(cloneAction, "clone history edits as new rows"));

            entry.Pane.Visible = true;

            string summary = errors.Count == 0 && warnings.Count == 0
                ? $"{moduleName}: all validation checks passed."
                : $"{moduleName}: {errors.Count} error(s), {warnings.Count} warning(s) — see the Validation Issues pane.";
            SheetOperations.ShowStatusConfirmation(summary);

            if (showBlockingNotice && (errors.Count > 0 || warnings.Count > 0))
            {
                ShowBlockingNotice(moduleName, errors.Count, warnings.Count, isClosing);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{moduleName}] Could not open the Validation Task Pane; falling back to the summary dialog", ex);
            ShowFallbackDialog(moduleName, errors, warnings, onProceedWithSave, isClosing, onDebugQuitAnyway);
        }
    }

    /// <summary>
    /// Hides and deletes every pane bound to <paramref name="wbObj"/> (matched by reference or workbook name),
    /// falling back to the active workbook when <paramref name="wbObj"/> is null. Call on the macro thread
    /// before force-closing a workbook so no orphaned pane survives the window.
    /// </summary>
    public static void HideAndDisposeForWorkbook(object? wbObj)
    {
        try
        {
            if (wbObj == null)
            {
                try { wbObj = ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook; } catch { }
            }

            string name = string.Empty;
            if (wbObj != null)
            {
                try { name = Convert.ToString(((dynamic)wbObj).Name) ?? string.Empty; } catch { }
            }

            var keys = _panes
                .Where(kv => (wbObj != null && ReferenceEquals(kv.Value.Workbook, wbObj))
                          || (name.Length > 0 && string.Equals(kv.Value.WorkbookName, name, StringComparison.OrdinalIgnoreCase)))
                .Select(kv => kv.Key)
                .ToList();

            foreach (var key in keys)
            {
                DisposeEntry(_panes[key]);
                _panes.Remove(key);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ValidationTaskPane] HideAndDisposeForWorkbook warning: {ex.Message}");
        }
    }

    /// <summary>Hides the active workbook window's pane (if any).</summary>
    public static void HideForActiveWorkbook()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            string key = BuildKey((object)xlApp, (object)wb);
            if (_panes.TryGetValue(key, out PaneEntry? entry))
            {
                try { entry.Pane.Visible = false; } catch { }
                return;
            }

            // Window key may differ (e.g. after a rename / Save As): fall back to the workbook name.
            string name = string.Empty;
            try { name = Convert.ToString((object)wb.Name) ?? string.Empty; } catch { }
            foreach (var e in _panes.Values.Where(p => string.Equals(p.WorkbookName, name, StringComparison.OrdinalIgnoreCase)))
            {
                try { e.Pane.Visible = false; } catch { }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ValidationTaskPane] HideForActiveWorkbook warning: {ex.Message}");
        }
    }

    /// <summary>Hides and deletes every tracked pane. Called from <c>EventsAutoOpen.AutoClose</c>.</summary>
    public static void CloseAll()
    {
        foreach (var entry in _panes.Values.ToList())
        {
            DisposeEntry(entry);
        }
        _panes.Clear();
    }

    // =========================================================================
    // 3. NAVIGATION (macro thread)
    // =========================================================================
    public static void NavigateToIssue(ValidationIssue issue, object? workbook = null)
    {
        if (issue == null) return;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = workbook;
            try { if (wb != null) _ = wb.Name; } catch { wb = null; } // workbook closed since the pane was filled
            wb ??= xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic? ws = null;
            if (!string.IsNullOrWhiteSpace(issue.SheetName))
            {
                foreach (dynamic sheet in wb.Sheets)
                {
                    string name = string.Empty;
                    try { name = Convert.ToString(sheet.Name) ?? string.Empty; } catch { }
                    if (string.Equals(name.Trim(), issue.SheetName.Trim(), StringComparison.OrdinalIgnoreCase))
                    {
                        ws = sheet;
                        break;
                    }
                }
            }

            if (ws == null)
            {
                SheetOperations.ShowStatusConfirmation(string.IsNullOrWhiteSpace(issue.SheetName)
                    ? $"Issue: {issue.Message}"
                    : $"Sheet '{issue.SheetName}' was not found. Issue: {issue.Message}");
                return;
            }

            // Hidden auxiliary sheets cannot be activated; report instead of throwing.
            int visibility = -1; // xlSheetVisible
            try { visibility = Convert.ToInt32(ws.Visible); } catch { }
            if (visibility != -1)
            {
                SheetOperations.ShowStatusConfirmation($"Sheet '{issue.SheetName}' is hidden. Issue: {issue.Message}");
                return;
            }

            try { wb.Activate(); } catch { }
            try { ws.Activate(); } catch { }

            if (issue.ExcelRow > 0)
            {
                dynamic targetCell = ws.Cells[issue.ExcelRow, issue.ExcelCol > 0 ? issue.ExcelCol : 1];
                try { xlApp.Goto(targetCell, false); }
                catch
                {
                    try { targetCell.Select(); } catch { }
                }
            }

            string location = issue.ExcelRow > 0 ? $"{issue.SheetName}!{issue.CellAddress}" : issue.SheetName;
            SheetOperations.ShowStatusConfirmation($"{issue.Severity} at {location}: {issue.Message}");
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ValidationTaskPane] Navigate to '{issue.SheetName}' {issue.CellAddress} failed: {ex.Message}");
        }
    }

    // =========================================================================
    // 4. MACRO DISPATCH & FALLBACK
    // =========================================================================
    private static Action? WrapMacro(Action? action, string label)
        => action == null ? null : () => QueueMacro(action, label);

    private static void QueueMacro(Action action, string label)
    {
        try
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { action(); }
                catch (Exception ex) { AppLogger.Error($"[ValidationTaskPane] '{label}' action failed", ex); }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[ValidationTaskPane] Could not queue '{label}' action", ex);
        }
    }

    private sealed class ExcelWindowOwner(IntPtr handle) : IWin32Window
    {
        public IntPtr Handle { get; } = handle;
    }

    /// <summary>Concise Excel-owned notice explaining why Save/Close did not complete. Never throws.</summary>
    private static void ShowBlockingNotice(string moduleName, int errorCount, int warningCount, bool isClosing)
    {
        try
        {
            SKSB_App.Core.RuntimeUI.DropDown.Arrow.HideAll();
            try { dynamic xlApp = ExcelDnaUtil.Application; xlApp.EnableEvents = true; } catch { }
            try { dynamic xlApp = ExcelDnaUtil.Application; xlApp.ScreenUpdating = true; } catch { }

            string caption;
            string text;
            MessageBoxIcon icon;
            if (errorCount > 0)
            {
                icon = MessageBoxIcon.Error;
                caption = $"{moduleName} — Save Blocked";
                string prefix = warningCount > 0
                    ? $"{errorCount} error(s) and {warningCount} warning(s) were found."
                    : $"{errorCount} error(s) were found.";
                text = $"{prefix}\n\nSaving is blocked until all errors are resolved. Use the Validation Issues panel at the bottom of Excel to click, navigate to, and fix each error.";
            }
            else
            {
                icon = MessageBoxIcon.Warning;
                caption = $"{moduleName} — Validation Warnings";
                string actionLabel = isClosing ? "Save & Close" : "Proceed with Save";
                text = $"{warningCount} warning(s) were found. Save is paused for review.\n\nUse the Validation Issues panel at the bottom of Excel to review or resolve them, or click \"{actionLabel}\" to continue.";
            }

            IntPtr hwnd = IntPtr.Zero;
            try { hwnd = ExcelDnaUtil.WindowHandle; } catch { }

            using (LoadingOverlay.Suspend())
            {
                if (hwnd != IntPtr.Zero)
                    MessageBox.Show(new ExcelWindowOwner(hwnd), text, caption, MessageBoxButtons.OK, icon);
                else
                    MessageBox.Show(text, caption, MessageBoxButtons.OK, icon);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[{moduleName}] Blocking notice failed: {ex.Message}");
        }
    }

    /// <summary>Legacy modal summary, used only if the Custom Task Pane cannot be created.</summary>
    private static void ShowFallbackDialog(
        string moduleName,
        List<ValidationIssue> errors,
        List<ValidationIssue> warnings,
        Action? onProceedWithSave,
        bool isClosing = false,
        Action? onDebugQuitAnyway = null)
    {
        if (errors.Count == 0 && warnings.Count == 0) return;
        try
        {
            using var dialog = new SummaryDialog(
                moduleName,
                errors.Select(i => i.ToDisplayString()).ToList(),
                warnings.Select(i => i.ToDisplayString()).ToList(),
                allowDebugQuit: isClosing && onDebugQuitAnyway != null);
            using (LoadingOverlay.Suspend())
            {
                dialog.ShowDialog();
            }
            if (errors.Count == 0 && dialog.Action == ValidationDialogAction.ProceedAnyway && onProceedWithSave != null)
            {
                QueueMacro(onProceedWithSave, "proceed with save");
            }
            else if (dialog.Action == ValidationDialogAction.DebugQuitAnyway && onDebugQuitAnyway != null)
            {
                QueueMacro(onDebugQuitAnyway, "debug quit");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{moduleName}] Validation summary fallback failed", ex);
            MessageBox.Show(
                string.Join("\n", errors.Concat(warnings).Take(30).Select(i => i.ToDisplayString())),
                $"{moduleName} - Data Integrity Check", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }
    }
}
