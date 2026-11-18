namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;

/// <summary>
/// Failsafe that makes users confirm before they modify or save a record dated before today.
/// Module-agnostic: Schedule Manager (working date, attendance punches), Vehicle Manager logs, Project Manager
/// daily summaries, ... all call the same helper so the wording and behaviour stay consistent.
/// <para>
/// Every prompt follows the CLAUDE.md modal-dialog rule: floating widgets are hidden, ScreenUpdating and
/// EnableEvents are switched back on for the duration of the dialog, and the caller's state is restored in a
/// <c>finally</c> block.
/// </para>
/// </summary>
public static class PastRecordGuard
{
    private const string Caption = "Past Record Warning";
    private const int MaxDatesListed = 5;

    // Edit approvals already given: "<session>|<workbook>|<yyyy-MM-dd>". Keyed by workbook so closing a workbook
    // (or opening another) starts a fresh session; keyed by date so moving to another past date asks again.
    private static readonly HashSet<string> _approvedEdits = new(StringComparer.OrdinalIgnoreCase);
    private static readonly object _lock = new();

    // =========================================================================
    // 1. SAVE CONFIRMATION
    // =========================================================================

    /// <summary>
    /// Returns true when <paramref name="recordDate"/> is today or later (no prompt), or when the user confirms saving
    /// a past record. <paramref name="todayOverride"/> exists for tests and replays.
    /// </summary>
    public static bool ConfirmPastRecordSave(DateOnly recordDate, string moduleName, string? details = null, DateOnly? todayOverride = null) =>
        ConfirmPastRecordsSave(new[] { recordDate }, moduleName, details, todayOverride);

    /// <summary>Multi-row variant: prompts once, listing the past dates found (oldest first).</summary>
    public static bool ConfirmPastRecordsSave(IEnumerable<DateOnly> dates, string moduleName, string? details = null, DateOnly? todayOverride = null)
    {
        var past = PastDates(dates, todayOverride);
        if (past.Count == 0) return true;

        string subject = past.Count == 1
            ? $"a past {moduleName} record ({Format(past[0])})"
            : $"{past.Count} past {moduleName} records ({DescribeDates(past)})";

        string message =
            $"You are about to save changes to {subject}.\n\n" +
            "Historical records should normally remain locked. Are you sure you want to proceed?" +
            (string.IsNullOrWhiteSpace(details) ? string.Empty : $"\n\n{details}");

        bool confirmed = Prompt(message);
        AppLogger.Info($"[PastRecordGuard] Save of {past.Count} past {moduleName} date(s) (oldest {Format(past[0])}) {(confirmed ? "confirmed" : "declined")}.");
        return confirmed;
    }

    // =========================================================================
    // 2. EDIT WARNING (once per session + date, not once per cell)
    // =========================================================================

    /// <summary>
    /// Warns when the user starts editing a past record. An approval is remembered for
    /// (<paramref name="sessionKey"/>, workbook, date), so it is asked once per past date rather than on every cell
    /// change; a decline is never remembered. Returns true when editing may continue.
    /// </summary>
    public static bool ConfirmPastRecordEdit(string sessionKey, DateOnly recordDate, string moduleName, DateOnly? todayOverride = null)
    {
        if (PastDates(new[] { recordDate }, todayOverride).Count == 0) return true;

        string key = EditKey(sessionKey, recordDate);
        lock (_lock)
        {
            if (_approvedEdits.Contains(key)) return true;
        }

        bool confirmed = Prompt(
            $"You are about to edit a past {moduleName} record ({Format(recordDate)}).\n\n" +
            "Historical records should normally remain locked. Are you sure you want to proceed?");

        if (confirmed)
        {
            lock (_lock) { _approvedEdits.Add(key); }
        }
        AppLogger.Info($"[PastRecordGuard] Edit of past {moduleName} record {Format(recordDate)} {(confirmed ? "confirmed" : "declined")}.");
        return confirmed;
    }

    /// <summary>Forgets edit approvals - for one session key, or all of them. Call when the date changes or the module unloads.</summary>
    public static void ResetEditSession(string? sessionKey = null)
    {
        lock (_lock)
        {
            if (sessionKey == null) { _approvedEdits.Clear(); return; }
            _approvedEdits.RemoveWhere(k => k.StartsWith(sessionKey + "|", StringComparison.OrdinalIgnoreCase));
        }
    }

    // =========================================================================
    // 3. HELPERS
    // =========================================================================
    private static List<DateOnly> PastDates(IEnumerable<DateOnly> dates, DateOnly? todayOverride)
    {
        DateOnly today = todayOverride ?? DateOnly.FromDateTime(DateTime.Today);
        return dates.Where(d => d < today).Distinct().OrderBy(d => d).ToList();
    }

    private static string Format(DateOnly d) => d.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);

    private static string DescribeDates(List<DateOnly> past)
    {
        string listed = string.Join(", ", past.Take(MaxDatesListed).Select(Format));
        return past.Count > MaxDatesListed ? $"{listed}, ..." : listed;
    }

    private static string EditKey(string sessionKey, DateOnly date)
    {
        string workbook = string.Empty;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            workbook = xlApp.ActiveWorkbook?.FullName?.ToString() ?? string.Empty;
        }
        catch { }
        return $"{sessionKey}|{workbook}|{Format(date)}";
    }

    private sealed class ExcelWindow : IWin32Window
    {
        public IntPtr Handle { get; }
        public ExcelWindow(IntPtr handle) => Handle = handle;
    }

    /// <summary>Yes / No warning dialog; No is the default button. Excel state is restored whatever happens.</summary>
    private static bool Prompt(string message)
    {
        dynamic? xlApp = null;
        bool origUpdating = true;
        bool origEvents = true;
        IntPtr owner = IntPtr.Zero;
        try
        {
            xlApp = ExcelDnaUtil.Application;
            try { Arrow.HideAll(); } catch { }
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = true; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = true; } catch { }
            try { owner = Arrow.GetExcelHwnd(xlApp); } catch { }

            var result = owner != IntPtr.Zero
                ? MessageBox.Show(new ExcelWindow(owner), message, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                : MessageBox.Show(message, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            return result == DialogResult.Yes;
        }
        finally
        {
            if (xlApp != null)
            {
                try { xlApp.EnableEvents = origEvents; } catch { }
                try { xlApp.ScreenUpdating = origUpdating; } catch { }
            }
        }
    }
}
