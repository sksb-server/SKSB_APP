namespace SKSB_App.Core.RuntimeUI.Dictionary;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.Guide;

/// <summary>
/// Entry point for the Collective Dictionary: shows the dialog and, whenever the user asks to
/// "Pick Text on Screen", closes it, runs the modal screen picker, resolves the capture to dictionary
/// entries, and re-opens the dialog on the matching Edit / Disambiguation / Update view.
/// </summary>
public static class DictionaryLauncher
{
    /// <param name="owner">Owner window for the dialog (Portal form or Excel main window wrapper).</param>
    /// <param name="excelCellResolver">
    /// Optional resolver from a screen point (pixels) to a Cell Guide snapshot. Called synchronously on the
    /// caller's thread (inside the Ribbon's QueueAsMacro callback for Excel).
    /// </param>
    public static void Show(IWin32Window? owner, Func<Point, CellGuideSnapshot?>? excelCellResolver = null)
    {
        string? langCode = null;
        IReadOnlyList<DictionaryPickerMatch>? pendingMatches = null;
        bool showNoMatchBanner = false;
        bool openUpdateView = false;

        while (true)
        {
            bool requestedPick;
            using (var dlg = new CollectiveDictionaryDialog(langCode, pendingMatches, showNoMatchBanner, openUpdateView))
            {
                if (owner != null) dlg.ShowDialog(owner);
                else dlg.ShowDialog();

                requestedPick = dlg.RequestedScreenPick;
                langCode = dlg.SelectedLangCode;
            }

            if (!requestedPick) break;

            var pick = ScreenTextPickerSession.Pick(langCode);
            if (pick.Cancelled)
            {
                pendingMatches = null;
                showNoMatchBanner = false;
                openUpdateView = true;
                continue;
            }

            List<DictionaryPickerMatch> matches = ResolveMatches(pick, langCode, excelCellResolver);
            pendingMatches = matches;
            showNoMatchBanner = matches.Count == 0;
            openUpdateView = matches.Count == 0;
        }
    }

    private static List<DictionaryPickerMatch> ResolveMatches(
        ScreenPickResult pick,
        string langCode,
        Func<Point, CellGuideSnapshot?>? excelCellResolver)
    {
        try
        {
            CellGuideSnapshot? excelCell = null;
            if (pick.IsExcelGridCell && excelCellResolver != null)
            {
                try { excelCell = excelCellResolver(pick.ScreenPoint); } catch { excelCell = null; }
            }

            return LocalizationTextMatcher.MatchFromProximityPick(pick, excelCell, langCode);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DictionaryLauncher] Match resolution failed: {ex.Message}");
            return new List<DictionaryPickerMatch>();
        }
    }
}
