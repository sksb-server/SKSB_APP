namespace SKSB_App.Core.RuntimeUI.Guide;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.Ribbon;

/// <summary>
/// Owns the floating, drag-dockable "Cell Guide" Custom Task Pane (one per workbook window).
/// The pane starts FLOATING so it never fights the bottom-docked Validation Issues pane; users can drag it
/// to any edge (Excel's native CTP docking). All public members must be called on Excel's macro thread.
/// </summary>
public static class GuideTaskPaneManager
{
    // =========================================================================
    // 1. PANE REGISTRY
    // =========================================================================
    private const int DefaultPaneWidth = 340;
    private const int DefaultPaneHeight = 480;

    private sealed class PaneEntry
    {
        public CustomTaskPane Pane = null!;
        public GuideTaskPaneControl Control = null!;
        public string WorkbookName = string.Empty;
        public string ModuleName = string.Empty;
    }

    private static readonly Dictionary<string, PaneEntry> _panes = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>O(1) fast-exit flag for <see cref="OnSelectionChanged"/>: true while at least one pane may be visible.</summary>
    private static bool _anyPaneVisible;
    private static bool _suppressPreferencePersist;

    private static CustomTaskPaneEvents_VisibleStateChangeEventHandler? _visibleStateChangedHandler; // pinned (CLAUDE.md Rule 4)
    private static CustomTaskPaneEvents_DockPositionStateChangeEventHandler? _dockPositionChangedHandler; // pinned (CLAUDE.md Rule 4)
    private static Action? _languageChangedHandler;                                                  // pinned (CLAUDE.md Rule 4)
    private static bool _languageHooked;

    private delegate void GuideSheetSelectionChangeDelegate(dynamic sh, dynamic target);
    private delegate void GuideSheetActivateDelegate(dynamic sh);
    private static dynamic? _xlAppHooked;
    private static GuideSheetSelectionChangeDelegate? _selectionHandler;                             // pinned (CLAUDE.md Rule 4)
    private static GuideSheetActivateDelegate? _activateHandler;                                     // pinned (CLAUDE.md Rule 4)

    private static string BuildKey(dynamic xlApp, dynamic wb)
    {
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
        try { if (_visibleStateChangedHandler != null) entry.Pane.VisibleStateChange -= _visibleStateChangedHandler; } catch { }
        try { if (_dockPositionChangedHandler != null) entry.Pane.DockPositionStateChange -= _dockPositionChangedHandler; } catch { }
        try { entry.Pane.Visible = false; } catch { }
        try { entry.Pane.Delete(); } catch { }
    }

    private static void PruneDeadPanes()
    {
        foreach (var key in _panes.Where(kv => !IsAlive(kv.Value)).Select(kv => kv.Key).ToList())
        {
            DisposeEntry(_panes[key]);
            _panes.Remove(key);
        }
    }

    private static void EnsureSelectionEventsHooked(dynamic xlApp)
    {
        if (_xlAppHooked != null) return;
        try
        {
            _selectionHandler ??= OnSelectionChanged;
            _activateHandler ??= OnSheetActivated;
            xlApp.SheetSelectionChange += _selectionHandler;
            xlApp.SheetActivate += _activateHandler;
            _xlAppHooked = xlApp;
        }
        catch (Exception ex)
        {
            try { if (_selectionHandler != null) xlApp.SheetSelectionChange -= _selectionHandler; } catch { }
            try { if (_activateHandler != null) xlApp.SheetActivate -= _activateHandler; } catch { }
            AppLogger.Warn($"[GuideTaskPane] Could not hook Excel selection events: {ex.Message}");
        }
    }

    private static PaneEntry GetOrCreate(string key, dynamic xlApp, dynamic wb)
    {
        PruneDeadPanes();
        if (_panes.TryGetValue(key, out var existing))
        {
            EnsureSelectionEventsHooked(xlApp);
            return existing;
        }

        HookLanguageChanged();
        EnsureSelectionEventsHooked(xlApp);
        _visibleStateChangedHandler ??= OnPaneVisibleStateChanged;
        _dockPositionChangedHandler ??= OnPaneDockPositionChanged;

        // Created without an explicit parent: Excel attaches it to the active window.
        CustomTaskPane pane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(GuideTaskPaneControl), Loc.Tip("guide.pane.title"));
        pane.DockPosition = MsoCTPDockPosition.msoCTPDockPositionFloating;
        try { pane.Width = DefaultPaneWidth; } catch { }
        try { pane.Height = DefaultPaneHeight; } catch { }
        pane.VisibleStateChange += _visibleStateChangedHandler;
        pane.DockPositionStateChange += _dockPositionChangedHandler;

        var entry = new PaneEntry
        {
            Pane = pane,
            Control = (GuideTaskPaneControl)pane.ContentControl
        };
        entry.Control.ConfigureHostWindow(
            ExcelDnaUtil.WindowHandle,
            pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionFloating);
        try { entry.WorkbookName = Convert.ToString(wb.Name) ?? string.Empty; } catch { }
        _panes[key] = entry;
        return entry;
    }

    private static PaneEntry? FindForActiveWindow()
    {
        if (_panes.Count == 0) return null;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return null;
            return _panes.TryGetValue(BuildKey(xlApp, wb), out PaneEntry? entry) ? entry : null;
        }
        catch
        {
            return null;
        }
    }

    private static void RecomputeAnyVisible()
    {
        bool any = false;
        foreach (var entry in _panes.Values)
        {
            try { if (entry.Pane.Visible) { any = true; break; } } catch { }
        }
        _anyPaneVisible = any;
    }

    // =========================================================================
    // 2. SHOW / HIDE / CLOSE
    // =========================================================================
    private static void PersistGuidePreference(bool showOnStartup)
    {
        if (_suppressPreferencePersist) return;
        try
        {
            var prefs = UserPreferences.Load();
            if (prefs.ShowGuideOnStartup == showOnStartup) return;
            prefs.ShowGuideOnStartup = showOnStartup;
            UserPreferences.Save(prefs);
        }
        catch { }
    }

    /// <summary>
    /// Suppresses persistence of visibility changes while a workbook or Excel is closing so
    /// native Custom Task Pane teardown never overwrites <see cref="UserPreferencesData.ShowGuideOnStartup"/>.
    /// </summary>
    public static void BeginShutdown()
    {
        _suppressPreferencePersist = true;
    }

    /// <summary>Shows (creating if needed) or hides the active workbook window's Cell Guide pane.</summary>
    public static void SetVisible(string moduleName, bool visible, bool persistPreference = true)
    {
        try
        {
            Loc.ReloadFromPreferences();
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            string key = BuildKey(xlApp, wb);

            if (persistPreference)
            {
                _suppressPreferencePersist = false;
            }

            if (!visible)
            {
                PruneDeadPanes();
                if (_panes.TryGetValue(key, out PaneEntry? hideEntry))
                {
                    if (!persistPreference)
                    {
                        bool prev = _suppressPreferencePersist;
                        _suppressPreferencePersist = true;
                        try { hideEntry.Pane.Visible = false; } catch { }
                        finally { _suppressPreferencePersist = prev; }
                    }
                    else
                    {
                        try { hideEntry.Pane.Visible = false; } catch { }
                    }
                }
                RecomputeAnyVisible();
                if (persistPreference) PersistGuidePreference(false);
                return;
            }

            var entry = GetOrCreate(key, xlApp, wb);
            entry.ModuleName = moduleName ?? string.Empty;
            entry.Control.SetModuleName(moduleName);
            RefreshEntry(entry, xlApp);
            if (!persistPreference)
            {
                bool prev = _suppressPreferencePersist;
                _suppressPreferencePersist = true;
                try { entry.Pane.Visible = true; }
                finally { _suppressPreferencePersist = prev; }
            }
            else
            {
                entry.Pane.Visible = true;
            }
            _anyPaneVisible = true;
            if (persistPreference) PersistGuidePreference(true);
            AppLogger.Info($"[{moduleName}] Cell Guide pane shown.");
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{moduleName}] Could not toggle the Cell Guide pane", ex);
        }
        finally
        {
            ManagerRibbonBase.InvalidateGuideToggleControl();
        }
    }

    /// <summary>True when the active workbook window's Cell Guide pane is visible (Ribbon toggle state).</summary>
    public static bool IsVisibleForActiveWorkbook()
    {
        var entry = FindForActiveWindow();
        if (entry == null) return false;
        try { return !entry.Control.IsDisposed && entry.Pane.Visible; } catch { return false; }
    }

    /// <summary>Hides and deletes every tracked pane. Called from <c>EventsAutoOpen.AutoClose</c>.</summary>
    public static void CloseAll()
    {
        _suppressPreferencePersist = true;
        foreach (var entry in _panes.Values.ToList())
        {
            DisposeEntry(entry);
        }
        _panes.Clear();
        _anyPaneVisible = false;

        if (_languageHooked && _languageChangedHandler != null)
        {
            try { Loc.LanguageChanged -= _languageChangedHandler; } catch { }
            _languageHooked = false;
        }

        if (_xlAppHooked != null)
        {
            try { if (_selectionHandler != null) _xlAppHooked.SheetSelectionChange -= _selectionHandler; } catch { }
            try { if (_activateHandler != null) _xlAppHooked.SheetActivate -= _activateHandler; } catch { }
            _xlAppHooked = null;
        }
    }

    // =========================================================================
    // 3. LIVE UPDATES
    // =========================================================================
    /// <summary>SheetSelectionChange hook (macro thread). Returns immediately when no Guide pane is visible.</summary>
    public static void OnSelectionChanged(dynamic sh, dynamic target)
    {
        if (!_anyPaneVisible) return;
        try
        {
            var entry = FindForActiveWindow();
            if (entry == null || entry.Control.IsDisposed) return;
            bool visible;
            try { visible = entry.Pane.Visible; } catch { return; }
            if (!visible) return;

            entry.Control.UpdateContext(CellContextInspector.Inspect(sh, target));
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[GuideTaskPane] Selection update failed: {ex.Message}");
        }
    }

    /// <summary>SheetActivate hook (macro thread). Refreshes the active pane with the newly activated sheet's selection.</summary>
    private static void OnSheetActivated(dynamic sh)
    {
        if (!_anyPaneVisible) return;
        try
        {
            var entry = FindForActiveWindow();
            if (entry == null || entry.Control.IsDisposed) return;
            bool visible;
            try { visible = entry.Pane.Visible; } catch { return; }
            if (!visible) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            RefreshEntry(entry, xlApp);
            entry.Control.AdvanceTip();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[GuideTaskPane] Sheet activation refresh failed: {ex.Message}");
        }
    }

    private static void RefreshEntry(PaneEntry entry, dynamic xlApp)
    {
        try
        {
            dynamic? sheet = null;
            dynamic? selection = null;
            try { sheet = xlApp.ActiveSheet; } catch { }
            try { selection = xlApp.Selection; } catch { }

            // Selection may be a shape / chart: only inspect real ranges.
            bool isRange = false;
            try { isRange = selection != null && selection!.Cells != null; } catch { }
            entry.Control.UpdateContext(isRange
                ? CellContextInspector.Inspect(sheet, selection)
                : CellGuideSnapshot.Empty());
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[GuideTaskPane] Initial refresh failed: {ex.Message}");
        }
    }

    /// <summary>Fires when the user closes the pane with its X button (or Excel hides it): sync the Ribbon toggle and persist preference.</summary>
    private static void OnPaneVisibleStateChanged(CustomTaskPane pane)
    {
        try
        {
            if (!_suppressPreferencePersist)
            {
                try
                {
                    bool visible = pane.Visible;
                    PersistGuidePreference(visible);
                }
                catch { }
            }
            RecomputeAnyVisible();
            ManagerRibbonBase.InvalidateGuideToggleControl();
        }
        catch { }
    }

    /// <summary>Fires when the user docks or undocks the pane: updates floating-idle transparency state.</summary>
    private static void OnPaneDockPositionChanged(CustomTaskPane pane)
    {
        try
        {
            if (pane.ContentControl is GuideTaskPaneControl control && !control.IsDisposed)
            {
                bool isFloating = pane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionFloating;
                control.SetFloatingState(isFloating);
            }
        }
        catch { }
    }

    private static void HookLanguageChanged()
    {
        if (_languageHooked) return;
        _languageHooked = true;
        // May fire on a thread-pool thread (prefs file watcher): marshal to Excel's macro thread.
        _languageChangedHandler ??= () =>
        {
            try { ExcelAsyncUtil.QueueAsMacro(RefreshAllLanguage); } catch { }
        };
        Loc.LanguageChanged += _languageChangedHandler;
    }

    private static void RefreshAllLanguage()
    {
        bool prevSuppress = _suppressPreferencePersist;
        _suppressPreferencePersist = true;
        try
        {
            PruneDeadPanes();
            dynamic? xlApp = null;
            try { xlApp = ExcelDnaUtil.Application; } catch { }

            string desiredTitle = Loc.Tip("guide.pane.title");

            foreach (var kv in _panes.ToList())
            {
                string key = kv.Key;
                PaneEntry entry = kv.Value;
                try
                {
                    bool wasVisible = false;
                    try { wasVisible = entry.Pane.Visible; } catch { }

                    string currentTitle = string.Empty;
                    try { currentTitle = entry.Pane.Title ?? string.Empty; } catch { }

                    if (!string.Equals(currentTitle, desiredTitle, StringComparison.Ordinal))
                    {
                        if (!wasVisible)
                        {
                            DisposeEntry(entry);
                            _panes.Remove(key);
                            continue;
                        }

                        var dockPos = MsoCTPDockPosition.msoCTPDockPositionFloating;
                        int width = DefaultPaneWidth;
                        int height = DefaultPaneHeight;
                        object? window = null;
                        try { dockPos = entry.Pane.DockPosition; } catch { }
                        try { width = entry.Pane.Width; } catch { }
                        try { height = entry.Pane.Height; } catch { }
                        try { window = entry.Pane.Window; } catch { }

                        string workbookName = entry.WorkbookName;
                        string moduleName = entry.ModuleName;
                        DisposeEntry(entry);
                        _panes.Remove(key);

                        _visibleStateChangedHandler ??= OnPaneVisibleStateChanged;
                        _dockPositionChangedHandler ??= OnPaneDockPositionChanged;
                        CustomTaskPane newPane;
                        if (window != null)
                        {
                            try
                            {
                                newPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(GuideTaskPaneControl), desiredTitle, window);
                            }
                            catch
                            {
                                newPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(GuideTaskPaneControl), desiredTitle);
                            }
                        }
                        else
                        {
                            newPane = CustomTaskPaneFactory.CreateCustomTaskPane(typeof(GuideTaskPaneControl), desiredTitle);
                        }

                        try { newPane.DockPosition = dockPos; } catch { }
                        try { newPane.Width = width; } catch { }
                        try { newPane.Height = height; } catch { }
                        newPane.VisibleStateChange += _visibleStateChangedHandler;
                        newPane.DockPositionStateChange += _dockPositionChangedHandler;

                        var newEntry = new PaneEntry
                        {
                            Pane = newPane,
                            Control = (GuideTaskPaneControl)newPane.ContentControl,
                            WorkbookName = workbookName,
                            ModuleName = moduleName
                        };
                        newEntry.Control.ConfigureHostWindow(
                            ExcelDnaUtil.WindowHandle,
                            newPane.DockPosition == MsoCTPDockPosition.msoCTPDockPositionFloating);
                        newEntry.Control.SetModuleName(moduleName);
                        _panes[key] = newEntry;

                        if (xlApp != null)
                        {
                            RefreshEntry(newEntry, xlApp);
                        }
                        try { newEntry.Pane.Visible = true; } catch { }
                    }
                    else
                    {
                        if (xlApp != null && wasVisible)
                        {
                            RefreshEntry(entry, xlApp);
                        }
                        entry.Control.RefreshLanguage();
                    }
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"[GuideTaskPane] Language refresh failed for pane '{key}': {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[GuideTaskPane] Language refresh failed: {ex.Message}");
        }
        finally
        {
            _suppressPreferencePersist = prevSuppress;
            try { RecomputeAnyVisible(); } catch { }
            try { ManagerRibbonBase.InvalidateGuideToggleControl(); } catch { }
        }
    }
}
