namespace SKSB_App.Core.RuntimeUI.Ribbon;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.BugReport;
using SKSB_App.Core.RuntimeUI.Dictionary;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Guide;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.Validation;

/// <summary>
/// Shared Ribbon tab for every Manager module: tab/group XML assembly plus the Load / Save,
/// Row Management and Debug callbacks. Modules override the descriptor properties, add their own
/// groups through <see cref="GetCustomGroupsXml"/> and declare only domain-specific button handlers.
///
/// IMPORTANT: this class is deliberately NOT generic. Excel resolves Ribbon callbacks by name through
/// the object's COM class interface (IDispatch), and the CLR refuses to expose a class interface for
/// any type that has a generic class anywhere in its inheritance chain - every callback (including
/// ones declared on the derived controller) would silently stop firing. The module's typed
/// <see cref="SaveContract{T}"/> is therefore consumed through the non-generic <see cref="ISaveContract"/>.
/// </summary>
[ComVisible(true)]
public abstract class ManagerRibbonBase : ExcelRibbon
{
    protected static IRibbonUI? RibbonUi { get; private set; }
    private static string? _activeTabId;

    private static bool _statusHooked;
    private static Action? _statusChangedHandler; // pinned static delegate (CLAUDE.md Rule 4)
    private static ISaveContract? _statusContract;

    private static bool _languageHooked;
    private static Action? _languageChangedHandler; // pinned static delegate (CLAUDE.md Rule 4)

    // =========================================================================
    // 1. MODULE DESCRIPTOR & CUSTOMIZATION HOOKS
    // =========================================================================
    protected abstract ISaveContract Contract { get; }
    protected abstract string TabId { get; }
    protected abstract string TabLabel { get; }

    /// <summary>The Load / Save group, normally produced by <see cref="BuildPersistenceGroupXml"/>.</summary>
    protected abstract string PersistenceGroupXml { get; }

    /// <summary>Extra module-specific groups, inserted between the persistence and row management groups.</summary>
    protected virtual string GetCustomGroupsXml() => string.Empty;

    protected virtual bool ShowRowManagementGroup => true;

    /// <summary>When false, an empty fetch result shows <see cref="EmptyLoadMessage"/> instead of populating.</summary>
    protected virtual bool PopulateWhenEmpty => false;
    protected virtual bool ShowEmptyLoadDialog => true;
    protected virtual string EmptyLoadMessage => Loc.T("ribbon.load.emptyBody", Loc.ModuleDisplayName(Contract.ModuleName));

    /// <summary>Decides which columns get a SUM when the Total Row is switched on.</summary>
    protected virtual bool ShouldSumColumnInTotalRow(string header) =>
        header.IndexOf("Cost", StringComparison.OrdinalIgnoreCase) >= 0 ||
        header.IndexOf("Amount", StringComparison.OrdinalIgnoreCase) >= 0;

    // =========================================================================
    // 2. RIBBON XML ASSEMBLY
    // =========================================================================
    public override string GetCustomUI(string ribbonId)
    {
        return $@"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='{TabId}' getLabel='GetTabLabel'>
                {BuildServerStatusGroupXml()}
                {PersistenceGroupXml}
                {GetCustomGroupsXml()}
                {(ShowRowManagementGroup ? BuildRowManagementGroupXml() : string.Empty)}
                {BuildSupportGroupXml()}
              </tab>
              {BuildDebugTabXml()}
            </tabs>
          </ribbon>
        </customUI>";
    }

    protected static string BuildPersistenceGroupXml(
        string groupId,
        string groupLabel,
        string loadId,
        string loadScreentip,
        string loadSupertip,
        string saveId,
        string saveScreentip,
        string saveSupertip,
        string extraControlsXml = "")
    {
        // The module-supplied English text is the fallback when no "ribbon.<controlId>.<kind>" key exists.
        _persistenceFallbackText[$"{groupId}.label"] = groupLabel;
        _persistenceFallbackText[$"{loadId}.screentip"] = loadScreentip;
        _persistenceFallbackText[$"{loadId}.supertip"] = loadSupertip;
        _persistenceFallbackText[$"{saveId}.screentip"] = saveScreentip;
        _persistenceFallbackText[$"{saveId}.supertip"] = saveSupertip;

        return $@"
                <group id='{groupId}' getLabel='GetPersistenceGroupLabel'>
                  <button id='{loadId}'
                          getLabel='GetLoadButtonLabel'
                          size='large'
                          imageMso='Refresh'
                          getScreentip='GetLoadButtonScreentip'
                          getSupertip='GetLoadButtonSupertip'
                          onAction='OnLoadClicked' />
                  <button id='{saveId}'
                          getLabel='GetSaveButtonLabel'
                          size='large'
                          imageMso='FileSave'
                          getScreentip='GetSaveButtonScreentip'
                          getSupertip='GetSaveButtonSupertip'
                          onAction='OnSaveClicked' />
                  <button id='btnValidateWorkbook'
                          getLabel='GetValidateButtonLabel'
                          size='large'
                          imageMso='Spelling'
                          getScreentip='GetValidateButtonScreentip'
                          getSupertip='GetValidateButtonSupertip'
                          onAction='OnValidateClicked' />
                  {extraControlsXml}
                </group>";
    }

    /// <summary>Online/Offline status display plus the Ping Server (go online) button.</summary>
    protected static string BuildServerStatusGroupXml()
    {
        return @"
                <group id='grpServerStatus' getLabel='GetServerStatusGroupLabel'>
                  <box id='boxServerStatus' boxStyle='vertical'>
                    <labelControl id='lblServerMode' getLabel='GetServerModeLabel' />
                    <labelControl id='lblServerSource' getLabel='GetServerSourceLabel' />
                    <labelControl id='lblServerPending' getLabel='GetServerPendingLabel' />
                  </box>
                  <button id='btnPingServer'
                          size='large'
                          getLabel='GetPingServerButtonLabel'
                          getImage='GetPingServerImage'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnPingServerClicked' />
                </group>";
    }

    protected static string BuildRowManagementGroupXml()
    {
        return @"
                <group id='grpRowDirection' getLabel='GetLocalizedLabel'>
                  <button id='btnInsertTableRow'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='TableInsertRowsBelow'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnInsertTableRowClicked' />
                  <toggleButton id='btnToggleInsertTop'
                                getLabel='GetLocalizedLabel'
                                size='large'
                                imageMso='TableInsertRowsAbove'
                                getScreentip='GetLocalizedScreentip'
                                getSupertip='GetLocalizedSupertip'
                                getPressed='GetInsertAtTopPressed'
                                onAction='OnToggleInsertAtTopClicked' />
                  <button id='btnCleanEmptyRows'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='TableDeleteRows'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnCleanEmptyRowsClicked' />
                  <button id='btnCleanDuplicateRows'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='RemoveDuplicates'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnCleanDuplicateRowsClicked' />
                  <button id='btnDeleteTableRow'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='TableRowsDelete'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnDeleteTableRowClicked' />
                  <toggleButton id='btnToggleTotalRow'
                                getLabel='GetLocalizedLabel'
                                size='large'
                                imageMso='AutoSum'
                                getScreentip='GetLocalizedScreentip'
                                getSupertip='GetLocalizedSupertip'
                                getPressed='GetTotalRowPressed'
                                onAction='OnToggleTotalRowClicked' />
                </group>";
    }

    protected static string BuildSupportGroupXml()
    {
        return @"
                <group id='grpSupport' getLabel='GetLocalizedLabel'>
                  <button id='btnReportBug'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='ReviewNewComment'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnReportBugClicked' />
                  <toggleButton id='btnToggleGuidePanel'
                                getLabel='GetLocalizedLabel'
                                size='large'
                                imageMso='Help'
                                getScreentip='GetLocalizedScreentip'
                                getSupertip='GetLocalizedSupertip'
                                getPressed='GetGuidePanelPressed'
                                onAction='OnToggleGuidePanelClicked' />
                  <button id='btnCollectiveDictionary'
                          getLabel='GetLocalizedLabel'
                          size='large'
                          imageMso='TranslateDocument'
                          getScreentip='GetLocalizedScreentip'
                          getSupertip='GetLocalizedSupertip'
                          onAction='OnCollectiveDictionaryClicked' />
                  <dropDown id='ddLanguage'
                            getLabel='GetLanguageDropdownLabel'
                            getItemCount='GetLanguageItemCount'
                            getItemID='GetLanguageItemId'
                            getItemLabel='GetLanguageItemLabel'
                            getSelectedItemIndex='GetLanguageSelectedIndex'
                            getScreentip='GetLanguageScreentip'
                            getSupertip='GetLanguageSupertip'
                            onAction='OnLanguageSelected' />
                </group>";
    }

    /// <summary>Ribbon control id -> localization key prefix for the shared (non-module) controls.</summary>
    private static readonly Dictionary<string, string> _sharedControlKeys = new(StringComparer.Ordinal)
    {
        ["btnPingServer"] = "ribbon.ping",
        ["grpRowDirection"] = "ribbon.rows.group",
        ["btnInsertTableRow"] = "ribbon.rows.insert",
        ["btnToggleInsertTop"] = "ribbon.rows.insertTop",
        ["btnCleanEmptyRows"] = "ribbon.rows.trim",
        ["btnCleanDuplicateRows"] = "ribbon.rows.dedupe",
        ["btnDeleteTableRow"] = "ribbon.rows.delete",
        ["btnToggleTotalRow"] = "ribbon.rows.total",
        ["grpSupport"] = "ribbon.support.group",
        ["btnReportBug"] = "ribbon.support.bug",
        ["btnToggleGuidePanel"] = "ribbon.support.guide",
        ["btnCollectiveDictionary"] = "ribbon.support.dictionary",
    };

    /// <summary>English text supplied by derived controllers through <see cref="BuildPersistenceGroupXml"/>.</summary>
    private static readonly ConcurrentDictionary<string, string> _persistenceFallbackText = new(StringComparer.Ordinal);

    protected static string BuildDebugTabXml()
    {
#if DEBUG
        return @"
              <tab id='tabSksbDebug' label='Debug'>
                <group id='grpDebugServer' label='Server &amp; Database'>
                  <toggleButton id='btnToggleOffline'
                                label='Simulate Server Offline'
                                size='large'
                                imageMso='HighImportance'
                                getPressed='GetSimulateOfflinePressed'
                                onAction='OnToggleOfflineClicked' />
                  <button id='btnMirrorReleaseData'
                          label='Mirror Data from Release'
                          size='large'
                          imageMso='DatabaseCopyDatabaseFile'
                          screentip='Mirror Release Data to Debug Server'
                          supertip='Resets the isolated Debug schema (debug.*) to an exact copy of the live Release database (public.*) and reloads this workbook.'
                          onAction='OnMirrorReleaseDataClicked' />
                </group>
                <group id='grpDebugTemplate' label='Template Tools'>
                  <button id='btnPurgeSheetData'
                          label='Purge Sheet Data'
                          size='large'
                          imageMso='ClearAll'
                          screentip='Purge All Sheet Data'
                          supertip='Clears all operational table rows, top-bar parameter values, and extra Daily Schedule board slots while preserving headers, formulas, formatting, and Lookups/Data sheets.'
                          onAction='OnPurgeSheetDataClicked' />
                </group>
              </tab>";
#else
        return string.Empty;
#endif
    }

    /// <summary>Escapes text for use inside a single-quoted Ribbon XML attribute.</summary>
    protected static string Esc(string? text) => SecurityElement.Escape(text ?? string.Empty) ?? string.Empty;

    // =========================================================================
    // 3. PERSISTENCE CALLBACKS
    // =========================================================================
    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        RibbonUi = ribbon;
        _activeTabId = TabId;
        _statusContract = Contract;
        HookLanguageChanged();
        if (_statusHooked) return;
        _statusHooked = true;
        _statusChangedHandler = () =>
        {
            try { RibbonUi?.Invalidate(); } catch { }
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try { _statusContract?.UpdateWindowCaption(); } catch { }
            });
        };
        ServerEnvironment.StatusChanged += _statusChangedHandler;
    }

    /// <summary>Activates this module's custom Manager tab on the Excel Ribbon.</summary>
    public static void ActivateManagerTab()
    {
        try
        {
            if (RibbonUi != null && !string.IsNullOrWhiteSpace(_activeTabId))
            {
                RibbonUi.ActivateTab(_activeTabId);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ManagerRibbonBase] Could not activate tab '{_activeTabId}': {ex.Message}");
        }
    }

    // =========================================================================
    // 3a. LOCALIZATION CALLBACKS
    // =========================================================================
    private static void HookLanguageChanged()
    {
        Loc.ReloadFromPreferences();
        Loc.WatchPreferences();
        if (_languageHooked) return;
        _languageHooked = true;
        CollectiveDictionaryRepository.SyncFromServerInBackground();
        // May fire on a thread-pool thread (prefs file watcher): marshal the Ribbon refresh to Excel's macro thread.
        _languageChangedHandler = () =>
        {
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try { RibbonUi?.Invalidate(); } catch { }
                    try
                    {
                        dynamic xlApp = ExcelDnaUtil.Application;
                        RuleRegistry.RefreshWorkbookHeaderNotes(xlApp.ActiveWorkbook);
                    }
                    catch { }
                });
            }
            catch { }
        };
        Loc.LanguageChanged += _languageChangedHandler;
    }

    /// <summary>Module display name in the active language ("Staff Manager" -> "Pengurus Kakitangan").</summary>
    protected string ModuleDisplayName => Loc.ModuleDisplayName(Contract.ModuleName);

    private static string ResolveSharedControlText(IRibbonControl control, string kind)
    {
        string id = control?.Id ?? string.Empty;
        if (!_sharedControlKeys.TryGetValue(id, out string? prefix)) return string.Empty;
        // Group entries already map to their full label key.
        string key = id.StartsWith("grp", StringComparison.Ordinal) ? prefix : $"{prefix}.{kind}";
        return Loc.T(key);
    }

    private static string ResolvePersistenceText(IRibbonControl control, string kind)
    {
        string id = control?.Id ?? string.Empty;
        if (Loc.TryT($"ribbon.{id}.{kind}", out string text)) return text;
        return _persistenceFallbackText.TryGetValue($"{id}.{kind}", out string? fallback) ? fallback : string.Empty;
    }

    public string GetLocalizedLabel(IRibbonControl control) => ResolveSharedControlText(control, "label");
    public string GetLocalizedScreentip(IRibbonControl control) => ResolveSharedControlText(control, "screentip");
    public string GetLocalizedSupertip(IRibbonControl control) => ResolveSharedControlText(control, "supertip");

    protected string ResolveCustomControlText(IRibbonControl control, string kind)
    {
        string id = control?.Id ?? string.Empty;
        if (string.IsNullOrEmpty(id)) return string.Empty;
        return Loc.T($"ribbon.custom.{id}.{kind}");
    }

    public string GetCustomLabel(IRibbonControl control) => ResolveCustomControlText(control, "label");
    public string GetCustomScreentip(IRibbonControl control) => ResolveCustomControlText(control, "screentip");
    public string GetCustomSupertip(IRibbonControl control) => ResolveCustomControlText(control, "supertip");

    public string GetPersistenceGroupLabel(IRibbonControl control) => ResolvePersistenceText(control, "label");
    public string GetLoadButtonLabel(IRibbonControl control) => Loc.T("ribbon.load.label");
    public string GetLoadButtonScreentip(IRibbonControl control) => ResolvePersistenceText(control, "screentip");
    public string GetLoadButtonSupertip(IRibbonControl control) => ResolvePersistenceText(control, "supertip");
    public string GetSaveButtonLabel(IRibbonControl control) => Loc.T("ribbon.save.label");
    public string GetSaveButtonScreentip(IRibbonControl control) => ResolvePersistenceText(control, "screentip");
    public string GetSaveButtonSupertip(IRibbonControl control) => ResolvePersistenceText(control, "supertip");
    public string GetValidateButtonLabel(IRibbonControl control) => Loc.T("ribbon.validate.label");
    public string GetValidateButtonScreentip(IRibbonControl control) => Loc.T("ribbon.validate.screentip");
    public string GetValidateButtonSupertip(IRibbonControl control) => Loc.T("ribbon.validate.supertip");

    public string GetLanguageDropdownLabel(IRibbonControl control) => Loc.T("language.label");
    public string GetLanguageScreentip(IRibbonControl control) => Loc.T("language.screentip");
    public string GetLanguageSupertip(IRibbonControl control) => Loc.T("language.supertip");

    public int GetLanguageItemCount(IRibbonControl control) => AppLanguageInfo.All.Count;

    public string GetLanguageItemId(IRibbonControl control, int index)
        => index >= 0 && index < AppLanguageInfo.All.Count
            ? "lang_" + AppLanguageInfo.ToCode(AppLanguageInfo.All[index]).Replace('-', '_')
            : $"lang_{index}";

    public string GetLanguageItemLabel(IRibbonControl control, int index)
        => index >= 0 && index < AppLanguageInfo.All.Count
            ? AppLanguageInfo.GetDisplayName(AppLanguageInfo.All[index])
            : string.Empty;

    public int GetLanguageSelectedIndex(IRibbonControl control)
        => AppLanguageInfo.IndexOf(Loc.CurrentLanguage);

    public void OnLanguageSelected(IRibbonControl control, string selectedId, int selectedIndex)
    {
        try
        {
            if (selectedIndex < 0 || selectedIndex >= AppLanguageInfo.All.Count) return;
            var language = AppLanguageInfo.All[selectedIndex];
            if (language == Loc.CurrentLanguage) return;
            Loc.SetLanguage(language); // raises LanguageChanged -> queued RibbonUi.Invalidate()
            try { RibbonUi?.Invalidate(); } catch { }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ManagerRibbonBase] Language switch failed: {ex.Message}");
        }
    }

    // =========================================================================
    // 3b. CELL GUIDE PANEL CALLBACKS
    // =========================================================================
    public bool GetGuidePanelPressed(IRibbonControl control)
    {
        try { return GuideTaskPaneManager.IsVisibleForActiveWorkbook(); }
        catch { return false; }
    }

    public void OnToggleGuidePanelClicked(IRibbonControl control, bool isPressed)
    {
        string moduleName = Contract.ModuleName;
        // Ribbon callback: create / show the Custom Task Pane on Excel's macro thread.
        ExcelAsyncUtil.QueueAsMacro(() => GuideTaskPaneManager.SetVisible(moduleName, isPressed));
    }

    /// <summary>Refreshes the Help Guide toggle (pane shown / closed with its X button / window switch).</summary>
    public static void InvalidateGuideToggleControl()
    {
        try { RibbonUi?.InvalidateControl("btnToggleGuidePanel"); } catch { }
    }

    // =========================================================================
    // 3c. SERVER STATUS CALLBACKS
    // =========================================================================
    public string GetTabLabel(IRibbonControl control)
    {
        string name = Loc.ModuleDisplayName(TabLabel);
        return ServerEnvironment.IsOffline ? Loc.T("ribbon.tab.offline", name) : name;
    }

    public string GetServerStatusGroupLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? Loc.T("ribbon.status.groupOffline") : Loc.T("ribbon.status.group");

    public string GetServerModeLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? Loc.T("ribbon.status.modeOffline") : Loc.T("ribbon.status.modeOnline");

    public string GetServerSourceLabel(IRibbonControl control)
#if DEBUG
        => ServerEnvironment.IsOffline ? Loc.T("ribbon.status.sourceOffline") : Loc.T("ribbon.status.sourceDebug");
#else
        => ServerEnvironment.IsOffline ? Loc.T("ribbon.status.sourceOffline") : Loc.T("ribbon.status.sourceServer");
#endif

    public string GetServerPendingLabel(IRibbonControl control)
        => Contract.HasLocalDraft ? Loc.T("ribbon.status.localDraft")
         : Contract.HasPendingUpload ? Loc.T("ribbon.status.pendingYes") : Loc.T("ribbon.status.pendingNone");

    /// <summary>Re-queries every Ribbon callback (status labels, draft / pending indicators).</summary>
    public static void InvalidateRibbon()
    {
        try { RibbonUi?.Invalidate(); } catch { }
    }

    public string GetPingServerButtonLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? Loc.T("ribbon.ping.labelOffline") : Loc.T("ribbon.ping.labelOnline");

    public object? GetPingServerImage(IRibbonControl control)
    {
        try
        {
            string mso = ServerEnvironment.IsOffline ? "HighImportance" : "Refresh";
            dynamic xlApp = ExcelDnaUtil.Application;
            return xlApp.CommandBars.GetImageMso(mso, 32, 32);
        }
        catch
        {
            return null;
        }
    }

    public void OnPingServerClicked(IRibbonControl control)
    {
        try
        {
#if DEBUG
            if (ServerEnvironment.SimulateOffline)
            {
                var answer = MessageBox.Show(
                    Loc.T("ribbon.ping.simulatedBody"),
                    Loc.T("ribbon.ping.simulatedTitle"), MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer == DialogResult.Yes) ServerEnvironment.SimulateOffline = false;
            }
#endif
            bool online = ServerEnvironment.PingServer(out string err);
            RibbonUi?.Invalidate();

            if (!online)
            {
                MessageBox.Show(
                    Loc.T("ribbon.ping.unreachableBody", err),
                    Loc.T("ribbon.ping.unreachableTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Ribbon callback: defer all worksheet / dialog work to the macro thread.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Contract.UpdateWindowCaption();
                    SheetOperations.ShowStatusConfirmation(Loc.T("ribbon.ping.backOnlineStatus", ModuleDisplayName));

                    var choice = MessageBox.Show(
                        Loc.T("ribbon.ping.backOnlineBody"),
                        Loc.T("ribbon.ping.backOnlineTitle", ModuleDisplayName),
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (choice == DialogResult.Yes)
                    {
                        Contract.Save();
                    }
                    else if (choice == DialogResult.No)
                    {
                        Contract.ClearPendingUpload();
                        Contract.ClearLocalDraft();
                        OnLoadClicked(control);
                    }
                    RibbonUi?.Invalidate();
                }
                catch (Exception ex)
                {
                    AppLogger.Error($"[{Contract.ModuleName}] Post-ping handling failed", ex);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{Contract.ModuleName}] Ping server failed", ex);
        }
    }

    public virtual void OnLoadClicked(IRibbonControl control)
    {
        try
        {
            if (Contract.HasLocalDraft)
            {
                var discard = MessageBox.Show(
                    Loc.T("ribbon.load.discardDraftBody"),
                    Loc.T("ribbon.load.discardDraftTitle", ModuleDisplayName),
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);
                if (discard != DialogResult.Yes) return;

                Contract.ClearLocalDraft();
                ValidationTaskPaneManager.HideForActiveWorkbook();
                RibbonUi?.Invalidate();
            }

            bool isEmpty;
            Action? populate;
            // Stable English operation key keeps LoadingTimeEstimator history independent of the UI language.
            string loadOperationKey = $"{Contract.ModuleName} — Loading from System";
            using (LoadingOverlay.BeginScope(Loc.T("ribbon.load.overlayTitle", ModuleDisplayName), Loc.T("ribbon.load.fetching"), operationKey: loadOperationKey))
            {
                (isEmpty, populate) = Contract.FetchForPopulate();
            }

            if (populate == null || (isEmpty && !PopulateWhenEmpty))
            {
                // Queued via QueueAsMacro (outside a synchronous scope) so lookup lists still refresh.
                Contract.SyncLookups?.Invoke(false);
                if (ShowEmptyLoadDialog)
                {
                    MessageBox.Show(EmptyLoadMessage, Loc.T("ribbon.load.emptyTitle", ModuleDisplayName), MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }

            // Worksheet mutation must run on Excel's macro thread, never inside the Ribbon callback.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    using (LoadingOverlay.BeginScope(Loc.T("ribbon.load.overlayTitle", ModuleDisplayName), Loc.T("ribbon.load.syncLookups"), operationKey: loadOperationKey))
                    using (SyncHelper.BeginSynchronousScope())
                    {
                        LoadingOverlay.Report(Loc.T("ribbon.load.syncLookups"));
                        Contract.SyncLookups?.Invoke(false);

                        LoadingOverlay.Report(Loc.T("ribbon.load.populating"));
                        populate();
                    }
                    Contract.UpdateWindowCaption();
                    SheetOperations.ShowStatusConfirmation(ServerEnvironment.IsOffline
                        ? Loc.T("ribbon.load.doneOffline", ModuleDisplayName)
                        : Loc.T("ribbon.load.doneOnline", ModuleDisplayName));
                    RibbonUi?.Invalidate();

                    // Overlay scope has ended and Excel state is restored: safe to show advisory dialogs.
                    try { Contract.NotifyAfterLoad(); }
                    catch (Exception hookEx) { AppLogger.Warn($"[{Contract.ModuleName}] OnAfterLoad hook failed: {hookEx.Message}"); }
                }
                catch (Exception ex)
                {
                    AppLogger.Error($"[{Contract.ModuleName}] Populate after load failed", ex);
                }
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(Loc.T("ribbon.load.errorBody", ModuleDisplayName, ex.Message), Loc.T("ribbon.load.errorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public virtual void OnSaveClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() => Contract.Save());
    }

    public virtual void OnValidateClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() => Contract.Validate());
    }

    // =========================================================================
    // 4. ROW MANAGEMENT CALLBACKS
    // =========================================================================
    public bool GetInsertAtTopPressed(IRibbonControl control)
        => TableRowManager.IsInsertAtTopEnabled;

    /// <summary>Refreshes the per-table toggle buttons (sheet switch / table selection change).</summary>
    public static void InvalidateTableToggleControls()
    {
        try { RibbonUi?.InvalidateControl("btnToggleInsertTop"); } catch { }
        try { RibbonUi?.InvalidateControl("btnToggleTotalRow"); } catch { }
        try { RibbonUi?.InvalidateControl("btnToggleGuidePanel"); } catch { }
    }

    public void OnToggleInsertAtTopClicked(IRibbonControl control, bool isPressed)
    {
        dynamic? activeTable = TableRowManager.ResolveActiveTable();
        string? activeTableName = null;
        try { activeTableName = activeTable?.Name?.ToString(); } catch { }

        if (!string.IsNullOrWhiteSpace(activeTableName))
            TableRowManager.SetInsertAtTopForTable(activeTableName, isPressed);
        else
            TableRowManager.IsInsertAtTopEnabled = isPressed;

        try { RibbonUi?.InvalidateControl(control.Id); } catch { }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic? table = TableRowManager.ResolveActiveTable();
                if (table != null)
                {
                    TableRowManager.CleanSingleTableEmptyRows(table);
                }
                else
                {
                    dynamic xlApp = ExcelDnaUtil.Application;
                    dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                    if (ws != null) TableRowManager.CleanActiveSheetTables(ws);
                }
            }
            catch { }

            bool hasTable = !string.IsNullOrWhiteSpace(activeTableName);
            string message = isPressed
                ? (hasTable ? Loc.T("ribbon.rows.insertTopOnTable", activeTableName!) : Loc.T("ribbon.rows.insertTopOn"))
                : (hasTable ? Loc.T("ribbon.rows.insertBottomOnTable", activeTableName!) : Loc.T("ribbon.rows.insertBottomOn"));
            SheetOperations.ShowStatusConfirmation(message);
        });
    }

    public void OnCleanEmptyRowsClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                int tablesCleaned;
                using (LoadingOverlay.BeginScope(Loc.T("ribbon.rows.trimOverlay", ModuleDisplayName), Loc.T("ribbon.rows.trimStep"),
                           operationKey: $"{Contract.ModuleName} — Trimming Empty Rows"))
                {
                    tablesCleaned = TableRowManager.CleanActiveSheetTables(ws);
                }
                SheetOperations.ShowStatusConfirmation(Loc.T("ribbon.rows.trimDone", tablesCleaned));
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.T("ribbon.rows.trimErrorBody", ex.Message), Loc.T("ribbon.rows.trimErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    /// <summary>Called on the macro thread after duplicate rows were removed (COM flags already restored).</summary>
    protected virtual void OnAfterDuplicateRowsCleaned(int removedCount) { }

    public void OnCleanDuplicateRowsClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                bool origEvents = true;
                bool origUpdating = true;
                int origCalculation = -4105; // xlCalculationAutomatic
                try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
                try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }

                int totalDuplicatesRemoved = 0;
                var warnings = new List<string>();

                try
                {
                    using (LoadingOverlay.BeginScope(Loc.T("ribbon.rows.dedupeOverlay", ModuleDisplayName), Loc.T("ribbon.rows.dedupeStep"),
                               operationKey: $"{Contract.ModuleName} — Removing Duplicates"))
                    {
                        foreach (dynamic table in ws.ListObjects)
                        {
                            totalDuplicatesRemoved += TableRowManager.CleanTableDuplicateRows(table, warnings);
                        }
                    }
                }
                finally
                {
                    // Restored before any modal dialog is shown below.
                    try { xlApp.Calculation = origCalculation; } catch { }
                    try { xlApp.EnableEvents = origEvents; } catch { }
                    try { xlApp.ScreenUpdating = origUpdating; } catch { }
                }

                if (totalDuplicatesRemoved > 0)
                {
                    try { OnAfterDuplicateRowsCleaned(totalDuplicatesRemoved); }
                    catch (Exception hookEx) { AppLogger.Warn($"[ManagerRibbonBase] OnAfterDuplicateRowsCleaned failed: {hookEx.Message}"); }
                    SheetOperations.ShowStatusConfirmation(Loc.T("ribbon.rows.dedupeStatus", totalDuplicatesRemoved));
                    MessageBox.Show(
                        Loc.T("ribbon.rows.dedupeBody", totalDuplicatesRemoved) + "\n\n" +
                        string.Join("\n", warnings),
                        Loc.T("ribbon.rows.dedupeTitle"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    SheetOperations.ShowStatusConfirmation(Loc.T("ribbon.rows.dedupeNoneStatus"));
                    MessageBox.Show(
                        Loc.T("ribbon.rows.dedupeNoneBody"),
                        Loc.T("ribbon.rows.dedupeNoneTitle"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.T("ribbon.rows.dedupeErrorBody", ex.Message), Loc.T("ribbon.rows.dedupeErrorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnInsertTableRowClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(TableRowManager.InsertRowCommand);
    }

    public void OnDeleteTableRowClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(TableRowManager.DeleteSelectedRowCommand);
    }

    public bool GetTotalRowPressed(IRibbonControl control)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? table = xlApp.ActiveCell?.ListObject;
            if (table != null)
            {
                return (bool)table.ShowTotals;
            }
        }
        catch { }
        return false;
    }

    public void OnToggleTotalRowClicked(IRibbonControl control, bool isPressed)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                dynamic? table = xlApp.ActiveCell?.ListObject;
                if (table == null && ws.ListObjects.Count > 0)
                {
                    table = ws.ListObjects[1]; // Fallback to first table on sheet
                }
                if (table == null) return;

                table.ShowTotals = isPressed;

                if (isPressed)
                {
                    // Configure default calculation for summable columns
                    int colCount = (int)table.ListColumns.Count;
                    for (int c = 1; c <= colCount; c++)
                    {
                        dynamic col = table.ListColumns[c];
                        string header = col.Name?.ToString() ?? string.Empty;

                        if (ShouldSumColumnInTotalRow(header))
                        {
                            col.TotalsCalculation = 1; // xlTotalsCalculationSum
                            try { col.Total.NumberFormat = "#,##0.00"; } catch { }
                        }
                    }
                }

                RibbonUi?.InvalidateControl(control.Id);
                SheetOperations.ShowStatusConfirmation(isPressed ? Loc.T("ribbon.rows.totalOn") : Loc.T("ribbon.rows.totalOff"));
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed toggling Total Row", ex);
            }
        });
    }

    // =========================================================================
    // 5. SUPPORT CALLBACKS
    // =========================================================================
    public void OnReportBugClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;

                string activeSheet = string.Empty;
                try { activeSheet = Convert.ToString(xlApp.ActiveSheet?.Name) ?? string.Empty; } catch { }

                // Capture while the workbook is still unobstructed by the dialog.
                byte[]? screenshot = BugReportRepository.CaptureScreenPng();

                // Modal dialog exception (CLAUDE.md Rule 5.2): Excel must repaint & respond while the dialog is open.
                try { xlApp.ScreenUpdating = true; } catch { }
                try { xlApp.EnableEvents = true; } catch { }
                try { Arrow.HideAll(); } catch { }

                var context = new BugReportContext
                {
                    SourceModule = Contract.ModuleName,
                    ActiveSheet = activeSheet,
                    AppVersion = BugReportRepository.ResolveAppVersion(),
                    IsOffline = ServerEnvironment.IsOffline,
                    InitialScreenshotPng = screenshot
                };

                using var dialog = new BugReportSubmitDialog(context);
                dialog.ShowDialog(new BugReportOwnerWindow(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[{Contract.ModuleName}] Report Bug dialog failed", ex);
                MessageBox.Show(Loc.T("common.bugReport.openFailed", ex.Message), Loc.T("ribbon.support.bug.label"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnCollectiveDictionaryClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;

                // Modal dialog exception (CLAUDE.md Rule 5.2): Excel must repaint & respond while the dialog is open.
                try { xlApp.ScreenUpdating = true; } catch { }
                try { xlApp.EnableEvents = true; } catch { }
                try { Arrow.HideAll(); } catch { }

                using var dialog = new CollectiveDictionaryDialog();
                dialog.ShowDialog(new BugReportOwnerWindow(ExcelDnaUtil.WindowHandle));
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[{Contract.ModuleName}] Collective Dictionary dialog failed", ex);
                MessageBox.Show(Loc.T("dict.openFailed", ex.Message), Loc.T("ribbon.support.dictionary.label"), MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    // =========================================================================
    // 6. DEBUG CALLBACKS
    // =========================================================================
    protected virtual void OnAfterPurgeSheetData() { }

#if DEBUG
    public bool GetSimulateOfflinePressed(IRibbonControl control)
        => ServerEnvironment.SimulateOffline;

    public void OnToggleOfflineClicked(IRibbonControl control, bool pressed)
        => ServerEnvironment.ToggleSimulatedOffline(RibbonUi, control.Id, pressed);

    public void OnMirrorReleaseDataClicked(IRibbonControl control)
    {
        if (ServerEnvironment.IsOffline)
        {
            MessageBox.Show("The server is offline. Disable 'Simulate Server Offline' or Ping Server first.",
                "Mirror Data from Release", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var answer = MessageBox.Show(
            "Reset the Debug database (schema 'debug') to an exact mirror of the live Release database (schema 'public')?\n\nAll current Debug-only database changes across all tables will be overwritten with live Release data, and this workbook will reload.",
            "Mirror Data from Release", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
        if (answer != DialogResult.Yes) return;

        if (!ServerEnvironment.MirrorReleaseToDebug(out int tables, out int rows, out string err))
        {
            MessageBox.Show($"Failed to mirror Release database:\n\n{err}", "Mirror Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        AppLogger.Info($"[Debug Server] Mirrored {tables} tables ({rows} total rows) from 'public' to 'debug'.");
        Contract.ClearPendingUpload();
        OnLoadClicked(control);
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try { SheetOperations.ShowStatusConfirmation($"[DEBUG SERVER] Mirrored {rows} row(s) across {tables} table(s) from Release and reloaded {Contract.ModuleName}."); }
            catch { }
        });
    }

    public void OnPurgeSheetDataClicked(IRibbonControl control)
    {
        var answer = MessageBox.Show(
            "Purge all sheet data?\n\nCLEARED: table rows (reduced to 1 clean row), top-bar parameter values, and extra Daily Schedule board slots.\nPRESERVED: headers, labels, formulas, formatting, validations, and the Lookups/Data sheets.",
            "Purge Sheet Data", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (answer != DialogResult.Yes) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SheetOperations.PurgeAllSheetsData();
            OnAfterPurgeSheetData();
            SheetOperations.ShowStatusConfirmation("Purged all sheet data (tables, top bars, and schedule board reset to clean template state).");
        });
    }
#endif
}
