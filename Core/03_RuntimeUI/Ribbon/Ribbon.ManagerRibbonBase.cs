namespace SKSB_App.Core.RuntimeUI.Ribbon;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Security;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;

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

    private static bool _statusHooked;
    private static Action? _statusChangedHandler; // pinned static delegate (CLAUDE.md Rule 4)
    private static ISaveContract? _statusContract;

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
    protected virtual string EmptyLoadMessage => $"No {Contract.ModuleName} records found in database or local snapshot.";

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
                {BuildDebugGroupXml()}
              </tab>
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
        return $@"
                <group id='{groupId}' label='{Esc(groupLabel)}'>
                  <button id='{loadId}'
                          label='Load from System'
                          size='large'
                          imageMso='Refresh'
                          screentip='{Esc(loadScreentip)}'
                          supertip='{Esc(loadSupertip)}'
                          onAction='OnLoadClicked' />
                  <button id='{saveId}'
                          label='Save to System'
                          size='large'
                          imageMso='FileSave'
                          screentip='{Esc(saveScreentip)}'
                          supertip='{Esc(saveSupertip)}'
                          onAction='OnSaveClicked' />
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
                          screentip='Ping PostgreSQL Server'
                          supertip='Tests the connection to the PostgreSQL server (192.168.0.20:5432) and brings the app back online when it is reachable.'
                          onAction='OnPingServerClicked' />
                </group>";
    }

    protected static string BuildRowManagementGroupXml()
    {
        return @"
                <group id='grpRowDirection' label='Row Management'>
                  <toggleButton id='btnToggleInsertTop'
                                label='Insert at Top'
                                size='large'
                                imageMso='TableInsertRowsAbove'
                                screentip='Toggle Top/Bottom Auto-Insert'
                                supertip='When enabled, typing in the first data row creates a new row at the top. When disabled, rows are appended at the bottom.'
                                getPressed='GetInsertAtTopPressed'
                                onAction='OnToggleInsertAtTopClicked' />
                  <button id='btnCleanEmptyRows'
                          label='Trim Empty Rows'
                          size='large'
                          imageMso='TableDeleteRows'
                          screentip='Trim Redundant Empty Rows'
                          supertip='Scans all tables on the active sheet and trims extra blank rows, leaving exactly one empty entry row.'
                          onAction='OnCleanEmptyRowsClicked' />
                  <button id='btnCleanDuplicateRows'
                          label='Remove Duplicates'
                          size='large'
                          imageMso='RemoveDuplicates'
                          screentip='Remove Duplicate Rows'
                          supertip='Scans all tables on the active sheet and removes identical duplicate rows, keeping the first occurrence.'
                          onAction='OnCleanDuplicateRowsClicked' />
                  <button id='btnDeleteTableRow'
                          label='Delete Row'
                          size='large'
                          imageMso='TableRowsDelete'
                          screentip='Delete Row'
                          supertip='Deletes selected row(s) in absolute master tables. Deletion is blocked on additive transaction tables.'
                          onAction='OnDeleteTableRowClicked' />
                  <toggleButton id='btnToggleTotalRow'
                                label='Total Row'
                                size='large'
                                imageMso='AutoSum'
                                screentip='Toggle Total Row'
                                supertip='Shows or hides the summary total row at the bottom of the active table.'
                                getPressed='GetTotalRowPressed'
                                onAction='OnToggleTotalRowClicked' />
                </group>";
    }

    protected static string BuildDebugGroupXml()
    {
#if DEBUG
        return @"
                <group id='grpDebugControls' label='Debug Controls'>
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
                </group>";
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
        _statusContract = Contract;
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

    // =========================================================================
    // 3a. SERVER STATUS CALLBACKS
    // =========================================================================
    public string GetTabLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? $"{TabLabel} [OFFLINE MODE]" : TabLabel;

    public string GetServerStatusGroupLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? "⚠ OFFLINE MODE" : "Server Status";

    public string GetServerModeLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? "STATUS: OFFLINE MODE" : "STATUS: ONLINE";

    public string GetServerSourceLabel(IRibbonControl control)
#if DEBUG
        => ServerEnvironment.IsOffline ? "Source: Local Offline Snapshot" : "Source: Debug Server (debug)";
#else
        => ServerEnvironment.IsOffline ? "Source: Local Offline Snapshot" : "Source: PostgreSQL Server";
#endif

    public string GetServerPendingLabel(IRibbonControl control)
        => Contract.HasPendingUpload ? "Pending Upload: YES (Uncommitted)" : "Pending Upload: None";

    public string GetPingServerButtonLabel(IRibbonControl control)
        => ServerEnvironment.IsOffline ? "Ping Server\n(Go Online)" : "Ping Server\n(Online)";

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
                    "Offline simulation is ON.\n\nDisable the simulation and reconnect to the server?",
                    "Simulated Offline", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (answer == DialogResult.Yes) ServerEnvironment.SimulateOffline = false;
            }
#endif
            bool online = ServerEnvironment.PingServer(out string err);
            RibbonUi?.Invalidate();

            if (!online)
            {
                MessageBox.Show(
                    $"The PostgreSQL server is still unreachable.\n\n{err}\n\nThe app remains in OFFLINE MODE.",
                    "Server Unreachable", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Ribbon callback: defer all worksheet / dialog work to the macro thread.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    Contract.UpdateWindowCaption();
                    SheetOperations.ShowStatusConfirmation($"Server reachable: {Contract.ModuleName} is back in Online Mode.");

                    var choice = MessageBox.Show(
                        "Server is reachable. The app is back in ONLINE MODE.\n\n" +
                        "  Yes:     Save the current worksheet records to the server now.\n" +
                        "  No:      Discard any pending offline upload and reload the latest records from the server.\n" +
                        "  Cancel:  Stay on the current sheet without syncing yet.",
                        $"{Contract.ModuleName} - Back Online",
                        MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

                    if (choice == DialogResult.Yes)
                    {
                        Contract.Save();
                    }
                    else if (choice == DialogResult.No)
                    {
                        Contract.ClearPendingUpload();
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
            Contract.SyncLookups?.Invoke(false);

            var (isEmpty, populate) = Contract.FetchForPopulate();
            if (populate == null || (isEmpty && !PopulateWhenEmpty))
            {
                if (ShowEmptyLoadDialog)
                {
                    MessageBox.Show(EmptyLoadMessage, $"Load {Contract.ModuleName}", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                return;
            }

            // Worksheet mutation must run on Excel's macro thread, never inside the Ribbon callback.
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    populate();
                    Contract.UpdateWindowCaption();
                    SheetOperations.ShowStatusConfirmation(ServerEnvironment.IsOffline
                        ? $"[OFFLINE MODE] {Contract.ModuleName} records loaded from local offline snapshot."
                        : $"{Contract.ModuleName} records loaded from server (Online Mode).");
                    RibbonUi?.Invalidate();
                }
                catch (Exception ex)
                {
                    AppLogger.Error($"[{Contract.ModuleName}] Populate after load failed", ex);
                }
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed loading {Contract.ModuleName} records:\n{ex.Message}", "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public virtual void OnSaveClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() => Contract.Save());
    }

    // =========================================================================
    // 4. ROW MANAGEMENT CALLBACKS
    // =========================================================================
    public bool GetInsertAtTopPressed(IRibbonControl control)
        => TableRowManager.IsInsertAtTopEnabled;

    public void OnToggleInsertAtTopClicked(IRibbonControl control, bool isPressed)
    {
        TableRowManager.IsInsertAtTopEnabled = isPressed;
        RibbonUi?.InvalidateControl(control.Id);

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic? ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws != null)
                {
                    TableRowManager.CleanActiveSheetTables(ws);
                }
            }
            catch { }

            SheetOperations.ShowStatusConfirmation(
                isPressed
                    ? "Insert at TOP active: Spawns at top, redundant bottom rows trimmed."
                    : "Insert at BOTTOM active: Spawns at bottom, redundant top rows trimmed."
            );
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

                int tablesCleaned = TableRowManager.CleanActiveSheetTables(ws);
                SheetOperations.ShowStatusConfirmation(
                    $"Cleaned empty rows across {tablesCleaned} table(s). Exactly 1 empty row retained."
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed cleaning empty rows:\n{ex.Message}", "Row Cleanup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

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
                    foreach (dynamic table in ws.ListObjects)
                    {
                        totalDuplicatesRemoved += TableRowManager.CleanTableDuplicateRows(table, warnings);
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
                    SheetOperations.ShowStatusConfirmation($"Removed {totalDuplicatesRemoved} duplicate row(s).");
                    MessageBox.Show(
                        $"Successfully removed {totalDuplicatesRemoved} duplicate row(s) from the active sheet.\n\n" +
                        string.Join("\n", warnings),
                        "Duplicates Removed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    SheetOperations.ShowStatusConfirmation("No duplicate rows found on the active sheet.");
                    MessageBox.Show(
                        "No duplicate rows found in any tables on the active sheet.",
                        "Clean Duplicates",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to remove duplicate rows:\n{ex.Message}", "Cleanup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
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
                SheetOperations.ShowStatusConfirmation(isPressed ? "Total row enabled." : "Total row hidden.");
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed toggling Total Row", ex);
            }
        });
    }

    // =========================================================================
    // 5. DEBUG CALLBACKS
    // =========================================================================
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

#endif
}
