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
              <tab id='{TabId}' label='{Esc(TabLabel)}'>
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
                                imageMso='AccessSetOfflineStatus'
                                getPressed='GetSimulateOfflinePressed'
                                onAction='OnToggleOfflineClicked' />
                  <toggleButton id='btnToggleMasterDebug'
                                label='Debug Logs Active'
                                size='large'
                                imageMso='TraceDependents'
                                getPressed='GetMasterDebugPressed'
                                onAction='OnToggleMasterDebugClicked' />
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
    public void OnRibbonLoad(IRibbonUI ribbon) => RibbonUi = ribbon;

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
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
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
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
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
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
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
                    SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                    {
                        foreach (dynamic table in ws.ListObjects)
                        {
                            totalDuplicatesRemoved += TableRowManager.CleanTableDuplicateRows(table, warnings);
                        }
                    }));
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
            dynamic table = xlApp.ActiveCell?.ListObject;
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
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                dynamic table = xlApp.ActiveCell?.ListObject;
                if (table == null && ws.ListObjects.Count > 0)
                {
                    table = ws.ListObjects[1]; // Fallback to first table on sheet
                }
                if (table == null) return;

                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
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
                }));

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

    public bool GetMasterDebugPressed(IRibbonControl control)
        => DebugDialog.DebugEnabled;

    public void OnToggleMasterDebugClicked(IRibbonControl control, bool pressed)
    {
        DebugDialog.DebugEnabled = pressed;
        RibbonUi?.InvalidateControl(control.Id);
        MessageBox.Show(
            pressed
                ? "Debug diagnostics are now ENABLED across all operations."
                : "Debug diagnostics are now MUTED.",
            "Debug Log Switch",
            MessageBoxButtons.OK,
            pressed ? MessageBoxIcon.Information : MessageBoxIcon.Warning
        );
    }
#endif
}
