namespace SKSB_App.VehicleManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
#if DEBUG
        string debugGroupXml = @"
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
        string debugGroupXml = string.Empty;
#endif

        return $@"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabVehicleManager' label='Vehicle Manager'>
                <group id='grpVehiclePersistence' label='Fleet Database Sync'>
                  <button id='btnLoadVehicles' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          screentip='Load Vehicle Records'
                          supertip='Fetches all vehicle master records and logs from PostgreSQL or offline storage.'
                          onAction='OnLoadClicked' />
                  <button id='btnSaveVehicles' 
                          label='Save to System' 
                          size='large' 
                          imageMso='FileSave' 
                          screentip='Save Vehicle Records'
                          supertip='Validates and commits fleet specifications and maintenance records.'
                          onAction='OnSaveClicked' />
                </group>

                <!-- ROW MANAGEMENT GROUP: TOGGLE DIRECTION & MANUAL CLEANUP BUTTON -->
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
                          supertip='Scans all tables on the active sheet and trims extra blank rows, leaving exactly one empty entry row based on the Insert at Top setting.'
                          onAction='OnCleanEmptyRowsClicked' />
                </group>

                {debugGroupXml}
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbonUi = ribbon;

    public void OnLoadClicked(IRibbonControl control)
    {
        if (!Guard.IsVehicleWorkbook()) return;

        try
        {
            Events.SyncVehicleLookupsData(isInteractive: false);

            var records = ServerSyncCommands.FetchVehicleRecords();
            if (records == null || records.Count == 0)
            {
                MessageBox.Show("No vehicle records found in database or local snapshot.",
                    "Load Vehicles", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SheetMapper.PopulateWorksheet(records);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed loading vehicle records:\n{ex.Message}", "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        if (!Guard.IsVehicleWorkbook()) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SaveEntry.Save(Events.VehicleSyncDescriptor);
        });
    }

    public bool GetInsertAtTopPressed(IRibbonControl control)
        => TableRowManager.IsInsertAtTopEnabled;

    public void OnToggleInsertAtTopClicked(IRibbonControl control, bool isPressed)
    {
        TableRowManager.IsInsertAtTopEnabled = isPressed;
        _ribbonUi?.InvalidateControl(control.Id);

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

#if DEBUG
    public bool GetSimulateOfflinePressed(IRibbonControl control)
        => ServerEnvironment.SimulateOffline;

    public void OnToggleOfflineClicked(IRibbonControl control, bool pressed)
    {
        ServerEnvironment.SimulateOffline = pressed;
        _ribbonUi?.InvalidateControl(control.Id);
        MessageBox.Show(
            pressed
                ? "Server is now SIMULATED OFFLINE.\nAll PostgreSQL database operations will fail immediately to test offline local stashing."
                : "Server is now ONLINE.\nNormal PostgreSQL database connectivity restored.",
            "Debug Server Mode",
            MessageBoxButtons.OK,
            pressed ? MessageBoxIcon.Warning : MessageBoxIcon.Information
        );
    }

    public bool GetMasterDebugPressed(IRibbonControl control)
        => DebugDialog.DebugEnabled;

    public void OnToggleMasterDebugClicked(IRibbonControl control, bool pressed)
    {
        DebugDialog.DebugEnabled = pressed;
        _ribbonUi?.InvalidateControl(control.Id);
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