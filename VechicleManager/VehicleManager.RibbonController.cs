namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Types;

[ComVisible(true)]
public class VehicleManagerRibbon : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabVehicleManager' label='Vehicle Manager'>
                <group id='grpVehiclePersistence' label='Fleet Database Sync'>
                  <button id='btnSaveVehicles' 
                          label='Save to System' 
                          size='large' 
                          imageMso='SaveObject' 
                          onAction='OnSaveClicked' />
                  <button id='btnLoadVehicles' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          onAction='OnLoadClicked' />
                  <button id='btnSyncLookups' 
                          label='Sync Lookups' 
                          size='large' 
                          imageMso='GroupCustomAttributes' 
                          onAction='OnSyncLookupsClicked' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbonUi = ribbon;

    public void OnSyncLookupsClicked(IRibbonControl control)
    {
        try
        {
            Events.SyncVehicleLookupsData();
            MessageBox.Show(
                "Fleet lookups refreshed successfully.\n\n• Lookups: Managed centrally in Core\n• Data: Live driver names updated in 'Data' tab",
                "Lookups Synced",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        // 1. MAIN UI THREAD: Synchronous read of Excel cells via C-API
        var vehicleList = SheetMapper.ExtractVehicleRecords();
        if (vehicleList == null || vehicleList.Count == 0)
        {
            return;
        }

        // 2. BACKGROUND WORKER: Offload disk serialization and PostgreSQL persistence
        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveVehiclesAsync(vehicleList);
                if (success)
                {
                    MessageBox.Show(
                        $"Successfully saved {count} vehicle record(s).\n\n• Storage: {path}",
                        "Save Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Failed to persist vehicles to PostgreSQL or local storage.",
                        "Persistence Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Save Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnLoadClicked(IRibbonControl control)
    {
        var records = ServerSyncCommands.FetchVehicleRecords();
        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No vehicle records found in storage.", "Load Vehicles", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SheetMapper.PopulateWorksheet(records);
            MessageBox.Show($"Successfully loaded {records.Count} vehicle record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}