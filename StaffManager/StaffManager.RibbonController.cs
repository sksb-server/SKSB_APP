namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Types;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabStaffManager' label='Staff Manager'>
                <group id='grpStaffPersistence' label='Staff Database Sync'>
                  <button id='btnSaveStaff' 
                          label='Save to System' 
                          size='large' 
                          imageMso='SaveObject' 
                          onAction='OnSaveClicked' />
                  <button id='btnLoadStaff' 
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
            Events.SyncStaffLookupsData();
            MessageBox.Show(
                "Staff lookups refreshed successfully.\n\n• Lookups: Enums managed centrally in Core\n• Data: Live vehicle plates updated in 'Data' sheet",
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
        var staffList = SheetMapper.ExtractStaffRecords();
        if (staffList == null || staffList.Count == 0)
        {
            return;
        }

        // 2. BACKGROUND WORKER: Offload disk serialization and PostgreSQL persistence
        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveStaffAsync(staffList);
                if (success)
                {
                    MessageBox.Show(
                        $"Successfully saved {count} staff record(s).\n\n• Storage: {path}",
                        "Save Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Failed to persist staff to PostgreSQL or local storage.",
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
        var records = ServerSyncCommands.FetchStaffRecords();
        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No staff records found in storage.", "Load Staff", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SheetMapper.PopulateWorksheet(records);
            MessageBox.Show($"Successfully loaded {records.Count} staff record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}