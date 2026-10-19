namespace SKSB_App.ProjectManager;

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
              <tab id='tabProjectManager' label='Project Manager'>
                <group id='grpProjectWorkbookHelper' label='Project Database Sync'>
                  <button id='btnSaveProjects' 
                          label='Save to System' 
                          size='large' 
                          imageMso='SaveObject' 
                          onAction='OnSaveClicked' />
                  <button id='btnLoadProjects' 
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
            Events.SyncProjectLookupsData();
            MessageBox.Show(
                "Project lookups refreshed successfully.\n\n• Lookups: Enums managed centrally in Core\n• Data: Live company entities updated in 'Data' sheet",
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
        // 1. MAIN UI THREAD: Read Excel cells safely via C-API
        var (projectList, companyList) = SheetMapper.ExtractProjectRecords();
        if (projectList == null || projectList.Count == 0)
        {
            return;
        }

        // 2. BACKGROUND WORKER: Offload disk serialization and PostgreSQL WorkbookHelper
        Task.Run(async () =>
        {
            try
            {
                // Save Companies first if any were modified
                if (companyList != null && companyList.Count > 0)
                {
                    await ServerSyncCommands.SaveCompaniesAsync(companyList);
                }

                var (success, count, path) = await ServerSyncCommands.SaveProjectsAsync(projectList);
                if (success)
                {
                    MessageBox.Show(
                        $"Successfully saved {count} project record(s).\n\n• Storage: {path}",
                        "Save Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Failed to persist projects to PostgreSQL or local storage.",
                        "WorkbookHelper Error",
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
        var records = ServerSyncCommands.FetchProjectRecords();
        var companies = ServerSyncCommands.FetchCompanyRecords() ?? new List<CompanyRecord>();

        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No project records found in storage.", "Load Projects", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SheetMapper.PopulateWorksheet(records, companies);
            MessageBox.Show($"Successfully loaded {records.Count} project record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}