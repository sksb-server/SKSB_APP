namespace SKSB_App.ProjectManager;

using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

[ComVisible(true)]
public class ProjectManagerRibbon : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabProjectManager' label='Project Manager'>
                <group id='grpProjectPersistence' label='Database &amp; Data Sync'>
                  <button id='btnSaveProjects' label='Save to System' size='large' imageMso='SaveObject' onAction='OnSaveClicked' />
                  <button id='btnLoadProjects' label='Load from System' size='large' imageMso='Refresh' onAction='OnLoadClicked' />
                  <separator id='sep1' />
                  <toggleButton id='btnToggleEditMode' label='Template Edit Mode' size='large' imageMso='DesignMode' getPressed='GetEditModePressed' onAction='OnToggleEditModeClicked' />
                  <button id='btnPurgeData' label='Clear All Sheet Data' size='large' imageMso='TableDeleteRows' getVisible='GetPurgeButtonVisible' onAction='OnPurgeDataClicked' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _ribbonUi = ribbon;
    }

    public bool GetEditModePressed(IRibbonControl control) => Events.IsTemplateEditMode;
    public bool GetPurgeButtonVisible(IRibbonControl control) => Events.IsTemplateEditMode;

    public void OnToggleEditModeClicked(IRibbonControl control, bool isPressed)
    {
        if (!isPressed && Events.IsTemplateEditMode)
        {
            var dialogResult = MessageBox.Show(
                "You are exiting Template Edit Mode.\n\nWould you like to save your template design changes to disk and reload live project data from the server?",
                "Failsafe: Exit Template Edit Mode", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (dialogResult != DialogResult.Yes)
            {
                Events.IsTemplateEditMode = true;
                _ribbonUi?.InvalidateControl("btnToggleEditMode");
                _ribbonUi?.InvalidateControl("btnPurgeData");
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    XlCall.Excel(XlCall.xlcSave);
                    WorkbookSession.PurgeAllSheetsData();

                    Events.IsTemplateEditMode = false;
                    _ribbonUi?.InvalidateControl("btnToggleEditMode");
                    _ribbonUi?.InvalidateControl("btnPurgeData");

                    if (ServerSyncCommands.LoadProjectsSilent())
                    {
                        WorkbookSession.ShowStatusConfirmation("Template saved. Live data reloaded from server.");
                        MessageBox.Show("Template changes saved successfully, and live project records have been reloaded.", "Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Template saved, but no existing project records were returned by the system.", "Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed during template exit sequence: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });
            return;
        }

        Events.IsTemplateEditMode = isPressed;
        MessageBox.Show($"Template Edit Mode is now ENABLED\n• Auto-save hijacked: OFF\n• Standard disk saving: ON", "Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        _ribbonUi?.InvalidateControl("btnPurgeData");
        _ribbonUi?.InvalidateControl("btnToggleEditMode");
    }

    public void OnPurgeDataClicked(IRibbonControl control)
    {
        if (MessageBox.Show("Are you sure you want to clear all rows on the sheets?", "Confirm Template Clear", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            WorkbookSession.PurgeAllSheetsData();
            MessageBox.Show("All sheet data has been purged.", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        // 1. MAIN THREAD: Extract data using the C API safely
        var projectList = ProjectSheetMapper.ExtractProjectRecords();

        if (projectList == null || projectList.Count == 0)
        {
            MessageBox.Show("No valid project records found to save.", "Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 2. BACKGROUND THREAD: Offload the database and file IO
        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveProjectsAsync(projectList);

                if (success)
                {
                    MessageBox.Show($"Successfully saved {count} project record(s).\n\n• Storage: {path}",
                                    "Save Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                else
                {
                    MessageBox.Show("Failed to persist records to the system.", "Project Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No saved project records found.", "Load Projects", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            ProjectSheetMapper.PopulateWorksheet(records);
            MessageBox.Show($"Successfully populated {records.Count} project record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}