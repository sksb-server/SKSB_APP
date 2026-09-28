namespace SKSB_App.StaffManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Services;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

[ComVisible(true)]
public class StaffManagerRibbon : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabStaffManager' label='Staff Manager'>
                <group id='grpStaffPersistence' label='Database &amp; Data Sync'>
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
                  <separator id='sep1' />
                  <toggleButton id='btnToggleEditMode'
                                label='Template Edit Mode'
                                size='large'
                                imageMso='DesignMode'
                                getPressed='GetEditModePressed'
                                onAction='OnToggleEditModeClicked' />
                  <button id='btnPurgeData'
                          label='Clear All Sheet Data'
                          size='large'
                          imageMso='TableDeleteRows'
                          getVisible='GetPurgeButtonVisible'
                          onAction='OnPurgeDataClicked' />
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

    public bool GetEditModePressed(IRibbonControl control)
    {
        return Events.IsTemplateEditMode;
    }

    public bool GetPurgeButtonVisible(IRibbonControl control)
    {
        return Events.IsTemplateEditMode;
    }

    public void OnSyncLookupsClicked(IRibbonControl control)
    {
        try
        {
            Events.SyncStaffEnums();
            Events.SyncStaffLookupsData();

            MessageBox.Show("Lookups synchronisation complete.\n\n" +
                            "• Static Enums (DepartmentList, LicenseClassList, etc.)\n" +
                            "• Dynamic Fleet Assets (VehicleIDList)\n\n" +
                            "Press Ctrl + F3 to inspect in Name Manager.",
                            "Lookups Synchronisation", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync trigger failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnToggleEditModeClicked(IRibbonControl control, bool isPressed)
    {
        if (!isPressed && Events.IsTemplateEditMode)
        {
            var dialogResult = MessageBox.Show(
                "You are exiting Template Edit Mode.\n\n" +
                "Would you like to save your template design changes to disk and reload live staff data from the server?\n\n" +
                "• Yes: Save template file, purge mock rows, and fetch server data.\n" +
                "• No / Cancel: Stay in Template Edit Mode.",
                "Failsafe: Exit Template Edit Mode",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

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

                    bool loaded = ServerSyncCommands.LoadStaffSilent();
                    if (loaded)
                    {
                        WorkbookSession.ShowStatusConfirmation("Template saved. Live data reloaded from server.");
                        MessageBox.Show("Template changes saved successfully, and live staff records have been reloaded.",
                                        "Staff Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Template saved, but no existing staff records were returned by the system.",
                                        "Staff Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Failed during template exit sequence: {ex.Message}",
                                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            });

            return;
        }

        Events.IsTemplateEditMode = isPressed;

        string state = "ENABLED\n• Auto-save hijacked: OFF\n• Standard disk saving: ON\n• Template wipe tools: AVAILABLE";
        MessageBox.Show($"Template Edit Mode is now {state}", "Staff Manager Configuration",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);

        _ribbonUi?.InvalidateControl("btnPurgeData");
        _ribbonUi?.InvalidateControl("btnToggleEditMode");
    }

    public void OnPurgeDataClicked(IRibbonControl control)
    {
        var confirm = MessageBox.Show(
            "Are you sure you want to clear all rows on the sheets?\n\n" +
            "This will purge data rows while preserving your headers, styling, and drop-down validations.",
            "Confirm Template Clear",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        if (confirm != DialogResult.Yes) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                WorkbookSession.PurgeAllSheetsData();
                MessageBox.Show("All sheet data has been purged. Template is now clean.",
                                "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Purge failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        var staffList = StaffSheetMapper.ExtractStaffRecords();
        if (staffList == null || staffList.Count == 0)
        {
            MessageBox.Show("No valid staff records found to save.",
                            "Staff Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveStaffAsync(staffList);
                if (!success || count == 0)
                {
                    MessageBox.Show("Failed to persist staff records to the system.",
                                    "Staff Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MessageBox.Show($"Successfully saved {count} staff record(s).\n\n• Storage: {path}",
                                "Save Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Save Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnLoadClicked(IRibbonControl control)
    {
        try
        {
            var records = ServerSyncCommands.FetchStaffRecords();
            if (records == null || records.Count == 0)
            {
                MessageBox.Show("No saved staff records found.", "Load Staff", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                StaffSheetMapper.PopulateWorksheet(records);
                MessageBox.Show($"Successfully populated {records.Count} staff record(s).",
                                "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Load Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}