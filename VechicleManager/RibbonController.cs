namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

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
                <group id='grpVehiclePersistence' label='Database &amp; Data Sync'>
                  <button id='btnSaveVehicles' label='Save to System' size='large' imageMso='SaveObject' onAction='OnSaveClicked' />
                  <button id='btnLoadVehicles' label='Load from System' size='large' imageMso='Refresh' onAction='OnLoadClicked' />
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
                "You are exiting Template Edit Mode.\n\nWould you like to save your template design changes to disk and reload live vehicle data from the server?",
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

                    if (ServerSyncCommands.LoadVehiclesSilent())
                    {
                        WorkbookSession.ShowStatusConfirmation("Template saved. Live data reloaded from server.");
                        MessageBox.Show("Template changes saved successfully, and live vehicle records have been reloaded.", "Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    else
                    {
                        MessageBox.Show("Template saved, but no existing vehicle records were returned by the system.", "Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        MessageBox.Show($"Template Edit Mode is now ENABLED\n• Auto-save hijacked: OFF\n• Standard disk saving: ON", "Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
        // 1. MAIN UI THREAD: Safely extract data using the C API
        var vehicleList = VehicleSheetMapper.ExtractVehicleRecords();
        if (vehicleList == null || vehicleList.Count == 0)
        {
            MessageBox.Show("No valid vehicle records found to save.",
                            "Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        // 2. BACKGROUND THREAD: Offload heavy database and file persistence
        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveVehiclesAsync(vehicleList);
                if (!success || count == 0)
                {
                    MessageBox.Show("Failed to persist vehicle records to the system.",
                                    "Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                MessageBox.Show($"Successfully saved {count} vehicle record(s).\n\n• Storage: {path}",
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
        var records = ServerSyncCommands.FetchVehicleRecords();
        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No saved vehicle records found.", "Load Vehicles", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            VehicleSheetMapper.PopulateWorksheet(records);
            MessageBox.Show($"Successfully populated {records.Count} vehicle record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}