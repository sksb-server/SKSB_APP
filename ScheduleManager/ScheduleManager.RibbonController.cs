namespace SKSB_App.ScheduleManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;
    public static DateOnly CurrentScheduleDate { get; private set; } = DateOnly.FromDateTime(DateTime.Today);

    public static void SetScheduleDate(DateOnly newDate)
    {
        CurrentScheduleDate = newDate;
        _ribbonUi?.InvalidateControl("ebScheduleDate");
    }

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabScheduleManager' label='Schedule Manager'>
                <group id='grpScheduleDate' label='Schedule Date'>
                  <editBox id='ebScheduleDate' 
                           label='Date (YYYY-MM-DD):' 
                           sizeString='2026-10-14'
                           getText='GetScheduleDateText' 
                           onChange='OnScheduleDateChanged' />
                  <button id='btnLoadDay' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          onAction='OnLoadClicked' />
                  <button id='btnSaveDay' 
                          label='Save to System' 
                          size='large' 
                          imageMso='SaveObject' 
                          onAction='OnSaveClicked' />
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

    public string GetScheduleDateText(IRibbonControl control) => CurrentScheduleDate.ToString("yyyy-MM-dd");

    public void OnScheduleDateChanged(IRibbonControl control, string text)
    {
        if (DateOnly.TryParse(text, out DateOnly newDate))
        {
            CurrentScheduleDate = newDate;
            WorkbookSession.ShowStatusConfirmation($"Switched to Date: {newDate:yyyy-MM-dd}");
            ExcelAsyncUtil.QueueAsMacro(() => OnLoadClicked(control));
        }
        else
        {
            MessageBox.Show("Please enter a valid date formatted as YYYY-MM-DD.", "Invalid Date", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            _ribbonUi?.InvalidateControl("ebScheduleDate");
        }
    }

    public void OnSyncLookupsClicked(IRibbonControl control)
    {
        try
        {
            Events.SyncScheduleLookupsData();
            MessageBox.Show(
                "Schedule lookup data refreshed successfully.\n\n• Lookups: Synced core enums\n• Data: Live Worker, Vehicle, and Project identifiers refreshed",
                "Lookups Synced",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Lookup sync failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        // 1. MAIN UI THREAD: Synchronous read of Excel cells via C-API with validation gate
        var schedule = SheetMapper.ExtractScheduleRecord(CurrentScheduleDate);
        if (schedule == null)
        {
            return; // Pre-save validation failed or no data
        }

        // 2. BACKGROUND WORKER: Offload serialization and PostgreSQL persistence
        Task.Run(async () =>
        {
            try
            {
                var (success, path) = await ServerSyncCommands.SaveScheduleAsync(schedule);
                if (success)
                {
                    Services.ClearPendingOfflineSnapshot<ScheduleRecord>();
                    MessageBox.Show(
                        $"Successfully saved schedule for {schedule.DailyScheduleDate:yyyy-MM-dd}.\n\n• Storage: {path}",
                        "Save Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Failed to persist schedule to database or local storage.",
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
        var record = ServerSyncCommands.FetchSchedule(CurrentScheduleDate);
        if (record == null)
        {
            MessageBox.Show($"No saved schedule found for {CurrentScheduleDate:yyyy-MM-dd}.", "Load Schedule", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SheetMapper.PopulateScheduleSheets(record);
            WorkbookSession.ShowStatusConfirmation($"Loaded schedule for {CurrentScheduleDate:yyyy-MM-dd}");
        });
    }
}