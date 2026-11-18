namespace SKSB_App.ScheduleManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ManagerRibbonBase
{
    // =========================================================================
    // 1. MODULE DESCRIPTOR
    // =========================================================================
    protected override ISaveContract Contract => Events.ScheduleSyncDescriptor;
    protected override string TabId => "tabScheduleManager";
    protected override string TabLabel => "Schedule Manager";

    protected override string PersistenceGroupXml => BuildPersistenceGroupXml(
        groupId: "grpSchedulePersistence",
        groupLabel: "Schedule Database Sync",
        loadId: "btnLoadSchedule",
        loadScreentip: "Load Schedule",
        loadSupertip: "Fetches the dispatch board, attendance entries, rollup, preferences and teams for the selected date from PostgreSQL or offline storage.",
        saveId: "btnSaveSchedule",
        saveScreentip: "Save Schedule",
        saveSupertip: "Validates the board (duplicate workers block the save), rolls punches up into daily attendance and commits everything.");

    // An empty day still populates the board and the reference tables.
    protected override bool PopulateWhenEmpty => true;
    protected override bool ShowEmptyLoadDialog => false;

    protected override bool ShouldSumColumnInTotalRow(string header) =>
        header.Contains("Shift", StringComparison.OrdinalIgnoreCase) ||
        header.Contains("Hours", StringComparison.OrdinalIgnoreCase);

    // =========================================================================
    // 2. SCHEDULE GROUPS
    // =========================================================================
    protected override string GetCustomGroupsXml() => @"
                <group id='grpScheduleDay' label='Working Day'>
                  <button id='btnPrevDay' label='Previous Day' size='large' imageMso='LeftArrow2' onAction='OnPreviousDayClicked'
                          screentip='Previous Day' supertip='Loads the schedule for the day before the one on the board.' />
                  <button id='btnToday' label='Today' size='large' imageMso='DateAndTimeInsert' onAction='OnTodayClicked'
                          screentip='Today' supertip='Loads the schedule for today.' />
                  <button id='btnNextDay' label='Next Day' size='large' imageMso='RightArrow2' onAction='OnNextDayClicked'
                          screentip='Next Day' supertip='Loads the schedule for the day after the one on the board.' />
                </group>
                <group id='grpScheduleAttendance' label='Attendance'>
                  <button id='btnRollup' label='Roll Up Attendance' size='large' imageMso='Refresh' onAction='OnRollupClicked'
                          screentip='Roll Up Attendance'
                          supertip='Rebuilds Daily Attendances and the board check-in times from the approved Attendance Entries.' />
                  <button id='btnSyncTerminal' label='Sync Thumbprint Terminal' size='large' imageMso='AccessRefreshAllLists' onAction='OnSyncTerminalClicked'
                          screentip='Sync Thumbprint Terminal'
                          supertip='Reads punches from the attendance terminal and appends the new ones to Attendance Entries (Source = Thumbprint).' />
                </group>";

    // =========================================================================
    // 3. EXCEL COMMANDS (must live on a non-generic type)
    // =========================================================================
    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro() => Events.RunDelayedAutoLoad();

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave() => Events.RunHotKeySave();

    // =========================================================================
    // 4. DAY NAVIGATION
    // =========================================================================
    public void OnPreviousDayClicked(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.GoToDate(Events.CurrentDate().AddDays(-1)));

    public void OnNextDayClicked(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.GoToDate(Events.CurrentDate().AddDays(1)));

    public void OnTodayClicked(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.GoToDate(DateOnly.FromDateTime(DateTime.Today)));

    // =========================================================================
    // 5. ATTENDANCE
    // =========================================================================
    public void OnRollupClicked(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.RefreshDerived(reloadContext: true, rollup: true));

    public void OnSyncTerminalClicked(IRibbonControl control)
    {
        // Read the date on the macro thread, talk to the terminal on a worker thread, write back on the macro thread.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            DateOnly date = Events.CurrentDate();
            AppLogger.Info($"[ScheduleManager] Syncing thumbprint terminal ({DeviceSyncManager.DefaultIp})...");
            string? addinDir = null;
            try { addinDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath); } catch { }
            try { ((dynamic)ExcelDnaUtil.Application).StatusBar = $"Syncing thumbprint terminal ({DeviceSyncManager.DefaultIp})..."; } catch { }
            Task.Run(async () =>
            {
                try
                {
                    var punches = await DeviceSyncManager.FetchPunchesAsync(since: date, addinDirectory: addinDir);
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                        var (added, duplicates, unknown) = SheetMapper.AppendTerminalPunches(punches);
                        MessageBox.Show(
                            $"Terminal {DeviceSyncManager.DefaultIp}: {punches.Count} punch(es) read.\n\n" +
                            $"  Added to Attendance Entries : {added}\n" +
                            $"  Already present / double scans : {duplicates}\n" +
                            $"  Unknown staff / other dates : {unknown}\n\n" +
                            "Use 'Save to System' to persist the new entries.",
                            "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex2)
                        {
                            AppLogger.Error("[ScheduleManager] Appending terminal punches failed", ex2);
                            MessageBox.Show($"Appending terminal punches failed:\n\n{ex2.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally
                        {
                            try { ((dynamic)ExcelDnaUtil.Application).StatusBar = false; } catch { }
                        }
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Error("[ScheduleManager] Terminal sync failed", ex);
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try { ((dynamic)ExcelDnaUtil.Application).StatusBar = false; } catch { }
                        MessageBox.Show(
                            $"Terminal sync failed:\n\n{ex.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            });
        });
    }
}
