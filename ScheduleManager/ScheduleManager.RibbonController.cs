namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Validation;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;
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
                  <box id='boxWorkingDay' boxStyle='vertical'>
                    <labelControl id='lblWorkingDate' getLabel='GetWorkingDateLabel' />
                    <labelControl id='lblWorkingDayOfWeek' getLabel='GetWorkingDayOfWeekLabel' />
                    <labelControl id='lblWorkingDayRelation' getLabel='GetWorkingDayRelationLabel' />
                  </box>
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
                          supertip='Reads all punches from the attendance terminal, saves new entries across all dates to the database, and refreshes the current day.' />
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

    private static DateOnly DisplayedDate => SheetMapper.LoadedDate ?? DateOnly.FromDateTime(DateTime.Today);

    public string GetWorkingDateLabel(IRibbonControl control) => $"Date: {DisplayedDate:dd MMM yyyy}";

    public string GetWorkingDayOfWeekLabel(IRibbonControl control) => $"Day: {DisplayedDate:dddd}";

    public string GetWorkingDayRelationLabel(IRibbonControl control)
    {
        int diff = DisplayedDate.DayNumber - DateOnly.FromDateTime(DateTime.Today).DayNumber;
        return diff switch
        {
            0 => "Status: Today",
            -1 => "Status: Yesterday (Past)",
            < -1 => $"Status: Past ({-diff}d ago)",
            1 => "Status: Tomorrow",
            _ => $"Status: Upcoming (+{diff}d)"
        };
    }

    internal static void RefreshWorkingDayDisplay()
    {
        try { RibbonUi?.InvalidateControl("lblWorkingDate"); } catch { }
        try { RibbonUi?.InvalidateControl("lblWorkingDayOfWeek"); } catch { }
        try { RibbonUi?.InvalidateControl("lblWorkingDayRelation"); } catch { }
    }

    // =========================================================================
    // 5. ATTENDANCE
    // =========================================================================
    protected override void OnAfterPurgeSheetData() => ScheduleBoardMapper.ResetSlotCount();

    protected override void OnAfterDuplicateRowsCleaned(int removedCount) => SheetMapper.OnDuplicateRowsCleaned();

    public void OnRollupClicked(IRibbonControl control) =>
        ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.RefreshDerived(reloadContext: true, rollup: true, ensureAdminEntries: true));

    public void OnSyncTerminalClicked(IRibbonControl control)
    {
        // Read state on the macro thread, talk to the terminal/database on a worker thread, write back on the macro thread.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            DateOnly date = Events.CurrentDate();
            bool isOnline = ServerEnvironment.IsOnline;
            string? addinDir = null;
            try { addinDir = Path.GetDirectoryName(ExcelDnaUtil.XllPath); } catch { }

            if (!isOnline) { SyncTerminalOffline(date, addinDir); return; }

            var knownStaffIds = Roster.AttendanceStaff.Select(s => s.Id).ToHashSet();
            var sheetThumbprintEntries = SheetMapper.ReadCurrentSheetEntries().Where(e => e.Source == AttendanceSource.Thumbprint).ToList();
            SetStatus($"Syncing all punches from thumbprint terminal ({DeviceSyncManager.DefaultIp})...");
            AppLogger.Info($"[ScheduleManager] Syncing all punches from thumbprint terminal ({DeviceSyncManager.DefaultIp})...");

            Task.Run(async () =>
            {
                try
                {
                    var punches = await DeviceSyncManager.FetchPunchesAsync(since: null, addinDirectory: addinDir);
                    if (punches.Count == 0)
                    {
                        ExcelAsyncUtil.QueueAsMacro(() =>
                        {
                            SetStatus(false);
                            MessageBox.Show($"Terminal {DeviceSyncManager.DefaultIp}: 0 punches found on the terminal.",
                                "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        });
                        return;
                    }

                    DateOnly firstDate = DateOnly.FromDateTime(punches.Min(p => p.Timestamp));
                    DateOnly lastDate = DateOnly.FromDateTime(punches.Max(p => p.Timestamp));
                    // Per-date Day Start Times: each working day runs from its own start to the next date's start.
                    var dayStarts = ServerSyncCommands.FetchDayStartTimesSince(firstDate.AddDays(-1));
                    TimeOnly DayStartFor(DateOnly d) => dayStarts.TryGetValue(d, out var t) ? t : DeviceSyncManager.DefaultDayStartTime;
                    var dbEntries = ServerSyncCommands.FetchThumbprintEntriesSince(firstDate.AddDays(-1));
                    var existing = dbEntries.Concat(sheetThumbprintEntries).ToList();

                    var (newEntries, duplicates, unknown) = DeviceSyncManager.BuildEntries(punches, existing, id => knownStaffIds.Contains(id), DayStartFor);
                    int savedToDb = newEntries.Count > 0 ? ServerDispatcher.InsertAttendanceEntries(newEntries) : 0;

                    int totalDays = punches.Select(p => DeviceSyncManager.WorkingDate(DateOnly.FromDateTime(p.Timestamp), TimeOnly.FromDateTime(p.Timestamp), DayStartFor)).Distinct().Count();
                    int savedDays = newEntries.Select(e => DeviceSyncManager.WorkingDate(e.Date, e.Time, DayStartFor)).Distinct().Count();

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        try
                        {
                            var allPersisted = dbEntries.Concat(newEntries).ToList();
                            var (totalAddedToSheet, addedForDate) = SheetMapper.MergePersistedEntries(allPersisted, date);
                            string msg =
                                $"Terminal {DeviceSyncManager.DefaultIp}: {punches.Count} punch(es) read, {firstDate:dd MMM yyyy} - {lastDate:dd MMM yyyy} ({totalDays} working day(s)).\n\n" +
                                $"  Saved to database (all dates) : {savedToDb}" + (savedToDb > 0 ? $" (across {savedDays} day(s))" : "") + "\n" +
                                $"  Added to Attendance Entries : {totalAddedToSheet} ({addedForDate} on {date:dd MMM yyyy})\n" +
                                $"  Already present / double scans : {duplicates}\n" +
                                $"  Unknown staff : {unknown}\n\n" +
                                $"Attendance Entries holds all dates and is filtered to {date:dd MMM yyyy} via the Date header (clear the Date filter to view all dates)." +
                                (addedForDate > 0 ? " The current day's board and Daily Attendances were refreshed; use 'Save to System' to persist the rolled-up Daily Attendances." : "");
                            SetStatus(false);
                            MessageBox.Show(msg, "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex2)
                        {
                            AppLogger.Error("[ScheduleManager] Merging terminal punches failed", ex2);
                            MessageBox.Show($"Entries were saved to the database, but refreshing the sheet failed:\n\n{ex2.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                        finally { SetStatus(false); }
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Error("[ScheduleManager] Terminal sync failed", ex);
                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        SetStatus(false);
                        MessageBox.Show($"Terminal sync failed:\n\n{ex.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    });
                }
            });
        });
    }

    private static void SetStatus(object value)
    {
        try { ((dynamic)ExcelDnaUtil.Application).StatusBar = value; } catch { }
    }

    /// <summary>Offline fallback: all punches are read from the terminal and appended to the sheet, filtered to the displayed date. Macro thread.</summary>
    private static void SyncTerminalOffline(DateOnly date, string? addinDir)
    {
        if (date < DateOnly.FromDateTime(DateTime.Today) &&
            !PastRecordGuard.ConfirmPastRecordEdit(SheetMapper.GuardSession, date, "Schedule Manager"))
        {
            SheetOperations.ShowStatusConfirmation("Terminal sync cancelled.");
            return;
        }
        AppLogger.Info($"[ScheduleManager] Offline: syncing all punches from thumbprint terminal ({DeviceSyncManager.DefaultIp})...");
        SetStatus($"Syncing all punches from thumbprint terminal ({DeviceSyncManager.DefaultIp})...");
        Task.Run(async () =>
        {
            try
            {
                var punches = await DeviceSyncManager.FetchPunchesAsync(since: null, addinDirectory: addinDir);
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        var (added, duplicates, unknown) = SheetMapper.AppendTerminalPunches(punches, date);
                        SetStatus(false);
                        MessageBox.Show(
                            $"Terminal {DeviceSyncManager.DefaultIp}: {punches.Count} punch(es) read (all dates).\n\n" +
                            $"  Added to Attendance Entries : {added}\n" +
                            $"  Already present / double scans : {duplicates}\n" +
                            $"  Unknown staff / other dates :{unknown}\n\n" +
                            $"Attendance Entries holds all dates and is filtered to {date:dd MMM yyyy} via the Date header (clear the Date filter to view all dates).\n" +
                            "Use 'Save to System' to persist the new entries.",
                            "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex2)
                    {
                        AppLogger.Error("[ScheduleManager] Appending terminal punches failed", ex2);
                        MessageBox.Show($"Appending terminal punches failed:\n\n{ex2.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally { SetStatus(false); }
                });
            }
            catch (Exception ex)
            {
                AppLogger.Error("[ScheduleManager] Terminal sync failed", ex);
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    SetStatus(false);
                    MessageBox.Show($"Terminal sync failed:\n\n{ex.Message}", "Thumbprint Sync", MessageBoxButtons.OK, MessageBoxIcon.Error);
                });
            }
        });
    }
}
