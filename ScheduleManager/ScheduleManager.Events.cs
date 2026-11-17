namespace SKSB_App.ScheduleManager;

using System;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.SheetMapper;

public class Events : ManagerEventsBase<ScheduleRecord>
{
    // =========================================================================
    // 1. MODULE SAVE CONTRACT
    // =========================================================================
    // Static so it is fully initialised before ManagerEventsBase's constructor reads Contract.
    public static readonly SaveContract<ScheduleRecord> ScheduleSyncDescriptor = new()
    {
        ModuleName = "Schedule Manager",
        ServerTableName = "daily_project_schedules",
        WindowCaption = "Schedule_Manager",
        MasterSheetAliases = SheetMapper.MasterSheetAliases,
        MasterTableName = SheetMapper.MasterTableName,
        EnumsToSync = new[]
        {
            typeof(Shifts),
            typeof(AttendanceStatus),
            typeof(AttendanceEntryType),
            typeof(AttendanceSource),
            typeof(Days),
            typeof(VehicleType),
            typeof(ProjectType)
        },
        ExtractAndValidate = () => SheetMapper.ExtractScheduleRecordWithDiagnostics(),
        IsDatasetEmpty = IsEmpty,
        // Reference tables (teams, preferences) and the board must load even before any day has been scheduled.
        PopulateStartupWhenEmpty = true,
        SaveToServer = schedule => (ServerSyncCommands.SaveScheduleSilent(schedule, out string err), err),
        PopulateWorksheet = schedule => SheetMapper.PopulateWorksheet(schedule),
        FetchFromServer = () => ServerSyncCommands.FetchSchedule(CurrentDate()),
        ValidateOfflinePayload = schedule => schedule != null && schedule.DailyScheduleDate.Year > 2000,
        FormatConflictSummary = (local, server) =>
            $"  Local Offline : {local.DailyScheduleDate:yyyy-MM-dd} - {local.Projects.Count} slot(s), {local.WorkerAttendance.Count} attendance row(s), {local.AttendanceEntries.Count} punch(es)\n" +
            $"  Server Copy   : {server.DailyScheduleDate:yyyy-MM-dd} - {server.Projects.Count} slot(s), {server.WorkerAttendance.Count} attendance row(s), {server.AttendanceEntries.Count} punch(es)"
    };

    protected override SaveContract<ScheduleRecord> Contract => ScheduleSyncDescriptor;

    private static bool IsEmpty(ScheduleRecord? s) =>
        s == null ||
        (s.Projects.Count == 0 && s.WorkerAttendance.Count == 0 && s.AttendanceEntries.Count == 0 &&
         s.Preferences.Count == 0 && s.Teams.Count == 0);

    /// <summary>The date typed on the board (today when the cell is empty or no workbook is open).</summary>
    internal static DateOnly CurrentDate()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb != null)
            {
                dynamic? board = SheetMapperHelper.ResolveSheetSafely((object)wb, ScheduleBoardMapper.SheetAliases, ScheduleBoardMapper.DefaultSheetName);
                if (board != null)
                {
                    DateOnly? d = ScheduleBoardMapper.ReadDate((object)board);
                    if (d.HasValue) return d.Value;
                }
            }
        }
        catch { }
        return DateOnly.FromDateTime(DateTime.Today);
    }

    // =========================================================================
    // 2. ADD-IN LIFECYCLE (board dropdowns + live observer)
    // =========================================================================
    // Pinned in a static field: ParameterCells keeps a reference but the delegate identity is needed to unregister.
    private static readonly Action<dynamic, dynamic> SheetObserver = SheetMapper.OnSheetChanged;
    private static bool _boardHooked;

    public override void AutoOpen()
    {
        base.AutoOpen();
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            if (_boardHooked) return;
            ScheduleBoardMapper.RegisterDropdowns();
            ParameterCells.RegisterChangeObserver(SheetObserver);
            _boardHooked = true;
        });
    }

    public override void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            ParameterCells.UnregisterChangeObserver(SheetObserver);
            CustomDropdowns.Clear();
            _boardHooked = false;
        });
        base.AutoClose();
    }

    // =========================================================================
    // 3. LOOKUP SYNCHRONISATION (runs on the Excel macro thread via the base class)
    // =========================================================================
    protected override void SyncCustomLookups(bool isInteractive)
    {
        try { ServerSyncCommands.SyncLookupLists(); }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager] Lookup sync warning: {ex.Message}"); }
    }

    // =========================================================================
    // 4. COMMAND ENTRY POINTS (Excel-DNA only registers commands on non-generic types)
    // =========================================================================
    public static void RunHotKeySave() => ExecuteHotKeySave();

    public static void RunDelayedAutoLoad() => ExecuteDelayedAutoLoad();
}
