namespace SKSB_App.ProjectManager;

using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookSave.Entry;
using System.Collections.Generic;

public class Events : ManagerEventsBase<List<ProjectRecord>>
{
    // =========================================================================
    // 1. MODULE SAVE CONTRACT
    // =========================================================================
    // Static so it is fully initialised before ManagerEventsBase's constructor reads Contract.
    public static readonly SaveContract<List<ProjectRecord>> ProjectSyncDescriptor = new()
    {
        ModuleName = "Project Manager",
        ServerTableName = "project_master",
        WindowCaption = "Project_Manager",
        MasterSheetAliases = SheetMapper.MasterSheetAliases,
        MasterTableName = SheetMapper.MasterTableName,
        EnumsToSync = new[]
        {
            typeof(ProjectType),
            typeof(ProjectDistrict),
            typeof(ProjectPriority),
            typeof(ProjectStatus),
            typeof(ProjectCategory),
            typeof(Shifts)
        },
        ExtractAndValidate = () => SheetMapper.ExtractProjectRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        // Companies and the summary selector must populate even before the first project exists.
        PopulateStartupWhenEmpty = true,
        SaveToServer = list => (ServerSyncCommands.SaveProjectsSilent(list, out string err), err),
        SaveToOffline = list => ServerSyncCommands.SaveProjectsOffline(list),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        FetchFromServer = () => ServerSyncCommands.FetchProjectRecords(),
        ValidateOfflinePayload = list => list.Count > 0 && list.TrueForAll(p => !string.IsNullOrWhiteSpace(p.ProjectCode)),
        FormatConflictSummary = (local, server) =>
            $"  Local Offline : {local.Count} project(s)\n" +
            $"  Server Copy   : {server.Count} project(s)"
    };

    protected override SaveContract<List<ProjectRecord>> Contract => ProjectSyncDescriptor;

    // =========================================================================
    // 3. LOOKUP SYNCHRONISATION (runs on the Excel macro thread via the base class)
    // =========================================================================
    protected override void SyncCustomLookups(bool isInteractive)
    {
        // Staff / vehicle / company / project lists come straight from PostgreSQL (offline snapshot as fallback),
        // so the Data sheet never shows placeholders just because other managers were not opened on this machine.
        ServerSyncCommands.SyncLookupLists();
    }

    // =========================================================================
    // 4. COMMAND ENTRY POINTS (Excel-DNA only registers commands on non-generic types)
    // =========================================================================
    public static void RunHotKeySave() => ExecuteHotKeySave();

    public static void RunDelayedAutoLoad() => ExecuteDelayedAutoLoad();
}
