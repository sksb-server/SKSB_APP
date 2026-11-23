namespace SKSB_App.VehicleManager;

using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System.Collections.Generic;
using System.Linq;

public class Events : ManagerEventsBase<List<VehicleRecord>>
{
    // =========================================================================
    // 1. MODULE SAVE CONTRACT
    // =========================================================================
    // Static so it is fully initialised before ManagerEventsBase's constructor reads Contract.
    public static readonly SaveContract<List<VehicleRecord>> VehicleSyncDescriptor = new()
    {
        ModuleName = "Vehicle Manager",
        ServerTableName = "vehicle_master",
        WindowCaption = "Vehicle_Manager",
        MasterSheetAliases = SheetMapper.AllVehicleAliases,
        MasterTableName = "AllVehiclesTable",
        EnumsToSync = new[]
        {
            typeof(DateFilterScope),
            typeof(VehicleType),
            typeof(ServiceType),
            typeof(MaintenanceStatus),
            typeof(VehicleStatus),
            typeof(VehiclePermitType),
            typeof(VehiclePermitStatus),
            typeof(RoadTaxStatus),
            typeof(InspectionStatus),
            typeof(InsuranceType),
            typeof(InsuranceStatus),
            typeof(InsuranceClaimStatus)
        },
        ExtractAndValidate = () => SheetMapper.ExtractVehicleRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        // Manual Load populates even when empty (RibbonController.PopulateWhenEmpty); startup must match,
        // otherwise fuel companies / workshops / cards never load until the user presses Load.
        PopulateStartupWhenEmpty = true,
        SaveToServer = list => (ServerSyncCommands.SaveVehiclesSilent(list, out string err), err),
        SaveToServerWithConflictCheck = (list, force) =>
        {
            bool ok = ServerSyncCommands.SaveVehiclesWithDelta(list, force, out string err, out var conflicts);
            return (ok, err, conflicts);
        },
        // "Keep Server Values" in the pane's Conflict Mode: re-baseline only the conflicting rows.
        OnAcceptServerValues = ServerSyncCommands.AcceptServerConflictBaselines,
        SaveToOffline = list => ServerSyncCommands.SaveVehiclesOffline(list),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        // Authoritative re-fetch after a delta save: brings back server-assigned child / workshop IDs and rows
        // added by other users, and refreshes every delta baseline (auxiliary tables are re-fetched by
        // PopulateWorksheet itself).
        RefreshAfterSave = savedList =>
        {
            var fresh = ServerSyncCommands.FetchVehicleRecords();
            SheetMapper.PopulateWorksheet(fresh ?? savedList);
        },
        FetchFromServer = () => ServerSyncCommands.FetchVehicleRecords(),
        ValidateOfflinePayload = list => list.Count > 0 && list.TrueForAll(v => !string.IsNullOrWhiteSpace(v.VehiclePlate)),
        FormatConflictSummary = (local, server) =>
        {
            int localClaims = local.Sum(v => v.InsuranceClaims?.Count ?? 0);
            int serverClaims = server.Sum(v => v.InsuranceClaims?.Count ?? 0);

            return $"  Local Offline : {local.Count} vehicle(s), {localClaims} insurance claim(s)\n" +
                   $"  Server Copy   : {server.Count} vehicle(s), {serverClaims} insurance claim(s)";
        }
    };

    protected override SaveContract<List<VehicleRecord>> Contract => VehicleSyncDescriptor;

    // =========================================================================
    // 1b. ADD-IN LIFECYCLE (date filter observers on top of the shared base)
    // =========================================================================
    public override void AutoOpen()
    {
        base.AutoOpen();
        SheetMapper.RegisterRollupHooks();
        SheetMapper.RegisterSummaryHooks();
    }

    public override void AutoClose()
    {
        SheetMapper.UnregisterRollupHooks();
        SheetMapper.UnregisterSummaryHooks();
        CustomDropdowns.Clear();
        base.AutoClose();
    }

    // =========================================================================
    // 2. LOOKUP SYNCHRONISATION (runs on the Excel macro thread via the base class)
    // =========================================================================
    protected override void SyncCustomLookups(bool isInteractive)
    {
        // Online Mode: staff come purely from PostgreSQL. Offline Mode: purely from the local snapshot.
        var staff = StaffLookupHelper.LoadCachedStaff();

        // 1. Staff IDs (ordered numerically)
        DataList.SyncProperty<StaffRecord, long>(s => s.StaffID, staff, "StaffIDList");

        // 2. StaffShortName into StaffNameList (fallback to StaffName if blank)
        DataList.SyncProperty<StaffRecord, string>(
            s => !string.IsNullOrWhiteSpace(s.StaffDetails.StaffShortName)
                ? s.StaffDetails.StaffShortName
                : s.StaffDetails.StaffName,
            staff, "StaffNameList");

        // 3. Dedicated StaffShortNameList (Staff ID / Staff Name ColumnSync on Odometers and Insurance Claims)
        DataList.SyncProperty<StaffRecord, string>(s => s.StaffDetails.StaffShortName, staff, "StaffShortNameList");
    }

    // =========================================================================
    // 3. COMMAND ENTRY POINTS (Excel-DNA only registers commands on non-generic types)
    // =========================================================================
    public static void RunHotKeySave() => ExecuteHotKeySave();

    public static void RunDelayedAutoLoad() => ExecuteDelayedAutoLoad();
}
