namespace SKSB_App.StaffManager;

using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;

public class Events : ManagerEventsBase<List<StaffRecord>>
{
    // =========================================================================
    // 1. MODULE SAVE CONTRACT
    // =========================================================================
    // Static so it is fully initialised before ManagerEventsBase's constructor reads Contract.
    public static readonly SaveContract<List<StaffRecord>> StaffSyncDescriptor = new()
    {
        ModuleName = "Staff Manager",
        ServerTableName = "staff_master",
        WindowCaption = "Staff_Manager",
        MasterSheetAliases = SheetMapper.MasterSheetAliases,
        MasterTableName = SheetMapper.MasterTableName,
        EnumsToSync = new[]
        {
            typeof(Department),
            typeof(ContractStatus),
            typeof(VehicleLicenseClass),
            typeof(VehicleLicenseStatus),
            typeof(SkillLicenseStatus),
            typeof(DateFilterScope)
        },
        ExtractAndValidate = () => SheetMapper.ExtractStaffRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        SaveToServer = list => (ServerSyncCommands.SaveStaffSilent(list, out string err), err),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        FetchFromServer = () => ServerSyncCommands.FetchStaffRecords(),
        ValidateOfflinePayload = list => list != null && list.Count > 0 && list.TrueForAll(s => s.StaffID > 0),
        FormatConflictSummary = (local, server) =>
        {
            int localVLic = local.Sum(s => s.VehicleLicenses?.Count ?? 0);
            int localSLic = local.Sum(s => s.SkillLicenses?.Count ?? 0);
            int serverVLic = server.Sum(s => s.VehicleLicenses?.Count ?? 0);
            int serverSLic = server.Sum(s => s.SkillLicenses?.Count ?? 0);

            return $"  Local Offline : {local.Count} staff(s), {localVLic} vehicle license(s), {localSLic} skill license(s)\n" +
                   $"  Server Copy   : {server.Count} staff(s), {serverVLic} vehicle license(s), {serverSLic} skill license(s)";
        }
    };

    protected override SaveContract<List<StaffRecord>> Contract => StaffSyncDescriptor;

    // =========================================================================
    // 2. LOOKUP SYNCHRONISATION (runs on the Excel macro thread via the base class)
    // =========================================================================
    protected override void SyncCustomLookups(bool isInteractive)
    {
        // External Vehicle Plates -> 'Data' sheet
        DataList.SyncProperty<VehicleRecord, string>(
            v => v.VehiclePlate,
            customListName: "VehiclePlateList"
        );

        // SkillLicenseTypesTable & SkillLicenseTypeList
        SheetMapper.PopulateSkillLicenseTypesTable(showDialogIfFailed: isInteractive);
    }

    // =========================================================================
    // 3. COMMAND ENTRY POINTS (Excel-DNA only registers commands on non-generic types)
    // =========================================================================
    public static void RunHotKeySave() => ExecuteHotKeySave();

    public static void RunDelayedAutoLoad() => ExecuteDelayedAutoLoad();
}
