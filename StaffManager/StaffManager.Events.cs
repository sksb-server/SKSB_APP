namespace SKSB_App.StaffManager;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
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
        SaveToOffline = list => ServerSyncCommands.SaveStaffOffline(list),
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
        // External Vehicle Plates -> 'Data' sheet. Online Mode: PostgreSQL only; Offline Mode: local snapshot only.
        DataList.SyncProperty<VehicleRecord, string>(
            v => v.VehiclePlate,
            LoadVehiclePlates(),
            customListName: "VehiclePlateList"
        );

        // SkillLicenseTypesTable & SkillLicenseTypeList
        SheetMapper.PopulateSkillLicenseTypesTable(showDialogIfFailed: isInteractive);
    }

    private static List<VehicleRecord> LoadVehiclePlates()
    {
        if (ServerEnvironment.IsOnline)
        {
            try
            {
                return new PostgreSqlService()
                    .Query<string>("SELECT vehicle_plate FROM vehicle_master ORDER BY vehicle_plate;")
                    .Select(p => new VehicleRecord { VehiclePlate = p })
                    .ToList();
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[StaffManager] vehicle_master query failed: {ex.Message}");
                if (!ServerEnvironment.IsOffline) return new List<VehicleRecord>();
            }
        }
        return OfflineSnapshot.Load<List<VehicleRecord>>() ?? new List<VehicleRecord>();
    }

    // =========================================================================
    // 3. COMMAND ENTRY POINTS (Excel-DNA only registers commands on non-generic types)
    // =========================================================================
    public static void RunHotKeySave() => ExecuteHotKeySave();

    public static void RunDelayedAutoLoad() => ExecuteDelayedAutoLoad();
}
