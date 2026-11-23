namespace SKSB_App.StaffManager;

using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
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
        SaveToServerWithConflictCheck = (list, force) =>
        {
            bool ok = ServerSyncCommands.SaveStaffWithDelta(list, force, out string err, out var conflicts);
            return (ok, err, conflicts);
        },
        // "Keep Server Values" in the pane's Conflict Mode: re-baseline only the conflicting rows.
        OnAcceptServerValues = ServerSyncCommands.AcceptServerConflictBaselines,
        SaveToOffline = list => ServerSyncCommands.SaveStaffOffline(list),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        // Authoritative re-fetch after a delta save: brings back server-assigned license IDs and rows added by
        // other users, and refreshes the delta baselines.
        RefreshAfterSave = savedList =>
        {
            var fresh = ServerSyncCommands.FetchStaffRecords();
            SheetMapper.PopulateWorksheet(fresh ?? savedList);
        },
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

    public override void AutoOpen()
    {
        base.AutoOpen();
        SheetMapper.RegisterSummaryHooks();
        SheetMapper.RegisterSkillTypeCascadeHooks();
    }

    public override void AutoClose()
    {
        SheetMapper.UnregisterSummaryHooks();
        SheetMapper.UnregisterSkillTypeCascadeHooks();
        CustomDropdowns.Clear();
        base.AutoClose();
    }

    // =========================================================================
    // 2. LOOKUP SYNCHRONISATION (runs on the Excel macro thread via the base class)
    // =========================================================================
    protected override void SyncCustomLookups(bool isInteractive)
    {
        // External Vehicle Plates -> 'Data' sheet. Online Mode: PostgreSQL only; Offline Mode: local snapshot only.
        // VehiclePlateList and VehiclePlateTypeList stay index-aligned ("PLATE — Type" dropdown labels).
        var vehicles = LoadVehiclePlates()
            .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
            .OrderBy(v => v.VehiclePlate.Trim(), StringComparer.OrdinalIgnoreCase)
            .ToList();
        if (vehicles.Count > 0)
        {
            DataList.SyncAlignedLists(new Dictionary<string, IEnumerable<string>>
            {
                ["VehiclePlateList"] = vehicles.Select(v => v.VehiclePlate.Trim()).ToList(),
                ["VehiclePlateTypeList"] = vehicles.Select(v => v.VehicleDetail.Type == VehicleType.None ? string.Empty : v.VehicleDetail.Type.ToString()).ToList(),
            });
        }

        // Emoji palette behind SkillLicenseTypesTable[Display Icon]
        DataList.SyncValues("DisplayIconList", SkillCatalog.DisplayIcons);

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
                    .Query<dynamic>("SELECT vehicle_plate, COALESCE(vehicle_type, '') AS vehicle_type FROM vehicle_master ORDER BY vehicle_plate;")
                    .Select(r =>
                    {
                        var rec = new VehicleRecord { VehiclePlate = (string?)r.vehicle_plate ?? string.Empty };
                        if (Enum.TryParse<VehicleType>((string?)r.vehicle_type, true, out var t)) rec.VehicleDetail.Type = t;
                        else rec.VehicleDetail.Type = VehicleType.None;
                        return rec;
                    })
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
