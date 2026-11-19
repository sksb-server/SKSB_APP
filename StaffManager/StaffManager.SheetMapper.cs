namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;

public static class SheetMapper
{
    public const string MasterSheetName = "All Staffs";
    public const string MasterTableName = "AllStaffsTable";

    internal static readonly string[] MasterSheetAliases = { "All Staffs", "AllStaffs", "Staff_Master", "Staffs", "All Staff", "All Workers" };
    private static readonly string[] VehicleLicenseAliases = { "Vehicle Licenses", "VehicleLicenses", "Vehicle License", "Staff_Licenses" };
    private static readonly string[] SkillLicenseAliases = { "Skill Licenses", "SkillLicenses", "Skill License", "Staff_Skills" };

    // =========================================================================
    // 1. DECLARATIVE WORKBOOK MAPPING (MASTER + CHILD TABLES)
    // =========================================================================
    private static readonly WorkbookMapping<StaffRecord, long> Mapping = new()
    {
        MasterSheetAliases = MasterSheetAliases,
        MasterSheetName = MasterSheetName,
        MasterTableName = MasterTableName,
        EntityLabel = "staff",
        MasterKeyLabel = "Staff ID",
        MasterKeySelector = s => s.StaffID,
        IsValidKey = id => id > 0,

        Children = new List<IChildTableBinding<StaffRecord, long>>
        {
            new ChildTableBinding<StaffRecord, VehicleLicense, long>
            {
                SheetAliases = VehicleLicenseAliases,
                DefaultSheetName = "Vehicle Licenses",
                TableName = "VehicleLicensesTable",
                ForeignKeySelector = l => l.StaffID,
                ParentCollectionSelector = s => s.VehicleLicenses,
                BackfillFromParent = (s, l) =>
                {
                    if (l.StaffID <= 0) l.StaffID = s.StaffID;
                    if (string.IsNullOrWhiteSpace(l.StaffName)) l.StaffName = s.StaffDetails.StaffName;
                },
                ChildPrimaryKeySelector = l => l.VehicleLicenseID,
                ChildPrimaryKeyLabel = "Vehicle License ID",
                ValidateForeignKeyExists = true
            },
            new ChildTableBinding<StaffRecord, SkillLicense, long>
            {
                SheetAliases = SkillLicenseAliases,
                DefaultSheetName = "Skill Licenses",
                TableName = "SkillLicensesTable",
                ForeignKeySelector = l => l.StaffID,
                ParentCollectionSelector = s => s.SkillLicenses,
                BackfillFromParent = (s, l) =>
                {
                    if (l.StaffID <= 0) l.StaffID = s.StaffID;
                    if (string.IsNullOrWhiteSpace(l.StaffName)) l.StaffName = s.StaffDetails.StaffName;
                },
                ChildPrimaryKeySelector = l => l.SkillLicenseID,
                ChildPrimaryKeyLabel = "Skill License ID",
                ValidateForeignKeyExists = true
            }
        },

        CustomValidator = (wb, masters, errors, warnings) =>
        {
            foreach (var s in masters)
            {
                var d = s.StaffDetails;
                if (d.ResignDate.HasValue && d.ResignDate.Value < d.JoinDate)
                {
                    errors.Add($"  [{MasterSheetName}]: Staff ID '{s.StaffID}' ({d.StaffName}) Resign Date ({d.ResignDate.Value:dd/MM/yyyy}) cannot be earlier than Join Date ({d.JoinDate:dd/MM/yyyy}).");
                }
            }
        },

        OrderMasterRecords = records => records
            .OrderBy(s => s.StaffDetails.ContractStatus == ContractStatus.Resigned ? 1 : 0)
            .ThenBy(s => s.StaffID)
            .ToList(),

        OnPopulateAuxiliaryTables = (wb, masters) => PopulateSkillLicenseTypesFromCache((object)wb),

        AuditTitle = "Staff Manager Populate Audit",
        MasterLabel = "Staff Profiles",
        InspectTables = new[] { "AllStaffsTable", "VehicleLicensesTable", "SkillLicensesTable", "SkillLicenseTypesTable" }
    };

    // =========================================================================
    // 2. MASTER EXTRACTION & POPULATE PIPELINES
    // =========================================================================
    public static (List<StaffRecord>? Data, List<string> Errors, List<string> Warnings) ExtractStaffRecordsWithDiagnostics()
    {
        var result = WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);
        if (result.Data != null) ServerSyncCommands.DeriveShortNames(result.Data); // blank Short Name -> first + last word of Staff Name
        return result;
    }

    public static void PopulateWorksheet(List<StaffRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 3. SKILL LICENSE TYPES (AUXILIARY REFERENCE TABLE)
    // =========================================================================
    /// <summary>Populate-time path: server in Online Mode, local snapshot in Offline Mode (decided inside the fetch).</summary>
    private static void PopulateSkillLicenseTypesFromCache(object wbObj)
    {
        try
        {
            // Online: PostgreSQL only. Offline: FetchSkillLicenseTypesFromDb reads the local snapshot.
            WriteSkillLicenseTypes(wbObj, ServerSyncCommands.FetchSkillLicenseTypesFromDb());
        }
        catch { }
    }

    public static void PopulateSkillLicenseTypesTable(bool showDialogIfFailed = false)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            // FetchSkillLicenseTypesFromDb already falls back to the offline cache only when the server is unreachable.
            WriteSkillLicenseTypes((object)wb, ServerSyncCommands.FetchSkillLicenseTypesFromDb());
        }
        catch (Exception ex)
        {
            if (showDialogIfFailed)
                System.Windows.Forms.MessageBox.Show($"Failed populating skill license types: {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    private static void WriteSkillLicenseTypes(object wbObj, List<SkillLicenseType>? types)
    {
        // Always populate the emoji palette so the 'Display Icon' dropdown works even on an empty table.
        DataList.SyncValues("DisplayIconList", SkillCatalog.DisplayIcons);
        if (types == null || types.Count == 0) return;

        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SkillLicenseAliases, "Skill Licenses");
        if (ws == null) return;

        TableMapper.PopulateTableSafe<SkillLicenseType>((object)ws, "SkillLicenseTypesTable", types);
        var names = types.Select(t => t.SkillLicenseName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
        DataList.SyncValues("SkillLicenseTypeList", names);
    }

    public static List<SkillLicenseType> ExtractSkillLicenseTypes()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return new List<SkillLicenseType>();

            dynamic? sWs = SheetMapperHelper.ResolveSheetSafely((object)wb, SkillLicenseAliases, "Skill Licenses");
            if (sWs == null) return new List<SkillLicenseType>();

            var extracted = TableMapper.ExtractFromTable<SkillLicenseType>((object)sWs, "SkillLicenseTypesTable");
            var distinct = extracted
                .Where(t => !string.IsNullOrWhiteSpace(t.SkillLicenseName))
                .GroupBy(t => t.SkillLicenseName.Trim(), StringComparer.OrdinalIgnoreCase)
                .Select(g => new
                {
                    Name = g.Key,
                    Icon = g.Select(t => (t.DisplayIcon ?? string.Empty).Trim()).FirstOrDefault(i => i.Length > 0) ?? string.Empty
                })
                .OrderBy(x => x.Name)
                .Select((x, idx) => new SkillLicenseType
                {
                    SkillLicenseTypeID = idx + 1,
                    SkillLicenseName = x.Name,
                    DisplayIcon = x.Icon
                })
                .ToList();

            if (distinct.Count > 0)
            {
                DataList.SyncValues("SkillLicenseTypeList", distinct.Select(d => d.SkillLicenseName));
            }

            return distinct;
        }
        catch
        {
            return new List<SkillLicenseType>();
        }
    }
}
