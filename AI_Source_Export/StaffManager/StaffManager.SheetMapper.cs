namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;

public static class SheetMapper
{
    public const string MasterSheetName = "All Staffs";
    public const string MasterTableName = "AllStaffsTable";

    private static readonly string[] MasterSheetAliases = { "All Staffs", "AllStaffs", "Staff_Master", "Staffs", "All Staff" };
    private static readonly string[] VehicleLicenseAliases = { "Vehicle Licenses", "VehicleLicenses", "Vehicle License" };
    private static readonly string[] SkillLicenseAliases = { "Skill Licenses", "SkillLicenses", "Skill License" };

    public static (List<StaffRecord>? Data, List<string> Errors, List<string> Warnings) ExtractStaffRecordsWithDiagnostics()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;
        if (wb == null) return (null, new List<string> { "No active workbook found." }, new List<string>());

        string masterSheetName = SheetOperations.ResolveSheetName(MasterSheetAliases, MasterSheetName);
        dynamic masterWs = wb.Sheets[masterSheetName];
        if (masterWs == null) return (null, new List<string> { $"Worksheet '{MasterSheetName}' was not found." }, new List<string>());

        // 1. Extract Parent Staff records dynamically by header
        List<StaffRecord> staffs = TableMapper.ExtractFromTable<StaffRecord>((object)masterWs, MasterTableName);

        // 2. Extract Child Collections across sheets
        ExtractChildRecords(wb, staffs);

        // 3. Run validation rules engine (statically bound via object cast)
        var (errors, warnings) = ValidateStaffDataset((object)masterWs, staffs);

        return (staffs, errors, warnings);
    }

    public static List<StaffRecord> ExtractStaffRecords()
    {
        var (data, _, _) = ExtractStaffRecordsWithDiagnostics();
        return data ?? new List<StaffRecord>();
    }

    public static void PopulateWorksheet(List<StaffRecord> records)
    {
        if (records == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            // 1. Populate Master Table
            string targetSheet = SheetOperations.ResolveSheetName(MasterSheetAliases, MasterSheetName);
            dynamic masterWs = wb.Sheets[targetSheet];
            if (masterWs != null)
            {
                try { masterWs.Unprotect("SKSB_Internal_Guard"); } catch { }
                TableMapper.PopulateTableSafe<StaffRecord>((object)masterWs, MasterTableName, records);
                SheetProtections.ApplyStyleDrivenProtection(masterWs);
            }

            // 2. Populate Child Tables
            PopulateChildTables(wb, records);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[StaffSheetMapper.PopulateWorksheet Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    public static void PopulateSkillLicenseTypesTable(bool showDialogIfFailed = false)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            string sLicSheet = SheetOperations.ResolveSheetName(SkillLicenseAliases, "Skill Licenses");
            dynamic ws = wb.Sheets[sLicSheet];
            if (ws == null) return;

            var types = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
            if (types.Count > 0)
            {
                try { ws.Unprotect("SKSB_Internal_Guard"); } catch { }
                TableMapper.PopulateTableSafe<SkillLicenseType>((object)ws, "SkillLicenseTypesTable", types);
                SheetProtections.ApplyStyleDrivenProtection(ws);
            }
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

    private static void ExtractChildRecords(dynamic wb, List<StaffRecord> staffs)
    {
        var staffMap = staffs
            .Where(s => s.StaffID > 0)
            .GroupBy(s => s.StaffID)
            .ToDictionary(g => g.Key, g => g.First());

        // Vehicle Licenses
        string vLicSheet = SheetOperations.ResolveSheetName(VehicleLicenseAliases, "Vehicle Licenses");
        try
        {
            dynamic ws = wb.Sheets[vLicSheet];
            if (ws != null)
            {
                List<VehicleLicense> vLicenses = TableMapper.ExtractFromTable<VehicleLicense>((object)ws, "VehicleLicensesTable");
                foreach (var lic in vLicenses)
                {
                    if (staffMap.TryGetValue(lic.StaffID, out StaffRecord? s) && s != null)
                    {
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        s.VehicleLicenses.Add(lic);
                    }
                }
            }
        }
        catch { }

        // Skill Licenses
        string sLicSheet = SheetOperations.ResolveSheetName(SkillLicenseAliases, "Skill Licenses");
        try
        {
            dynamic ws = wb.Sheets[sLicSheet];
            if (ws != null)
            {
                List<SkillLicense> sLicenses = TableMapper.ExtractFromTable<SkillLicense>((object)ws, "SkillLicensesTable");
                foreach (var lic in sLicenses)
                {
                    if (staffMap.TryGetValue(lic.StaffID, out StaffRecord? s) && s != null)
                    {
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        s.SkillLicenses.Add(lic);
                    }
                }
            }
        }
        catch { }
    }

    private static void PopulateChildTables(dynamic wb, List<StaffRecord> records)
    {
        // 1. Vehicle Licenses
        string vLicSheet = SheetOperations.ResolveSheetName(VehicleLicenseAliases, "Vehicle Licenses");
        try
        {
            dynamic ws = wb.Sheets[vLicSheet];
            if (ws != null)
            {
                try { ws.Unprotect("SKSB_Internal_Guard"); } catch { }
                var allVLicenses = records.SelectMany(s => s.VehicleLicenses.Select(lic =>
                {
                    if (string.IsNullOrWhiteSpace(lic.StaffName))
                        lic.StaffName = s.StaffDetails.StaffName;
                    return lic;
                })).ToList();

                TableMapper.PopulateTableSafe<VehicleLicense>((object)ws, "VehicleLicensesTable", allVLicenses);
                SheetProtections.ApplyStyleDrivenProtection(ws);
            }
        }
        catch { }

        // 2. Skill Licenses & Lookup Reference List
        string sLicSheet = SheetOperations.ResolveSheetName(SkillLicenseAliases, "Skill Licenses");
        try
        {
            dynamic ws = wb.Sheets[sLicSheet];
            if (ws != null)
            {
                try { ws.Unprotect("SKSB_Internal_Guard"); } catch { }
                var allSLicenses = records.SelectMany(s => s.SkillLicenses.Select(lic =>
                {
                    if (string.IsNullOrWhiteSpace(lic.StaffName))
                        lic.StaffName = s.StaffDetails.StaffName;
                    return lic;
                })).ToList();

                TableMapper.PopulateTableSafe<SkillLicense>((object)ws, "SkillLicensesTable", allSLicenses);

                // Populate side reference list (SkillLicenseTypesTable)
                var types = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
                if (types.Count > 0)
                {
                    TableMapper.PopulateTableSafe<SkillLicenseType>((object)ws, "SkillLicenseTypesTable", types);
                }

                SheetProtections.ApplyStyleDrivenProtection(ws);
            }
        }
        catch { }
    }

    private static (List<string> Errors, List<string> Warnings) ValidateStaffDataset(object masterWsObj, List<StaffRecord> staffs)
    {
        var errors = new List<string>();
        var warnings = new List<string>();

        dynamic masterWs = masterWsObj;
        dynamic? table = null;
        try { table = masterWs.ListObjects[MasterTableName]; } catch { }
        if (table == null && masterWs.ListObjects.Count > 0) table = masterWs.ListObjects[1];

        if (table != null && table.DataBodyRange != null)
        {
            int rowCount = (int)table.DataBodyRange.Rows.Count;
            int startRow = (int)table.DataBodyRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;
            int colSpan = (int)table.ListColumns.Count;

            var (_, ruleErrors, ruleWarnings) = RulesExtractor.EvaluateAndExtract(
                worksheet: (object)masterWs,
                startRow: startRow,
                maxRows: rowCount,
                startCol: startCol,
                columnSpan: colSpan,
                driverColIndex: 1
            );

            errors.AddRange(ruleErrors);
            warnings.AddRange(ruleWarnings);
        }

        // Check for duplicate primary keys
        var idGroups = staffs.GroupBy(s => s.StaffID);
        foreach (var group in idGroups)
        {
            if (group.Count() > 1)
            {
                errors.Add($"Duplicate Staff ID '{group.Key}' detected {group.Count()} times in master table.");
            }
        }

        return (errors, warnings);
    }
}