namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

public static class SheetMapper
{
    public const string MasterSheetName = "All Staffs";
    public const string MasterTableName = "AllStaffsTable";

    private static readonly string[] MasterSheetAliases = { "All Staffs", "AllStaffs", "Staff_Master", "Staffs", "All Staff", "All Workers" };
    private static readonly string[] VehicleLicenseAliases = { "Vehicle Licenses", "VehicleLicenses", "Vehicle License", "Staff_Licenses" };
    private static readonly string[] SkillLicenseAliases = { "Skill Licenses", "SkillLicenses", "Skill License", "Staff_Skills" };

    // =========================================================================
    // 1. MASTER EXTRACTION & VALIDATION PIPELINE
    // =========================================================================
    public static (List<StaffRecord>? Data, List<string> Errors, List<string> Warnings) ExtractStaffRecordsWithDiagnostics()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<string> { "No active workbook found." }, new List<string>());

            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs == null)
            {
                return (null, new List<string> { $"Worksheet '{MasterSheetName}' was not found in workbook." }, new List<string>());
            }

            // 1. Extract Master Staff Records (runs EvaluateAndHealTable automatically inside TableMapper)
            List<StaffRecord> staffs = TableMapper.ExtractFromTable<StaffRecord>((object)masterWs, MasterTableName);
            foreach (var s in staffs)
            {
                s.StaffDetails.StaffID = s.StaffID;
            }

            // 2. Extract Child License Collections & Check Foreign Key Integrity
            var childErrors = ExtractChildRecords((object)wb, staffs);

            // 3. Validate All Tables Across All Sheets (AllStaffsTable, VehicleLicensesTable, SkillLicensesTable, SkillLicenseTypesTable)
            var (errors, warnings) = RulesExtractor.ValidateAllWorkbookTables((object)wb);
            errors.AddRange(childErrors);

            // 4. Primary Key & Date Integrity Checks on Master Table
            var idGroups = staffs
                .Where(s => s.StaffID > 0)
                .GroupBy(s => s.StaffID);

            foreach (var group in idGroups)
            {
                if (group.Count() > 1)
                {
                    errors.Add($"  [{MasterSheetName}]: Duplicate Staff ID '{group.Key}' detected {group.Count()} times in '{MasterTableName}'.");
                }
            }

            foreach (var s in staffs)
            {
                if (s.StaffDetails.ResignDate.HasValue && s.StaffDetails.ResignDate.Value < s.StaffDetails.JoinDate)
                {
                    errors.Add($"  [{MasterSheetName}]: Staff ID '{s.StaffID}' ({s.StaffDetails.StaffName}) Resign Date ({s.StaffDetails.ResignDate.Value:dd/MM/yyyy}) cannot be earlier than Join Date ({s.StaffDetails.JoinDate:dd/MM/yyyy}).");
                }
            }

            return (staffs, errors, warnings);
        }
        catch (Exception ex)
        {
            return (null, new List<string> { $"Failed extracting staff records: {ex.Message}" }, new List<string>());
        }
    }

    public static List<StaffRecord> ExtractStaffRecords()
    {
        var (data, _, _) = ExtractStaffRecordsWithDiagnostics();
        return data ?? new List<StaffRecord>();
    }

    // =========================================================================
    // 2. MASTER POPULATE PIPELINE
    // =========================================================================
    public static void PopulateWorksheet(List<StaffRecord> records)
    {
        if (records == null) return;

        var sortedList = records
            .OrderBy(s => s.StaffDetails.ContractStatus == ContractStatus.Resigned ? 1 : 0)
            .ThenBy(s => s.StaffID)
            .ToList();

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        bool origUpdating = true;
        int origCalculation = -4105; // xlCalculationAutomatic
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }

        try
        {
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            // 1. Populate Master Table
            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs != null)
            {
                SheetProtections.ExecuteWithBypass((object)masterWs, new Action(() =>
                {
                    TableMapper.PopulateTableSafe<StaffRecord>((object)masterWs, MasterTableName, sortedList);
                }));
            }

            // 2. Populate Child Tables & Auxiliary SkillLicenseTypesTable
            PopulateChildTables((object)wb, sortedList);

#if DEBUG
            ShowPopulateDiagnostics((object)wb, sortedList);
#endif
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[StaffSheetMapper.PopulateWorksheet Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.Calculation = origCalculation; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    // =========================================================================
    // 3. CHILD & AUXILIARY TABLE OPERATIONS
    // =========================================================================
    private static List<string> ExtractChildRecords(object wbObj, List<StaffRecord> staffs)
    {
        var childErrors = new List<string>();
        dynamic wb = wbObj;

        var staffMap = staffs
            .Where(s => s.StaffID > 0)
            .GroupBy(s => s.StaffID)
            .ToDictionary(g => g.Key, g => g.First());

        // 1. Vehicle Licenses
        dynamic? vWs = ResolveSheetSafely(wbObj, VehicleLicenseAliases, "Vehicle Licenses");
        if (vWs != null)
        {
            try
            {
                List<VehicleLicense> vLicenses = TableMapper.ExtractFromTable<VehicleLicense>((object)vWs, "VehicleLicensesTable");
                var seenIds = new HashSet<long>();

                foreach (var lic in vLicenses)
                {
                    if (lic.StaffID <= 0) continue;

                    if (lic.VehicleLicenseID > 0 && !seenIds.Add(lic.VehicleLicenseID))
                    {
                        childErrors.Add($"  [Vehicle Licenses]: Duplicate Vehicle License ID '{lic.VehicleLicenseID}'.");
                    }

                    if (staffMap.TryGetValue(lic.StaffID, out StaffRecord? s) && s != null)
                    {
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        s.VehicleLicenses.Add(lic);
                    }
                    else
                    {
                        childErrors.Add($"  [Vehicle Licenses]: Staff ID '{lic.StaffID}' does not exist in '{MasterSheetName}'.");
                    }
                }
            }
            catch { }
        }

        // 2. Skill Licenses
        dynamic? sWs = ResolveSheetSafely(wbObj, SkillLicenseAliases, "Skill Licenses");
        if (sWs != null)
        {
            try
            {
                List<SkillLicense> sLicenses = TableMapper.ExtractFromTable<SkillLicense>((object)sWs, "SkillLicensesTable");
                var seenIds = new HashSet<long>();

                foreach (var lic in sLicenses)
                {
                    if (lic.StaffID <= 0) continue;

                    if (lic.SkillLicenseID > 0 && !seenIds.Add(lic.SkillLicenseID))
                    {
                        childErrors.Add($"  [Skill Licenses]: Duplicate Skill License ID '{lic.SkillLicenseID}'.");
                    }

                    if (staffMap.TryGetValue(lic.StaffID, out StaffRecord? s) && s != null)
                    {
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        s.SkillLicenses.Add(lic);
                    }
                    else
                    {
                        childErrors.Add($"  [Skill Licenses]: Staff ID '{lic.StaffID}' does not exist in '{MasterSheetName}'.");
                    }
                }
            }
            catch { }
        }

        return childErrors;
    }

    private static void PopulateChildTables(object wbObj, List<StaffRecord> records)
    {
        // 1. Vehicle Licenses
        dynamic? vWs = ResolveSheetSafely(wbObj, VehicleLicenseAliases, "Vehicle Licenses");
        if (vWs != null)
        {
            try
            {
                SheetProtections.ExecuteWithBypass((object)vWs, new Action(() =>
                {
                    var allVLicenses = records.SelectMany(s => s.VehicleLicenses.Select(lic =>
                    {
                        if (lic.StaffID <= 0) lic.StaffID = s.StaffID;
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        return lic;
                    })).ToList();

                    TableMapper.PopulateTableSafe<VehicleLicense>((object)vWs, "VehicleLicensesTable", allVLicenses);
                }));
            }
            catch { }
        }

        // 2. Skill Licenses & Side Reference Table (SkillLicenseTypesTable)
        dynamic? sWs = ResolveSheetSafely(wbObj, SkillLicenseAliases, "Skill Licenses");
        if (sWs != null)
        {
            try
            {
                SheetProtections.ExecuteWithBypass((object)sWs, new Action(() =>
                {
                    var allSLicenses = records.SelectMany(s => s.SkillLicenses.Select(lic =>
                    {
                        if (lic.StaffID <= 0) lic.StaffID = s.StaffID;
                        if (string.IsNullOrWhiteSpace(lic.StaffName))
                            lic.StaffName = s.StaffDetails.StaffName;
                        return lic;
                    })).ToList();

                    TableMapper.PopulateTableSafe<SkillLicense>((object)sWs, "SkillLicensesTable", allSLicenses);

                    // Load SkillLicenseTypes from local cache (already refreshed during FetchStaffRecords)
                    var savedTypes = OfflineSnapshot.Load<List<SkillLicenseType>>();
                    if (savedTypes == null || savedTypes.Count == 0)
                    {
                        savedTypes = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
                    }

                    if (savedTypes != null && savedTypes.Count > 0)
                    {
                        TableMapper.PopulateTableSafe<SkillLicenseType>((object)sWs, "SkillLicenseTypesTable", savedTypes);
                        var names = savedTypes.Select(t => t.SkillLicenseName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
                        DataList.SyncValues("SkillLicenseTypeList", names);
                    }
                }));
            }
            catch { }
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

            dynamic? ws = ResolveSheetSafely((object)wb, SkillLicenseAliases, "Skill Licenses");
            if (ws == null) return;

            var types = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
            if (types.Count == 0)
            {
                types = OfflineSnapshot.Load<List<SkillLicenseType>>() ?? new List<SkillLicenseType>();
            }

            if (types.Count > 0)
            {
                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
                    TableMapper.PopulateTableSafe<SkillLicenseType>((object)ws, "SkillLicenseTypesTable", types);
                }));

                var names = types.Select(t => t.SkillLicenseName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
                DataList.SyncValues("SkillLicenseTypeList", names);
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

    public static List<SkillLicenseType> ExtractSkillLicenseTypes()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return new List<SkillLicenseType>();

            dynamic? sWs = ResolveSheetSafely((object)wb, SkillLicenseAliases, "Skill Licenses");
            if (sWs == null) return new List<SkillLicenseType>();

            var extracted = TableMapper.ExtractFromTable<SkillLicenseType>((object)sWs, "SkillLicenseTypesTable");
            var distinct = extracted
                .Where(t => !string.IsNullOrWhiteSpace(t.SkillLicenseName))
                .Select(t => t.SkillLicenseName.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(n => n)
                .Select((name, idx) => new SkillLicenseType
                {
                    SkillLicenseTypeID = idx + 1,
                    SkillLicenseName = name
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

    internal static dynamic? ResolveSheetSafely(object wbObj, string[] aliases, string defaultName)
    {
        dynamic wb = wbObj;
        string resolvedName = SheetOperations.ResolveSheetName(aliases, defaultName);
        try
        {
            dynamic s = wb.Sheets[resolvedName];
            if (s != null) return s;
        }
        catch { }

        foreach (dynamic ws in wb.Sheets)
        {
            string name = ws.Name?.ToString() ?? string.Empty;
            if (aliases.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase)))
            {
                return ws;
            }
        }
        return null;
    }

#if DEBUG
    private static void ShowPopulateDiagnostics(object wbObj, List<StaffRecord> records)
    {
        try
        {
            if (!DebugDialog.DebugEnabled) return;

            dynamic wb = wbObj;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== STAFF MANAGER POPULATE AUDIT ===");
            sb.AppendLine($"Total Staff Profiles Loaded: {records.Count}");
            sb.AppendLine();

            string[] inspectTables = { "AllStaffsTable", "VehicleLicensesTable", "SkillLicensesTable", "SkillLicenseTypesTable" };
            foreach (var tName in inspectTables)
            {
                dynamic? tbl = null;
                string sName = "Unknown";
                foreach (dynamic ws in wb.Sheets)
                {
                    try { tbl = ws.ListObjects[tName]; if (tbl != null) { sName = ws.Name?.ToString() ?? "Unknown"; break; } } catch { }
                }

                if (tbl != null)
                {
                    int cols = (int)tbl.ListColumns.Count;
                    int rows = tbl.DataBodyRange != null ? (int)tbl.DataBodyRange.Rows.Count : 0;
                    int startCol = (int)tbl.HeaderRowRange.Column;
                    int startRow = (int)tbl.HeaderRowRange.Row;
                    sb.AppendLine($"  [{tName}] on Sheet '{sName}':");
                    sb.AppendLine($"      Dimensions : {rows} data row(s) × {cols} column(s)");
                    sb.AppendLine($"      Coordinates: Row {startRow}, Columns {startCol} to {startCol + cols - 1}");
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine($"  [{tName}]: NOT FOUND on any worksheet");
                    sb.AppendLine();
                }
            }

            DebugDialog.ShowLog("Staff Manager - Populate Audit", sb.ToString());
        }
        catch { }
    }
#endif
}