namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

public static partial class SheetMapper
{
    public const string MasterSheetName = "All Staffs";
    public const string MasterTableName = "AllStaffsTable";

    internal static readonly string[] MasterSheetAliases = { "All Staffs", "AllStaffs", "Staff_Master", "Staffs", "All Staff", "All Workers" };
    private static readonly string[] VehicleLicenseAliases = { "Vehicle Licenses", "VehicleLicenses", "Vehicle License", "Staff_Licenses" };
    private static readonly string[] SkillLicenseAliases = { "Skill Licenses", "SkillLicenses", "Skill License", "Staff_Skills" };
    private static readonly string[] StaffSalaryAliases = { "Staff Salaries", "StaffSalaries", "Salaries" };
    private const string StaffSalariesSheetName = "Staff Salaries";
    private const string StaffSalaryTableName = "StaffSalaryTable";
    private static readonly string[] StaffLeaveAliases = { "Staff Leaves", "StaffLeaves", "Leaves" };
    private const string StaffLeavesSheetName = "Staff Leaves";
    private const string StaffLeavesTableName = "StaffLeavesTable";

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
                ValidateForeignKeyExists = true,
                OrderChildRecords = list => list.OrderByDescending(l => l.RenewDate).ThenByDescending(l => l.VehicleLicenseID).ToList()
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
                ValidateForeignKeyExists = true,
                OrderChildRecords = list => list.OrderByDescending(l => l.RenewDate).ThenByDescending(l => l.SkillLicenseID).ToList()
            },
            new ChildTableBinding<StaffRecord, StaffSalary, long>
            {
                SheetAliases = StaffSalaryAliases,
                DefaultSheetName = StaffSalariesSheetName,
                TableName = StaffSalaryTableName,
                ForeignKeySelector = w => w.StaffID,
                ParentCollectionSelector = s => s.Salaries,
                BackfillFromParent = (s, w) =>
                {
                    if (w.StaffID <= 0) w.StaffID = s.StaffID;
                    if (string.IsNullOrWhiteSpace(w.StaffName)) w.StaffName = s.StaffDetails.StaffName;
                },
                ChildPrimaryKeySelector = w => w.SalaryID,
                ChildPrimaryKeyLabel = "Salary ID",
                ValidateForeignKeyExists = true,
                OrderChildRecords = list => list.OrderByDescending(w => w.EffectiveDate).ThenBy(w => w.StaffID).ToList()
            },
            new ChildTableBinding<StaffRecord, StaffLeave, long>
            {
                SheetAliases = StaffLeaveAliases,
                DefaultSheetName = StaffLeavesSheetName,
                TableName = StaffLeavesTableName,
                ForeignKeySelector = l => l.StaffID,
                ParentCollectionSelector = s => s.Leaves,
                BackfillFromParent = (s, l) =>
                {
                    if (l.StaffID <= 0) l.StaffID = s.StaffID;
                    if (string.IsNullOrWhiteSpace(l.StaffName)) l.StaffName = s.StaffDetails.StaffName;
                },
                ChildPrimaryKeySelector = l => l.LeaveID,
                ChildPrimaryKeyLabel = "Leave ID",
                ValidateForeignKeyExists = false,
                OrderChildRecords = list => list.OrderByDescending(l => l.LeaveDate).ThenByDescending(l => l.LeaveID).ToList()
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

        OnPopulateAuxiliaryTables = (wb, masters) =>
        {
            PopulateSkillLicenseTypesFromCache((object)wb);
            ApplySalaryRowFormats((object)wb, masters);
            InitializeStaffSummaryContext((object)wb, masters);
        },

        AuditTitle = "Staff Manager Populate Audit",
        MasterLabel = "Staff Profiles",
        InspectTables = new[] { "AllStaffsTable", "VehicleLicensesTable", "SkillLicensesTable", "SkillLicenseTypesTable", StaffSalaryTableName, StaffLeavesTableName }
    };

    // =========================================================================
    // 2. MASTER EXTRACTION & POPULATE PIPELINES
    // =========================================================================
    public static (List<StaffRecord>? Data, List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ExtractStaffRecordsWithDiagnostics()
    {
        var result = WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);
        if (result.Data != null) ServerSyncCommands.DeriveShortNames(result.Data); // blank Short Name -> first + last word of Staff Name
        return result;
    }

    public static void PopulateWorksheet(List<StaffRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 2b. PER-ROW SALARY NUMBER FORMATS ("/mth" vs "/day")
    // =========================================================================
    /// <summary>
    /// Post-populate: applies the monthly / daily salary number format per row to AllStaffsTable[Current Salary],
    /// StaffSalaryTable[Salary Rate] and StaffSalaryTable[Salary Before]. Key columns are read in one bulk call;
    /// contiguous rows sharing a format are written as one range. Blank / em-dash cells are left untouched. Never
    /// selects or activates.
    /// </summary>
    private static void ApplySalaryRowFormats(object wbObj, List<StaffRecord> masters)
    {
        try
        {
            var basisById = masters
                .Where(s => s.StaffID > 0)
                .GroupBy(s => s.StaffID)
                .ToDictionary(g => g.Key, g => g.First().StaffDetails.CurrentSalaryBasis);

            // Salary Before basis keyed by (Staff ID, Date): unique per staff after SQL Payroll reconstruction.
            var beforeBasisByKey = new Dictionary<(long, DateOnly), SalaryBasis>();
            foreach (var s in masters.Where(s => s.StaffID > 0))
            {
                s.ResolveSalaryBeforeBases();
                foreach (var row in s.Salaries.Where(r => r.SalaryBefore.HasValue))
                    beforeBasisByKey[(s.StaffID, row.EffectiveDate)] = row.SalaryBeforeBasis;
            }

            RunFrozen(() =>
            {
                // a. AllStaffsTable[Current Salary]: basis looked up by Staff ID.
                dynamic? masterWs = SheetMapperHelper.ResolveSheetSafely(wbObj, MasterSheetAliases, MasterSheetName);
                if (masterWs != null)
                {
                    ApplyPerRowFormat((object)masterWs, MasterTableName, "Staff ID", "Current Salary", key =>
                    {
                        long id = ToLong(key);
                        return id > 0 && basisById.TryGetValue(id, out var b) ? b : SalaryBasis.Monthly;
                    });
                }

                dynamic? salaryWs = SheetMapperHelper.ResolveSheetSafely(wbObj, StaffSalaryAliases, StaffSalariesSheetName);
                if (salaryWs != null)
                {
                    // b. StaffSalaryTable[Salary Rate]: basis read from the row's own New Salary Basis cell.
                    ApplyPerRowFormat((object)salaryWs, StaffSalaryTableName, "New Salary Basis", "Salary Rate", key =>
                        Enum.TryParse<SalaryBasis>(Convert.ToString(key)?.Trim(), ignoreCase: true, out var b) ? b : SalaryBasis.Monthly);

                    // c. StaffSalaryTable[Salary Before]: basis of the preceding revision (SalaryBeforeBasis).
                    ApplySalaryBeforeFormats((object)salaryWs, beforeBasisByKey);
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[StaffManager] Applying salary row formats failed: {ex.Message}");
        }
    }

    /// <summary>
    /// Formats StaffSalaryTable[Salary Before] per row with <see cref="StaffSalary.SalaryBeforeBasis"/>, matched by the
    /// row's (Staff ID, Date). Staff ID / Date are bulk-read once; formats are written in contiguous runs.
    /// </summary>
    private static void ApplySalaryBeforeFormats(object wsObj, IReadOnlyDictionary<(long, DateOnly), SalaryBasis> beforeBasisByKey)
    {
        dynamic ws = wsObj;
        dynamic? table = null;
        try { table = ws.ListObjects[StaffSalaryTableName]; } catch { }
        if (table == null) return;

        dynamic? idBody = null, dateBody = null;
        try { idBody = table.ListColumns["Staff ID"].DataBodyRange; } catch { }
        try { dateBody = table.ListColumns["Date"].DataBodyRange; } catch { }
        if (idBody == null || dateBody == null) return;

        int rows = (int)idBody.Rows.Count;
        if (rows <= 0) return;

        object[,] ids = TableMapper.GetRange2DValues(idBody, rows, 1);
        object[,] dates = TableMapper.GetRange2DValues(dateBody, rows, 1);
        int i0 = ids.GetLowerBound(0), ic = ids.GetLowerBound(1);
        int d0 = dates.GetLowerBound(0), dc = dates.GetLowerBound(1);

        var basisAt = new SalaryBasis[rows];
        for (int i = 0; i < rows; i++)
        {
            long id = ToLong(ids[i0 + i, ic]);
            DateOnly? date = ToDateOnly(dates[d0 + i, dc]);
            basisAt[i] = id > 0 && date.HasValue && beforeBasisByKey.TryGetValue((id, date.Value), out var b) && b != SalaryBasis.None
                ? b
                : SalaryBasis.Monthly;
        }

        ApplyPerRowFormat(wsObj, StaffSalaryTableName, "Staff ID", "Salary Before", _ => SalaryBasis.Monthly, rowIndex => basisAt[rowIndex]);
    }

    private static DateOnly? ToDateOnly(object? v) => v switch
    {
        double d => DateOnly.FromDateTime(DateTime.FromOADate(d)),
        DateTime dt => DateOnly.FromDateTime(dt),
        string s when DateTime.TryParse(s, out var parsed) => DateOnly.FromDateTime(parsed),
        _ => null
    };

    private static long ToLong(object? v) => v switch
    {
        double d => (long)d,
        long l => l,
        int i => i,
        _ => long.TryParse(Convert.ToString(v)?.Trim(), out long n) ? n : 0
    };

    private static void ApplyPerRowFormat(object wsObj, string tableName, string keyHeader, string valueHeader,
        Func<object?, SalaryBasis> basisFor, Func<int, SalaryBasis>? basisAtRow = null)
    {
        dynamic ws = wsObj;
        dynamic? table = null;
        try { table = ws.ListObjects[tableName]; } catch { }
        if (table == null) return;

        dynamic? keyBody = null, valueBody = null;
        try { keyBody = table.ListColumns[keyHeader].DataBodyRange; } catch { }
        try { valueBody = table.ListColumns[valueHeader].DataBodyRange; } catch { }
        if (keyBody == null || valueBody == null) return;

        int rows = (int)valueBody.Rows.Count;
        if (rows <= 0) return;

        object[,] keys = TableMapper.GetRange2DValues(keyBody, rows, 1);
        object[,] values = TableMapper.GetRange2DValues(valueBody, rows, 1);
        int k0 = keys.GetLowerBound(0), kc = keys.GetLowerBound(1);
        int v0 = values.GetLowerBound(0), vc = values.GetLowerBound(1);

        string? FormatAt(int i)
        {
            object? raw = values[v0 + i, vc];
            if (raw == null) return null;
            if (raw is string s && (s.Trim().Length == 0 || s.Trim() == "—" || s.Trim() == "-")) return null;
            SalaryBasis basis = basisAtRow != null ? basisAtRow(i) : basisFor(keys[k0 + i, kc]);
            return basis == SalaryBasis.Daily ? RuleRegistry.WageDailyFormat : RuleRegistry.WageMonthlyFormat;
        }

        // Contiguous runs of the same format -> one NumberFormat write per run.
        int runStart = -1;
        string? runFmt = null;
        for (int i = 0; i <= rows; i++)
        {
            string? fmt = i < rows ? FormatAt(i) : null;
            if (fmt == runFmt && i < rows) continue;

            if (runFmt != null && runStart >= 0)
            {
                try
                {
                    dynamic runRange = ws.Range[valueBody.Cells[runStart + 1, 1], valueBody.Cells[i, 1]];
                    runRange.NumberFormat = runFmt;
                }
                catch { }
            }
            runStart = i;
            runFmt = fmt;
        }
    }

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

        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SkillLicenseAliases, "Skill Licenses");
        if (types == null || types.Count == 0)
        {
            if (ws != null) _skillTypeNameByRow = SnapshotSkillLicenseTypeRows(ws);
            return;
        }
        if (ws == null) return;

        TableMapper.PopulateTableSafe<SkillLicenseType>((object)ws, "SkillLicenseTypesTable", types);
        var names = types.Select(t => t.SkillLicenseName).Where(n => !string.IsNullOrWhiteSpace(n)).ToList();
        DataList.SyncValues("SkillLicenseTypeList", names);
        _skillTypeNameByRow = SnapshotSkillLicenseTypeRows(ws);
    }

    // =========================================================================
    // 4. SKILL LICENSE TYPE RENAME CASCADE (Excel-side)
    // =========================================================================
    private const string SkillTypeTableName = "SkillLicenseTypesTable";
    private const string SkillLicensesTableName = "SkillLicensesTable";
    private const string SkillTypeHeader = "Skill License Type";

    private static Dictionary<int, string> _skillTypeNameByRow = new();
    private static readonly Action<dynamic, dynamic> _skillTypeObserver = OnSkillLicenseTypesChanged;
    private static bool _skillTypeCascadeHooked;
    private static bool _isCascadingSkillTypeRename;

    internal static void RegisterSkillTypeCascadeHooks()
    {
        if (_skillTypeCascadeHooked) return;
        _skillTypeCascadeHooked = true;
        ParameterCells.RegisterChangeObserver(_skillTypeObserver);
    }

    internal static void UnregisterSkillTypeCascadeHooks()
    {
        ParameterCells.UnregisterChangeObserver(_skillTypeObserver);
        _skillTypeCascadeHooked = false;
    }

    private static dynamic? GetTypeColumnBody(dynamic ws, string tableName)
    {
        try
        {
            dynamic table = ws.ListObjects[tableName];
            return table.ListColumns[SkillTypeHeader].DataBodyRange;
        }
        catch { return null; }
    }

    private static Dictionary<int, string> SnapshotSkillLicenseTypeRows(dynamic ws)
    {
        var result = new Dictionary<int, string>();
        try
        {
            dynamic? body = GetTypeColumnBody(ws, SkillTypeTableName);
            if (body == null) return result;

            int rows = (int)body.Rows.Count;
            int firstRow = (int)body.Row;
            object[,] grid = TableMapper.GetRange2DValues(body, rows, 1);
            int lower = grid.GetLowerBound(0);
            for (int i = 0; i < rows; i++)
            {
                string text = Convert.ToString(grid[lower + i, grid.GetLowerBound(1)])?.Trim() ?? string.Empty;
                result[firstRow + i] = text;
            }
        }
        catch { }
        return result;
    }

    private static void OnSkillLicenseTypesChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (_isCascadingSkillTypeRename) return;
            if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;
            if (!SheetOperations.IsModuleWorkbook(null, Events.StaffSyncDescriptor)) return;

            string sheetName = (string)sh.Name;
            if (!SkillLicenseAliases.Any(a => a.Equals(sheetName, StringComparison.OrdinalIgnoreCase))) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? typeColBody = GetTypeColumnBody(sh, SkillTypeTableName);
            if (typeColBody == null || xlApp.Intersect(target, typeColBody) == null) return;

            int editedRow = (int)target.Row;
            bool isSingleCell = (int)target.Rows.Count == 1 && (int)target.Columns.Count == 1;

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    dynamic app = ExcelDnaUtil.Application;
                    dynamic? wb = app.ActiveWorkbook;
                    if (wb == null) return;
                    dynamic? ws = SheetMapperHelper.ResolveSheetSafely((object)wb, SkillLicenseAliases, "Skill Licenses");
                    if (ws == null) return;

                    Dictionary<int, string> currentByRow = SnapshotSkillLicenseTypeRows(ws);

                    if (isSingleCell
                        && _skillTypeNameByRow.Count == currentByRow.Count
                        && _skillTypeNameByRow.TryGetValue(editedRow, out string? oldName)
                        && !string.IsNullOrWhiteSpace(oldName)
                        && currentByRow.TryGetValue(editedRow, out string? newName)
                        && !string.IsNullOrWhiteSpace(newName)
                        && !oldName.Equals(newName, StringComparison.OrdinalIgnoreCase))
                    {
                        bool oldStillExists = currentByRow
                            .Any(kv => kv.Key != editedRow && kv.Value.Equals(oldName, StringComparison.OrdinalIgnoreCase));
                        if (!oldStillExists)
                        {
                            int updatedCount = 0;
                            _isCascadingSkillTypeRename = true;
                            try
                            {
                                RunFrozen(() =>
                                {
                                    dynamic? bodyCol = GetTypeColumnBody(ws, SkillLicensesTableName);
                                    if (bodyCol != null)
                                    {
                                        int rowCount = (int)bodyCol.Rows.Count;
                                        object[,] grid = TableMapper.GetRange2DValues(bodyCol, rowCount, 1);
                                        int r0 = grid.GetLowerBound(0), c0 = grid.GetLowerBound(1);
                                        for (int i = 0; i < rowCount; i++)
                                        {
                                            string cell = Convert.ToString(grid[r0 + i, c0])?.Trim() ?? string.Empty;
                                            if (cell.Equals(oldName, StringComparison.OrdinalIgnoreCase))
                                            {
                                                grid[r0 + i, c0] = newName;
                                                updatedCount++;
                                            }
                                        }
                                        if (updatedCount > 0) bodyCol.Value2 = grid;
                                    }

                                    DataList.SyncValues("SkillLicenseTypeList",
                                        currentByRow.Values.Where(v => v.Length > 0).Distinct(StringComparer.OrdinalIgnoreCase).ToList());
                                });
                            }
                            finally { _isCascadingSkillTypeRename = false; }

                            if (updatedCount > 0)
                                SheetOperations.ShowStatusConfirmation($"Renamed skill license type '{oldName}' -> '{newName}' ({updatedCount} license row(s) updated).");
                        }
                    }

                    _skillTypeNameByRow = currentByRow;
                }
                catch (Exception ex)
                {
                    AppLogger.Warn($"[StaffManager] Skill type rename cascade warning: {ex.Message}");
                }
            });
        }
        catch { }
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
