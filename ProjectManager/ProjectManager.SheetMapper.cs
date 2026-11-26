namespace SKSB_App.ProjectManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

internal static partial class SheetMapper
{
    // =========================================================================
    // 1. SHEET / TABLE IDENTITY
    // =========================================================================
    public const string MasterSheetName = "General Projects";
    public const string MasterTableName = "GeneralProjectsTable";

    public const string GeneralTableName = "GeneralProjectsTable";
    public const string InternalTableName = "InternalProjectsTable";

    internal sealed record ProjectSheet(bool IsInternal, string[] Aliases, string DefaultSheetName, string TableName);

    /// <summary>
    /// General Projects carries every operational category (iJob / MWO / PVT / Other) in its 'Category' column;
    /// Internal Projects is the lightweight non-operational sheet (no Category / company / revenue columns).
    /// </summary>
    internal static readonly ProjectSheet GeneralSheet =
        new(false, new[] { "General Projects", "GeneralProjects", "General" }, "General Projects", GeneralTableName);

    internal static readonly ProjectSheet InternalSheet =
        new(true, new[] { "Internal Projects", "InternalProjects", "Internal" }, "Internal Projects", InternalTableName);

    internal static readonly ProjectSheet[] ProjectSheets = { GeneralSheet, InternalSheet };

    internal static readonly string[] MasterSheetAliases = GeneralSheet.Aliases;
    internal static readonly string[] CompaniesAliases = { "Companies", "Company" };

    // =========================================================================
    // 2. DECLARATIVE WORKBOOK MAPPING (2 MASTER TABLES, NO CHILD TABLES)
    // =========================================================================
    private static readonly WorkbookMapping<ProjectRecord, string> Mapping = new()
    {
        MasterSheetAliases = MasterSheetAliases,
        MasterSheetName = MasterSheetName,
        MasterTableName = MasterTableName,
        EntityLabel = "project",
        MasterKeyLabel = "Project Code",
        MasterKeySelector = p => p.ProjectCode,
        KeyComparer = StringComparer.OrdinalIgnoreCase,
        IsValidKey = code => !string.IsNullOrWhiteSpace(code),

        // Both project tables are read / written by the module; primary-key uniqueness is then checked across them.
        CustomMasterExtractor = ExtractAllMasterTables,
        CustomMasterPopulator = PopulateAllMasterTables,

        PreValidationHook = (wb, masters, errors, warnings) => ExtractCompaniesForValidation((object)wb, errors),
        CustomValidator = (wb, masters, errors, warnings) => ValidateProjects(masters, errors, warnings),

        OnPopulateAuxiliaryTables = (wb, masters) => PopulateCompanies((object)wb),

        RecalculateAfterPopulate = true,

        ExcludeValidationTables = new[] { "ProjectSummaryTable" },

        HealTableOrder = new[]
        {
            "CompaniesTable",
            GeneralTableName,
            InternalTableName
        },

        AuditTitle = "Project Manager Populate Audit",
        MasterLabel = "Projects",
        InspectTables = new[]
        {
            GeneralTableName, InternalTableName, "CompaniesTable"
        }
    };

    // =========================================================================
    // 3. EXTRACT / POPULATE ENTRY POINTS
    // =========================================================================
    internal static (List<ProjectRecord>? Data, List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ExtractProjectRecordsWithDiagnostics()
        => WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);

    internal static void PopulateWorksheet(List<ProjectRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 4. MULTI-SHEET MASTER EXTRACTION
    // =========================================================================
    private static List<ProjectRecord> ExtractAllMasterTables(object wbObj, List<string> errors)
    {
        var all = new List<ProjectRecord>();

        foreach (var sheet in ProjectSheets)
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, sheet.Aliases, sheet.DefaultSheetName);
            if (ws == null)
            {
                errors.Add($"  Worksheet '{sheet.DefaultSheetName}' was not found in workbook.");
                continue;
            }

            try
            {
                var rows = TableMapper.ExtractFromTable<ProjectRecord>((object)ws, sheet.TableName);
                foreach (var record in rows)
                {
                    record.ProjectCode = (record.ProjectCode ?? string.Empty).Trim().ToUpperInvariant();
                    if (string.IsNullOrWhiteSpace(record.ProjectCode)) continue; // reported by the required-field rule

                    if (sheet.IsInternal)
                    {
                        // Non-operational: never auto-scheduled, never billed to a client SQL Account project.
                        record.Category = ProjectCategory.Internal;
                        record.ProjectDetails.AutoSchedule = false;
                        record.ProjectDetails.SqlAccountProjectCode = "----";
                    }
                    else if (record.Category == ProjectCategory.Internal)
                    {
                        // The row's own 'Category' column is authoritative on General Projects, but Internal belongs elsewhere.
                        errors.Add($"  [{sheet.DefaultSheetName}]: Project '{record.ProjectCode}' has Category 'Internal'; " +
                                   $"move it to the '{InternalSheet.DefaultSheetName}' sheet.");
                    }

                    RuleCalculator.ApplyRecordCalculations(record); // Actual Project Cost, Profit / Margin
                    all.Add(record);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[ProjectManager] Failed extracting '{sheet.TableName}'", ex);
                errors.Add($"  [{sheet.DefaultSheetName}]: Failed reading '{sheet.TableName}': {ex.Message}");
            }
        }

        return all;
    }

    // =========================================================================
    // 5. MULTI-SHEET MASTER POPULATION
    // =========================================================================
    private static void PopulateAllMasterTables(object wbObj, List<ProjectRecord> records)
    {
        foreach (var sheet in ProjectSheets)
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, sheet.Aliases, sheet.DefaultSheetName);
            if (ws == null) continue;

            var forSheet = sheet.IsInternal
                ? records.Where(r => r.Category == ProjectCategory.Internal)
                         .OrderBy(r => r.ProjectCode, StringComparer.OrdinalIgnoreCase)
                         .ToList()
                : records.Where(r => r.Category != ProjectCategory.Internal)
                         .OrderBy(r => CategoryOrder(r.Category))
                         .ThenBy(r => r.ProjectCode, StringComparer.OrdinalIgnoreCase)
                         .ToList();

            foreach (var record in forSheet) RuleCalculator.ApplyRecordCalculations(record);

            try
            {
                // Actual Project Cost / Profit / Margin are written from the C# calculation above.
                TableMapper.PopulateTableSafe<ProjectRecord>((object)ws, sheet.TableName, forSheet);
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[ProjectManager] Failed populating '{sheet.TableName}'", ex);
            }
        }
    }

    private static int CategoryOrder(ProjectCategory c) => c switch
    {
        ProjectCategory.iJob => 0,
        ProjectCategory.MWO => 1,
        ProjectCategory.PVT => 2,
        ProjectCategory.Other => 3,
        _ => 4
    };

    // =========================================================================
    // 6. DOMAIN VALIDATION
    // =========================================================================
    private static void ValidateProjects(List<ProjectRecord> masters, List<string> errors, List<string> warnings)
    {
        var companyIds = _extractedCompanies != null && _extractedCompanies.Count > 0
            ? new HashSet<long>(_extractedCompanies.Select(c => c.CompanyID))
            : null;

        foreach (var p in masters)
        {
            // Internal Projects: no Category / company requirement.
            if (p.Category == ProjectCategory.Internal) continue;

            string where = $"  [{GeneralSheet.DefaultSheetName}]: Project '{p.ProjectCode}'";

            if (p.Category is not (ProjectCategory.iJob or ProjectCategory.MWO or ProjectCategory.PVT or ProjectCategory.Other))
                errors.Add($"{where} must have a Category of iJob, MWO, PVT or Other.");

            long id = p.ProjectDetails.CompanyID;
            if (id <= 0)
                errors.Add($"{where} requires a Company ID.");
            else if (companyIds != null && !companyIds.Contains(id))
                errors.Add($"{where} references Company ID '{id}' which does not exist in 'CompaniesTable'.");
        }
    }

    // =========================================================================
    // 7. SHEET-CHANGE OBSERVER (auto Project Code + category-aware SQL Account Project)
    // =========================================================================
    private const string HdrProjectCode = "Project Code";
    private const string HdrCategory = "Category";
    private const string HdrSqlAccount = "SQL Account Project";
    private const string HdrCompanyId = "Company ID";
    private const string HdrCompanyName = "Company Name";
    private const string InternalPrefix = "INT-";

    /// <summary>Category-generated SQL Account Project values; a cell holding one of these may be re-defaulted.</summary>
    private static readonly HashSet<string> GeneratedSqlAccountDefaults =
        new(StringComparer.OrdinalIgnoreCase) { "2016 - C6-MWO", "2016 - PRIV - SHORT", "----" };

    private static bool _isAutoFilling;

    private readonly record struct ChangedArea(int FirstRow, int RowCount, int FirstCol, int ColCount);

    /// <summary>
    /// Registered with <c>ParameterCells</c>; runs on the Excel event thread. Only reads COM here: the changed areas
    /// are captured as plain coordinates and every write is deferred to <see cref="ExcelAsyncUtil.QueueAsMacro"/>.
    /// </summary>
    internal static void OnSheetChanged(dynamic sheet, dynamic target)
    {
        if (_isAutoFilling || SheetOperations.IsTemplateEditMode) return;

        try
        {
            object ws = sheet;
            string? tableName = FindProjectTableOnSheet(ws);
            if (tableName == null) return;

            dynamic wb = ((dynamic)ws).Parent;
            if (SheetOperations.IsMasterTemplateFile((object)wb)) return;
            if (!SheetOperations.IsModuleWorkbook((object)wb, Events.ProjectSyncDescriptor)) return;

            var areas = new List<ChangedArea>();
            dynamic targetAreas = target.Areas;
            int areaCount = Convert.ToInt32(targetAreas.Count);
            for (int i = 1; i <= areaCount; i++)
            {
                dynamic a = targetAreas[i];
                areas.Add(new ChangedArea(
                    Convert.ToInt32(a.Row), Convert.ToInt32(a.Rows.Count),
                    Convert.ToInt32(a.Column), Convert.ToInt32(a.Columns.Count)));
            }
            if (areas.Count == 0) return;

            ExcelAsyncUtil.QueueAsMacro(() => AutoFillProjectRows(ws, tableName, areas));
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ProjectManager] OnSheetChanged skipped: {ex.Message}");
        }
    }

    private static string? FindProjectTableOnSheet(object ws)
    {
        foreach (string name in new[] { GeneralTableName, InternalTableName })
        {
            try
            {
                dynamic? lo = ((dynamic)ws).ListObjects[name];
                if (lo != null) return name;
            }
            catch { }
        }
        return null;
    }

    // -------------------------------------------------------------------------
    // 7a. MACRO-CONTEXT AUTO-FILL (frozen Excel state, bulk 2D reads / column writes)
    // -------------------------------------------------------------------------
    private static void AutoFillProjectRows(object wsObj, string tableName, List<ChangedArea> areas)
    {
        if (_isAutoFilling) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool prevEvents = true, prevScreen = true;
        object? prevCalc = null;
        try { prevEvents = (bool)xlApp.EnableEvents; } catch { }
        try { prevScreen = (bool)xlApp.ScreenUpdating; } catch { }
        try { prevCalc = xlApp.Calculation; } catch { }

        _isAutoFilling = true;
        try
        {
            try { xlApp.EnableEvents = false; } catch { }
            try { xlApp.ScreenUpdating = false; } catch { }
            try { xlApp.Calculation = -4135; } catch { } // xlCalculationManual

            dynamic ws = wsObj;
            dynamic table = ws.ListObjects[tableName];
            dynamic? body = table.DataBodyRange;
            if (body == null) return;

            var grid = TableGrid.Read(table, body);
            if (grid == null || grid.RowCount == 0) return;

            int codeCol = grid.Col(HdrProjectCode);
            if (codeCol < 1) return;

            // Changed data rows (1-based within the body) and the set of changed column headers.
            var changedRows = new SortedSet<int>();
            var changedHeaders = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var a in areas)
            {
                int r0 = Math.Max(a.FirstRow, grid.FirstRow);
                int r1 = Math.Min(a.FirstRow + a.RowCount - 1, grid.FirstRow + grid.RowCount - 1);
                int c0 = Math.Max(a.FirstCol, grid.FirstCol);
                int c1 = Math.Min(a.FirstCol + a.ColCount - 1, grid.FirstCol + grid.ColCount - 1);
                if (r0 > r1 || c0 > c1) continue;

                for (int r = r0; r <= r1; r++) changedRows.Add(r - grid.FirstRow + 1);
                for (int c = c0; c <= c1; c++) changedHeaders.Add(grid.Headers[c - grid.FirstCol]);
            }
            if (changedRows.Count == 0) return;

            var codes = CollectKnownProjectCodes((object)ws.Parent, tableName, grid, codeCol);

            if (tableName == GeneralTableName)
                AutoFillGeneralRows((object)ws.Parent, grid, changedRows, changedHeaders, codes);
            else
                AutoFillInternalRows(grid, changedRows, codes);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[ProjectManager] Auto-fill on '{tableName}' failed", ex);
        }
        finally
        {
            try { if (prevCalc != null) xlApp.Calculation = prevCalc; } catch { }
            try { xlApp.ScreenUpdating = prevScreen; } catch { }
            try { xlApp.EnableEvents = prevEvents; } catch { }
            _isAutoFilling = false;
        }
    }

    private static void AutoFillGeneralRows(object wbObj, TableGrid grid, SortedSet<int> rows, HashSet<string> changedHeaders, HashSet<string> codes)
    {
        int codeCol = grid.Col(HdrProjectCode);
        int catCol = grid.Col(HdrCategory);
        int sqlCol = grid.Col(HdrSqlAccount);
        int idCol = grid.Col(HdrCompanyId);
        int nameCol = grid.Col(HdrCompanyName);
        if (catCol < 1) return;

        bool categoryChanged = changedHeaders.Contains(HdrCategory);
        bool companyChanged = changedHeaders.Contains(HdrCompanyId) || changedHeaders.Contains(HdrCompanyName);
        if (!categoryChanged && !companyChanged) return;

        CompanySqlLookup? companies = null;
        CompanySqlLookup Companies() => companies ??= CompanySqlLookup.Load(wbObj);
        string CompanyDefault(int r) => sqlCol < 1 ? string.Empty
            : Companies().Find(idCol > 0 ? grid.Text(r, idCol) : string.Empty, nameCol > 0 ? grid.Text(r, nameCol) : string.Empty);

        bool codeDirty = false, sqlDirty = false;

        foreach (int r in rows)
        {
            if (!Enum.TryParse<GeneralProjectCategory>(grid.Text(r, catCol), ignoreCase: true, out var category)
                || category == GeneralProjectCategory.None)
                continue;

            string sqlCurrent = sqlCol > 0 ? grid.Text(r, sqlCol) : string.Empty;

            if (categoryChanged)
            {
                // a. Project Code: saved codes are permanent; unsaved rows follow their category.
                string prefix = CategoryPrefix(category);
                string current = grid.Text(r, codeCol).Trim().ToUpperInvariant();
                bool keep = current.Length > 0
                    && (ServerSyncCommands.IsSavedProjectCode(current) || current.StartsWith(prefix, StringComparison.OrdinalIgnoreCase));
                if (!keep)
                {
                    codes.Remove(current);
                    string next = NextCode(prefix, codes);
                    codes.Add(next);
                    grid.Set(r, codeCol, next);
                    codeDirty = true;
                }

                // b. SQL Account Project: only replace blanks / values a category or company produced earlier.
                if (sqlCol > 0)
                {
                    string companyDefault = CompanyDefault(r);
                    bool replaceable = sqlCurrent.Length == 0
                        || GeneratedSqlAccountDefaults.Contains(sqlCurrent)
                        || (companyDefault.Length > 0 && string.Equals(sqlCurrent, companyDefault, StringComparison.OrdinalIgnoreCase));
                    if (replaceable)
                    {
                        string target = category switch
                        {
                            GeneralProjectCategory.MWO => "2016 - C6-MWO",
                            GeneralProjectCategory.PVT => "2016 - PRIV - SHORT",
                            GeneralProjectCategory.Other => "----",
                            _ => companyDefault
                        };
                        if (!string.Equals(target, sqlCurrent, StringComparison.Ordinal))
                        {
                            grid.Set(r, sqlCol, target);
                            sqlCurrent = target;
                            sqlDirty = true;
                        }
                    }
                }
            }

            // c. iJob company picked after the category: fill a blank SQL Account Project from the company.
            if (companyChanged && category == GeneralProjectCategory.iJob && sqlCol > 0 && sqlCurrent.Length == 0)
            {
                string companyDefault = CompanyDefault(r);
                if (companyDefault.Length > 0)
                {
                    grid.Set(r, sqlCol, companyDefault);
                    sqlDirty = true;
                }
            }
        }

        if (codeDirty) grid.WriteColumn(codeCol);
        if (sqlDirty) grid.WriteColumn(sqlCol);
    }

    private static void AutoFillInternalRows(TableGrid grid, SortedSet<int> rows, HashSet<string> codes)
    {
        int codeCol = grid.Col(HdrProjectCode);
        bool dirty = false;

        foreach (int r in rows)
        {
            if (grid.Text(r, codeCol).Trim().Length > 0) continue;
            if (!grid.RowHasAnyValue(r)) continue;

            string next = NextCode(InternalPrefix, codes);
            codes.Add(next);
            grid.Set(r, codeCol, next);
            dirty = true;
        }

        if (dirty) grid.WriteColumn(codeCol);
    }

    // -------------------------------------------------------------------------
    // 7b. CODE GENERATION
    // -------------------------------------------------------------------------
    private static string CategoryPrefix(GeneralProjectCategory c) => c switch
    {
        GeneralProjectCategory.iJob => "IJOB-",
        GeneralProjectCategory.MWO => "MWO-",
        GeneralProjectCategory.PVT => "PVT-",
        _ => "OTH-"
    };

    private static string NextCode(string prefix, IEnumerable<string> existing)
    {
        int max = 0;
        foreach (string code in existing)
        {
            if (!code.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)) continue;
            if (int.TryParse(code.AsSpan(prefix.Length), NumberStyles.None, CultureInfo.InvariantCulture, out int n) && n > max)
                max = n;
        }
        return prefix + (max + 1).ToString("D4", CultureInfo.InvariantCulture);
    }

    /// <summary>Project codes in both project tables plus every code in the server baseline (normalized).</summary>
    private static HashSet<string> CollectKnownProjectCodes(object wbObj, string currentTable, TableGrid currentGrid, int currentCodeCol)
    {
        var codes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        void Add(string? c)
        {
            string n = (c ?? string.Empty).Trim().ToUpperInvariant();
            if (n.Length > 0) codes.Add(n);
        }

        for (int r = 1; r <= currentGrid.RowCount; r++) Add(currentGrid.Text(r, currentCodeCol));

        var other = currentTable == GeneralTableName ? InternalSheet : GeneralSheet;
        try
        {
            dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, other.Aliases, other.DefaultSheetName);
            if (ws != null)
            {
                dynamic table = ws.ListObjects[other.TableName];
                dynamic? body = table.DataBodyRange;
                if (body != null)
                {
                    var grid = TableGrid.Read(table, body);
                    int col = grid?.Col(HdrProjectCode) ?? 0;
                    if (grid != null && col > 0)
                        for (int r = 1; r <= grid.RowCount; r++) Add(grid.Text(r, col));
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ProjectManager] Could not read '{other.TableName}' codes: {ex.Message}");
        }

        foreach (string c in ServerSyncCommands.GetBaselineProjectCodes()) Add(c);
        return codes;
    }

    // -------------------------------------------------------------------------
    // 7c. BULK TABLE SNAPSHOT (one Value2 read; dirty columns written back as single 2D arrays)
    // -------------------------------------------------------------------------
    private sealed class TableGrid
    {
        private readonly dynamic _body;
        private readonly object[,] _values;
        private readonly int _rowBase, _colBase;

        public int FirstRow { get; }
        public int FirstCol { get; }
        public int RowCount { get; }
        public int ColCount { get; }
        public string[] Headers { get; }

        private TableGrid(dynamic body, object[,] values, string[] headers, int firstRow, int firstCol, int rows, int cols)
        {
            _body = body;
            _values = values;
            _rowBase = values.GetLowerBound(0);
            _colBase = values.GetLowerBound(1);
            Headers = headers;
            FirstRow = firstRow;
            FirstCol = firstCol;
            RowCount = rows;
            ColCount = cols;
        }

        public static TableGrid? Read(dynamic table, dynamic body)
        {
            int rows = Convert.ToInt32(body.Rows.Count);
            int cols = Convert.ToInt32(body.Columns.Count);
            if (rows <= 0 || cols <= 0) return null;

            object[,] header = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, cols);
            var headers = new string[cols];
            int hr = header.GetLowerBound(0), hc = header.GetLowerBound(1);
            for (int c = 0; c < cols; c++)
                headers[c] = c + hc <= header.GetUpperBound(1) ? (header[hr, c + hc]?.ToString() ?? string.Empty).Trim() : string.Empty;

            object[,] values = TableMapper.GetRange2DValues(body, rows, cols);
            return new TableGrid(body, values, headers, Convert.ToInt32(body.Row), Convert.ToInt32(body.Column), rows, cols);
        }

        /// <summary>1-based column index of <paramref name="header"/>, or 0 when absent.</summary>
        public int Col(string header)
        {
            for (int c = 0; c < Headers.Length; c++)
                if (string.Equals(Headers[c], header, StringComparison.OrdinalIgnoreCase)) return c + 1;
            return 0;
        }

        public object? Get(int r, int c) => _values[r - 1 + _rowBase, c - 1 + _colBase];

        public string Text(int r, int c) => c < 1 ? string.Empty : Convert.ToString(Get(r, c), CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;

        public void Set(int r, int c, object? value) => _values[r - 1 + _rowBase, c - 1 + _colBase] = value;

        public bool RowHasAnyValue(int r)
        {
            for (int c = 1; c <= ColCount; c++)
                if (Text(r, c).Length > 0) return true;
            return false;
        }

        /// <summary>Writes one whole column back in a single 2D assignment (never cell-by-cell).</summary>
        public void WriteColumn(int c)
        {
            var column = new object?[RowCount, 1];
            for (int r = 1; r <= RowCount; r++) column[r - 1, 0] = Get(r, c);
            dynamic target = _body.Columns[c];
            target.Value2 = column;
        }
    }

    // -------------------------------------------------------------------------
    // 7d. COMPANIES TABLE LOOKUP (iJob SQL Account Project by Company ID / Short Company Name)
    // -------------------------------------------------------------------------
    private sealed class CompanySqlLookup
    {
        private readonly Dictionary<string, string> _byId = new(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _byName = new(StringComparer.OrdinalIgnoreCase);

        public static CompanySqlLookup Load(object wbObj)
        {
            var lookup = new CompanySqlLookup();
            try
            {
                dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, CompaniesAliases, "Companies");
                if (ws == null) return lookup;

                dynamic table = ws.ListObjects["CompaniesTable"];
                dynamic? body = table.DataBodyRange;
                if (body == null) return lookup;

                var grid = TableGrid.Read(table, body);
                if (grid == null) return lookup;

                int idCol = grid.Col("Company ID");
                int nameCol = grid.Col("Short Company Name");
                int sqlCol = grid.Col("iJob SQL Account Project");
                if (sqlCol < 1) return lookup;

                for (int r = 1; r <= grid.RowCount; r++)
                {
                    string sql = grid.Text(r, sqlCol);
                    string id = NormalizeId(grid.Text(r, idCol));
                    string name = grid.Text(r, nameCol);
                    if (id.Length > 0) lookup._byId.TryAdd(id, sql);
                    if (name.Length > 0) lookup._byName.TryAdd(name, sql);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[ProjectManager] Could not read CompaniesTable for SQL Account defaults: {ex.Message}");
            }
            return lookup;
        }

        /// <summary>Company ID wins; Company Name covers the moment before the ID column has been synced.</summary>
        public string Find(string companyId, string companyName)
        {
            string id = NormalizeId(companyId);
            if (id.Length > 0 && _byId.TryGetValue(id, out var byId)) return byId;
            if (companyName.Length > 0 && _byName.TryGetValue(companyName.Trim(), out var byName)) return byName;
            return string.Empty;
        }

        // Value2 returns numeric IDs as doubles ("12" vs "12.0"): compare as integers when possible.
        private static string NormalizeId(string raw) =>
            double.TryParse(raw, NumberStyles.Float, CultureInfo.InvariantCulture, out double d) && d > 0
                ? ((long)d).ToString(CultureInfo.InvariantCulture)
                : raw.Trim();
    }
}
