namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Validation;

// =============================================================================
// 1. DECLARATIVE CHILD TABLE BINDING
// =============================================================================

/// <summary>Type-erased view of a child binding so differently-typed children can share one list.</summary>
public interface IChildTableBinding<TMaster, TKey> where TKey : notnull
{
    string TableName { get; }

    /// <summary>Extracts the child table and attaches each row to its parent in <paramref name="masterMap"/>.</summary>
    void ExtractInto(
        object wbObj,
        IReadOnlyDictionary<TKey, TMaster> masterMap,
        Func<TKey, bool> isValidKey,
        string masterSheetName,
        string masterKeyLabel,
        List<string> errors);

    /// <summary>Flattens every parent's child collection and writes it to the child table.</summary>
    void Populate(object wbObj, IList<TMaster> masters);
}

/// <summary>
/// Describes one child ListObject (e.g. "VehicleLicensesTable") owned by a master record through a foreign key.
/// </summary>
public sealed class ChildTableBinding<TMaster, TChild, TKey> : IChildTableBinding<TMaster, TKey>
    where TChild : new()
    where TKey : notnull
{
    public string[] SheetAliases { get; set; } = Array.Empty<string>();
    public string DefaultSheetName { get; set; } = string.Empty;
    public string TableName { get; set; } = string.Empty;

    public Func<TChild, TKey> ForeignKeySelector { get; set; } = null!;
    public Func<TMaster, IList<TChild>> ParentCollectionSelector { get; set; } = null!;

    /// <summary>Copies parent identity onto the child (e.g. StaffID / StaffName, VehiclePlate) on both extract and populate.</summary>
    public Action<TMaster, TChild>? BackfillFromParent { get; set; }

    /// <summary>When set, duplicate positive child IDs are reported as errors.</summary>
    public Func<TChild, long>? ChildPrimaryKeySelector { get; set; }
    public string ChildPrimaryKeyLabel { get; set; } = "ID";

    /// <summary>When true, a child whose foreign key is missing from the master table is reported as an error.</summary>
    public bool ValidateForeignKeyExists { get; set; }

    public void ExtractInto(
        object wbObj,
        IReadOnlyDictionary<TKey, TMaster> masterMap,
        Func<TKey, bool> isValidKey,
        string masterSheetName,
        string masterKeyLabel,
        List<string> errors)
    {
        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetAliases, DefaultSheetName);
        if (ws == null) return;

        try
        {
            List<TChild> children = TableMapper.ExtractFromTable<TChild>((object)ws, TableName);
            var seenIds = new HashSet<long>();

            foreach (var child in children)
            {
                TKey fk = ForeignKeySelector(child);
                if (fk is null || !isValidKey(fk)) continue;

                if (ChildPrimaryKeySelector != null)
                {
                    long childId = ChildPrimaryKeySelector(child);
                    if (childId > 0 && !seenIds.Add(childId))
                    {
                        errors.Add($"  [{DefaultSheetName}]: Duplicate {ChildPrimaryKeyLabel} '{childId}'.");
                    }
                }

                if (masterMap.TryGetValue(fk, out TMaster? parent) && parent != null)
                {
                    BackfillFromParent?.Invoke(parent, child);
                    ParentCollectionSelector(parent).Add(child);
                }
                else if (ValidateForeignKeyExists)
                {
                    errors.Add($"  [{DefaultSheetName}]: {masterKeyLabel} '{fk}' does not exist in '{masterSheetName}'.");
                }
            }
        }
        catch (Exception ex)
        {
            // Surface the failure instead of swallowing it: a silently-empty child collection would be
            // treated as "all children deleted" by the full-batch server reconciliation.
            AppLogger.Error($"[WorkbookMapper] Failed extracting '{TableName}'", ex);
            errors.Add($"  [{DefaultSheetName}]: Failed reading '{TableName}': {ex.Message}");
        }
    }

    public void Populate(object wbObj, IList<TMaster> masters)
    {
        dynamic? ws = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetAliases, DefaultSheetName);
        if (ws == null) return;

        try
        {
            var allChildren = new List<TChild>();
            foreach (var master in masters)
            {
                IList<TChild>? owned = ParentCollectionSelector(master);
                if (owned == null) continue;

                foreach (var child in owned)
                {
                    BackfillFromParent?.Invoke(master, child);
                    allChildren.Add(child);
                }
            }

            SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
            {
                TableMapper.PopulateTableSafe<TChild>((object)ws, TableName, allChildren);
            }));
            SheetProtections.ApplyStyleDrivenProtection((object)ws);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[WorkbookMapper] Failed populating '{TableName}'", ex);
        }
    }
}

// =============================================================================
// 2. DECLARATIVE WORKBOOK MAPPING (MASTER + CHILDREN + HOOKS)
// =============================================================================

/// <summary>
/// Describes a module workbook: the master table, its child tables and the optional customization hooks
/// used by <see cref="WorkbookMapper"/>. Declared once per module as a static readonly descriptor.
/// </summary>
public sealed class WorkbookMapping<TMaster, TKey>
    where TMaster : new()
    where TKey : notnull
{
    // ---- Master table identity ----
    public string[] MasterSheetAliases { get; set; } = Array.Empty<string>();
    public string MasterSheetName { get; set; } = string.Empty;
    public string MasterTableName { get; set; } = string.Empty;

    /// <summary>Lower-case noun used in failure messages, e.g. "staff" or "vehicle".</summary>
    public string EntityLabel { get; set; } = "master";

    /// <summary>Display name of the master primary key, e.g. "Staff ID" or "Vehicle Plate".</summary>
    public string MasterKeyLabel { get; set; } = "ID";

    public Func<TMaster, TKey> MasterKeySelector { get; set; } = null!;
    public IEqualityComparer<TKey>? KeyComparer { get; set; }

    /// <summary>Filters out uninitialised keys (e.g. StaffID &lt;= 0, blank plate) on both master and child rows.</summary>
    public Func<TKey, bool>? IsValidKey { get; set; }

    public List<IChildTableBinding<TMaster, TKey>> Children { get; set; } = new();

    // ---- Extraction hooks ----
    public Action<TMaster>? NormalizeMasterRecord { get; set; }

    /// <summary>
    /// Runs after child extraction but BEFORE the workbook-wide rule validation, for steps that may repair
    /// the sheet first (e.g. restoring a reference row that is still in use). Args: workbook, masters, errors, warnings.
    /// </summary>
    public Action<dynamic, List<TMaster>, List<string>, List<string>>? PreValidationHook { get; set; }

    /// <summary>Domain-specific rules, run last. Args: workbook, masters, errors, warnings.</summary>
    public Action<dynamic, List<TMaster>, List<string>, List<string>>? CustomValidator { get; set; }

    // ---- Populate hooks ----
    public Func<List<TMaster>, List<TMaster>>? OrderMasterRecords { get; set; }

    /// <summary>Runs just before the master table is written. Arg: the master worksheet.</summary>
    public Action<dynamic>? OnBeforeMasterPopulate { get; set; }

    /// <summary>Runs after master and child tables are written. Args: workbook, ordered masters.</summary>
    public Action<dynamic, List<TMaster>>? OnPopulateAuxiliaryTables { get; set; }

    /// <summary>Forces a full recalculation after populate so formula columns are current before healing.</summary>
    public bool RecalculateAfterPopulate { get; set; }

    /// <summary>Tables to run through <see cref="RuleEvaluator.EvaluateAndHealTable"/> after populate, in order (repeats allowed).</summary>
    public string[] HealTableOrder { get; set; } = Array.Empty<string>();

    // ---- DEBUG populate audit ----
    public string AuditTitle { get; set; } = "Populate Audit";
    public string MasterLabel { get; set; } = "Records";
    public string[] InspectTables { get; set; } = Array.Empty<string>();
}

// =============================================================================
// 3. GENERIC EXTRACT / POPULATE PIPELINES
// =============================================================================
public static class WorkbookMapper
{
    // =========================================================================
    // 3a. MASTER + CHILD EXTRACTION & VALIDATION PIPELINE
    // =========================================================================
    public static (List<TMaster>? Data, List<string> Errors, List<string> Warnings) ExtractMasterAndChildrenWithDiagnostics<TMaster, TKey>(
        WorkbookMapping<TMaster, TKey> map)
        where TMaster : new()
        where TKey : notnull
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<string> { "No active workbook found." }, new List<string>());
            object wbObj = wb;

            dynamic? masterWs = SheetMapperHelper.ResolveSheetSafely(wbObj, map.MasterSheetAliases, map.MasterSheetName);
            if (masterWs == null)
            {
                return (null, new List<string> { $"Worksheet '{map.MasterSheetName}' was not found in workbook." }, new List<string>());
            }

            // 1. Extract master records (runs EvaluateAndHealTable automatically inside TableMapper)
            List<TMaster> masters = TableMapper.ExtractFromTable<TMaster>((object)masterWs, map.MasterTableName);

            // 2. Normalise master records
            if (map.NormalizeMasterRecord != null)
            {
                foreach (var m in masters) map.NormalizeMasterRecord(m);
            }

            // 3. Extract child collections and attach them to their parent
            Func<TKey, bool> isValidKey = map.IsValidKey ?? (_ => true);
            var masterMap = new Dictionary<TKey, TMaster>(map.KeyComparer);
            foreach (var m in masters)
            {
                TKey key = map.MasterKeySelector(m);
                if (key is null || !isValidKey(key)) continue;
                masterMap.TryAdd(key, m); // first occurrence wins; duplicates are reported in step 5
            }

            var hookErrors = new List<string>();
            var hookWarnings = new List<string>();
            foreach (var child in map.Children)
            {
                child.ExtractInto(wbObj, masterMap, isValidKey, map.MasterSheetName, map.MasterKeyLabel, hookErrors);
            }

            map.PreValidationHook?.Invoke(wbObj, masters, hookErrors, hookWarnings);

            // 4. Validate all tables across all sheets
            var (errors, warnings) = RulesExtractor.ValidateAllWorkbookTables(wbObj);
            errors.AddRange(hookErrors);
            warnings.AddRange(hookWarnings);

            // 5. Master primary key uniqueness
            var keyGroups = masters
                .Select(map.MasterKeySelector)
                .Where(k => k is not null && isValidKey(k))
                .GroupBy(k => k, map.KeyComparer);

            foreach (var group in keyGroups)
            {
                int occurrences = group.Count();
                if (occurrences > 1)
                {
                    errors.Add($"  [{map.MasterSheetName}]: Duplicate {map.MasterKeyLabel} '{group.Key}' detected {occurrences} times in '{map.MasterTableName}'.");
                }
            }

            // 6. Domain-specific rules
            map.CustomValidator?.Invoke(wbObj, masters, errors, warnings);

            return (masters, errors, warnings);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[WorkbookMapper] Failed extracting {map.EntityLabel} records", ex);
            return (null, new List<string> { $"Failed extracting {map.EntityLabel} records: {ex.Message}" }, new List<string>());
        }
    }

    // =========================================================================
    // 3b. MASTER + CHILD POPULATE PIPELINE
    // =========================================================================
    public static void PopulateMasterAndChildren<TMaster, TKey>(
        WorkbookMapping<TMaster, TKey> map,
        List<TMaster> records)
        where TMaster : new()
        where TKey : notnull
    {
        if (records == null) return;

        // 2. Module-defined ordering (pure in-memory, done before any COM work)
        List<TMaster> ordered = map.OrderMasterRecords?.Invoke(records) ?? records;

        // 1. Standard COM freeze for the whole pass
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        bool origUpdating = true;
        int origCalculation = -4105; // xlCalculationAutomatic
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }

        object? auditWorkbook = null;
        try
        {
            dynamic? wb = xlApp.ActiveWorkbook;
            if (wb == null) return;
            object wbObj = wb;

            // 3 + 4. Master table
            dynamic? masterWs = SheetMapperHelper.ResolveSheetSafely(wbObj, map.MasterSheetAliases, map.MasterSheetName);
            if (masterWs != null)
            {
                SheetProtections.ExecuteWithBypass((object)masterWs, new Action(() =>
                {
                    map.OnBeforeMasterPopulate?.Invoke((object)masterWs);
                    TableMapper.PopulateTableSafe<TMaster>((object)masterWs, map.MasterTableName, ordered);
                }));
                SheetProtections.ApplyStyleDrivenProtection((object)masterWs);
            }

            // 5. Child tables
            foreach (var child in map.Children)
            {
                child.Populate(wbObj, ordered);
            }

            // 6. Auxiliary / reference tables
            map.OnPopulateAuxiliaryTables?.Invoke(wbObj, ordered);

            // 7. Post-populate calculation and ordered self-healing pass
            if (map.RecalculateAfterPopulate)
            {
                try { xlApp.CalculateFull(); }
                catch (Exception exCalc) { AppLogger.Warn($"[WorkbookMapper] CalculateFull warning: {exCalc.Message}"); }
            }

            foreach (string tableName in map.HealTableOrder)
            {
                foreach (dynamic sheet in wb.Worksheets)
                {
                    try
                    {
                        dynamic tbl = sheet.ListObjects[tableName];
                        if (tbl != null)
                        {
                            RuleEvaluator.EvaluateAndHealTable(sheet, tbl);

                            // Calculation is frozen to manual: refresh dependants so the next table
                            // in the chain heals against current formula results.
                            try { xlApp.Calculate(); } catch { }
                            break;
                        }
                    }
                    catch { } // Table is not on this sheet
                }
            }

            auditWorkbook = wbObj;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[WorkbookMapper] Failed populating {map.EntityLabel} records", ex);
        }
        finally
        {
            try { xlApp.Calculation = origCalculation; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }

        // 8. DEBUG audit (no-op in Release). Shown only after the freeze is lifted: a modal dialog
        //    must never open while ScreenUpdating / EnableEvents are suppressed.
        if (auditWorkbook != null)
        {
            SheetMapperHelper.ShowPopulateDiagnostics(auditWorkbook, map.AuditTitle, map.MasterLabel, ordered.Count, map.InspectTables);
        }
    }
}
