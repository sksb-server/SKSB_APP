namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
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

            TableMapper.PopulateTableSafe<TChild>((object)ws, TableName, allChildren);
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

    // ---- Multi-table master override (e.g. one master table per category sheet) ----
    /// <summary>
    /// When set, replaces single-table master extraction: the module reads every master table itself
    /// (and may add its own errors). Args: workbook, errors. The standard child / validation pipeline still runs.
    /// </summary>
    public Func<object, List<string>, List<TMaster>>? CustomMasterExtractor { get; set; }

    /// <summary>When set, replaces the single-table master write. Args: workbook, ordered masters.</summary>
    public Action<object, List<TMaster>>? CustomMasterPopulator { get; set; }

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

    /// <summary>Runs [ColumnCalculate] / [ColumnRollup] over the masters (after children are attached on extract, before populate).</summary>
    public Action<object, List<TMaster>>? ApplyMasterCalculations { get; set; }

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

            var customErrors = new List<string>();
            List<TMaster> masters;
            if (map.CustomMasterExtractor != null)
            {
                masters = map.CustomMasterExtractor(wbObj, customErrors);
            }
            else
            {
                dynamic? masterWs = SheetMapperHelper.ResolveSheetSafely(wbObj, map.MasterSheetAliases, map.MasterSheetName);
                if (masterWs == null)
                {
                    return (null, new List<string> { $"Worksheet '{map.MasterSheetName}' was not found in workbook." }, new List<string>());
                }

                // 1. Extract master records (runs EvaluateAndHealTable automatically inside TableMapper)
                masters = TableMapper.ExtractFromTable<TMaster>((object)masterWs, map.MasterTableName);
            }

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

            var hookErrors = new List<string>(customErrors);
            var hookWarnings = new List<string>();
            foreach (var child in map.Children)
            {
                child.ExtractInto(wbObj, masterMap, isValidKey, map.MasterSheetName, map.MasterKeyLabel, hookErrors);
            }

            // 3b. Child collections are attached: evaluate declarative rollups / calculations on each master
            map.ApplyMasterCalculations?.Invoke(wbObj, masters);

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
    // 3a. POST-FREEZE CONDITIONAL FORMATTING REFRESH
    // =========================================================================
    /// <summary>
    /// Icon Set / formula-driven conditional formats evaluated while Calculation was manual and
    /// ScreenUpdating off can keep stale visuals. Marks populated table bodies dirty, recalculates
    /// once, then forces a repaint. Must run with EnableEvents still false (no Worksheet_Change).
    /// Writes nothing to cells, so it is unaffected by validation-based locking.
    /// </summary>
    private static void RefreshConditionalFormatting<TMaster, TKey>(
        dynamic xlApp, object wbObj, WorkbookMapping<TMaster, TKey> map)
        where TMaster : new()
        where TKey : notnull
    {
        try
        {
#if DEBUG
            LogCfDiagnostics(wbObj, map.MasterTableName, "BEFORE refresh");
#endif
            NudgeIconSetRules(xlApp, wbObj);

#if DEBUG
            LogCfDiagnostics(wbObj, map.MasterTableName, "AFTER inline refresh");
#endif

            // Excel can defer CF rendering until the current macro returns. Queue a second nudge
            // that runs on the macro thread afterwards, with events suppressed again.
            try
            {
                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    dynamic app = ExcelDnaUtil.Application;
                    bool prevEvents = true;
                    try { prevEvents = (bool)app.EnableEvents; app.EnableEvents = false; } catch { }
                    try
                    {
                        NudgeIconSetRules(app, wbObj);
#if DEBUG
                        LogCfDiagnostics(wbObj, map.MasterTableName, "AFTER deferred refresh");
#endif
                    }
                    catch (Exception exDef)
                    {
                        AppLogger.Warn($"[WorkbookMapper] Deferred CF refresh warning: {exDef.Message}");
                    }
                    finally
                    {
                        try { app.EnableEvents = prevEvents; } catch { }
                    }
                });
            }
            catch (Exception exQ)
            {
                AppLogger.Warn($"[WorkbookMapper] Could not queue deferred CF refresh: {exQ.Message}");
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[WorkbookMapper] CF refresh warning: {ex.Message}");
        }
    }

    /// <summary>
    /// Rule-level (not cell-level) nudge: re-assigns each Icon Set rule's AppliesTo on every table
    /// body so Excel invalidates its cached icon state, then recalculates. Writes no cell values.
    /// </summary>
    private static void NudgeIconSetRules(dynamic xlApp, object wbObj)
    {
        dynamic wb = wbObj;
        foreach (dynamic sheet in wb.Worksheets)
        {
            try
            {
                foreach (dynamic tbl in sheet.ListObjects)
                {
                    try
                    {
                        dynamic? body = tbl.DataBodyRange;
                        if (body == null) continue;
                        try { body.Dirty(); } catch { }

                        dynamic fcs = body.FormatConditions;
                        int count = (int)fcs.Count;
                        for (int i = 1; i <= count; i++)
                        {
                            try
                            {
                                dynamic fc = fcs[i];
                                if ((int)fc.Type == 6) // xlIconSets
                                {
                                    fc.ModifyAppliesToRange(fc.AppliesTo);
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }
            }
            catch { }
        }

        try { xlApp.Calculate(); } catch { }
    }

#if DEBUG
    /// <summary>DEBUG: logs CF state for the first Expired row of the master table.</summary>
    private static void LogCfDiagnostics(object wbObj, string masterTableName, string label)
    {
        try
        {
            dynamic wb = wbObj;
            dynamic? tbl = null;
            dynamic? ws = null;
            foreach (dynamic sheet in wb.Worksheets)
            {
                try { tbl = sheet.ListObjects[masterTableName]; ws = sheet; if (tbl != null) break; } catch { }
            }
            if (tbl == null || ws == null) return;

            dynamic? body = tbl!.DataBodyRange;
            if (body == null) return;

            // Bulk reads: headers + whole body in 2 COM calls
            object[,] headers = TableMapper.GetRange2DValues(tbl.HeaderRowRange, 1, (int)tbl.ListColumns.Count);
            int rows = (int)body.Rows.Count;
            int cols = (int)body.Columns.Count;
            object[,] data = TableMapper.GetRange2DValues(body, rows, cols);

            int statusCol = -1;
            for (int c = 1; c <= cols; c++)
                if (string.Equals(headers[1, c]?.ToString(), "Status", StringComparison.OrdinalIgnoreCase)) { statusCol = c; break; }
            if (statusCol < 0) { AppLogger.Info($"[CF-DIAG {label}] No Status column"); return; }

            int expiredRow = -1;
            for (int r = 1; r <= rows; r++)
                if (string.Equals(data[r, statusCol]?.ToString(), "Expired", StringComparison.OrdinalIgnoreCase)) { expiredRow = r; break; }
            if (expiredRow < 0) { AppLogger.Info($"[CF-DIAG {label}] No Expired row in {masterTableName}"); return; }

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"[CF-DIAG {label}] {masterTableName} first Expired row = body row {expiredRow}");

            for (int c = 1; c <= cols; c++)
            {
                string h = headers[1, c]?.ToString() ?? "";
                if (!h.Contains("Expiry", StringComparison.OrdinalIgnoreCase)) continue;
                try
                {
                    dynamic cell = body.Cells[expiredRow, c];
                    string color = "n/a";
                    try { color = cell.DisplayFormat.Interior.Color.ToString(); } catch { }
                    sb.AppendLine($"  Cell {cell.Address[false, false]} '{h}': Value2={cell.Value2}, Text='{cell.Text}', Formula='{cell.Formula}', DisplayFormat.Interior.Color={color}");
                }
                catch (Exception exCell) { sb.AppendLine($"  Cell col {c} read failed: {exCell.Message}"); }
            }

            dynamic fcs = body.FormatConditions;
            int count = (int)fcs.Count;
            sb.AppendLine($"  FormatConditions on body: {count}");
            for (int i = 1; i <= count; i++)
            {
                try
                {
                    dynamic fc = fcs[i];
                    int type = (int)fc.Type;
                    string line = $"  [{i}] Type={type} Priority={fc.Priority} StopIfTrue={fc.StopIfTrue} AppliesTo={fc.AppliesTo.Address}";
                    if (type == 6)
                    {
                        try
                        {
                            dynamic crit = fc.IconCriteria;
                            int cc = (int)crit.Count;
                            var parts = new List<string>();
                            for (int k = 1; k <= cc; k++)
                            {
                                dynamic ic = crit[k];
                                parts.Add($"#{k}(Type={ic.Type},Op={ic.Operator},Value={ic.Value})");
                            }
                            line += " IconCriteria=" + string.Join(" ", parts);
                        }
                        catch (Exception exIc) { line += $" IconCriteria read failed: {exIc.Message}"; }
                    }
                    sb.AppendLine(line);
                }
                catch (Exception exFc) { sb.AppendLine($"  [{i}] read failed: {exFc.Message}"); }
            }

            AppLogger.Info(sb.ToString());
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[CF-DIAG {label}] failed: {ex.Message}");
        }
    }
#endif

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
        if (map.ApplyMasterCalculations != null)
        {
            try
            {
                dynamic calcApp = ExcelDnaUtil.Application;
                dynamic? activeWb = calcApp.ActiveWorkbook;
                if (activeWb != null) map.ApplyMasterCalculations((object)activeWb, records);
            }
            catch (Exception exCalc) { AppLogger.Warn($"[WorkbookMapper] ApplyMasterCalculations warning: {exCalc.Message}"); }
        }

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
            dynamic? masterWs = map.CustomMasterPopulator != null
                ? null
                : SheetMapperHelper.ResolveSheetSafely(wbObj, map.MasterSheetAliases, map.MasterSheetName);
            if (map.CustomMasterPopulator != null)
            {
                map.CustomMasterPopulator(wbObj, ordered);
            }
            else if (masterWs != null)
            {
                map.OnBeforeMasterPopulate?.Invoke((object)masterWs);
                TableMapper.PopulateTableSafe<TMaster>((object)masterWs, map.MasterTableName, ordered);
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
                            try { RuleEvaluator.EvaluateAndHealTable(sheet, tbl); }
                            catch (Exception exHeal) { AppLogger.Warn($"[WorkbookMapper] Heal failed for '{tableName}': {exHeal.Message}"); }

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
            // Lift calculation + screen freeze first, but keep EnableEvents off until the
            // repaint pass is done so it cannot raise Worksheet_Change loops.
            try { xlApp.Calculation = origCalculation; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            if (auditWorkbook != null)
            {
                RefreshConditionalFormatting(xlApp, auditWorkbook, map);
            }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }

        // 8. DEBUG audit (no-op in Release). Shown only after the freeze is lifted: a modal dialog
        //    must never open while ScreenUpdating / EnableEvents are suppressed.
        if (auditWorkbook != null)
        {
            SheetMapperHelper.ShowPopulateDiagnostics(auditWorkbook, map.AuditTitle, map.MasterLabel, ordered.Count, map.InspectTables);
        }
    }
}
