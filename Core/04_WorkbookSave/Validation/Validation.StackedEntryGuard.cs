namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text.Json.Serialization;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types.Attributes;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using TableSnapshot = (string WorkbookKey, string SheetName, string TableName, string DisplayLabel, int ColCount, string[] Headers, bool[] EditableCols, int IdColIndex, System.Collections.Generic.Dictionary<string, int> OriginalFingerprints, System.Collections.Generic.Dictionary<string, object?[]> OriginalRowsById, System.Collections.Generic.List<object?[]> OriginalRows);

/// <summary>
/// Failsafe for stacked / append-only history tables (licenses, maintenance, insurance, permits, road tax, ...).
/// Warns the user when they edit a row that was loaded from (or last saved to) the database: at edit time, and
/// before save as cell-level <see cref="ValidationIssue"/> warnings in the Validation Task Pane (with one-click
/// "Revert History Edits" / "Clone Edits as New Rows" actions). Existing rows should only be edited to fix mistakes.
/// <para>
/// Rows are matched by their auto-increment ID when the table has one; otherwise by a fingerprint of the
/// user-editable cells (locked / formula / <c>[JsonIgnore]</c> UI-only columns excluded), so sorting and adding new
/// rows never count as modifications.
/// </para>
/// </summary>
public static class StackedEntryGuard
{
    private const string Caption = "Stacked Entry Reminder";
    private const string FieldSeparator = "\u001F";

    // =========================================================================
    // 1. STACKED TABLE REGISTRY
    // =========================================================================

    private static readonly Dictionary<string, string> StackedTables = new(StringComparer.OrdinalIgnoreCase)
    {
        ["VehicleLicensesTable"] = "Vehicle License",
        ["SkillLicensesTable"] = "Skill License",
        ["MaintenancesTable"] = "Vehicle Maintenance",
        ["ServicesTable"] = "Vehicle Service",
        ["InsurancesTable"] = "Vehicle Insurance",
        ["InsuranceClaimsTable"] = "Insurance Claim",
        ["PermitsTable"] = "Vehicle Permit",
        ["RoadTaxesTable"] = "Road Tax",
        ["InspectionsTable"] = "Vehicle Inspection",
        ["OdometersTable"] = "Odometer Log",
        ["AttendanceEntriesTable"] = "Attendance Entry",
    };

    public static bool IsStackedTable(string? tableName) =>
        !string.IsNullOrWhiteSpace(tableName) && StackedTables.ContainsKey(tableName!);

    // =========================================================================
    // 2. SNAPSHOT STATE
    // =========================================================================

    private sealed record PendingSelection(
        string SnapshotKey,
        string SheetName,
        string TableName,
        string DisplayLabel,
        int ExcelRow,
        string Fingerprint,
        bool IsOriginalRow,
        object[,] RowValues,
        bool[] EditableCols,
        PropertyRuleDescriptor?[] ColRules);

    private const string SnapshotsKey = "SKSB_STACKED_ENTRY_SNAPSHOTS";
    private const string ApprovedKey = "SKSB_STACKED_ENTRY_APPROVED";

    // Snapshots / approvals live in AppDomain data (BCL types only) so every .xll - each loaded in its own
    // AssemblyLoadContext with its own copy of Core.dll - shares one state.
    private static Dictionary<string, TableSnapshot> SharedSnapshots()
    {
        lock (AppDomain.CurrentDomain)
        {
            if (AppDomain.CurrentDomain.GetData(SnapshotsKey) is Dictionary<string, TableSnapshot> existing) return existing;
            var created = new Dictionary<string, TableSnapshot>(StringComparer.OrdinalIgnoreCase);
            AppDomain.CurrentDomain.SetData(SnapshotsKey, created);
            return created;
        }
    }

    private static HashSet<string> SharedApprovedEdits()
    {
        lock (AppDomain.CurrentDomain)
        {
            if (AppDomain.CurrentDomain.GetData(ApprovedKey) is HashSet<string> existing) return existing;
            var created = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            AppDomain.CurrentDomain.SetData(ApprovedKey, created);
            return created;
        }
    }

    private static bool TryGetSnapshot(string key, out TableSnapshot snapshot)
    {
        var map = SharedSnapshots();
        lock (map) { return map.TryGetValue(key, out snapshot); }
    }

    private static PendingSelection? _lastSelected;
    private static volatile bool _editPromptPending;

    // =========================================================================
    // 3. SNAPSHOT CAPTURE
    // =========================================================================

    /// <summary>
    /// Records the original rows / editable-cell fingerprints of a stacked table. Call only after the table has been
    /// fully healed (synced UI helpers + state guards settled), otherwise the heal pass reads as a user edit.
    /// </summary>
    public static void CaptureTableSnapshot(dynamic worksheet, dynamic table)
    {
        try
        {
            if (worksheet == null || table == null) return;
            string tableName = table!.Name?.ToString() ?? string.Empty;
            if (!StackedTables.TryGetValue(tableName, out string? label)) return;

            string wbKey = GetWorkbookKey((object?)worksheet!.Parent);
            string sheetName = worksheet.Name?.ToString() ?? string.Empty;

            int colCount = (int)table.ListColumns.Count;
            int headerRow = (int)table.HeaderRowRange.Row;
            int startCol = (int)table.HeaderRowRange.Column;

            dynamic? body = table.DataBodyRange;
            int rowCount = 0;
            if (body != null) { try { rowCount = (int)body.Rows.Count; } catch { } }

            object[,] headerGrid = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, colCount);
            var headers = new string[colCount];
            for (int c = 0; c < colCount; c++)
                headers[c] = SafeGet(headerGrid, 1, c + 1)?.ToString()?.Trim() ?? string.Empty;
            PropertyRuleDescriptor?[] colRules = ResolveColRules(tableName, headers);

            // Formula detection from the first body row in one bulk read (insert row when the table is empty).
            string[] formulas = rowCount > 0
                ? ReadRowFormulas((object)body!.Rows[1], colCount)
                : ReadRowFormulas((object)worksheet.Range(worksheet.Cells[headerRow + 1, startCol], worksheet.Cells[headerRow + 1, startCol + colCount - 1]), colCount);

            var editableCols = new bool[colCount];
            int idColIndex = -1;
            for (int c = 0; c < colCount; c++)
            {
                var rule = colRules[c];
                bool locked = rule?.IsLocked == true;
                bool isUiOnly = rule?.Property?.IsDefined(typeof(JsonIgnoreAttribute), true) == true;
                editableCols[c] = !locked && !isUiOnly && !formulas[c].StartsWith("=", StringComparison.Ordinal);
                if (idColIndex < 0 && rule?.DefaultRule?.Kind == DefaultKind.AutoIncrement) idColIndex = c;
            }

            var fingerprints = new Dictionary<string, int>(StringComparer.Ordinal);
            var rowsById = new Dictionary<string, object?[]>(StringComparer.OrdinalIgnoreCase);
            var originalRows = new List<object?[]>();
            if (body != null && rowCount > 0)
            {
                object[,] grid = TableMapper.GetRange2DValues(body, rowCount, colCount);
                for (int r = 1; r <= rowCount; r++)
                {
                    object?[] row = RowFromGrid(grid, r, colCount);
                    string fp = Fingerprint(row, editableCols, colRules);
                    if (string.IsNullOrEmpty(fp)) continue;

                    fingerprints.TryGetValue(fp, out int n);
                    fingerprints[fp] = n + 1;
                    originalRows.Add(row);

                    if (idColIndex >= 0)
                    {
                        string idKey = NormalizeId(row[idColIndex]);
                        if (idKey.Length > 0) rowsById[idKey] = row;
                    }
                }
            }

            string key = SnapshotKey(wbKey, tableName);
            var snapshots = SharedSnapshots();
            lock (snapshots)
            {
                snapshots[key] = (wbKey, sheetName, tableName, label, colCount, headers, editableCols, idColIndex,
                    fingerprints, rowsById, originalRows);
            }
            var approved = SharedApprovedEdits();
            lock (approved) { approved.Remove(key); }
            var last = _lastSelected;
            if (last != null && string.Equals(last.SnapshotKey, key, StringComparison.OrdinalIgnoreCase))
                _lastSelected = null;
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StackedEntryGuard] CaptureTableSnapshot failed", ex);
        }
    }

    /// <summary>Re-captures every stacked table of <paramref name="wb"/> (call after populate / heal and after a successful save).</summary>
    public static void RefreshWorkbookSnapshots(dynamic? wb)
    {
        if (wb == null) return;
        try
        {
            foreach (dynamic ws in wb.Worksheets)
            {
                try
                {
                    foreach (dynamic tbl in ws.ListObjects)
                    {
                        string name = tbl.Name?.ToString() ?? string.Empty;
                        if (IsStackedTable(name)) CaptureTableSnapshot(ws, tbl);
                    }
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StackedEntryGuard] RefreshWorkbookSnapshots failed", ex);
        }
    }

    public static void ClearAll()
    {
        var snapshots = SharedSnapshots();
        lock (snapshots) { snapshots.Clear(); }
        var approved = SharedApprovedEdits();
        lock (approved) { approved.Clear(); }
        _lastSelected = null;
    }

    // =========================================================================
    // 4. PRE-EDIT SELECTION CAPTURE
    // =========================================================================

    /// <summary>Remembers the selected stacked-table row (values + fingerprint) before the user types into it.</summary>
    public static void OnSelectionChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (target == null || SheetOperations.IsTemplateEditMode) return;

            dynamic singleCell = target!.Cells[1, 1];
            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }
            if (table == null) { _lastSelected = null; return; }

            string tableName = table.Name?.ToString() ?? string.Empty;
            if (!IsStackedTable(tableName)) { _lastSelected = null; return; }

            string key = SnapshotKey(GetWorkbookKey((object?)sh.Parent), tableName);
            if (!TryGetSnapshot(key, out TableSnapshot snapshot) || snapshot.OriginalFingerprints.Count == 0) { _lastSelected = null; return; }

            dynamic? body = table.DataBodyRange;
            if (body == null) { _lastSelected = null; return; }
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp.Intersect(singleCell, body) == null) { _lastSelected = null; return; }

            int excelRow = (int)singleCell.Row;
            int dataRowIndex = excelRow - (int)body.Row + 1;
            if (dataRowIndex < 1 || dataRowIndex > (int)body.Rows.Count) { _lastSelected = null; return; }

            PropertyRuleDescriptor?[] colRules = ResolveColRules(snapshot.TableName, snapshot.Headers);
            object[,] rowValues = TableMapper.GetRange2DValues(table.ListRows[dataRowIndex].Range, 1, snapshot.ColCount);
            string fp = Fingerprint(RowFromGrid(rowValues, 1, snapshot.ColCount), snapshot.EditableCols, colRules);
            bool isOriginal = !string.IsNullOrEmpty(fp) && snapshot.OriginalFingerprints.ContainsKey(fp);

            string sheetName = sh?.Name?.ToString() ?? string.Empty;
            _lastSelected = new PendingSelection(key, sheetName, tableName, snapshot.DisplayLabel, excelRow, fp, isOriginal,
                rowValues, snapshot.EditableCols, colRules);
        }
        catch (Exception ex)
        {
            _lastSelected = null;
            AppLogger.Error("[StackedEntryGuard] OnSelectionChanged failed", ex);
        }
    }

    // =========================================================================
    // 5. ON-EDIT GUARD
    // =========================================================================

    /// <summary>Warns (once per table per load/save session) when an original stacked-table row is modified.</summary>
    public static void OnSheetChanged(dynamic sh, dynamic target)
    {
        try
        {
            if (target == null || SheetOperations.IsTemplateEditMode || _editPromptPending) return;

            PendingSelection? pending = _lastSelected;
            if (pending == null || !pending.IsOriginalRow) return;

            string sheetName = sh?.Name?.ToString() ?? string.Empty;
            if (!string.Equals(sheetName, pending.SheetName, StringComparison.OrdinalIgnoreCase)) return;
            if (!string.Equals(SnapshotKey(GetWorkbookKey((object?)sh!.Parent), pending.TableName), pending.SnapshotKey, StringComparison.OrdinalIgnoreCase)) return;

            var approvedSet = SharedApprovedEdits();
            lock (approvedSet)
            {
                if (approvedSet.Contains(pending.SnapshotKey)) return;
            }

            dynamic? table = null;
            try { table = sh.ListObjects[pending.TableName]; } catch { }
            if (table == null) return;
            dynamic? body = table.DataBodyRange;
            if (body == null) return;

            int bodyRow = (int)body.Row;
            int rowCount = (int)body.Rows.Count;
            int dataRowIndex = pending.ExcelRow - bodyRow + 1;
            if (dataRowIndex < 1 || dataRowIndex > rowCount) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic rowRange = table.ListRows[dataRowIndex].Range;
            if (xlApp.Intersect(target, rowRange) == null) return;

            int colCount = pending.EditableCols.Length;
            object[,] grid = TableMapper.GetRange2DValues(body, rowCount, colCount);
            string newFp = Fingerprint(RowFromGrid(grid, dataRowIndex, colCount), pending.EditableCols, pending.ColRules);
            if (string.Equals(newFp, pending.Fingerprint, StringComparison.Ordinal)) return;

            // If the original entry still exists as often as it was loaded (e.g. rows were shifted), nothing was overwritten.
            if (!TryGetSnapshot(pending.SnapshotKey, out TableSnapshot snapshot)) return;
            var current = new Dictionary<string, int>(StringComparer.Ordinal);
            for (int r = 1; r <= rowCount; r++)
            {
                string fp = Fingerprint(RowFromGrid(grid, r, colCount), pending.EditableCols, pending.ColRules);
                if (string.IsNullOrEmpty(fp)) continue;
                current.TryGetValue(fp, out int n);
                current[fp] = n + 1;
            }
            snapshot.OriginalFingerprints.TryGetValue(pending.Fingerprint, out int origCount);
            current.TryGetValue(pending.Fingerprint, out int currCount);
            if (currCount >= origCount) return;

            _editPromptPending = true;
            dynamic worksheet = sh;
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    string label = pending.DisplayLabel;
                    bool confirmed = Prompt(
                        $"You are editing an existing {label} entry.\n\n" +
                        "To preserve history, please add a new row for renewals or new events instead of overwriting a previous entry. " +
                        "Only edit existing entries to fix mistakes.\n\n" +
                        "Are you sure you want to edit this existing entry?");

                    if (confirmed)
                    {
                        lock (approvedSet) { approvedSet.Add(pending.SnapshotKey); }
                        AppLogger.Info($"[StackedEntryGuard] Edit of existing {label} entry (row {pending.ExcelRow}) confirmed.");
                        return;
                    }

                    RevertRow(worksheet, pending);
                    AppLogger.Info($"[StackedEntryGuard] Edit of existing {label} entry (row {pending.ExcelRow}) declined and reverted.");
                    SheetOperations.ShowStatusConfirmation($"Edit to existing {label} entry reverted. Add a new row for a new entry.");
                }
                catch (Exception ex)
                {
                    AppLogger.Error("[StackedEntryGuard] Edit prompt failed", ex);
                }
                finally
                {
                    _editPromptPending = false;
                }
            });
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StackedEntryGuard] OnSheetChanged failed", ex);
        }
    }

    private static void RevertRow(dynamic worksheet, PendingSelection pending)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic? table = null;
        try { table = worksheet.ListObjects[pending.TableName]; } catch { }
        if (table == null) return;

        bool origEvents = true;
        bool origUpdating = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
        try
        {
            int startCol = (int)table.HeaderRowRange.Column;
            for (int c = 0; c < pending.EditableCols.Length; c++)
            {
                if (!pending.EditableCols[c]) continue;
                try
                {
                    dynamic cell = worksheet.Cells[pending.ExcelRow, startCol + c];
                    string formula = cell.Formula?.ToString() ?? string.Empty;
                    if (formula.StartsWith("=", StringComparison.Ordinal)) continue;
                    cell.Value2 = SafeGet(pending.RowValues, 1, c + 1);
                }
                catch { }
            }
        }
        finally
        {
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
        }

        try { RuleEvaluator.EvaluateAndHealRow(worksheet, table, pending.ExcelRow); } catch { }
    }

    // =========================================================================
    // 6. PRE-SAVE CELL-LEVEL WARNINGS
    // =========================================================================

    /// <summary>
    /// Returns one <see cref="ValidationSeverity.Warning"/> per modified cell of an original (loaded / last saved)
    /// stacked-history row in <paramref name="wb"/>. Each issue carries <see cref="ValidationIssue.OriginalValue"/> and
    /// <see cref="ValidationIssue.RowKey"/> so the Validation Task Pane can revert or clone the edits.
    /// </summary>
    public static List<ValidationIssue> CollectStackedEditWarnings(dynamic? wb)
    {
        var issues = new List<ValidationIssue>();
        object? wbObj = wb;
        if (wbObj == null) return issues;
        try
        {
            string wbKey = GetWorkbookKey(wbObj);
            List<TableSnapshot> snapshots;
            var snapshotMap = SharedSnapshots();
            lock (snapshotMap)
            {
                snapshots = snapshotMap.Values
                    .Where(s => string.Equals(s.WorkbookKey, wbKey, StringComparison.OrdinalIgnoreCase) && s.OriginalRows.Count > 0)
                    .ToList();
            }

            foreach (var snapshot in snapshots)
            {
                try
                {
                    CollectTableWarnings(wbObj, snapshot, issues);
                }
                catch (Exception exTable)
                {
                    AppLogger.Warn($"[StackedEntryGuard] Stacked-edit scan skipped for '{snapshot.TableName}': {exTable.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("[StackedEntryGuard] CollectStackedEditWarnings failed", ex);
        }

        if (issues.Count > 0)
            AppLogger.Info($"[StackedEntryGuard] {issues.Count} modified history cell(s) flagged.");
        return issues;
    }

    private static void CollectTableWarnings(object wbObj, TableSnapshot snapshot, List<ValidationIssue> issues)
    {
        dynamic? table = ResolveTable(wbObj, snapshot.SheetName, snapshot.TableName);
        if (table == null) return;
        if ((int)table.ListColumns.Count != snapshot.ColCount) return; // structure changed since capture

        dynamic? body = table.DataBodyRange;
        int rowCount = 0;
        if (body != null) { try { rowCount = (int)body.Rows.Count; } catch { } }
        if (body == null || rowCount == 0) return;

        int bodyRow = (int)body.Row;
        int startCol = (int)table.HeaderRowRange.Column;
        int colCount = snapshot.ColCount;
        object[,] grid = TableMapper.GetRange2DValues(body, rowCount, colCount);
        PropertyRuleDescriptor?[] colRules = ResolveColRules(snapshot.TableName, snapshot.Headers);

        if (snapshot.IdColIndex >= 0)
        {
            for (int r = 1; r <= rowCount; r++)
            {
                object?[] curr = RowFromGrid(grid, r, colCount);
                string idKey = NormalizeId(curr[snapshot.IdColIndex]);
                if (idKey.Length == 0 || !snapshot.OriginalRowsById.TryGetValue(idKey, out object?[]? origRow)) continue;
                AddCellDiffs(issues, snapshot, colRules, origRow, curr, bodyRow + r - 1, startCol, idKey);
            }
            return;
        }

        // No ID column: compare fingerprint multisets and pair each unmatched current row with its best missing original.
        var remaining = new Dictionary<string, int>(snapshot.OriginalFingerprints, StringComparer.Ordinal);
        var unmatched = new List<(int Row, object?[] Values)>();
        for (int r = 1; r <= rowCount; r++)
        {
            object?[] curr = RowFromGrid(grid, r, colCount);
            string fp = Fingerprint(curr, snapshot.EditableCols, colRules);
            if (string.IsNullOrEmpty(fp)) continue;
            if (remaining.TryGetValue(fp, out int n) && n > 0) { remaining[fp] = n - 1; continue; }
            unmatched.Add((r, curr));
        }
        if (unmatched.Count == 0) return;

        var missing = new List<(string Fp, object?[] Values)>();
        foreach (var orig in snapshot.OriginalRows)
        {
            string fp = Fingerprint(orig, snapshot.EditableCols, colRules);
            if (remaining.TryGetValue(fp, out int n) && n > 0)
            {
                remaining[fp] = n - 1;
                missing.Add((fp, orig));
            }
        }
        if (missing.Count == 0) return;

        foreach (var (r, curr) in unmatched)
        {
            if (missing.Count == 0) break;
            int bestIdx = -1, bestScore = 0;
            for (int i = 0; i < missing.Count; i++)
            {
                int score = 0;
                for (int c = 0; c < colCount; c++)
                {
                    if (!snapshot.EditableCols[c]) continue;
                    if (string.Equals(Normalize(missing[i].Values[c], colRules[c]), Normalize(curr[c], colRules[c]), StringComparison.Ordinal)) score++;
                }
                if (score > bestScore) { bestScore = score; bestIdx = i; }
            }
            if (bestIdx < 0) continue; // nothing in common: a genuinely new row

            var (origFp, origRow) = missing[bestIdx];
            missing.RemoveAt(bestIdx);
            AddCellDiffs(issues, snapshot, colRules, origRow, curr, bodyRow + r - 1, startCol, origFp);
        }
    }

    private static void AddCellDiffs(
        List<ValidationIssue> issues,
        TableSnapshot snapshot,
        PropertyRuleDescriptor?[] colRules,
        object?[] origRow,
        object?[] currRow,
        int excelRow,
        int startCol,
        string rowKey)
    {
        for (int c = 0; c < snapshot.ColCount; c++)
        {
            if (!snapshot.EditableCols[c] || c == snapshot.IdColIndex) continue;
            var rule = colRules[c];
            object? origVal = origRow[c];
            object? currVal = currRow[c];
            if (string.Equals(Normalize(origVal, rule), Normalize(currVal, rule), StringComparison.Ordinal)) continue;

            issues.Add(new ValidationIssue(
                ValidationSeverity.Warning,
                snapshot.SheetName,
                snapshot.TableName,
                excelRow,
                startCol + c,
                snapshot.Headers[c],
                $"History entry modified: '{FormatForMessage(origVal, rule)}' -> '{FormatForMessage(currVal, rule)}' (add a new row for renewals/new events)",
                IsStackedHistoryEdit: true,
                OriginalValue: origVal,
                RowKey: rowKey));
        }
    }

    private static string FormatForMessage(object? val, PropertyRuleDescriptor? rule)
    {
        if (string.IsNullOrEmpty(Normalize(val, rule))) return "(blank)";
        if (rule?.IsDateColumn == true && RuleDataCache.TryParseDate(val, out DateOnly d))
            return d.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture);
        if (val is double or float or decimal or int or long or short)
            return Math.Round(Convert.ToDouble(val, CultureInfo.InvariantCulture), 6).ToString("G15", CultureInfo.InvariantCulture);
        return Convert.ToString(val, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
    }

    // =========================================================================
    // 7. TASK-PANE ACTIONS: REVERT / CLONE AS NEW ROWS (macro thread)
    // =========================================================================

    /// <summary>Writes every flagged cell's original value back, then heals the touched rows. Returns the reverted cell count.</summary>
    public static int RevertStackedEdits(dynamic? wb, IReadOnlyList<ValidationIssue> stackedIssues)
    {
        object? wbObj = wb;
        if (wbObj == null || stackedIssues == null) return 0;
        var groups = stackedIssues
            .Where(i => i.IsStackedHistoryEdit && i.ExcelRow > 0 && i.ExcelCol > 0)
            .GroupBy(i => (i.SheetName, i.TableName))
            .ToList();
        if (groups.Count == 0) return 0;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true, origEvents = true;
        object? origCalc = null;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }
        try { origCalc = (object)xlApp.Calculation; } catch { }

        int reverted = 0;
        try
        {
            try { xlApp.ScreenUpdating = false; } catch { }
            try { xlApp.EnableEvents = false; } catch { }
            try { xlApp.Calculation = -4135; } catch { }

            foreach (var group in groups)
            {
                try
                {
                    dynamic? table = ResolveTable(wbObj, group.Key.SheetName, group.Key.TableName);
                    if (table == null) continue;
                    dynamic ws = table.Parent;
                    TryGetSnapshot(SnapshotKey(GetWorkbookKey(wbObj), group.Key.TableName), out TableSnapshot snapshot);
                    Dictionary<string, int> idToRow = BuildIdRowMap(table, snapshot);
                    int startCol = (int)table.HeaderRowRange.Column;

                    var touchedRows = new SortedSet<int>();
                    foreach (var issue in group)
                    {
                        int excelRow = ResolveIssueRow(issue, snapshot, idToRow);
                        if (excelRow <= 0) continue;
                        try
                        {
                            dynamic cell = ws.Cells[excelRow, issue.ExcelCol];
                            string formula = cell.Formula?.ToString() ?? string.Empty;
                            if (formula.StartsWith("=", StringComparison.Ordinal)) continue;
                            cell.Value2 = issue.OriginalValue;
                            reverted++;
                            touchedRows.Add(excelRow);
                        }
                        catch (Exception exCell)
                        {
                            AppLogger.Warn($"[StackedEntryGuard] Revert of {group.Key.TableName} {issue.CellAddress} failed: {exCell.Message}");
                        }
                    }

                    foreach (int row in touchedRows)
                    {
                        try { RuleEvaluator.EvaluateAndHealRow(ws, table, row); } catch { }
                    }
                }
                catch (Exception exTable)
                {
                    AppLogger.Warn($"[StackedEntryGuard] Revert skipped for '{group.Key.TableName}': {exTable.Message}");
                }
            }
        }
        finally
        {
            try { if (origCalc != null) xlApp.Calculation = origCalc; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }

        AppLogger.Info($"[StackedEntryGuard] Reverted {reverted} modified history cell(s).");
        SheetOperations.ShowStatusConfirmation(reverted > 0
            ? $"Reverted {reverted} edited history cell(s) to their saved values."
            : "No history edits could be reverted.");
        return reverted;
    }

    /// <summary>
    /// For each modified history row: restores the original row in place and appends the user's edited values as a
    /// new row at the bottom of the table (with a fresh ID). Returns the cloned row count.
    /// </summary>
    public static int CloneStackedEditsAsNewRows(dynamic? wb, IReadOnlyList<ValidationIssue> stackedIssues)
    {
        object? wbObj = wb;
        if (wbObj == null || stackedIssues == null) return 0;
        var groups = stackedIssues
            .Where(i => i.IsStackedHistoryEdit && i.ExcelRow > 0)
            .GroupBy(i => (i.SheetName, i.TableName))
            .ToList();
        if (groups.Count == 0) return 0;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true, origEvents = true;
        object? origCalc = null;
        try { origUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { origEvents = (bool)xlApp.EnableEvents; } catch { }
        try { origCalc = (object)xlApp.Calculation; } catch { }

        int cloned = 0;
        try
        {
            try { xlApp.ScreenUpdating = false; } catch { }
            try { xlApp.EnableEvents = false; } catch { }
            try { xlApp.Calculation = -4135; } catch { }

            foreach (var group in groups)
            {
                try
                {
                    cloned += CloneTableEdits(wbObj, group.Key.SheetName, group.Key.TableName, group.ToList());
                }
                catch (Exception exTable)
                {
                    AppLogger.Warn($"[StackedEntryGuard] Clone skipped for '{group.Key.TableName}': {exTable.Message}");
                }
            }
        }
        finally
        {
            try { if (origCalc != null) xlApp.Calculation = origCalc; } catch { }
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }

        AppLogger.Info($"[StackedEntryGuard] Cloned {cloned} modified history row(s) as new rows.");
        SheetOperations.ShowStatusConfirmation(cloned > 0
            ? $"Restored {cloned} history row(s) and added the edits as new row(s) at the bottom of the table."
            : "No history edits could be cloned as new rows.");
        return cloned;
    }

    private static int CloneTableEdits(object wbObj, string sheetName, string tableName, List<ValidationIssue> issues)
    {
        dynamic? table = ResolveTable(wbObj, sheetName, tableName);
        if (table == null) return 0;
        if (!TryGetSnapshot(SnapshotKey(GetWorkbookKey(wbObj), tableName), out TableSnapshot snapshot)) return 0;
        if ((int)table.ListColumns.Count != snapshot.ColCount) return 0;

        dynamic ws = table.Parent;
        int colCount = snapshot.ColCount;
        PropertyRuleDescriptor?[] colRules = ResolveColRules(snapshot.TableName, snapshot.Headers);

        bool origTotals = false;
        try { origTotals = (bool)table.ShowTotals; if (origTotals) table.ShowTotals = false; } catch { }

        int cloned = 0;
        try
        {
            dynamic? body = table.DataBodyRange;
            int rowCount = 0;
            if (body != null) { try { rowCount = (int)body.Rows.Count; } catch { } }
            if (body == null || rowCount == 0) return 0;
            int bodyRow = (int)body.Row;

            // One bulk read: max ID + ID -> Excel row map (rows may have been sorted since validation).
            object[,] grid = TableMapper.GetRange2DValues(body, rowCount, colCount);
            long maxId = 0;
            var idToRow = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            if (snapshot.IdColIndex >= 0)
            {
                for (int r = 1; r <= rowCount; r++)
                {
                    string idKey = NormalizeId(SafeGet(grid, r, snapshot.IdColIndex + 1));
                    if (idKey.Length == 0) continue;
                    idToRow[idKey] = bodyRow + r - 1;
                    if (long.TryParse(idKey, NumberStyles.Integer, CultureInfo.InvariantCulture, out long id) && id > maxId) maxId = id;
                }
            }

            var targets = issues
                .Select(i => (Row: ResolveIssueRow(i, snapshot, idToRow), Key: i.RowKey ?? string.Empty))
                .Where(t => t.Row > 0)
                .Distinct()
                .ToList();

            foreach (var (excelRow, rowKey) in targets)
            {
                int dataIdx = excelRow - bodyRow + 1;
                if (dataIdx < 1 || dataIdx > rowCount) continue;

                // 1. User's edited values (bulk row read).
                dynamic rowRange = table.ListRows[dataIdx].Range;
                object?[] editedValues = RowFromGrid(TableMapper.GetRange2DValues(rowRange, 1, colCount), 1, colCount);
                string[] rowFormulas = ReadRowFormulas((object)rowRange, colCount);

                // 2. Untouched original row.
                object?[]? origRow = null;
                if (snapshot.IdColIndex >= 0 && rowKey.Length > 0)
                    snapshot.OriginalRowsById.TryGetValue(rowKey, out origRow);
                else if (rowKey.Length > 0)
                    origRow = snapshot.OriginalRows.FirstOrDefault(o => string.Equals(Fingerprint(o, snapshot.EditableCols, colRules), rowKey, StringComparison.Ordinal));
                if (origRow == null) continue;

                // Already identical to the original (e.g. a repeated click): nothing to clone.
                if (string.Equals(Fingerprint(editedValues, snapshot.EditableCols, colRules), Fingerprint(origRow, snapshot.EditableCols, colRules), StringComparison.Ordinal))
                    continue;

                // 3. Restore the historical row in place.
                int startCol = (int)table.HeaderRowRange.Column;
                for (int c = 0; c < colCount; c++)
                {
                    if (!snapshot.EditableCols[c] || rowFormulas[c].StartsWith("=", StringComparison.Ordinal)) continue;
                    if (string.Equals(Normalize(origRow[c], colRules[c]), Normalize(editedValues[c], colRules[c]), StringComparison.Ordinal)) continue;
                    try { ws.Cells[excelRow, startCol + c].Value2 = origRow[c]; } catch { }
                }
                try { RuleEvaluator.EvaluateAndHealRow(ws, table, excelRow); } catch { }

                // 4. Append a new row at the bottom of the ListObject.
                dynamic newListRow = table.ListRows.Add();
                dynamic newRange = newListRow.Range;
                int newExcelRow = (int)newRange.Row;
                string[] newFormulas = ReadRowFormulas((object)newRange, colCount);

                // 5. Write the edited values (fresh ID for the auto-increment column).
                for (int c = 0; c < colCount; c++)
                {
                    if (newFormulas[c].StartsWith("=", StringComparison.Ordinal)) continue;
                    object? value;
                    if (c == snapshot.IdColIndex) value = (double)(++maxId);
                    else if (colRules[c]?.IsLocked == true) continue;
                    else value = editedValues[c];
                    if (value == null) continue;
                    try { ws.Cells[newExcelRow, startCol + c].Value2 = value; } catch { }
                }

                // 6. Settle synced columns, defaults and formats on the new row.
                try { RuleEvaluator.EvaluateAndHealRow(ws, table, newExcelRow); } catch { }
                cloned++;
            }
        }
        finally
        {
            if (origTotals) { try { table.ShowTotals = true; } catch { } }
        }
        return cloned;
    }

    private static Dictionary<string, int> BuildIdRowMap(dynamic table, TableSnapshot snapshot)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        if (snapshot.Headers == null || snapshot.IdColIndex < 0) return map;
        try
        {
            dynamic? body = table.DataBodyRange;
            if (body == null) return map;
            int rowCount = (int)body.Rows.Count;
            int bodyRow = (int)body.Row;
            object[,] grid = TableMapper.GetRange2DValues(body, rowCount, snapshot.ColCount);
            for (int r = 1; r <= rowCount; r++)
            {
                string idKey = NormalizeId(SafeGet(grid, r, snapshot.IdColIndex + 1));
                if (idKey.Length > 0) map[idKey] = bodyRow + r - 1;
            }
        }
        catch { }
        return map;
    }

    /// <summary>Current Excel row of an issue: re-located by ID for ID tables (survives sorting), else the recorded row.</summary>
    private static int ResolveIssueRow(ValidationIssue issue, TableSnapshot snapshot, Dictionary<string, int> idToRow)
    {
        if (snapshot.Headers != null && snapshot.IdColIndex >= 0 && !string.IsNullOrEmpty(issue.RowKey))
            return idToRow.TryGetValue(issue.RowKey!, out int row) ? row : 0;
        return issue.ExcelRow;
    }

    private static dynamic? ResolveTable(object wbObj, string sheetName, string tableName)
    {
        dynamic wb = wbObj;
        try
        {
            dynamic ws = wb.Worksheets[sheetName];
            dynamic tbl = ws.ListObjects[tableName];
            if (tbl != null) return tbl;
        }
        catch { }

        try
        {
            foreach (dynamic ws in wb.Worksheets)
            {
                try
                {
                    dynamic tbl = ws.ListObjects[tableName];
                    if (tbl != null) return tbl;
                }
                catch { }
            }
        }
        catch { }
        return null;
    }

    // =========================================================================
    // 8. HELPERS
    // =========================================================================

    private static string SnapshotKey(string wbKey, string tableName) => $"{wbKey}|{tableName}";

    private static string GetWorkbookKey(object? wbObj)
    {
        if (wbObj == null) return string.Empty;
        dynamic wb = wbObj;
        try
        {
            string? full = wb.FullName?.ToString();
            if (!string.IsNullOrEmpty(full)) return full!;
        }
        catch { }
        try { return wb.Name?.ToString() ?? string.Empty; } catch { }
        return string.Empty;
    }

    private static PropertyRuleDescriptor?[] ResolveColRules(string tableName, string[] headers)
    {
        var rules = new PropertyRuleDescriptor?[headers.Length];
        for (int c = 0; c < headers.Length; c++)
        {
            if (string.IsNullOrWhiteSpace(headers[c])) continue;
            try { rules[c] = RuleRegistry.GetRule(tableName, headers[c]); } catch { }
        }
        return rules;
    }

    /// <summary>Bulk-reads the <c>Formula</c> of a single-row range (one COM call). Missing entries are empty strings.</summary>
    private static string[] ReadRowFormulas(object rowRangeObj, int colCount)
    {
        var result = new string[colCount];
        for (int c = 0; c < colCount; c++) result[c] = string.Empty;
        try
        {
            dynamic rowRange = rowRangeObj;
            object? f = rowRange.Formula;
            if (f is object[,] arr)
            {
                int r0 = arr.GetLowerBound(0), c0 = arr.GetLowerBound(1);
                int n = Math.Min(colCount, arr.GetUpperBound(1) - c0 + 1);
                for (int c = 0; c < n; c++) result[c] = Convert.ToString(arr[r0, c0 + c], CultureInfo.InvariantCulture) ?? string.Empty;
            }
            else if (colCount > 0)
            {
                result[0] = Convert.ToString(f, CultureInfo.InvariantCulture) ?? string.Empty;
            }
        }
        catch { }
        return result;
    }

    private static object? SafeGet(object[,] grid, int row, int col)
    {
        if (row < grid.GetLowerBound(0) || row > grid.GetUpperBound(0)) return null;
        if (col < grid.GetLowerBound(1) || col > grid.GetUpperBound(1)) return null;
        return grid[row, col];
    }

    /// <summary>0-based copy of one 1-based grid row.</summary>
    private static object?[] RowFromGrid(object[,] grid, int row, int colCount)
    {
        var values = new object?[colCount];
        for (int c = 0; c < colCount; c++) values[c] = SafeGet(grid, row, c + 1);
        return values;
    }

    /// <summary>Normalized positive integer ID ("" when blank / non-positive / not numeric).</summary>
    private static string NormalizeId(object? val)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        double d;
        if (val is double or float or decimal or int or long or short)
            d = Convert.ToDouble(val, CultureInfo.InvariantCulture);
        else if (!double.TryParse(Convert.ToString(val, CultureInfo.InvariantCulture)?.Trim(), NumberStyles.Float, CultureInfo.InvariantCulture, out d))
            return string.Empty;
        long id = (long)Math.Round(d);
        return id > 0 ? id.ToString(CultureInfo.InvariantCulture) : string.Empty;
    }

    /// <summary>Normalized fingerprint of the editable cells of one 0-based row; empty when every editable cell is blank.</summary>
    private static string Fingerprint(object?[] row, bool[] editableCols, PropertyRuleDescriptor?[] colRules)
    {
        var parts = new List<string>(editableCols.Length);
        bool anyValue = false;
        for (int c = 0; c < editableCols.Length; c++)
        {
            if (!editableCols[c]) continue;
            string norm = Normalize(c < row.Length ? row[c] : null, c < colRules.Length ? colRules[c] : null);
            if (norm.Length > 0) anyValue = true;
            parts.Add(norm);
        }
        return anyValue ? string.Join(FieldSeparator, parts) : string.Empty;
    }

    /// <summary>Canonical comparison form: equivalent blanks / dates / numbers / enum spellings never false-diff.</summary>
    private static string Normalize(object? val, PropertyRuleDescriptor? rule = null)
    {
        if (val == null || SheetMapperHelper.IsComError(val)) return string.Empty;
        if (val is string raw)
        {
            string t = raw.Trim();
            if (t.Length == 0 || t == "-" || t == "—") return string.Empty;
        }

        if (rule?.IsDateColumn == true && RuleDataCache.TryParseDate(val, out DateOnly date))
        {
            return date.Year <= 1900 ? string.Empty : date.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        if (val is double or float or decimal or int or long or short)
        {
            double d = Math.Round(Convert.ToDouble(val, CultureInfo.InvariantCulture), 6);
            if (d == 0d) d = 0d; // collapses -0.0
            if (Math.Abs(d) < 1e-6 && rule != null &&
                (rule.EffectiveZeroBehavior == ZeroBehavior.Blank || rule.EffectiveZeroBehavior == ZeroBehavior.Dash))
                return string.Empty;
            return d.ToString("G15", CultureInfo.InvariantCulture);
        }

        string s = Convert.ToString(val, CultureInfo.InvariantCulture)?.Trim() ?? string.Empty;
        if (rule != null && (Nullable.GetUnderlyingType(rule.PropertyType) ?? rule.PropertyType).IsEnum)
        {
            s = s.Replace(' ', '_').Replace('-', '_').ToLowerInvariant();
            if (s == "none") return string.Empty;
        }
        return s;
    }

    private sealed class ExcelWindow : IWin32Window
    {
        public IntPtr Handle { get; }
        public ExcelWindow(IntPtr handle) => Handle = handle;
    }

    /// <summary>Yes / No warning dialog; No is the default button. Excel state is restored whatever happens.</summary>
    private static bool Prompt(string message)
    {
        dynamic? xlApp = null;
        bool origUpdating = true;
        bool origEvents = true;
        IntPtr owner = IntPtr.Zero;
        try
        {
            xlApp = ExcelDnaUtil.Application;
            try { Arrow.HideAll(); } catch { }
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = true; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = true; } catch { }
            try { owner = Arrow.GetExcelHwnd(xlApp); } catch { }

            using var _ = LoadingOverlay.Suspend();

            var result = owner != IntPtr.Zero
                ? MessageBox.Show(new ExcelWindow(owner), message, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2)
                : MessageBox.Show(message, Caption, MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
            return result == DialogResult.Yes;
        }
        finally
        {
            if (xlApp != null)
            {
                try { xlApp.EnableEvents = origEvents; } catch { }
                try { xlApp.ScreenUpdating = origUpdating; } catch { }
            }
        }
    }
}
