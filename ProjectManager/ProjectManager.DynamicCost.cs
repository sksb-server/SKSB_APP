namespace SKSB_App.ProjectManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;

internal static class DynamicCostToggle
{
    // =========================================================================
    // 1. CONSTANTS & STATE
    // =========================================================================
    private const string ActualProjectCostHeader = "Actual Project Cost";
    private const string ProfitMarginHeader = "Profit / Margin";
    private const string ProjectCodeHeader = "Project Code";
    private const string NullDash = "\u2014";

    private static readonly string[] GeneralComponents = { "Fuel Cost", "Labour Cost", "Claim Cost" };
    private static readonly string[] InternalComponents = { "Labour Cost", "Claim Cost" };

    private static readonly object StateLock = new();
    private static readonly Dictionary<string, HashSet<string>> SelectedByTable = new(StringComparer.OrdinalIgnoreCase);

    // =========================================================================
    // 2. LIFECYCLE REGISTRATION
    // =========================================================================
    internal static void Register()
    {
        lock (StateLock)
        {
            SelectedByTable[SheetMapper.GeneralTableName] = new HashSet<string>(GeneralComponents, StringComparer.OrdinalIgnoreCase);
            SelectedByTable[SheetMapper.InternalTableName] = new HashSet<string>(InternalComponents, StringComparer.OrdinalIgnoreCase);
        }

        RuleCalculator.ClearColumnExpressionOverrides(SheetMapper.GeneralTableName);
        RuleCalculator.ClearColumnExpressionOverrides(SheetMapper.InternalTableName);
        CustomDropdowns.Register(AppliesTo, Build, allowOnLockedCell: true);
    }

    internal static void Unregister()
    {
        CustomDropdowns.Clear();
        RuleCalculator.ClearColumnExpressionOverrides(SheetMapper.GeneralTableName);
        RuleCalculator.ClearColumnExpressionOverrides(SheetMapper.InternalTableName);
    }

    // =========================================================================
    // 3. CUSTOM DROPDOWN PROVIDER
    // =========================================================================
    private static string[]? GetComponentsForTable(string tableName)
    {
        if (string.Equals(tableName, SheetMapper.GeneralTableName, StringComparison.OrdinalIgnoreCase))
            return GeneralComponents;
        if (string.Equals(tableName, SheetMapper.InternalTableName, StringComparison.OrdinalIgnoreCase))
            return InternalComponents;
        return null;
    }

    private static bool AppliesTo(dynamic singleCell)
    {
        if (SheetOperations.IsTemplateEditMode) return false;

        try
        {
            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }
            if (table == null) return false;

            string tableName = table.Name?.ToString() ?? string.Empty;
            if (GetComponentsForTable(tableName) == null) return false;

            if (!Commands.IsInsideTableDataBody(singleCell, table)) return false;

            dynamic ws = singleCell.Worksheet;
            int headerRow = (int)table.HeaderRowRange.Row;
            int col = (int)singleCell.Column;
            string header = ws.Cells[headerRow, col].Value2?.ToString()?.Trim() ?? string.Empty;
            if (!string.Equals(header, ActualProjectCostHeader, StringComparison.OrdinalIgnoreCase))
                return false;

            dynamic wb = ws.Parent;
            if (SheetOperations.IsMasterTemplateFile(wb)) return false;

            return SheetOperations.IsModuleWorkbook(wb, Events.ProjectSyncDescriptor);
        }
        catch
        {
            return false;
        }
    }

    private static CustomDropdownResult? Build(dynamic singleCell)
    {
        try
        {
            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }
            if (table == null) return null;

            string tableName = table.Name?.ToString() ?? string.Empty;
            string[]? components = GetComponentsForTable(tableName);
            if (components == null) return null;

            string[] activeSnapshot;
            lock (StateLock)
            {
                if (!SelectedByTable.TryGetValue(tableName, out var selected))
                {
                    selected = new HashSet<string>(components, StringComparer.OrdinalIgnoreCase);
                    SelectedByTable[tableName] = selected;
                }
                activeSnapshot = components.Where(selected.Contains).ToArray();
            }

            var options = new List<DropdownOption>(components.Length);
            for (int i = 0; i < components.Length; i++)
            {
                options.Add(new DropdownOption
                {
                    Display = components[i],
                    Raw = components[i],
                    Section = i,
                    IsValid = true
                });
            }

            return new CustomDropdownResult(
                options,
                (cell, opt) => OnToggle(cell, tableName, opt.Raw),
                IsMultiSelect: true,
                InitialChecked: activeSnapshot,
                AllowOnLockedCell: true);
        }
        catch (Exception ex)
        {
            AppLogger.Error("[DynamicCostToggle] Build failed", ex);
            return null;
        }
    }

    // =========================================================================
    // 4. TOGGLE & BULK TABLE RECALCULATION
    // =========================================================================
    private static void OnToggle(dynamic cell, string tableName, string component)
    {
        string[]? components = GetComponentsForTable(tableName);
        if (components == null || string.IsNullOrWhiteSpace(component)) return;

        List<string> active;
        lock (StateLock)
        {
            if (!SelectedByTable.TryGetValue(tableName, out var selected))
            {
                selected = new HashSet<string>(components, StringComparer.OrdinalIgnoreCase);
                SelectedByTable[tableName] = selected;
            }

            if (!selected.Remove(component)) selected.Add(component);
            active = components.Where(selected.Contains).ToList();
        }

        string? expr = active.Count == components.Length
            ? null
            : active.Count == 0
                ? "0"
                : string.Join(" + ", active.Select(c => $"[{c}]"));

        RuleCalculator.SetColumnExpressionOverride(tableName, ActualProjectCostHeader, expr);
        RecalculateTable(cell, tableName);

        string summary = active.Count == 0 ? "None (0)" : string.Join(" + ", active);
        SheetOperations.ShowStatusConfirmation($"Actual Project Cost tally ({tableName}): {summary}");
    }

    private static void RecalculateTable(dynamic cell, string tableName)
    {
        dynamic? xlApp = null;
        try { xlApp = ExcelDnaUtil.Application; } catch { }

        bool origEvents = true;
        bool origScreenUpdating = true;
        int origCalc = -4105;

        try { if (xlApp != null) origEvents = (bool)xlApp.EnableEvents; } catch { }
        try { if (xlApp != null) origScreenUpdating = (bool)xlApp.ScreenUpdating; } catch { }
        try { if (xlApp != null) origCalc = (int)xlApp.Calculation; } catch { }

        try
        {
            try { if (xlApp != null) xlApp.EnableEvents = false; } catch { }
            try { if (xlApp != null) xlApp.ScreenUpdating = false; } catch { }
            try { if (xlApp != null) xlApp.Calculation = -4135; } catch { }

            dynamic? table = null;
            try { table = cell.ListObject; } catch { }
            if (table == null) return;

            dynamic? body = null;
            try { body = table.DataBodyRange; } catch { }
            if (body == null) return;

            int rowCount = (int)body.Rows.Count;
            int colCount = (int)table.ListColumns.Count;
            if (rowCount <= 0 || colCount <= 0) return;

            object[,] hdrRaw = TableMapper.GetRange2DValues(table.HeaderRowRange, 1, colCount);
            int hdrRowLb = hdrRaw.GetLowerBound(0);
            int hdrColLb = hdrRaw.GetLowerBound(1);
            var colHeaders = new string[colCount + 1];
            for (int c = 1; c <= colCount; c++)
                colHeaders[c] = hdrRaw[hdrRowLb, hdrColLb + c - 1]?.ToString()?.Trim() ?? string.Empty;

            int costCol = FindColumnByHeader(colHeaders, ActualProjectCostHeader);
            if (costCol < 1) return;

            PropertyRuleDescriptor? costRule = RuleRegistry.GetRule(tableName, ActualProjectCostHeader);
            if (costRule == null) return;

            int profitCol = FindColumnByHeader(colHeaders, ProfitMarginHeader);
            PropertyRuleDescriptor? profitRule = profitCol >= 1 ? RuleRegistry.GetRule(tableName, ProfitMarginHeader) : null;
            int projectCodeCol = FindColumnByHeader(colHeaders, ProjectCodeHeader);

            object[,] bodyData = TableMapper.GetRange2DValues(body, rowCount, colCount);
            int rowLb = bodyData.GetLowerBound(0);
            int colLb = bodyData.GetLowerBound(1);

            var costOut = new object?[rowCount, 1];
            object?[,]? profitOut = profitCol >= 1 && profitRule != null ? new object?[rowCount, 1] : null;
            var rawObjects = new object?[colCount + 1];

            for (int r = 1; r <= rowCount; r++)
            {
                int srcR = rowLb + r - 1;
                for (int c = 1; c <= colCount; c++)
                    rawObjects[c] = bodyData[srcR, colLb + c - 1];

                bool isPopulated = (projectCodeCol >= 1 && HasNonBlankText(rawObjects[projectCodeCol]))
                    || HasAnyUserData(rawObjects, colCount, costCol, profitCol);

                if (!isPopulated)
                {
                    costOut[r - 1, 0] = rawObjects[costCol];
                    if (profitOut != null) profitOut[r - 1, 0] = rawObjects[profitCol];
                    continue;
                }

                decimal? costVal = RuleCalculator.EvaluateColumnRule(costRule, colHeaders, rawObjects, _ => 0m);
                object costCellVal = costVal.HasValue ? (double)costVal.Value : NullDash;
                rawObjects[costCol] = costCellVal;
                costOut[r - 1, 0] = costCellVal;

                if (profitOut != null && profitRule != null)
                {
                    decimal? profitVal = RuleCalculator.EvaluateColumnRule(profitRule, colHeaders, rawObjects, _ => 0m);
                    profitOut[r - 1, 0] = profitVal.HasValue ? (double)profitVal.Value : NullDash;
                }
            }

            table.ListColumns[costCol].DataBodyRange.Value2 = costOut;
            if (profitOut != null)
                table.ListColumns[profitCol].DataBodyRange.Value2 = profitOut;
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[DynamicCostToggle] RecalculateTable({tableName}) failed", ex);
        }
        finally
        {
            try { if (xlApp != null) xlApp.Calculation = origCalc; } catch { }
            try { if (xlApp != null) xlApp.ScreenUpdating = origScreenUpdating; } catch { }
            try { if (xlApp != null) xlApp.EnableEvents = origEvents; } catch { }
        }
    }

    private static int FindColumnByHeader(string[] headers, string target)
    {
        string normTarget = NormalizeHeader(target);
        for (int c = 1; c < headers.Length; c++)
        {
            if (string.Equals(headers[c], target, StringComparison.OrdinalIgnoreCase) ||
                string.Equals(NormalizeHeader(headers[c]), normTarget, StringComparison.OrdinalIgnoreCase))
                return c;
        }
        return -1;
    }

    private static string NormalizeHeader(string text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        Span<char> buf = stackalloc char[text.Length];
        int len = 0;
        foreach (char ch in text)
        {
            if (char.IsLetterOrDigit(ch)) buf[len++] = char.ToUpperInvariant(ch);
        }
        return new string(buf[..len]);
    }

    private static bool HasNonBlankText(object? value)
    {
        if (value == null) return false;
        string s = value.ToString()?.Trim() ?? string.Empty;
        return s.Length > 0 && s != NullDash;
    }

    private static bool HasAnyUserData(object?[] rawObjects, int colCount, int costCol, int profitCol)
    {
        for (int c = 1; c <= colCount; c++)
        {
            if (c == costCol || c == profitCol) continue;
            if (HasNonBlankText(rawObjects[c])) return true;
        }
        return false;
    }
}
