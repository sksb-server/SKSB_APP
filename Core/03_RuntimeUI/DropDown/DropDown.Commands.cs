namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.RuntimeUI.RuleEngine;

public static class Commands
{
    public const string DropDownStyleName = "DropDown";

    public static void ApplyDropdownStyle()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic selection = xlApp.Selection;
        if (selection == null) return;

        EnsureNamedStyleExists(xlApp, DropDownStyleName);
        selection.Style = DropDownStyleName;

        try
        {
            selection.Validation.Delete();
        }
        catch { }
    }

    public static void InspectAndConfigure(dynamic cell)
    {
        try
        {
            using var form = new ConfiguratorForm(cell);
            form.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed opening DropDown configurator: {ex.Message}", "Dropdown Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public static bool IsDropdownCell(dynamic cell)
    {
        if (cell == null) return false;

        try
        {
            dynamic singleCell = cell.Cells[1, 1];

            // 1. Primary path: Check if cell belongs to a ListObject table governed by C# RuleRegistry
            dynamic? cellTable = null;
            try { cellTable = singleCell.ListObject; } catch { }

            if (cellTable != null)
            {
                if (!IsInsideTableDataBody(singleCell, cellTable))
                    return false;

                var rule = RuleRegistry.GetRuleForCell(singleCell);
                if (rule != null && rule.HasDropdown)
                {
                    return true;
                }
            }

            // 2. Fallback path: Legacy style + cell config check (for non-mapped tables or top parameter cells)
            string currentStyleName = singleCell.Style?.Name?.ToString() ?? string.Empty;
            if (!string.Equals(currentStyleName, DropDownStyleName, StringComparison.OrdinalIgnoreCase))
            {
                return false;
            }

            var cfg = Metadata.ReadConfig(singleCell);
            return !cfg.IsEmpty;
        }
        catch
        {
            return false;
        }
    }

    public static bool IsInsideTableDataBody(dynamic singleCell, dynamic cellTable)
    {
        try
        {
            int row = (int)singleCell.Row;
            int headerRow = (int)cellTable.HeaderRowRange.Row;
            if (row <= headerRow) return false;

            dynamic? bodyRange = cellTable.DataBodyRange;
            if (bodyRange == null)
            {
                return row == headerRow + 1;
            }

            int bodyStartRow = (int)bodyRange.Row;
            int bodyEndRow = bodyStartRow + (int)bodyRange.Rows.Count - 1;
            return row >= bodyStartRow && row <= bodyEndRow;
        }
        catch
        {
            return false;
        }
    }

    public static bool IsInsideSameTableDataBody(dynamic cell)
    {
        try
        {
            dynamic? cellTable = null;
            try { cellTable = cell.ListObject; } catch { }
            if (cellTable == null) return false;

            return IsInsideTableDataBody(cell, cellTable);
        }
        catch
        {
            return false;
        }
    }

    public static string ResolveContextListFormula(dynamic targetCell)
    {
        try
        {
            var rule = RuleRegistry.GetRuleForCell(targetCell);
            if (rule != null && rule.Sources.Length > 0)
            {
                return "=" + string.Join(", ", rule.Sources);
            }

            dynamic sheet = targetCell.Worksheet;
            dynamic wb = sheet.Parent;
            int col = (int)targetCell.Column;

            dynamic? listObj = null;
            try { listObj = targetCell.ListObject; } catch { }
            if (listObj != null)
            {
                string tableName = listObj.Name?.ToString() ?? string.Empty;
                int headerRow = (int)listObj.HeaderRowRange.Row;
                string colHeader = sheet.Cells[headerRow, col].Value2?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(tableName) && !string.IsNullOrEmpty(colHeader))
                {
                    return $"={tableName}[{colHeader}]";
                }
            }

            string header = sheet.Cells[1, col].Value2?.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(header))
            {
                string cleanHeader = Regex.Replace(header, @"[^\w]", "");
                string candidateListName = cleanHeader.EndsWith("List", StringComparison.OrdinalIgnoreCase)
                    ? cleanHeader
                    : $"{cleanHeader}List";

                try
                {
                    dynamic match = wb.Names.Item(candidateListName);
                    if (match != null) return $"={candidateListName}";
                }
                catch { }
            }

            return string.Empty;
        }
        catch { return string.Empty; }
    }

    public static void EnsureNamedStyleExists(dynamic xlApp, string styleName)
    {
        try
        {
            dynamic styles = xlApp.ActiveWorkbook.Styles;
            foreach (dynamic s in styles)
            {
                if (string.Equals(s.Name?.ToString(), styleName, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            dynamic newStyle = styles.Add(styleName);
            newStyle.IncludeNumber = false;
            newStyle.IncludeFont = true;
            newStyle.IncludeAlignment = true;
            newStyle.IncludeBorder = true;
            newStyle.IncludePatterns = true;

            newStyle.Font.Name = "Segoe UI";
            newStyle.Font.Size = 10;
            newStyle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 244, 250));
        }
        catch { }
    }

    private static Config? _copiedConfig;

    public static void CopySelectedCellMetadata()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic activeCell = xlApp.ActiveCell;
            if (activeCell == null) return;

            var rule = RuleRegistry.GetRuleForCell(activeCell);
            if (rule != null)
            {
                MessageBox.Show(
                    $"Column '{rule.HeaderName}' on table '{rule.TableName}' is managed automatically by C# [ColumnRule] attributes.\n\n" +
                    $"  Sources: {(rule.Sources.Length > 0 ? string.Join(", ", rule.Sources) : "(None / Default Only)")}\n" +
                    $"  Default: {rule.DefaultRule?.Kind.ToString() ?? "None"}\n" +
                    $"  Syncs: {rule.SyncRules.Count}\n" +
                    $"  State Guards: {rule.StateGuards.Count}",
                    "C# Attribute Rule Active",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                return;
            }

            var cfg = Metadata.ReadConfig(activeCell);
            if (cfg.IsEmpty)
            {
                MessageBox.Show("Selected cell has no custom dropdown configuration to copy.",
                                "Copy Dropdown Metadata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            _copiedConfig = cfg;
            MessageBox.Show($"Copied dropdown configuration from {activeCell.Address}!",
                            "Metadata Copied", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to copy metadata: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public static void PasteSelectedCellMetadata()
    {
        if (_copiedConfig == null || _copiedConfig.IsEmpty)
        {
            MessageBox.Show("No legacy dropdown configuration has been copied yet.",
                            "Paste Dropdown Metadata", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic selection = xlApp.Selection;
            if (selection == null) return;

            EnsureNamedStyleExists(xlApp, DropDownStyleName);

            foreach (dynamic cell in selection.Cells)
            {
                cell.Style = DropDownStyleName;
                Metadata.WriteConfig(cell, _copiedConfig);
                try { cell.Validation.Delete(); } catch { }
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to paste metadata: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}