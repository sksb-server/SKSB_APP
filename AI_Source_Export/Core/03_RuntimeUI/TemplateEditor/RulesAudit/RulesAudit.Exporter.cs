namespace SKSB_App.Core.RuntimeUI.TemplateEditor.RulesAudit;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;

public static class RulesAuditExporter
{
    [ExcelCommand(Name = "ExportTemplateRulesToMarkdown")]
    public static void ExportAuditReport()
    {
        try
        {
            RuleRegistry.EnsureInitialized();
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("No active workbook found.", "Export Rules Audit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string rawWbName = wb.Name?.ToString() ?? "Template";
            string baseName = Path.GetFileNameWithoutExtension(rawWbName);
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                baseName = baseName.Replace(c, '_');
            }

            string targetFileName = $"{baseName}_Rules.md";
            var sb = new StringBuilder();

            sb.AppendLine($"# {baseName} Rules Audit");
            sb.AppendLine($"**Workbook:** `{rawWbName}`");
            sb.AppendLine($"**Generated:** {DateTime.Now:dd-MM-yyyy HH:mm:ss}");
            sb.AppendLine();

            int tableCount = 0;
            int sheetHeaderCellCount = 0;

            foreach (dynamic ws in wb.Sheets)
            {
                string sheetName = ws.Name?.ToString() ?? "Sheet";

                if (sheetName.Equals("Lookups", StringComparison.OrdinalIgnoreCase) ||
                    sheetName.Equals("Data", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                int minTableStartRow = int.MaxValue;
                if (ws.ListObjects.Count > 0)
                {
                    foreach (dynamic t in ws.ListObjects)
                    {
                        int hRow = (int)t.HeaderRowRange.Row;
                        if (hRow < minTableStartRow) minTableStartRow = hRow;
                    }
                }

                // 1. Scan Top Parameter Cells
                var topCells = ScanTopAnchorCells(ws, minTableStartRow);
                if (topCells.Count > 0)
                {
                    sb.AppendLine("---");
                    sb.AppendLine($"## SHEET_HEADER_CELLS: Sheet=`{sheetName}`");
                    sb.AppendLine();
                    sb.AppendLine("| CellAddress | Row | Col | Style | NumberFormat | Formula | SourceList / NativeFormula | NativeType | NativeAlert | NativeIgnoreBlank |");
                    sb.AppendLine("| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :---: | :---: |");

                    foreach (var tc in topCells)
                    {
                        sheetHeaderCellCount++;
                        sb.AppendLine(
                            $"| {tc.Address} " +
                            $"| {tc.Row} " +
                            $"| {tc.Col} " +
                            $"| {tc.Style} " +
                            $"| {RulesAuditHelper.EscapePipe(tc.NumberFormat)} " +
                            $"| {RulesAuditHelper.EscapePipe(tc.Formula)} " +
                            $"| {RulesAuditHelper.EscapePipe(tc.SourceOrFormula)} " +
                            $"| {tc.NativeType} " +
                            $"| {tc.NativeAlert} " +
                            $"| {(tc.NativeIgnoreBlank ? "TRUE" : "FALSE")} |"
                        );
                    }
                    sb.AppendLine();
                }

                if (ws.ListObjects.Count == 0) continue;

                // 2. Scan Tables
                foreach (dynamic table in ws.ListObjects)
                {
                    tableCount++;
                    string tableName = table.Name?.ToString() ?? "Table";
                    dynamic headerRange = table.HeaderRowRange;
                    int headerRow = (int)headerRange.Row;
                    int startCol = (int)headerRange.Column;
                    int colCount = (int)table.ListColumns.Count;
                    int archetypeRow = headerRow + 1;

                    sb.AppendLine("---");
                    sb.AppendLine($"## TABLE: Sheet=`{sheetName}` | Table=`{tableName}` | HeaderRow={headerRow} | ArchetypeRow={archetypeRow} | StartCol={startCol} | Columns={colCount}");
                    sb.AppendLine();
                    sb.AppendLine("### Column Rules");
                    sb.AppendLine("| Col | Row | Address | Header | Style | NumberFormat | EngineType | InCellFormula | Sources | Syncs | Transform | DefaultRule | StateGuards | IsMulti | IsWarning | NativeType | NativeFormula | NativeAlert | NativeIgnoreBlank |");
                    sb.AppendLine("| :---: | :---: | :---: | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :--- | :---: | :---: | :--- | :--- | :---: | :---: |");

                    for (int c = 1; c <= colCount; c++)
                    {
                        int absCol = startCol + c - 1;
                        dynamic headerCell = ws.Cells[headerRow, absCol];
                        dynamic archetypeCell = ws.Cells[archetypeRow, absCol];

                        string headerText = headerCell.Value2?.ToString()?.Trim() ?? $"Col_{c}";
                        string colLetter = RulesAuditHelper.GetExcelColumnLetter(absCol);
                        string cellAddress = $"{colLetter}{archetypeRow}";
                        string styleName = RulesAuditHelper.GetStyleName(archetypeCell);
                        string numFormat = RulesAuditHelper.GetCellNumberFormat(archetypeCell);

                        string cellFormula = archetypeCell.Formula2?.ToString() ?? string.Empty;
                        if (string.IsNullOrWhiteSpace(cellFormula))
                            cellFormula = archetypeCell.Formula?.ToString() ?? string.Empty;

                        // Identify C# Attribute-Driven Rule
                        var csharpRule = RuleRegistry.GetRule(tableName, headerText);
                        bool hasCSharp = csharpRule != null;

                        var cfg = Metadata.ReadConfig(archetypeCell);
                        bool isCustomDropdown = !cfg.IsEmpty;

                        bool hasNativeValidation = false;
                        string nativeType = string.Empty;
                        string nativeFormula = string.Empty;
                        string nativeAlert = string.Empty;
                        bool nativeIgnoreBlank = true;

                        try
                        {
                            dynamic val = archetypeCell.Validation;
                            if (val != null)
                            {
                                int vt = (int)val.Type;
                                hasNativeValidation = true;
                                nativeType = RulesAuditHelper.GetValidationTypeName(vt);
                                nativeFormula = val.Formula1?.ToString() ?? string.Empty;
                                nativeAlert = ((int)val.AlertStyle) switch { 1 => "Stop", 2 => "Warning", _ => "Info" };
                                nativeIgnoreBlank = (bool)val.IgnoreBlank;
                            }
                        }
                        catch { }

                        string engineType = hasCSharp ? "CSharpAttribute"
                                          : isCustomDropdown ? "LegacyDropdown"
                                          : hasNativeValidation ? "NativeValidation"
                                          : (!string.IsNullOrWhiteSpace(cellFormula) && cellFormula.StartsWith("=")) ? "Formula"
                                          : "StandardInput";

                        // Populate payload either from C# Attributes or Legacy Metadata
                        string sources = hasCSharp ? string.Join(", ", csharpRule!.Sources) : cfg.Source;
                        string syncs = hasCSharp ? string.Join(" ; ", csharpRule!.SyncRules.Select(s => $"{s.TargetHeader}<-{s.LookupSource}")) : cfg.SyncLeft;
                        string transform = hasCSharp ? csharpRule!.Transform.ToString() : cfg.Transform;
                        string defaultRule = hasCSharp && csharpRule!.DefaultRule != null ? csharpRule.DefaultRule.Kind.ToString() : cfg.Filter;
                        string guards = hasCSharp && csharpRule!.StateGuards.Count > 0 ? $"{csharpRule.StateGuards.Count} Guard(s)" : "";
                        bool isMulti = hasCSharp ? csharpRule!.IsMultiSelect : cfg.IsMultiSelect;
                        bool isWarning = hasCSharp ? csharpRule!.IsWarningOnly : cfg.IsWarningOnly;

                        sb.AppendLine(
                            $"| {colLetter} " +
                            $"| {archetypeRow} " +
                            $"| {cellAddress} " +
                            $"| {RulesAuditHelper.EscapePipe(headerText)} " +
                            $"| {styleName} " +
                            $"| {RulesAuditHelper.EscapePipe(numFormat)} " +
                            $"| {engineType} " +
                            $"| {RulesAuditHelper.EscapePipe(cellFormula)} " +
                            $"| {RulesAuditHelper.EscapePipe(sources)} " +
                            $"| {RulesAuditHelper.EscapePipe(syncs)} " +
                            $"| {transform} " +
                            $"| {RulesAuditHelper.EscapePipe(defaultRule)} " +
                            $"| {guards} " +
                            $"| {(isMulti ? "TRUE" : "FALSE")} " +
                            $"| {(isWarning ? "TRUE" : "FALSE")} " +
                            $"| {nativeType} " +
                            $"| {RulesAuditHelper.EscapePipe(nativeFormula)} " +
                            $"| {nativeAlert} " +
                            $"| {(nativeIgnoreBlank ? "TRUE" : "FALSE")} |"
                        );
                    }
                    sb.AppendLine();

                    // 3. Scan Conditional Formatting
                    try
                    {
                        dynamic bodyRange = table.DataBodyRange;
                        dynamic targetRange = bodyRange ?? ws.Range[ws.Cells[archetypeRow, startCol], ws.Cells[archetypeRow + 500, startCol + colCount - 1]];
                        dynamic formatConditions = targetRange.FormatConditions;
                        int cfCount = (int)formatConditions.Count;

                        if (cfCount > 0)
                        {
                            sb.AppendLine("### Conditional Formatting");
                            sb.AppendLine("| AppliesTo | RuleType | Operator | Formula1 | Formula2 | Severity | StopIfTrue |");
                            sb.AppendLine("| :--- | :--- | :--- | :--- | :--- | :---: | :---: |");

                            for (int i = 1; i <= cfCount; i++)
                            {
                                dynamic fc = formatConditions[i];
                                int cfType = (int)fc.Type;
                                string ruleType = cfType == 2 ? "Expression" : cfType == 1 ? "CellValue" : $"Type_{cfType}";

                                string op = string.Empty;
                                try { op = ((int)fc.Operator).ToString(); } catch { }

                                string f1 = string.Empty;
                                try { f1 = fc.Formula1?.ToString() ?? string.Empty; } catch { }

                                string f2 = string.Empty;
                                try { f2 = fc.Formula2?.ToString() ?? string.Empty; } catch { }

                                string severity = "None";
                                try
                                {
                                    int bgOle = Convert.ToInt32(fc.Interior.Color);
                                    int r = bgOle & 0xFF;
                                    int g = (bgOle >> 8) & 0xFF;
                                    int b = (bgOle >> 16) & 0xFF;
                                    if ((r > 200 && g < 150 && b < 150) || (r > 230 && g > 200 && b > 200 && r > g && r > b))
                                        severity = "Error";
                                    else if (r > 200 && g > 180 && b < 180)
                                        severity = "Warning";
                                }
                                catch { }

                                string appliesTo = string.Empty;
                                try { appliesTo = fc.AppliesTo?.Address?.ToString() ?? string.Empty; } catch { }

                                bool stopIfTrue = false;
                                try { stopIfTrue = (bool)fc.StopIfTrue; } catch { }

                                sb.AppendLine(
                                    $"| {RulesAuditHelper.EscapePipe(appliesTo)} " +
                                    $"| {ruleType} " +
                                    $"| {op} " +
                                    $"| {RulesAuditHelper.EscapePipe(f1)} " +
                                    $"| {RulesAuditHelper.EscapePipe(f2)} " +
                                    $"| {severity} " +
                                    $"| {(stopIfTrue ? "TRUE" : "FALSE")} |"
                                );
                            }
                            sb.AppendLine();
                        }
                    }
                    catch { }
                }
            }

            if (tableCount == 0 && sheetHeaderCellCount == 0)
            {
                MessageBox.Show("No ListObjects or configured template cells found.", "Export Rules", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            string exportPath = FolderRegistry.Data.GetFilePath(targetFileName);
            FolderRegistry.Data.EnsureExists();
            File.WriteAllText(exportPath, sb.ToString(), Encoding.UTF8);
            Clipboard.SetText(sb.ToString());
            MessageBox.Show(
                $"Audit export complete!\n\n  Tables: {tableCount}\n  Header/Date Parameter Cells: {sheetHeaderCellCount}\n  Saved file: {exportPath}\n  Copied to Clipboard.",
                "Export Complete", MessageBoxButtons.OK, MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to export audit report:\n{ex.Message}", "Export Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static List<RulesAuditHelper.HeaderCellInfo> ScanTopAnchorCells(dynamic ws, int maxRowLimit)
    {
        var list = new List<RulesAuditHelper.HeaderCellInfo>();
        try
        {
            int scanRows = Math.Min(3, maxRowLimit - 1);
            if (scanRows < 1) return list;

            for (int r = 1; r <= scanRows; r++)
            {
                for (int c = 1; c <= 7; c++)
                {
                    dynamic cell = ws.Cells[r, c];
                    string sName = RulesAuditHelper.GetStyleName(cell);
                    string numFormat = RulesAuditHelper.GetCellNumberFormat(cell);
                    string formula = cell.Formula?.ToString() ?? string.Empty;

                    bool isSpecialStyle = sName.Equals("DisplayDate", StringComparison.OrdinalIgnoreCase) ||
                                          sName.Equals("DisplayRange", StringComparison.OrdinalIgnoreCase);

                    bool hasVal = false;
                    string valType = string.Empty;
                    string f1 = string.Empty;
                    string alert = string.Empty;
                    bool ignoreBlank = true;

                    try
                    {
                        dynamic v = cell.Validation;
                        if (v != null)
                        {
                            int vt = (int)v.Type;
                            hasVal = true;
                            valType = RulesAuditHelper.GetValidationTypeName(vt);
                            f1 = v.Formula1?.ToString() ?? string.Empty;
                            alert = ((int)v.AlertStyle) == 2 ? "Warning" : "Stop";
                            ignoreBlank = (bool)v.IgnoreBlank;
                        }
                    }
                    catch { }

                    if (isSpecialStyle || hasVal || (!string.IsNullOrWhiteSpace(formula) && formula.StartsWith("=")))
                    {
                        list.Add(new RulesAuditHelper.HeaderCellInfo
                        {
                            Address = cell.Address?.ToString() ?? $"R{r}C{c}",
                            Row = r,
                            Col = c,
                            Style = sName,
                            NumberFormat = numFormat,
                            Formula = formula,
                            SourceOrFormula = f1,
                            NativeType = valType,
                            NativeAlert = alert,
                            NativeIgnoreBlank = ignoreBlank
                        });
                    }
                }
            }
        }
        catch { }
        return list;
    }
}