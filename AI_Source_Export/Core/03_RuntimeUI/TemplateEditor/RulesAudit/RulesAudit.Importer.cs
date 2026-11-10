namespace SKSB_App.Core.RuntimeUI.TemplateEditor.RulesAudit;

using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public static class RulesAuditImporter
{
    [ExcelCommand(Name = "ImportTemplateRulesFromMarkdown")]
    public static void ImportAuditReport()
    {
        var diag = new StringBuilder();
        diag.AppendLine("=== RULES AUDIT IMPORT DIAGNOSTIC LOG ===");
        diag.AppendLine($"Started at: {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        diag.AppendLine();

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null)
            {
                MessageBox.Show("No active workbook found.", "Import Rules Audit", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            string dataDir = FolderRegistry.Data.EnsureExists();
            using var openFileDialog = new OpenFileDialog
            {
                Title = "Select Rules Audit Markdown File to Import",
                InitialDirectory = dataDir,
                Filter = "Markdown Files (*.md)|*.md|Text Files (*.txt)|*.txt|All Files (*.*)|*.*",
                FilterIndex = 1,
                CheckFileExists = true,
                Multiselect = false,
                RestoreDirectory = true
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK) return;
            string selectedFilePath = openFileDialog.FileName;
            diag.AppendLine($"Source File: {selectedFilePath}");

            string markdownContent = File.ReadAllText(selectedFilePath, Encoding.UTF8);
            if (string.IsNullOrWhiteSpace(markdownContent))
            {
                MessageBox.Show("The selected file is empty.", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            Commands.EnsureNamedStyleExists(xlApp, Commands.DropDownStyleName);

            var lines = markdownContent.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None);
            string currentSheet = string.Empty;
            dynamic? currentWs = null;
            int archetypeRow = 3;

            bool readingHeaderCells = false;
            bool readingCols = false;
            bool readingCf = false;

            var headerColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var tableColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            var cfColMap = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            int appliedHeaderCells = 0;
            int appliedRulesCount = 0;
            int appliedCfCount = 0;
            var clearedSheetsCf = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            bool origAlerts = true;
            try { origAlerts = (bool)xlApp.DisplayAlerts; xlApp.DisplayAlerts = false; } catch { }

            try
            {
                foreach (var rawLine in lines)
                {
                    string line = rawLine.Replace(@"\vert{}", "|").Trim();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    if (line.StartsWith("## SHEET_HEADER_CELLS:"))
                    {
                        currentSheet = RulesAuditHelper.ExtractHeaderVal(line, "Sheet");
                        currentWs = RulesAuditHelper.ResolveSheetSafely(wb, currentSheet);
                        if (currentWs != null) { try { currentWs.Activate(); } catch { } }

                        readingHeaderCells = true;
                        readingCols = false;
                        readingCf = false;
                        headerColMap.Clear();
                        diag.AppendLine($"\n--- SECTION: SHEET_HEADER_CELLS on [{currentSheet}] ---");
                        continue;
                    }

                    if (line.StartsWith("## TABLE:"))
                    {
                        currentSheet = RulesAuditHelper.ExtractHeaderVal(line, "Sheet");
                        currentWs = RulesAuditHelper.ResolveSheetSafely(wb, currentSheet);
                        if (currentWs != null) { try { currentWs.Activate(); } catch { } }

                        int.TryParse(RulesAuditHelper.ExtractHeaderVal(line, "ArchetypeRow"), out archetypeRow);
                        readingHeaderCells = false;
                        readingCols = false;
                        readingCf = false;
                        tableColMap.Clear();
                        diag.AppendLine($"\n--- SECTION: TABLE on [{currentSheet}], ArchetypeRow={archetypeRow} ---");
                        continue;
                    }

                    if (line.StartsWith("### Column Rules"))
                    {
                        readingCols = true;
                        readingHeaderCells = false;
                        readingCf = false;
                        tableColMap.Clear();
                        continue;
                    }

                    if (line.StartsWith("### Conditional Formatting"))
                    {
                        readingCf = true;
                        readingHeaderCells = false;
                        readingCols = false;
                        cfColMap.Clear();
                        continue;
                    }

                    if (readingHeaderCells && line.StartsWith("|") && line.IndexOf("CellAddress", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        headerColMap = BuildHeaderMap(line);
                        continue;
                    }
                    if (readingCols && line.StartsWith("|") && line.IndexOf("Header", StringComparison.OrdinalIgnoreCase) >= 0 && line.IndexOf("EngineType", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        tableColMap = BuildHeaderMap(line);
                        continue;
                    }
                    if (readingCf && line.StartsWith("|") && line.IndexOf("AppliesTo", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        cfColMap = BuildHeaderMap(line);
                        continue;
                    }
                    if (line.StartsWith("| :---") || line.StartsWith("| Col |") || line.StartsWith("| CellAddress |") || line.StartsWith("| AppliesTo |"))
                    {
                        continue;
                    }

                    var rawTokens = line.Split('|');
                    var cells = rawTokens.Select(c => RulesAuditHelper.UnescapePipe(c.Trim())).ToArray();
                    if (currentWs == null) continue;

                    // 1. Process Sheet Header Parameter Cells
                    if (readingHeaderCells && line.StartsWith("|") && cells.Length >= 5)
                    {
                        string cellAddress = GetField(cells, headerColMap, "CellAddress", 1);
                        int row = int.TryParse(GetField(cells, headerColMap, "Row", 2), out int r) ? r : 1;
                        int col = int.TryParse(GetField(cells, headerColMap, "Col", 3), out int cl) ? cl : RulesAuditHelper.ColumnLetterToIndex(Regex.Replace(cellAddress, @"[^\w]", ""));
                        string styleName = GetField(cells, headerColMap, "Style", 4);
                        string numFormat = GetField(cells, headerColMap, "NumberFormat", 5);
                        string inCellFormula = GetField(cells, headerColMap, "Formula", 6);
                        string sourceOrNative = GetField(cells, headerColMap, "SourceList / NativeFormula", 7);
                        string nativeType = GetField(cells, headerColMap, "NativeType", 8);
                        string nativeAlert = GetField(cells, headerColMap, "NativeAlert", 9);
                        bool nativeIgnoreBlank = GetField(cells, headerColMap, "NativeIgnoreBlank", 10).Equals("TRUE", StringComparison.OrdinalIgnoreCase);

                        try
                        {
                            dynamic targetCell = currentWs.Cells[row, col];
                            if (!string.IsNullOrWhiteSpace(styleName) && !styleName.Equals("Normal", StringComparison.OrdinalIgnoreCase))
                            {
                                Commands.EnsureNamedStyleExists(xlApp, styleName);
                                targetCell.Style = styleName;
                            }
                            if (!string.IsNullOrWhiteSpace(inCellFormula) && inCellFormula.StartsWith("="))
                            {
                                try { targetCell.Formula2 = inCellFormula; }
                                catch { targetCell.Formula = inCellFormula; }
                            }
                            ApplyNumberFormatSafely(targetCell, numFormat);

                            if (!string.IsNullOrWhiteSpace(nativeType) && !string.IsNullOrWhiteSpace(sourceOrNative))
                            {
                                ApplyNativeValidationSafely(wb, currentWs, targetCell, nativeType, nativeAlert, sourceOrNative, nativeIgnoreBlank, diag, cellAddress);
                            }
                            appliedHeaderCells++;
                        }
                        catch (Exception exHdr)
                        {
                            diag.AppendLine($"  -> Header Cell ERROR on {cellAddress}: {exHdr.Message}");
                        }
                    }

                    // 2. Process Table Column Rules
                    if (readingCols && line.StartsWith("|") && cells.Length >= 10)
                    {
                        string colLetter = GetField(cells, tableColMap, "Col", 1);
                        int explicitRow = int.TryParse(GetField(cells, tableColMap, "Row", 2), out int rNum) ? rNum : archetypeRow;
                        string styleName = GetField(cells, tableColMap, "Style", 4);
                        string numFormat = GetField(cells, tableColMap, "NumberFormat", 5);
                        string engineType = GetField(cells, tableColMap, "EngineType", 6);
                        string inCellFormula = GetField(cells, tableColMap, "InCellFormula", 7);
                        string sourceList = GetField(cells, tableColMap, "Sources", 8);
                        string syncFormula = GetField(cells, tableColMap, "Syncs", 9);
                        string transformExpr = GetField(cells, tableColMap, "Transform", 10);
                        string defaultRule = GetField(cells, tableColMap, "DefaultRule", 11);
                        bool isMulti = GetField(cells, tableColMap, "IsMulti", 13).Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                        bool isWarning = GetField(cells, tableColMap, "IsWarning", 14).Equals("TRUE", StringComparison.OrdinalIgnoreCase);
                        string nativeType = GetField(cells, tableColMap, "NativeType", 15);
                        string nativeFormula = GetField(cells, tableColMap, "NativeFormula", 16);
                        string nativeAlert = GetField(cells, tableColMap, "NativeAlert", 17);
                        bool nativeIgnoreBlank = GetField(cells, tableColMap, "NativeIgnoreBlank", 18).Equals("TRUE", StringComparison.OrdinalIgnoreCase);

                        try
                        {
                            int colIndex = RulesAuditHelper.ColumnLetterToIndex(colLetter);
                            dynamic archetypeCell = currentWs.Cells[explicitRow, colIndex];

                            try { archetypeCell.Validation.Delete(); } catch { }

                            // If this is a C# Attribute column, avoid injecting legacy metadata.
                            if (engineType.Equals("CSharpAttribute", StringComparison.OrdinalIgnoreCase))
                            {
                                Metadata.PurgeConfigIfPresent(archetypeCell);
                            }
                            else if (engineType.Equals("LegacyDropdown", StringComparison.OrdinalIgnoreCase) || !string.IsNullOrWhiteSpace(sourceList))
                            {
                                Commands.EnsureNamedStyleExists(xlApp, Commands.DropDownStyleName);
                                archetypeCell.Style = Commands.DropDownStyleName;
                                var cfg = new Config
                                {
                                    Source = sourceList,
                                    SyncLeft = syncFormula,
                                    Filter = defaultRule,
                                    Transform = transformExpr,
                                    IsMultiSelect = isMulti,
                                    IsWarningOnly = isWarning
                                };
                                Metadata.WriteConfig(archetypeCell, cfg);
                            }
                            else
                            {
                                Metadata.PurgeConfigIfPresent(archetypeCell);
                            }

                            if (!string.IsNullOrWhiteSpace(inCellFormula) && inCellFormula.StartsWith("="))
                            {
                                try { archetypeCell.Formula2 = inCellFormula; }
                                catch { archetypeCell.Formula = inCellFormula; }
                            }

                            ApplyNumberFormatSafely(currentWs.Columns[colIndex], numFormat);
                            ApplyNumberFormatSafely(archetypeCell, numFormat);

                            if (engineType.Equals("NativeValidation", StringComparison.OrdinalIgnoreCase) && !string.IsNullOrWhiteSpace(nativeFormula))
                            {
                                ApplyNativeValidationSafely(wb, currentWs, archetypeCell, nativeType, nativeAlert, nativeFormula, nativeIgnoreBlank, diag, $"{colLetter}{explicitRow}");
                            }

                            appliedRulesCount++;
                        }
                        catch (Exception exCol)
                        {
                            diag.AppendLine($"  -> Col Error on {colLetter}{explicitRow}: {exCol.Message}");
                        }
                    }

                    // 3. Process Conditional Formatting
                    if (readingCf && line.StartsWith("|") && cells.Length >= 5)
                    {
                        string appliesTo = GetField(cells, cfColMap, "AppliesTo", 1);
                        string ruleType = GetField(cells, cfColMap, "RuleType", 2);
                        string op = GetField(cells, cfColMap, "Operator", 3);
                        string f1 = GetField(cells, cfColMap, "Formula1", 4);
                        string f2 = GetField(cells, cfColMap, "Formula2", 5);
                        string severity = GetField(cells, cfColMap, "Severity", 6);
                        bool stopIfTrue = GetField(cells, cfColMap, "StopIfTrue", 7).Equals("TRUE", StringComparison.OrdinalIgnoreCase);

                        if (string.IsNullOrWhiteSpace(appliesTo) || string.IsNullOrWhiteSpace(f1)) continue;

                        try
                        {
                            if (!clearedSheetsCf.Contains(currentSheet))
                            {
                                try { currentWs.Cells.FormatConditions.Delete(); } catch { }
                                clearedSheetsCf.Add(currentSheet);
                            }

                            dynamic targetRange = currentWs.Range[appliesTo];
                            int xlCfType = ruleType.Equals("CellValue", StringComparison.OrdinalIgnoreCase) ? 1 : 2;
                            dynamic newRule = null;

                            if (xlCfType == 2)
                            {
                                newRule = targetRange.FormatConditions.Add(Type: 2, Formula1: f1);
                            }
                            else
                            {
                                int opCode = int.TryParse(op, out int o) ? o : 1;
                                newRule = targetRange.FormatConditions.Add(Type: 1, Operator: opCode, Formula1: f1, Formula2: f2);
                            }

                            if (newRule != null)
                            {
                                if (severity.Equals("Error", StringComparison.OrdinalIgnoreCase))
                                {
                                    newRule.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 226, 226));
                                }
                                else if (severity.Equals("Warning", StringComparison.OrdinalIgnoreCase))
                                {
                                    newRule.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(254, 243, 199));
                                }
                                newRule.StopIfTrue = stopIfTrue;
                                appliedCfCount++;
                            }
                        }
                        catch (Exception exCf)
                        {
                            diag.AppendLine($"  -> CF Error on {appliesTo}: {exCf.Message}");
                        }
                    }
                }
            }
            finally
            {
                try { xlApp.DisplayAlerts = origAlerts; } catch { }
            }

            diag.AppendLine();
            diag.AppendLine($"Summary: Parameter Cells={appliedHeaderCells}, Column Rules={appliedRulesCount}, CF Rules={appliedCfCount}");
            DebugDialog.ShowLog("Import Diagnostics Log", diag.ToString());

            MessageBox.Show(
                $"Import completed successfully from:\n{Path.GetFileName(selectedFilePath)}\n\n" +
                $"  Configured Parameter Cells : {appliedHeaderCells}\n" +
                $"  Applied Table Column Rules : {appliedRulesCount}\n" +
                $"  Applied Conditional Rules  : {appliedCfCount}\n\n" +
                "All rules, validations, formulas, and visual warnings are active.",
                "Import Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to import rules audit:\n{ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private static Dictionary<string, int> BuildHeaderMap(string headerLine)
    {
        var map = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
        var tokens = headerLine.Split('|');
        for (int i = 0; i < tokens.Length; i++)
        {
            string clean = tokens[i].Trim();
            if (!string.IsNullOrEmpty(clean) && !clean.StartsWith(":") && !clean.StartsWith("-"))
            {
                map[clean] = i;
            }
        }
        return map;
    }

    private static string GetField(string[] cells, Dictionary<string, int> map, string headerName, int fallbackIndex)
    {
        if (map.TryGetValue(headerName, out int index) && index >= 0 && index < cells.Length)
        {
            return cells[index]?.Trim() ?? string.Empty;
        }
        return SafeGet(cells, fallbackIndex);
    }

    private static void ApplyNumberFormatSafely(dynamic rangeObj, string? explicitFormat)
    {
        try
        {
            string fmt = (explicitFormat ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(fmt))
            {
                fmt = "General";
            }
            string currentFmt = rangeObj.NumberFormat?.ToString() ?? string.Empty;
            if (!string.Equals(currentFmt, fmt, StringComparison.OrdinalIgnoreCase))
            {
                rangeObj.NumberFormatLocal = fmt;
                rangeObj.NumberFormat = fmt;
            }
        }
        catch { }
    }

    private static string SafeGet(string[] arr, int index)
    {
        if (index < 0 || index >= arr.Length) return string.Empty;
        return arr[index] ?? string.Empty;
    }

    private static void ApplyNativeValidationSafely(
        dynamic wb,
        dynamic ws,
        dynamic cell,
        string nativeType,
        string nativeAlert,
        string formula1,
        bool ignoreBlank,
        StringBuilder diag,
        string cellLabel)
    {
        try
        {
            try { ws.Activate(); } catch { }
            try { cell.Select(); } catch { }
            try { cell.Validation.Delete(); } catch { }

            int xlValType = RulesAuditHelper.GetValidationTypeCode(nativeType);
            int xlAlert = string.Equals(nativeAlert, "Warning", StringComparison.OrdinalIgnoreCase) ? 2 : 1;
            string raw = (formula1 ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(raw)) return;
            string unequals = raw.TrimStart('=').Trim();
            if (string.IsNullOrWhiteSpace(unequals)) return;

            if (xlValType == 3) // List
            {
                string listFormula = raw.StartsWith("=") ? ("=" + unequals) : raw;
                if (raw.StartsWith("="))
                {
                    bool exists = false;
                    try
                    {
                        dynamic nm = wb.Names.Item(unequals);
                        if (nm != null) exists = true;
                    }
                    catch { exists = false; }
                    if (!exists)
                    {
                        diag.AppendLine($"  [Validation Notice] {cellLabel}: Skipped List because Name '{unequals}' does not exist in Name Manager.");
                        return;
                    }
                }
                cell.Validation.Add(Type: 3, AlertStyle: xlAlert, Formula1: listFormula);
            }
            else if (xlValType == 7) // CustomFormula
            {
                cell.Validation.Add(Type: 7, AlertStyle: xlAlert, Formula1: "=" + unequals);
            }
            else if (xlValType == 1 || xlValType == 2) // WholeNumber, Decimal
            {
                string numStr = double.TryParse(unequals, NumberStyles.Any, CultureInfo.InvariantCulture, out double dVal)
                     ? (dVal == 0 ? "0" : dVal.ToString(CultureInfo.InvariantCulture))
                     : "0";
                int op = (numStr == "0") ? 7 : 5;
                cell.Validation.Add(Type: xlValType, AlertStyle: xlAlert, Operator: op, Formula1: numStr);
            }
            else if (xlValType == 4) // Date
            {
                string dateStr;
                if (DateTime.TryParseExact(unequals, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime dtIso))
                {
                    dateStr = $"=DATE({dtIso.Year},{dtIso.Month},{dtIso.Day})";
                }
                else if (raw.StartsWith("="))
                {
                    dateStr = "=" + unequals;
                }
                else
                {
                    dateStr = raw;
                }
                cell.Validation.Add(Type: 4, AlertStyle: xlAlert, Operator: 5, Formula1: dateStr);
            }
            else
            {
                cell.Validation.Add(Type: xlValType, AlertStyle: xlAlert, Formula1: raw);
            }

            cell.Validation.IgnoreBlank = ignoreBlank;
        }
        catch (Exception exVal)
        {
            diag.AppendLine($"  [Validation Exception] {cellLabel} (Type={nativeType}, Formula={formula1}): {exVal.Message}");
        }
    }
}