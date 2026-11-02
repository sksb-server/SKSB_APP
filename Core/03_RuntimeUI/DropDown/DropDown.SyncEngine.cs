namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Globalization;
using System.Text;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookOpen.Sheet;

public static class SyncEngine
{
    public static void ExecuteSync(dynamic singleCell, string syncFormula, string rawVal, string displayVal)
    {
        if (string.IsNullOrWhiteSpace(syncFormula)) return;
        string cellVal = !string.IsNullOrEmpty(rawVal) ? rawVal : (singleCell.Value2?.ToString()?.Trim() ?? string.Empty);
        if (string.IsNullOrEmpty(cellVal)) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic sheet = singleCell.Worksheet;
        int row = (int)singleCell.Row;
        int col = (int)singleCell.Column;

        var statements = syncFormula.Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
        bool origEvents = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        var debugAudit = new StringBuilder();
        debugAudit.AppendLine("==========================================================================");
        debugAudit.AppendLine(" AUTO-FILL SYNC ENGINE AUDIT");
        debugAudit.AppendLine("==========================================================================");
        debugAudit.AppendLine($"Trigger Cell : {singleCell.Address} (Row {row}, Col {col}) on '{sheet.Name}'");
        debugAudit.AppendLine($"Input Value  : '{cellVal}' (Display: '{displayVal}')");
        debugAudit.AppendLine();

        int executedCount = 0;

        try
        {
            foreach (var rawStmt in statements)
            {
                string stmt = rawStmt.Trim();
                if (string.IsNullOrEmpty(stmt)) continue;

                int targetCol = col - 1;
                string targetColLabel = "Left (Col - 1)";

                if (stmt.StartsWith("@target[", StringComparison.OrdinalIgnoreCase))
                {
                    int closeBracket = stmt.IndexOf(']');
                    if (closeBracket > 8)
                    {
                        string targetHeader = stmt.Substring(8, closeBracket - 8).Trim();
                        stmt = stmt.Substring(closeBracket + 1).Trim();
                        try
                        {
                            dynamic listObj = singleCell.ListObject;
                            if (listObj != null)
                            {
                                targetCol = (int)listObj.ListColumns[targetHeader].Range.Column;
                                targetColLabel = $"Target [{targetHeader}] -> Col {targetCol}";
                            }
                        }
                        catch { }
                    }
                }
                else if (stmt.Contains("@to_right", StringComparison.OrdinalIgnoreCase) || stmt.Contains("@right", StringComparison.OrdinalIgnoreCase))
                {
                    targetCol = col + 1;
                    targetColLabel = $"Right (Col + 1) -> Col {targetCol}";
                    stmt = stmt.Replace("@to_right", "", StringComparison.OrdinalIgnoreCase)
                               .Replace("@right", "", StringComparison.OrdinalIgnoreCase);
                }
                else if (stmt.Contains("@to_left", StringComparison.OrdinalIgnoreCase) || stmt.Contains("@left", StringComparison.OrdinalIgnoreCase))
                {
                    targetCol = col - 1;
                    targetColLabel = $"Left (Col - 1)";
                    stmt = stmt.Replace("@to_left", "", StringComparison.OrdinalIgnoreCase)
                               .Replace("@left", "", StringComparison.OrdinalIgnoreCase);
                }

                bool isNum = double.TryParse(cellVal, NumberStyles.Any, CultureInfo.InvariantCulture, out _) && cellVal.Length < 10;
                string rawToken = isNum ? cellVal : $"\"{cellVal.Replace("\"", "\"\"")}\"";
                string displayToken = $"\"{displayVal.Replace("\"", "\"\"")}\"";

                string preparedStmt = stmt.Replace("@item", rawToken)
                                          .Replace("@display", displayToken)
                                          .Trim().TrimStart('=');

                // CRITICAL FIX: Resolve cross-sheet table references before evaluating
                preparedStmt = Metadata.ResolveTableReferences(xlApp, preparedStmt);

                debugAudit.AppendLine($"--- Statement {++executedCount} ---");
                debugAudit.AppendLine($"Direction / Target : {targetColLabel}");
                debugAudit.AppendLine($"Formula Evaluated  : {preparedStmt}");
                object? resolvedVal = Metadata.EvaluateFormulaSafely(xlApp, preparedStmt);

                // Type-tolerant retry: If unquoted lookup against text range returned empty, try quoted
                if ((resolvedVal == null || string.IsNullOrWhiteSpace(resolvedVal.ToString())) && isNum)
                {
                    string quotedStmt = stmt.Replace("@item", $"\"{cellVal.Replace("\"", "\"\"")}\"")
                                            .Replace("@display", displayToken)
                                            .Trim().TrimStart('=');

                    quotedStmt = Metadata.ResolveTableReferences(xlApp, quotedStmt);
                    resolvedVal = Metadata.EvaluateFormulaSafely(xlApp, quotedStmt);
                }

                if (resolvedVal == null)
                {
                    debugAudit.AppendLine("Result             : <NULL>");
                    continue;
                }

                string resStr = resolvedVal.ToString()?.Trim() ?? string.Empty;
                bool isComError = (resolvedVal is int iErr && iErr < -2000000000) ||
                                  (resolvedVal is long lErr && lErr < -2000000000) ||
                                  resStr.StartsWith("#") ||
                                  (long.TryParse(resStr, out long num) && num < -2000000000);

                debugAudit.AppendLine($"Result Type        : {resolvedVal.GetType().Name}");
                debugAudit.AppendLine($"Result Value       : '{resStr}'");
                debugAudit.AppendLine($"Is COM Error?      : {isComError}");

                if (!isComError && !string.IsNullOrEmpty(resStr))
                {
                    try
                    {
                        sheet.Cells[row, targetCol].Value2 = resolvedVal;
                        debugAudit.AppendLine($"Writeback Status   : SUCCESS -> Wrote '{resStr}' to Cell R{row}C{targetCol}");
                    }
                    catch (Exception exWrite)
                    {
                        debugAudit.AppendLine($"Writeback Status   : FAILED -> {exWrite.Message}");
                    }
                }
                else
                {
                    debugAudit.AppendLine("Writeback Status   : SKIPPED (Result was empty or evaluated to an Excel Error)");
                }
                debugAudit.AppendLine();
            }
        }
        catch (Exception ex)
        {
            debugAudit.AppendLine($"Global Execution Exception: {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
#if DEBUG
            if (DebugDialog.DebugEnabled && !DebugDialog.SuppressForSession && executedCount > 0)
            {
                DebugDialog.ShowLog("Auto-Fill Sync Audit", debugAudit.ToString());
            }
#endif
        }
    }
}