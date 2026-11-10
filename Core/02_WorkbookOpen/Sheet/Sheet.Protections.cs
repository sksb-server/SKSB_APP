namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;

public static class SheetProtections
{
    private const string ProtectionPassword = "SKSB_Internal_Guard";

    // Whitelist of styles that represent user-editable inputs
    public static readonly HashSet<string> UnlockedStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "DropDown",
        "Unlocked",
        "StandardInput",
        "DisplayDate",
        "DisplayRange",
        "Comma",        // Standard Excel Comma style (#,##0.00)
        "Comma [0]",     // Comma style without decimals (#,##0)
        "Cost"
    };

    public static bool IsUnlockedStyle(string? styleName) =>
        !string.IsNullOrWhiteSpace(styleName) && UnlockedStyles.Contains(styleName.Trim());

    public static bool IsProtected(dynamic worksheet)
    {
        try { return (bool)worksheet.ProtectContents; } catch { return false; }
    }

    public static void Unprotect(dynamic worksheet)
    {
        try { worksheet.Unprotect(ProtectionPassword); } catch { }
    }

    /// <summary>
    /// Executes an action that modifies table structure by temporarily lifting protection.
    /// </summary>
    public static void ExecuteWithBypass(dynamic worksheet, Action action)
    {
        if (worksheet == null || action == null) return;
        // Sheet protection is no longer used; lift any legacy protection and never re-apply it.
        if (IsProtected(worksheet)) Unprotect(worksheet);
        action();
    }

    /// <summary>
    /// Unlocks table structure for native sorting, while locking non-editable columns
    /// using Excel Custom Data Validation (=FALSE) with a custom alert popup.
    /// </summary>
    public static void ApplyStyleDrivenProtection(dynamic worksheet)
    {
        if (SheetOperations.IsTemplateEditMode)
        {
            try { worksheet.Unprotect(ProtectionPassword); } catch { }
            try
            {
                worksheet.Cells.Locked = false;
                foreach (dynamic table in worksheet.ListObjects)
                {
                    table.Range.Locked = false;
                    foreach (dynamic col in table.ListColumns)
                    {
                        try { col.Range.Validation.Delete(); } catch { }
                    }
                }
            }
            catch { }
            return;
        }

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

        try
        {
            try { worksheet.Unprotect(ProtectionPassword); } catch { }

            // Configure all tables on the worksheet (no sheet protection is applied)
            foreach (dynamic table in worksheet.ListObjects)
            {
                dynamic bodyRange = table.DataBodyRange;
                int headerRow = (int)table.HeaderRowRange.Row;
                int startCol = (int)table.HeaderRowRange.Column;
                int colCount = (int)table.ListColumns.Count;
                int bodyRows = bodyRange != null ? (int)bodyRange.Rows.Count : 1;

                for (int c = 1; c <= colCount; c++)
                {
                    int currentCol = startCol + c - 1;
                    dynamic archetypeCell = worksheet.Cells[headerRow + 1, currentCol];
                    string styleName = archetypeCell.Style?.Name?.ToString() ?? string.Empty;

                    dynamic colBody = worksheet.Range[
                        worksheet.Cells[headerRow + 1, currentCol],
                        worksheet.Cells[headerRow + bodyRows, currentCol]
                    ];

                    if (!UnlockedStyles.Contains(styleName))
                    {
                        // Apply custom =FALSE Stop alert to calculated / read-only columns
                        try
                        {
                            colBody.Validation.Delete();
                            colBody.Validation.Add(
                                Type: 7,       // xlValidateCustom
                                AlertStyle: 1, // xlValidAlertStop
                                Formula1: "=FALSE"
                            );
                            colBody.Validation.IgnoreBlank = true;
                            colBody.Validation.ErrorMessage = "This column is auto-calculated and cannot be modified directly.";
                            colBody.Validation.ErrorTitle = "Protected Field";
                        }
                        catch { }
                    }
                    else
                    {
                        // Clean up any stale =FALSE validation if the column style was changed to editable
                        try
                        {
                            if (colBody.Validation?.Formula1 == "=FALSE")
                            {
                                colBody.Validation.Delete();
                            }
                        }
                        catch { }
                    }
                }
            }

        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ApplyStyleDrivenProtection Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    private static void UnlockIfStylePermits(dynamic cell)
    {
        try
        {
            string styleName = cell.Style?.Name?.ToString() ?? string.Empty;
            if (UnlockedStyles.Contains(styleName))
            {
                cell.Locked = false;
            }
        }
        catch { }
    }
}