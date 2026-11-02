namespace SKSB_App.Core.WorkbookOpen.Sheet;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;

public static class SheetProtections
{
    private const string ProtectionPassword = "SKSB_Internal_Guard";

    private static readonly HashSet<string> UnlockedStyles = new(StringComparer.OrdinalIgnoreCase)
    {
        "DropDown",
        "Unlocked",
        "StandardInput",
        "DisplayDate",
        "DisplayRange"
    };

    public static bool IsProtected(dynamic worksheet)
    {
        try { return (bool)worksheet.ProtectContents; } catch { return false; }
    }

    public static void Unprotect(dynamic worksheet)
    {
        try { worksheet.Unprotect(ProtectionPassword); } catch { }
    }

    public static void ExecuteWithBypass(dynamic worksheet, Action action)
    {
        if (worksheet == null || action == null) return;
        bool wasProtected = IsProtected(worksheet);
        if (wasProtected) Unprotect(worksheet);
        try
        {
            action();
        }
        finally
        {
            if (wasProtected) ApplyStyleDrivenProtection(worksheet);
        }
    }

    public static void ApplyStyleDrivenProtection(dynamic worksheet)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            xlApp.ScreenUpdating = false;

            try { worksheet.Unprotect(ProtectionPassword); } catch { }

            // 1. Lock all background cells, headers, and sheet layout by default
            worksheet.Cells.Locked = true;

            // 2. Unlock top parameter cells (B1 Date, C1 Range)
            UnlockIfStylePermits(worksheet.Cells[1, 2]);
            UnlockIfStylePermits(worksheet.Cells[1, 3]);

            // 3. Configure all ListObjects
            foreach (dynamic table in worksheet.ListObjects)
            {
                dynamic bodyRange = table.DataBodyRange;
                int headerRow = (int)table.HeaderRowRange.Row;
                int startCol = (int)table.HeaderRowRange.Column;
                int colCount = (int)table.ListColumns.Count;

                if (bodyRange != null)
                {
                    // Clean up any legacy =FALSE validation rules from previous experiments
                    for (int c = 1; c <= colCount; c++)
                    {
                        try
                        {
                            dynamic colRange = table.ListColumns[c].Range;
                            if (colRange.Validation?.Formula1 == "=FALSE")
                            {
                                colRange.Validation.Delete();
                            }
                        }
                        catch { }
                    }

                    // Unlock table data body so Excel's native header sort & filter commands work smoothly
                    bodyRange.Locked = false;
                }
                else
                {
                    // Table with 0 data rows: unlock archetype row so users can begin entering data
                    for (int c = 1; c <= colCount; c++)
                    {
                        worksheet.Cells[headerRow + 1, startCol + c - 1].Locked = false;
                    }
                }
            }

            // 4. Apply sheet protection
            worksheet.Protect(
                Password: ProtectionPassword,
                DrawingObjects: false,
                Contents: true,
                Scenarios: true,
                UserInterfaceOnly: true,
                AllowFormattingCells: false,
                AllowFormattingColumns: true,
                AllowFormattingRows: true,
                AllowInsertingRows: true,
                AllowDeletingRows: true,
                AllowSorting: true,
                AllowFiltering: true
            );
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ApplyStyleDrivenProtection Error] {ex.Message}");
        }
        finally
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                xlApp.ScreenUpdating = true;
            }
            catch { }
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