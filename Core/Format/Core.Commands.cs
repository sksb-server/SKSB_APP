namespace SKSB_App.Core;

using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;

public static class Commands
{
    /// <summary>
    /// Formats the selected range to mirror a reference cell's style and Data Validation list formula,
    /// while turning off the Error Alert to support multi-select.
    /// </summary>
    /// <param name="styleReferenceAddress">
    /// Optional cell coordinate holding the target style & validation formula 
    /// (e.g. "Lookups!$A$1" or "Master!$C$2"). If null, falls back to a default named list.
    /// </param>
    public static void ApplyStyleAndDropdownToSelection(string? styleReferenceAddress = null)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic selection = xlApp.Selection;

        if (selection == null) return;

        string listSource = "=DaysOfWeekList"; // Default fallback list

        // 1. If a style reference cell is provided, extract its style and validation formula
        if (!string.IsNullOrWhiteSpace(styleReferenceAddress))
        {
            try
            {
                dynamic refCell = xlApp.Range[styleReferenceAddress];

                // Extract the List validation formula from the template cell
                if (refCell.Validation != null && refCell.Validation.Type == 3) // 3 = xlValidateList
                {
                    listSource = refCell.Validation.Formula1;
                }

                // Copy formatting/style from the reference cell to the target selection
                refCell.Copy();
                // 13 = xlPasteFormats (copies font, fill color, borders, alignment)
                selection.PasteSpecial(-4122 /* xlPasteFormats */);
                xlApp.CutCopyMode = false;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Style Copy Warning] {ex.Message}");
            }
        }
        else
        {
            // Fallback visual cue if no reference cell is specified
            selection.Interior.Color = ColorTranslator.ToOle(Color.AliceBlue);
        }

        // 2. Clear existing validation
        selection.Validation.Delete();

        // 3. Apply the List Validation (3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween)
        selection.Validation.Add(3, 1, 1, listSource);

        // 4. CRITICAL: Configure multi-select parameters
        selection.Validation.IgnoreBlank = true;
        selection.Validation.InCellDropdown = true;
        selection.Validation.ShowError = false; // Prevents Excel from rejecting concatenated strings
    }

    public static void ShowChecklistForActiveCell(dynamic targetCell)
    {
        try
        {
            string validationFormula = targetCell.Validation.Formula1; 
            
            dynamic xlApp = ExcelDnaUtil.Application; 
            dynamic validationRange = xlApp.Evaluate(validationFormula); 
            
            if (validationRange == null) return; 

            string[] availableOptions;
            if (validationRange is object[,] rangeValues) 
            {
                availableOptions = rangeValues.Cast<object>()
                                              .Where(o => o != null)
                                              .Select(o => o.ToString()!)
                                              .ToArray(); 
            }
            else
            {
                return; 
            }

            string currentText = targetCell.Value2?.ToString() ?? string.Empty; 
            var selectedItems = currentText.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(d => d.Trim())
                                           .ToHashSet(StringComparer.OrdinalIgnoreCase); 

            using var form = new Form
            {
                Text = "Select Options",
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                StartPosition = FormStartPosition.Manual,
                Width = 240,
                Height = 280,
                TopMost = true,
                ShowInTaskbar = false
            }; 

            form.Location = Cursor.Position; 

            var checkList = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true }; 
            foreach (string option in availableOptions)
            {
                checkList.Items.Add(option, selectedItems.Contains(option)); 
            }

            var btnDone = new Button { Text = "Apply", Dock = DockStyle.Bottom, Height = 32 }; 
            btnDone.Click += (_, _) =>
            {
                var chosen = checkList.CheckedItems.Cast<string>(); 
                targetCell.Value2 = string.Join(", ", chosen); 
                form.Close(); 
            };

            form.Controls.Add(checkList); 
            form.Controls.Add(btnDone); 
            form.ShowDialog(); 
        }
        catch { }
    }
}