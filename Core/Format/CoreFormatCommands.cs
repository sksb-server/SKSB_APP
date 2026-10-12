namespace SKSB_App.Core.UI;

using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;

public static class FormatCommands
{
    public static void ApplyCustomDropdownToSelection(string listSource)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic selection = xlApp.Selection;

        if (selection == null) return;

        // Clear existing validation
        selection.Validation.Delete();

        // Add standard List Validation (3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween)
        selection.Validation.Add(3, 1, 1, listSource);

        // CRITICAL: Disable the error alert to permit multiple comma-separated items
        selection.Validation.IgnoreBlank = true;
        selection.Validation.InCellDropdown = true;
        selection.Validation.ShowError = false;

        // Mark the cell with a subtle background color to denote it is a WinForms Multi-Select cell
        selection.Interior.Color = ColorTranslator.ToOle(Color.AliceBlue);
    }

    public static void ShowChecklistForActiveCell(dynamic targetCell)
    {
        try
        {
            // Extract the validation formula (e.g. "=DaysOfWeekList")
            string validationFormula = targetCell.Validation.Formula1;

            dynamic xlApp = ExcelDnaUtil.Application;
            // Evaluate the validation range to extract permitted items dynamically
            dynamic validationRange = xlApp.Evaluate(validationFormula);

            if (validationRange == null) return;

            string[] availableOptions;

            // Extract the list of options from the evaluated range
            if (validationRange is object[,] rangeValues)
            {
                availableOptions = rangeValues.Cast<object>()
                                              .Where(o => o != null)
                                              .Select(o => o.ToString())
                                              .ToArray();
            }
            else
            {
                return; // Invalid format
            }

            // Parse currently selected items in the cell
            string currentText = targetCell.Value2?.ToString() ?? string.Empty;
            var selectedItems = currentText.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(d => d.Trim())
                                           .ToHashSet(StringComparer.OrdinalIgnoreCase);

            // Construct the WinForms popup
            using var form = new Form
            {
                Text = "Select Options",
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                StartPosition = FormStartPosition.Manual,
                Width = 220,
                Height = 260,
                TopMost = true,
                ShowInTaskbar = false
            };

            form.Location = Cursor.Position;

            var checkList = new CheckedListBox { Dock = DockStyle.Fill, CheckOnClick = true };

            foreach (string option in availableOptions)
            {
                bool isChecked = selectedItems.Contains(option);
                checkList.Items.Add(option, isChecked);
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
        catch
        {
            // If the cell doesn't have list validation, quietly ignore the double click
        }
    }
}