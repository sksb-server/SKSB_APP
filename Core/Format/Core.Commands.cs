namespace SKSB_App.Core.Format;

using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookManager;

public static class Commands
{
    public const string DropDownStyleName = "DropDown";

    /// <summary>
    /// Ensures the custom Named Style "DropDown" exists in the workbook and applies it to the active selection.
    /// </summary>
    public static void ApplyCustomDropdownStyle(string listSource = "=DaysOfWeekList")
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic selection = xlApp.Selection;

        if (selection == null) return;

        EnsureNamedStyleExists(xlApp, DropDownStyleName);

        // 1. Assign the custom Named Style
        selection.Style = DropDownStyleName;

        // 2. Clear previous validation and assign multi-select validation
        selection.Validation.Delete();
        selection.Validation.Add(3, 1, 1, listSource); // 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
        selection.Validation.IgnoreBlank = true;
        selection.Validation.InCellDropdown = true;
        selection.Validation.ShowError = false; // Bypasses Excel block to support multiple values
    }

    /// <summary>
    /// Checks whether a cell possesses the "DropDown" named style OR has list validation.
    /// </summary>
    public static bool IsCustomDropdownCell(dynamic cell)
    {
        if (cell == null) return false;

        try
        {
            dynamic singleCell = cell.Cells[1, 1];

            // 1. Try checking the Style name
            try
            {
                dynamic style = singleCell.Style;
                if (style != null)
                {
                    string sName = style.Name?.ToString() ?? string.Empty;
                    if (string.Equals(sName, DropDownStyleName, StringComparison.OrdinalIgnoreCase))
                        return true;
                }
            }
            catch { }

            // 2. Fallback: check if the cell has List validation (Type 3 = xlValidateList)
            try
            {
                if (singleCell.Validation != null && singleCell.Validation.Type == 3)
                {
                    return true;
                }
            }
            catch { }

            return false;
        }
        catch
        {
            return false;
        }
    }

    /// <summary>
    /// Template Edit Mode Dialog: Allows designing and editing validation lists, formulas, and styles.
    /// </summary>
    /// <summary>
    /// Template Edit Mode Dialog: Fully DPI-aware (scales cleanly at 125%, 150%, 200%).
    /// </summary>
    public static void InspectAndEditCell(dynamic cell)
    {
        try
        {
            string currentList = string.Empty;
            bool isMultiSelect = true;

            try
            {
                if (cell.Validation != null && cell.Validation.Type == 3)
                {
                    currentList = cell.Validation.Formula1 ?? string.Empty;
                    isMultiSelect = !cell.Validation.ShowError;
                }
            }
            catch { }

            // Default to a helpful prompt if empty
            if (string.IsNullOrWhiteSpace(currentList))
            {
                currentList = "=StaffNamesList";
            }

            string cellAddress = cell.Address?.ToString() ?? "Selected Cell";

            using var form = new Form
            {
                Text = $"Template Editor - {cellAddress}",
                FormBorderStyle = FormBorderStyle.FixedDialog,
                StartPosition = FormStartPosition.CenterScreen,
                MaximizeBox = false,
                MinimizeBox = false,
                TopMost = true,
                AutoScaleMode = AutoScaleMode.Dpi,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Padding = new Padding(12)
            };

            var tableLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                ColumnCount = 1,
                RowCount = 4
            };
            tableLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

            var lblInfo = new Label
            {
                Text = $"Target Style: '{DropDownStyleName}'\n" +
                       "Enter source list formula (e.g. =StaffNamesList, =DepartmentList, =LicenseTypesList):",
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 8)
            };

            var txtFormula = new TextBox
            {
                Text = currentList,
                Font = new Font("Segoe UI", 10f),
                Width = 340,
                Margin = new Padding(0, 0, 0, 10)
            };

            var chkMultiSelect = new CheckBox
            {
                Text = "Allow Multi-Selections (Checklist popup)",
                Checked = isMultiSelect,
                AutoSize = true,
                Margin = new Padding(0, 0, 0, 14)
            };

            var btnPanel = new FlowLayoutPanel
            {
                FlowDirection = FlowDirection.RightToLeft,
                AutoSize = true,
                AutoSizeMode = AutoSizeMode.GrowAndShrink,
                Dock = DockStyle.Fill,
                Margin = new Padding(0)
            };

            var btnCancel = new Button
            {
                Text = "Cancel",
                AutoSize = true,
                Padding = new Padding(10, 4, 10, 4),
                DialogResult = DialogResult.Cancel,
                Margin = new Padding(6, 0, 0, 0)
            };

            var btnSave = new Button
            {
                Text = "Apply to Cell",
                AutoSize = true,
                Padding = new Padding(10, 4, 10, 4),
                DialogResult = DialogResult.OK
            };

            btnSave.Click += (_, _) =>
            {
                string formula = txtFormula.Text.Trim();
                if (!string.IsNullOrEmpty(formula))
                {
                    // Ensure the style exists and is assigned
                    dynamic xlApp = ExcelDnaUtil.Application;
                    EnsureNamedStyleExists(xlApp, DropDownStyleName);
                    cell.Style = DropDownStyleName;

                    // Apply the validation formula
                    cell.Validation.Delete();
                    cell.Validation.Add(3, 1, 1, formula);
                    cell.Validation.IgnoreBlank = true;
                    cell.Validation.InCellDropdown = true;
                    cell.Validation.ShowError = !chkMultiSelect.Checked;
                }
                form.Close();
            };

            btnCancel.Click += (_, _) => form.Close();

            btnPanel.Controls.Add(btnCancel);
            btnPanel.Controls.Add(btnSave);

            tableLayout.Controls.Add(lblInfo, 0, 0);
            tableLayout.Controls.Add(txtFormula, 0, 1);
            tableLayout.Controls.Add(chkMultiSelect, 0, 2);
            tableLayout.Controls.Add(btnPanel, 0, 3);

            form.Controls.Add(tableLayout);
            form.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed editing cell validation: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Operational Mode Action: Shows checklist picker or standard dropdown picker.
    /// </summary>
    public static void ShowChecklistForActiveCell(dynamic targetCell)
    {
        try
        {
            string validationFormula = targetCell.Validation.Formula1;
            if (string.IsNullOrWhiteSpace(validationFormula)) return;

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

    private static void EnsureNamedStyleExists(dynamic xlApp, string styleName)
    {
        try
        {
            dynamic styles = xlApp.ActiveWorkbook.Styles;
            foreach (dynamic s in styles)
            {
                if (string.Equals(s.Name?.ToString(), styleName, StringComparison.OrdinalIgnoreCase))
                    return;
            }

            // Create style if absent
            dynamic newStyle = styles.Add(styleName);
            newStyle.IncludeNumber = false;
            newStyle.IncludeFont = true;
            newStyle.IncludeAlignment = true;
            newStyle.IncludeBorder = true;
            newStyle.IncludePatterns = true;

            // Visual properties for DropDown cells
            newStyle.Font.Name = "Segoe UI";
            newStyle.Font.Size = 10;
            newStyle.Interior.Color = ColorTranslator.ToOle(Color.FromArgb(240, 244, 250)); // Soft slate tint
        }
        catch { }
    }

    /// <summary>
    /// Executes when an operational user navigates to a DropDown-styled cell.
    /// Handles both multi-select checklist overlays and dynamic list refreshing.
    /// </summary>
    public static void OnCellActivatedForSelection(dynamic targetCell)
    {
        try
        {
            // If the cell was configured with multi-select (ShowError == false),
            // automatically display the checklist picker so the user doesn't need to double-click.
            if (targetCell.Validation != null && targetCell.Validation.ShowError == false)
            {
                ShowChecklistForActiveCell(targetCell);
            }
            else
            {
                // For single-select items, you can dynamically update targetCell.Validation.Formula1
                // based on the active row's data before the user clicks the native dropdown arrow.
                RefreshRowSpecificValidationList(targetCell);
            }
        }
        catch { }
    }

    private static void RefreshRowSpecificValidationList(dynamic targetCell)
    {
        // Example: If options depend on Staff_ID in column A of the same row:
        try
        {
            dynamic row = targetCell.Row;
            dynamic sheet = targetCell.Worksheet;
            string staffId = sheet.Cells[row, 1].Value2?.ToString() ?? string.Empty;

            // In-memory lookup: fetch valid items for this specific staff member
            // string validItemsCsv = "Van,Pickup,Lorry_1T";
            // targetCell.Validation.Modify(3, 1, 1, validItemsCsv);
        }
        catch { }
    }

    private static bool _isSanitizing = false;

    /// <summary>
    /// Validates single or multi-select cells against their validation formula.
    /// Flags invalid values, prompts the user, and strips out impossible tokens.
    /// </summary>
    public static void ValidateAndSanitizeCell(dynamic cell)
    {
        // Guard against recursive SheetChange calls when we modify the cell value
        if (_isSanitizing) return;

        try
        {
            if (cell == null || cell.Cells.Count > 1) return;
            if (!IsCustomDropdownCell(cell)) return;

            string cellText = cell.Value2?.ToString() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cellText)) return;

            // 1. Retrieve the list validation formula (e.g. "=DepartmentList" or "=Lookups!$M$2#")
            string validationFormula = string.Empty;
            try
            {
                if (cell.Validation != null && cell.Validation.Type == 3) // 3 = xlValidateList
                {
                    validationFormula = cell.Validation.Formula1 ?? string.Empty;
                }
            }
            catch { return; }

            if (string.IsNullOrWhiteSpace(validationFormula)) return;

            // 2. Evaluate the allowed list from Excel's engine
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic evaluatedRange = xlApp.Evaluate(validationFormula);
            if (evaluatedRange == null) return;

            var allowedOptions = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            if (evaluatedRange is object[,] rangeMatrix)
            {
                int rows = rangeMatrix.GetLength(0);
                int cols = rangeMatrix.GetLength(1);
                for (int r = 0; r < rows; r++)
                {
                    for (int c = 0; c < cols; c++)
                    {
                        string val = rangeMatrix[r, c]?.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(val))
                            allowedOptions.Add(val);
                    }
                }
            }
            else if (evaluatedRange is string singleVal && !string.IsNullOrWhiteSpace(singleVal))
            {
                allowedOptions.Add(singleVal.Trim());
            }

            if (allowedOptions.Count == 0) return;

            // 3. Tokenize entered values by comma
            var enteredTokens = cellText.Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                                        .Select(t => t.Trim())
                                        .Where(t => !string.IsNullOrEmpty(t))
                                        .Distinct(StringComparer.OrdinalIgnoreCase)
                                        .ToList();

            var validTokens = new List<string>();
            var invalidTokens = new List<string>();

            foreach (var token in enteredTokens)
            {
                if (allowedOptions.Contains(token))
                {
                    // Keep exact casing from allowed list if available
                    string matched = allowedOptions.First(opt => opt.Equals(token, StringComparison.OrdinalIgnoreCase));
                    validTokens.Add(matched);
                }
                else
                {
                    invalidTokens.Add(token);
                }
            }

            // 4. If any invalid tokens are detected, warn and sanitize
            if (invalidTokens.Count > 0)
            {
                _isSanitizing = true;
                try
                {
                    string cleanedString = string.Join(", ", validTokens);
                    cell.Value2 = cleanedString;

                    string cellAddress = cell.Address?.ToString() ?? "Cell";
                    string rejectedList = string.Join(", ", invalidTokens.Select(t => $"'{t}'"));

                    MessageBox.Show(
                        $"[Data Validation Error]\n\n" +
                        $"Location: {cellAddress}\n" +
                        $"Invalid Value(s): {rejectedList}\n\n" +
                        $"These values are not allowed in the current list and have been removed.\n" +
                        (validTokens.Count > 0
                            ? $"Retained valid entries: {cleanedString}"
                            : "The cell has been cleared."),
                        "Invalid Entry Rejected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
                finally
                {
                    _isSanitizing = false;
                }
            }
        }
        catch (Exception ex)
        {
            _isSanitizing = false;
            System.Diagnostics.Debug.WriteLine($"[ValidateAndSanitizeCell Error] {ex.Message}");
        }
    }
}