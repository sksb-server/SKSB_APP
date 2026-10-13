namespace SKSB_App.Core.TemplateEditor;

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
    /// Applies the DropDown style and binds the context-appropriate list formula based on the column header.
    /// </summary>
    public static void ApplyCustomDropdownStyle(string? explicitListSource = null)
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic selection = xlApp.Selection;

        if (selection == null) return;

        EnsureNamedStyleExists(xlApp, DropDownStyleName);

        // 1. Assign the custom Named Style
        selection.Style = DropDownStyleName;

        // 2. Resolve default list formula based on column header if not explicitly supplied
        string listSource = explicitListSource ?? ResolveContextListFormula(selection);

        try
        {
            selection.Validation.Delete();
            selection.Validation.Add(3, 1, 1, listSource); // 3 = xlValidateList, 1 = xlValidAlertStop, 1 = xlBetween
            selection.Validation.IgnoreBlank = true;
            selection.Validation.InCellDropdown = true;
            selection.Validation.ShowError = false; // Multi-select support
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - ApplyValidation Error] {ex.Message}");
        }
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

            // 1. Check style name
            try
            {
                dynamic style = singleCell.Style;
                if (style != null)
                {
                    string currentStyleName = style.Name?.ToString() ?? string.Empty;
                    if (string.Equals(currentStyleName, DropDownStyleName, StringComparison.OrdinalIgnoreCase))
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
    /// Automatically ensures that a cell styled with 'DropDown' has list validation attached.
    /// </summary>
    public static void EnsureDropdownValidationAttached(dynamic cell)
    {
        try
        {
            dynamic singleCell = cell.Cells[1, 1];

            bool hasValidation = false;
            try
            {
                if (singleCell.Validation != null && singleCell.Validation.Type == 3)
                {
                    hasValidation = true;
                }
            }
            catch { }

            if (!hasValidation)
            {
                string formula = ResolveContextListFormula(singleCell);
                singleCell.Validation.Delete();
                singleCell.Validation.Add(3, 1, 1, formula);
                singleCell.Validation.IgnoreBlank = true;
                singleCell.Validation.InCellDropdown = true;
                singleCell.Validation.ShowError = false;
            }
        }
        catch { }
    }

    /// <summary>
    /// Template Edit Mode Dialog: Allows designing and editing validation lists, formulas, and styles.
    /// </summary>
    public static void InspectAndEditCell(dynamic cell)
    {
        try
        {
            dynamic singleCell = cell.Cells[1, 1];
            string currentList = string.Empty;
            bool isMultiSelect = true;

            try
            {
                if (singleCell.Validation != null && singleCell.Validation.Type == 3)
                {
                    currentList = singleCell.Validation.Formula1 ?? string.Empty;
                    isMultiSelect = !singleCell.Validation.ShowError;
                }
            }
            catch { }

            // If empty, auto-resolve from the column header
            if (string.IsNullOrWhiteSpace(currentList))
            {
                currentList = ResolveContextListFormula(singleCell);
            }

            string cellAddress = singleCell.Address?.ToString() ?? "Selected Cell";

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
                       "Source list formula (e.g. =StaffNamesList, =LicenseTypesList, =LicenseClassList):",
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
                    dynamic xlApp = ExcelDnaUtil.Application;
                    EnsureNamedStyleExists(xlApp, DropDownStyleName);
                    singleCell.Style = DropDownStyleName;

                    singleCell.Validation.Delete();
                    singleCell.Validation.Add(3, 1, 1, formula);
                    singleCell.Validation.IgnoreBlank = true;
                    singleCell.Validation.InCellDropdown = true;
                    singleCell.Validation.ShowError = !chkMultiSelect.Checked;
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
            MessageBox.Show($"Failed editing cell validation: {ex.Message}", "Template Editor Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    /// <summary>
    /// Operational Mode Action: Shows checklist picker or standard dropdown picker.
    /// </summary>
    public static void ShowChecklistForActiveCell(dynamic targetCell)
    {
        try
        {
            dynamic singleCell = targetCell.Cells[1, 1];
            string validationFormula = string.Empty;

            try
            {
                if (singleCell.Validation != null && singleCell.Validation.Type == 3)
                {
                    validationFormula = singleCell.Validation.Formula1 ?? string.Empty;
                }
            }
            catch { return; }

            if (string.IsNullOrWhiteSpace(validationFormula)) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic validationRange = xlApp.Evaluate(validationFormula);

            if (validationRange == null) return;

            string[] availableOptions;
            if (validationRange is object[,] rangeValues)
            {
                availableOptions = rangeValues.Cast<object>()
                                              .Where(o => o != null)
                                              .Select(o => o.ToString()!.Trim())
                                              .Where(s => !string.IsNullOrEmpty(s))
                                              .ToArray();
            }
            else
            {
                return;
            }

            string currentText = singleCell.Value2?.ToString() ?? string.Empty;
            var selectedItems = currentText.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                                           .Select(d => d.Trim())
                                           .ToHashSet(StringComparer.OrdinalIgnoreCase);

            using var form = new Form
            {
                Text = "Select Options",
                FormBorderStyle = FormBorderStyle.FixedToolWindow,
                StartPosition = FormStartPosition.Manual,
                Width = 260,
                Height = 320,
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
                singleCell.Value2 = string.Join(", ", chosen);
                form.Close();
            };

            form.Controls.Add(checkList);
            form.Controls.Add(btnDone);
            form.ShowDialog();
        }
        catch { }
    }

    /// <summary>
    /// Inspects the column header above the cell to auto-determine the list formula.
    /// </summary>
    public static string ResolveContextListFormula(dynamic targetCell)
    {
        try
        {
            dynamic sheet = targetCell.Worksheet;
            int col = targetCell.Column;
            string headerText = sheet.Cells[1, col].Value2?.ToString()?.Trim() ?? string.Empty;

            return headerText switch
            {
                "Staff Name" => "=StaffNamesList",
                "License Type" or "License Types" => "=LicenseTypesList",
                "License_Class" or "License Class" => "=LicenseClassList",
                "Vehicle Type" => "=VehicleTypeList",
                "Status" => "=LicenseStatusList",
                "Department" => "=DepartmentList",
                "Shift" or "Preferred Shift" => "=ShiftList",
                "Pay Type" => "=PayTypeList",
                "Company" => "=CompanyNamesList",
                "Project Type" => "=ProjectTypeList",
                _ => "=DepartmentList"
            };
        }
        catch
        {
            return "=DepartmentList";
        }
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
}