namespace SKSB_App.Core.TemplateEditor;

using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.DropDown;

public static class Commands
{
    public const string DropDownStyleName = "DropDown";

    public static void ApplyCustomDropdownStyle()
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

    public static bool IsCustomDropdownCell(dynamic cell)
    {
        if (cell == null) return false;

        try
        {
            dynamic singleCell = cell.Cells[1, 1];
            if (singleCell.Style != null)
            {
                string currentStyleName = singleCell.Style.Name?.ToString() ?? string.Empty;
                if (string.Equals(currentStyleName, DropDownStyleName, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            return false;
        }
        catch
        {
            return false;
        }
    }

    public static void InspectAndEditCell(dynamic cell)
    {
        try
        {
            using var form = new ConfiguratorForm(cell);
            form.ShowDialog();
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed opening editor: {ex.Message}", "Template Editor", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public static string ResolveContextListFormula(dynamic targetCell)
    {
        try
        {
            dynamic sheet = targetCell.Worksheet;
            int col = targetCell.Column;
            string headerText = sheet.Cells[1, col].Value2?.ToString()?.Trim() ?? string.Empty;

            return headerText switch
            {
                "Worker Name" => "=AllWorkersTable[Worker Name]",
                "Worker ID" => "=AllWorkersTable[Worker ID]",
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
}