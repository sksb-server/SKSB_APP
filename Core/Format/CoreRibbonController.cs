namespace SKSB_App.Core.UI;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;

[ComVisible(true)]
public class CoreFormatRibbon : ExcelRibbon
{
    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui'>
          <ribbon>
            <tabs>
              <!-- Create the new 'format' ribbon tab -->
              <tab id='tabFormat' label='format'>
                <group id='grpCellFormat' label='Custom Formats'>
                  <!-- Format Cell Button -->
                  <button id='btnFormatDropdown' 
                          label='Multi-Select Dropdown' 
                          size='large' 
                          imageMso='DataValidation' 
                          onAction='OnFormatDropdownClicked' />
                </group>
                <group id='grpSync' label='Core Data'>
                  <!-- Sync Enums Button -->
                  <button id='btnSyncCoreEnums' 
                          label='Sync All Enums' 
                          size='large' 
                          imageMso='Refresh' 
                          onAction='OnSyncEnumsClicked' />
                </group>
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnFormatDropdownClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                // Prompts user for the list source (e.g., =DaysOfWeekList)
                string listSource = Microsoft.VisualBasic.Interaction.InputBox(
                    "Enter the list source formula (e.g., =DaysOfWeekList or =Lookups!$A$2#):",
                    "Multi-Select Validation Source", "=DaysOfWeekList");

                if (!string.IsNullOrWhiteSpace(listSource))
                {
                    FormatCommands.ApplyCustomDropdownToSelection(listSource);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Formatting failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnSyncEnumsClicked(IRibbonControl control)
    {
        try
        {
            // Sync all core enums directly from the format ribbon
            TypesToExcel.Sync(
                typeof(Department),
                typeof(Shift),
                typeof(PayType),
                typeof(ContractStatus),
                typeof(LeaveType),
                typeof(Session),
                typeof(DayOfWeek)
            );

            MessageBox.Show("Core Enums Synchronised to 'Lookups' sheet successfully.",
                            "Sync Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync trigger failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}