namespace SKSB_App.Core;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
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
                // Specify the reference cell whose style and validation list you want to adopt
                // Example: "Lookups!$A$2" or a styled template cell
                Commands.ApplyStyleAndDropdownToSelection("Lookups!$A$2");

                MessageBox.Show("Selected cell(s) formatted with reference style and multi-select validation.",
                                "Format Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            // Automatically discovers and syncs every public enum in Core.Types
            TypesToExcel.SyncAllCoreEnums(typeof(Department));

            MessageBox.Show("All Core Enums synchronised to 'Lookups' successfully.",
                            "Sync Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync trigger failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}