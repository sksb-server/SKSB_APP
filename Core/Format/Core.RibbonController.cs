namespace SKSB_App.Core.Format;

using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Services;
using SKSB_App.Core.WorkbookManager;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
        return @"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabFormat' label='Format'>
                <group id='grpCellFormat' label='Template &amp; Formatting'>
                  <!-- Exposed Template Edit Mode Toggle Button -->
                  <toggleButton id='btnToggleTemplateMode'
                                label='Template Edit Mode'
                                size='large'
                                imageMso='DesignMode'
                                getPressed='GetTemplateModePressed'
                                onAction='OnToggleTemplateModeClicked' />
                  <separator id='sep1' />
                  <!-- Custom Dropdown Style Button -->
                  <button id='btnFormatDropdown' 
                          label='Make Dropdown Cell' 
                          size='large' 
                          imageMso='DataValidation' 
                          onAction='OnFormatDropdownClicked' />
                  <!-- Data Purge Button (Only visible in Template Edit Mode) -->
                  <button id='btnPurgeTemplateData' 
                          label='Purge Template Data' 
                          size='large' 
                          imageMso='TableDeleteRows' 
                          getVisible='GetPurgeVisible' 
                          onAction='OnPurgeTemplateDataClicked' />
                </group>
                <group id='grpSync' label='Core Data'>
                  <!-- Exposed Manual Sync All Enums Button -->
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

    public void OnRibbonLoad(IRibbonUI ribbon)
    {
        _ribbonUi = ribbon;
    }

    public static void RefreshRibbonUI()
    {
        _ribbonUi?.InvalidateControl("btnToggleTemplateMode");
        _ribbonUi?.InvalidateControl("btnPurgeTemplateData");
    }

    public bool GetTemplateModePressed(IRibbonControl control)
    {
        return WorkbookSession.IsTemplateEditMode;
    }

    public bool GetPurgeVisible(IRibbonControl control)
    {
        return WorkbookSession.IsTemplateEditMode;
    }

    public void OnToggleTemplateModeClicked(IRibbonControl control, bool isPressed)
    {
        WorkbookSession.IsTemplateEditMode = isPressed;

        if (isPressed)
        {
            // 1. Unhide Lookups sheet
            WorkbookSession.SetLookupsVisibility(visible: true);

            // 2. Sync enums
            TypesToExcel.SyncAllCoreEnums();

            MessageBox.Show(
                "Template Edit Mode is now ENABLED.\n\n" +
                "• 'Lookups' worksheet is now VISIBLE.\n" +
                "• All Core Enums have been synchronized.\n" +
                "• Double-clicking dropdown cells will open the Template Inspector.",
                "Template Mode",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        else
        {
            // Hide Lookups sheet back down
            WorkbookSession.SetLookupsVisibility(visible: false);

            MessageBox.Show(
                "Template Edit Mode is now DISABLED.\n\n" +
                "• 'Lookups' worksheet is now HIDDEN.\n" +
                "• Switched to Operational Mode.",
                "Operational Mode",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        RefreshRibbonUI();
    }

    public void OnFormatDropdownClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                Commands.ApplyCustomDropdownStyle();
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed applying DropDown style: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnSyncEnumsClicked(IRibbonControl control)
    {
        try
        {
            TypesToExcel.SyncAllCoreEnums();
            MessageBox.Show(
                "All Core Enums have been synchronized to the 'Lookups' sheet successfully.\n\n" +
                "Press Ctrl + F3 to inspect the registered *List names.",
                "Sync Complete",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync trigger failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnPurgeTemplateDataClicked(IRibbonControl control)
    {
        var confirm = MessageBox.Show(
            "Clear all row data on this template?",
            "Purge Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        if (confirm == DialogResult.Yes)
        {
            ExcelAsyncUtil.QueueAsMacro(() => WorkbookSession.PurgeAllSheetsData());
        }
    }
}