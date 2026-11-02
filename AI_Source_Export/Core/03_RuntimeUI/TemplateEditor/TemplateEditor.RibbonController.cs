namespace SKSB_App.Core.RuntimeUI.TemplateEditor;

using System;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.TemplateEditor.RulesAudit;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
#if DEBUG
        string debugGroupXml = @"
                <group id='grpGlobalDebug' label='Diagnostics Control'>
                  <toggleButton id='btnToggleMasterDebug'
                                label='Debug Logs Active'
                                size='large'
                                imageMso='VisualBasic'
                                getPressed='GetMasterDebugPressed'
                                onAction='OnToggleMasterDebugClicked' />
                </group>";
#else
        string debugGroupXml = string.Empty;
#endif

        return $@"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabTemplateEditor' label='Template Editor'>
                <group id='grpTemplateDesign' label='Template Design &amp; Validation'>
                  <toggleButton id='btnToggleTemplateMode'
                                label='Template Edit Mode'
                                size='large'
                                imageMso='DesignMode'
                                getPressed='GetTemplateModePressed'
                                onAction='OnToggleTemplateModeClicked' />
                  <separator id='sep1' />
                  <button id='btnFormatDropdown'
                            label='Make Dropdown Cell'
                            size='large'
                            imageMso='DataValidation'
                            onAction='OnFormatDropdownClicked' />
                  <button id='btnCopyMetadata'
                            label='Copy Config'
                            size='normal'
                            imageMso='Copy'
                            onAction='OnCopyMetadataClicked' />
                  <button id='btnPasteMetadata'
                            label='Paste Config'
                            size='normal'
                            imageMso='Paste'
                            onAction='OnPasteMetadataClicked' />
                  <separator id='sep2' />
                  <button id='btnExportRules'
                            label='Export Rules Audit'
                            size='large'
                            imageMso='ExportExcel'
                            onAction='OnExportRulesClicked' />
                  <button id='btnImportRules'
                            label='Import Rules Audit'
                            size='large'
                            imageMso='ImportExcel'
                            onAction='OnImportRulesClicked' />
                  <button id='btnPurgeTemplateData'
                            label='Purge Template Data'
                            size='large'
                            imageMso='TableDeleteRows'
                            getVisible='GetPurgeVisible'
                            onAction='OnPurgeTemplateDataClicked' />
                </group>
                <group id='grpSync' label='Core Lookups &amp; Data'>
                  <button id='btnSyncCoreEnums'
                            label='Sync All Enums'
                            size='large'
                            imageMso='Refresh'
                            onAction='OnSyncEnumsClicked' />
                  <button id='btnSyncCoreData'
                            label='Sync All Data'
                            size='large'
                            imageMso='AccessRefreshAllLists'
                            onAction='OnSyncDataClicked' />
                </group>
                {debugGroupXml}
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
        try
        {
            _ribbonUi?.InvalidateControl("btnToggleTemplateMode");
            _ribbonUi?.InvalidateControl("btnPurgeTemplateData");
#if DEBUG
            _ribbonUi?.InvalidateControl("btnToggleMasterDebug");
#endif
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[RibbonController.RefreshRibbonUI Error] {ex.Message}");
        }
    }

    public bool GetTemplateModePressed(IRibbonControl control) => SheetOperations.IsTemplateEditMode;
    public bool GetPurgeVisible(IRibbonControl control) => SheetOperations.IsTemplateEditMode;

#if DEBUG
    public bool GetMasterDebugPressed(IRibbonControl control) => DebugDialog.DebugEnabled;

    public void OnToggleMasterDebugClicked(IRibbonControl control, bool isPressed)
    {
        DebugDialog.DebugEnabled = isPressed;
        _ribbonUi?.InvalidateControl("btnToggleMasterDebug");
        MessageBox.Show(
            isPressed
                ? "Debug diagnostics are now ENABLED across all operations."
                : "Debug diagnostics are now MUTED. No popup inspection logs will appear.",
            "Debug Log Switch",
            MessageBoxButtons.OK,
            isPressed ? MessageBoxIcon.Information : MessageBoxIcon.Warning
        );
    }
#endif

    public void OnToggleTemplateModeClicked(IRibbonControl control, bool isPressed)
    {
        SheetOperations.IsTemplateEditMode = isPressed;
        SheetOperations.SetAuxiliarySheetsVisibility(isPressed);

        if (isPressed)
        {
            EnumList.SyncAllCoreEnums();
            DataList.SyncAllTemplateData();
        }

        RefreshRibbonUI();
    }

    public void OnFormatDropdownClicked(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.ApplyDropdownStyle);
    public void OnCopyMetadataClicked(IRibbonControl control) => Commands.CopySelectedCellMetadata();
    public void OnPasteMetadataClicked(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(Commands.PasteSelectedCellMetadata);
    public void OnExportRulesClicked(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(RulesAuditExporter.ExportAuditReport);
    public void OnImportRulesClicked(IRibbonControl control) => ExcelAsyncUtil.QueueAsMacro(RulesAuditImporter.ImportAuditReport);
    public void OnPurgeTemplateDataClicked(IRibbonControl control)
    {
        var confirm = MessageBox.Show(
            "Clear all operational data rows on this template?\n\n(Headers, styles, and the 'Lookups' and 'Data' sheets will be preserved)",
            "Template Editor - Purge Template Data",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning
        );

        if (confirm == DialogResult.Yes)
        {
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SheetOperations.PurgeAllSheetsData();
                MessageBox.Show("All sheet data has been purged. Template is now clean.",
                                "Template Editor", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }
    }

    public void OnSyncEnumsClicked(IRibbonControl control) => EnumList.SyncAllCoreEnums();
    public void OnSyncDataClicked(IRibbonControl control)
    {
        DataList.SyncAllTemplateData();
        MessageBox.Show("All entity data synchronized to 'Data' sheet.",
                        "Template Editor", MessageBoxButtons.OK, MessageBoxIcon.Information);
    }
}