namespace SKSB_App.WorkerManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.WorkbookHelper;

[ComVisible(true)]
public class RibbonController : ExcelRibbon
{
    private static IRibbonUI? _ribbonUi;

    public override string GetCustomUI(string ribbonId)
    {
#if DEBUG
        string debugGroupXml = @"
                <group id='grpDebugControls' label='Debug Controls'>
                  <toggleButton id='btnToggleOffline' 
                                label='Simulate Server Offline' 
                                size='large' 
                                imageMso='AccessSetOfflineStatus' 
                                getPressed='GetSimulateOfflinePressed'
                                onAction='OnToggleOfflineClicked' />
                </group>";
#else
        string debugGroupXml = string.Empty;
#endif

        return $@"
        <customUI xmlns='http://schemas.microsoft.com/office/2009/07/customui' onLoad='OnRibbonLoad'>
          <ribbon>
            <tabs>
              <tab id='tabWorkerManager' label='Worker Manager'>
                <group id='grpWorkerPersistence' label='Worker Database Sync'>
                  <button id='btnSaveWorker' 
                          label='Save to System' 
                          size='large' 
                          imageMso='SaveObject' 
                          onAction='OnSaveClicked' />
                  <button id='btnLoadWorker' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          onAction='OnLoadClicked' />
                  <button id='btnSyncLookups' 
                          label='Sync Lookups' 
                          size='large' 
                          imageMso='GroupCustomAttributes' 
                          onAction='OnSyncLookupsClicked' />
                </group>
                <group id='grpErpIntegration' label='SQL Payroll Integration'>
                  <button id='btnImportWorkers' 
                          label='Import Workers from SQL Payroll' 
                          size='large' 
                          imageMso='ImportExcel' 
                          onAction='OnImportWorkersClicked' />
                </group>
                {debugGroupXml}
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbonUi = ribbon;

    public void OnImportWorkersClicked(IRibbonControl control)
    {
        try
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "Select SQL Payroll Export File (Employee.xlsx)",
                Filter = "Excel Files (*.xlsx;*.xls)|*.xlsx;*.xls|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                CheckFileExists = true,
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            string selectedFilePath = openFileDialog.FileName;
            string fileNameOnly = Path.GetFileName(selectedFilePath);

            var importResult = Importer.ImportWorkers(selectedFilePath);

            if (importResult.ImportedWorkers.Count == 0 && importResult.SkippedErrors.Count == 0)
            {
                MessageBox.Show(
                    "No operational project workers were found in the selected file.\n\n" +
                    "Ensure the file contains the Code, Name, Handphone/Mobile, and Department columns.",
                    "Import Results",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (importResult.ImportedWorkers.Count > 0)
                {
                    SheetMapper.PopulateWorksheet(importResult.ImportedWorkers);
                }

                using var summaryDialog = new GenericImportResult(importResult.ToGenericReport(), fileNameOnly);
                summaryDialog.ShowDialog();
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Import failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            EventHooks.Save(Events.WorkerSyncDescriptor);
        });
    }

    public void OnLoadClicked(IRibbonControl control)
    {
        var records = ServerSyncCommands.FetchWorkerRecords();
        if (records == null || records.Count == 0)
        {
            MessageBox.Show("No worker records found in storage.", "Load Workers", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        ExcelAsyncUtil.QueueAsMacro(() =>
            SheetMapper.PopulateWorksheet(records)
        );
    }

    public void OnSyncLookupsClicked(IRibbonControl control)
    {
        try
        {
            Events.SyncWorkerLookupsData(isInteractive: true);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

#if DEBUG
    public bool GetSimulateOfflinePressed(IRibbonControl control)
        => EventHooks.IsSimulatedOffline();

    public void OnToggleOfflineClicked(IRibbonControl control, bool pressed)
        => EventHooks.ToggleSimulatedOffline(_ribbonUi, control.Id, pressed);
#endif
}