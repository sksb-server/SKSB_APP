namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.Types;
using SKSB_App.Core.DialogBoxes;
using System;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;

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
            string fileNameOnly = System.IO.Path.GetFileName(selectedFilePath);

            // 1. Process file and gather metrics/errors
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

            // 2. Populate Excel table on the UI macro thread
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (importResult.ImportedWorkers.Count > 0)
                {
                    SheetMapper.PopulateWorksheet(importResult.ImportedWorkers);
                }

                // 3. Display summary dialog with 'View Skipped Errors' option
                using var summaryDialog = new ImportResultDialog(importResult, fileNameOnly);
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
        List<WorkerRecord> workerList = SheetMapper.ExtractWorkerRecords();
        if (workerList == null || workerList.Count == 0)
        {
            return;
        }

        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveWorkerAsync(workerList);
                if (success)
                {
                    // Update sheet with newly generated IDs & default values
                    ExcelAsyncUtil.QueueAsMacro(() => SheetMapper.PopulateWorksheet(workerList));

                    SaveNotificationDialog.ShowSuccess(
                        "Worker Manager",
                        $"Successfully saved {count} worker record(s).\n\n• Storage: {path}\n• All auto-generated IDs and default dates have been applied to your sheets."
                    );
                }
                else
                {
                    SaveNotificationDialog.ShowError(
                        "Worker Manager",
                        "Failed to persist workers to PostgreSQL or local storage."
                    );
                }
            }
            catch (Exception ex)
            {
                SaveNotificationDialog.ShowError("Worker Manager", $"Save Error: {ex.Message}");
            }
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
            Events.SyncWorkerLookupsData();
            MessageBox.Show(
                "Worker lookups refreshed successfully.\n\n• Lookups: Enums managed centrally in Core\n• Data: Live vehicle plates updated in 'Data' sheet",
                "Lookups Synced",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Sync failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}