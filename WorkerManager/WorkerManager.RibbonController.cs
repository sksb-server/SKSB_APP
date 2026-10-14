namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Types;
using SKSB_App.WorkerManager;
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
            // 1. Spawns Windows File Picker dialog
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

            // 2. Parse and filter workers using ERPBridge
            var workers = Importer.ImportWorkers(selectedFilePath);

            if (workers.Count == 0)
            {
                MessageBox.Show(
                    "No valid operational workers were found in the selected file.\n\n" +
                    "Ensure the file contains the Code, Name, Handphone, and Department columns.",
                    "Import Results",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            // 3. Auto-populate into the active Excel worksheet table via Excel-DNA macro queue
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SheetMapper.PopulateWorksheet(workers);
                MessageBox.Show(
                    $"Successfully imported and populated {workers.Count} active worker(s) from:\n{selectedFilePath}\n\n" +
                    "Review the table entries and click 'Save to System' when ready to sync.",
                    "SQL Payroll Import Success",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Import failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
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

    public void OnSaveClicked(IRibbonControl control)
    {
        var workerList = SheetMapper.ExtractWorkerRecords();
        if (workerList == null || workerList.Count == 0) return;

        Task.Run(async () =>
        {
            try
            {
                var (success, count, path) = await ServerSyncCommands.SaveWorkerAsync(workerList);
                if (success)
                {
                    MessageBox.Show(
                        $"Successfully saved {count} worker record(s).\n\n• Storage: {path}",
                        "Save Success",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    MessageBox.Show(
                        "Failed to persist workers to PostgreSQL or local storage.",
                        "Persistence Error",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Save Error: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
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
        {
            SheetMapper.PopulateWorksheet(records);
            MessageBox.Show($"Successfully loaded {records.Count} worker record(s).", "Load Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
        });
    }
}