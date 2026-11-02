namespace SKSB_App.StaffManager;

using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;

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
                  <toggleButton id='btnToggleMasterDebug'
                                label='Debug Logs Active'
                                size='large'
                                imageMso='TraceDependents'
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
              <tab id='tabStaffManager' label='Staff Manager'>
                <group id='grpStaffPersistence' label='Staff Database Sync'>
                  <button id='btnLoadStaff' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          screentip='Load Staff Records'
                          supertip='Fetches all staff master records and licenses from PostgreSQL or offline storage.'
                          onAction='OnLoadClicked' />
                  <button id='btnSaveStaff' 
                          label='Save to System' 
                          size='large' 
                          imageMso='FileSave' 
                          screentip='Save Staff Records'
                          supertip='Validates and commits employee profiles and licenses.'
                          onAction='OnSaveClicked' />
                </group>

                <group id='grpErpIntegration' label='SQL Payroll Integration'>
                  <button id='btnImportStaffs' 
                          label='Import from SQL Payroll' 
                          size='large' 
                          imageMso='ImportExcel' 
                          screentip='Import from SQL Payroll'
                          supertip='Imports employee records from an exported SQL Payroll Excel or CSV file while retaining existing licenses.'
                          onAction='OnImportStaffsClicked' />
                </group>

                <group id='grpRowDirection' label='Row Management'>
                  <toggleButton id='btnToggleInsertTop' 
                                label='Insert at Top' 
                                size='large' 
                                imageMso='TableInsertRowsAbove' 
                                screentip='Toggle Top/Bottom Auto-Insert'
                                supertip='When enabled, typing in the first data row creates a new row at the top. When disabled, rows are appended at the bottom.'
                                getPressed='GetInsertAtTopPressed'
                                onAction='OnToggleInsertAtTopClicked' />
                  <button id='btnCleanEmptyRows'
                          label='Trim Empty Rows'
                          size='large'
                          imageMso='TableDeleteRows'
                          screentip='Trim Redundant Empty Rows'
                          supertip='Scans all tables on the active sheet and trims extra blank rows, leaving exactly one empty entry row.'
                          onAction='OnCleanEmptyRowsClicked' />
                </group>

                {debugGroupXml}
              </tab>
            </tabs>
          </ribbon>
        </customUI>";
    }

    public void OnRibbonLoad(IRibbonUI ribbon) => _ribbonUi = ribbon;

    public void OnLoadClicked(IRibbonControl control)
    {
        if (!Guard.IsStaffWorkbook()) return;

        try
        {
            Events.SyncStaffLookupsData(isInteractive: false);

            var records = ServerSyncCommands.FetchStaffRecords();
            if (records == null || records.Count == 0)
            {
                MessageBox.Show("No staff records found in database or local snapshot.",
                    "Load Staffs", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                SheetMapper.PopulateWorksheet(records);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed loading staff records:\n{ex.Message}", "Load Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        if (!Guard.IsStaffWorkbook()) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SaveEntry.Save(Events.StaffSyncDescriptor);
        });
    }

    public void OnImportStaffsClicked(IRibbonControl control)
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

            var importResult = Importer.ImportStaffs(selectedFilePath);

            if (importResult.ImportedStaffs.Count == 0 && importResult.SkippedErrors.Count == 0)
            {
                MessageBox.Show(
                    "No operational staff records were found in the selected file.\n\n" +
                    "Ensure the file contains the Code, Name, Handphone/Mobile, and Department columns.",
                    "Import Results",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning
                );
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                if (importResult.ImportedStaffs.Count > 0)
                {
                    var (existingStaffs, _, _) = SheetMapper.ExtractStaffRecordsWithDiagnostics();
                    var existingLookup = (existingStaffs ?? new List<StaffRecord>()).ToDictionary(w => w.StaffID);

                    foreach (var imported in importResult.ImportedStaffs)
                    {
                        if (existingLookup.TryGetValue(imported.StaffID, out var current))
                        {
                            imported.VehicleLicenses = current.VehicleLicenses;
                            imported.SkillLicenses = current.SkillLicenses;
                        }
                    }

                    SheetMapper.PopulateWorksheet(importResult.ImportedStaffs);
                }

                using var summaryDialog = new ImportResultDialog(importResult.ToGenericReport(), fileNameOnly);
                summaryDialog.ShowDialog();
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Import failed: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    public bool GetInsertAtTopPressed(IRibbonControl control)
        => TableRowManager.IsInsertAtTopEnabled;

    public void OnToggleInsertAtTopClicked(IRibbonControl control, bool isPressed)
    {
        TableRowManager.IsInsertAtTopEnabled = isPressed;
        _ribbonUi?.InvalidateControl(control.Id);

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws != null)
                {
                    TableRowManager.CleanActiveSheetTables(ws);
                }
            }
            catch { }

            SheetOperations.ShowStatusConfirmation(
                isPressed
                    ? "Insert at TOP active: Spawns at top, redundant bottom rows trimmed."
                    : "Insert at BOTTOM active: Spawns at bottom, redundant top rows trimmed."
            );
        });
    }

    public void OnCleanEmptyRowsClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                int tablesCleaned = TableRowManager.CleanActiveSheetTables(ws);
                SheetOperations.ShowStatusConfirmation(
                    $"Cleaned empty rows across {tablesCleaned} table(s). Exactly 1 empty row retained."
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed cleaning empty rows:\n{ex.Message}", "Row Cleanup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

#if DEBUG
    public bool GetSimulateOfflinePressed(IRibbonControl control)
        => ServerEnvironment.SimulateOffline;

    public void OnToggleOfflineClicked(IRibbonControl control, bool pressed)
    {
        ServerEnvironment.SimulateOffline = pressed;
        _ribbonUi?.InvalidateControl(control.Id);
        MessageBox.Show(
            pressed
                ? "Server is now SIMULATED OFFLINE.\nAll PostgreSQL database operations will fail immediately to test offline local stashing."
                : "Server is now ONLINE.\nNormal PostgreSQL database connectivity restored.",
            "Debug Server Mode",
            MessageBoxButtons.OK,
            pressed ? MessageBoxIcon.Warning : MessageBoxIcon.Information
        );
    }

    public bool GetMasterDebugPressed(IRibbonControl control)
        => DebugDialog.DebugEnabled;

    public void OnToggleMasterDebugClicked(IRibbonControl control, bool pressed)
    {
        DebugDialog.DebugEnabled = pressed;
        _ribbonUi?.InvalidateControl(control.Id);
        MessageBox.Show(
            pressed
                ? "Debug diagnostics are now ENABLED across all operations."
                : "Debug diagnostics are now MUTED.",
            "Debug Log Switch",
            MessageBoxButtons.OK,
            pressed ? MessageBoxIcon.Information : MessageBoxIcon.Warning
        );
    }
#endif
}