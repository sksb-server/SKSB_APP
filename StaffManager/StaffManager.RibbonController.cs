namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ManagerRibbonBase
{
    // =========================================================================
    // 1. MODULE DESCRIPTOR
    // =========================================================================
    protected override ISaveContract Contract => Events.StaffSyncDescriptor;
    protected override string TabId => "tabStaffManager";
    protected override string TabLabel => "Staff Manager";

    protected override string PersistenceGroupXml => BuildPersistenceGroupXml(
        groupId: "grpStaffPersistence",
        groupLabel: "Staff Database Sync",
        loadId: "btnLoadStaff",
        loadScreentip: "Load Staff Records",
        loadSupertip: "Fetches all staff master records and licenses from PostgreSQL or offline storage.",
        saveId: "btnSaveStaff",
        saveScreentip: "Save Staff Records",
        saveSupertip: "Validates and commits employee profiles and licenses.");

    protected override string GetCustomGroupsXml() => @"
                <group id='grpErpIntegration' label='SQL Payroll Integration'>
                  <button id='btnImportStaffs'
                          label='Import from&#10;SQL Payroll'
                          size='large'
                          imageMso='ImportExcel'
                          screentip='Import from SQL Payroll'
                          supertip='Imports employee records from an exported SQL Payroll Excel or CSV file while retaining existing licenses.'
                          onAction='OnImportStaffsClicked' />
                </group>";

    protected override string EmptyLoadMessage => "No staff records found in database or local snapshot.";

    // =========================================================================
    // 2. EXCEL COMMANDS (must live on a non-generic type)
    // =========================================================================
    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro() => Events.RunDelayedAutoLoad();

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave() => Events.RunHotKeySave();

    // =========================================================================
    // 3. SQL PAYROLL IMPORT CALLBACK
    // =========================================================================
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
                var (existingStaffs, _, _) = SheetMapper.ExtractStaffRecordsWithDiagnostics();

                var cleanStaffs = TableRowManager.FilterDuplicates(
                    incomingRecords: importResult.ImportedStaffs,
                    keySelector: s => $"{s.StaffID}",
                    displaySelector: s => $"Staff ID {s.StaffID} ({s.StaffDetails.StaffName})",
                    entityLabel: "Staff Profile",
                    duplicateWarnings: out var duplicateNotices,
                    existingRecords: existingStaffs
                );

                if (cleanStaffs == null)
                {
                    SheetOperations.ShowStatusConfirmation("Staff import cancelled by user.");
                    return;
                }

                if (cleanStaffs.Count > 0)
                {
                    var existingLookup = (existingStaffs ?? new List<StaffRecord>())
                        .Where(w => w.StaffID > 0)
                        .GroupBy(w => w.StaffID)
                        .ToDictionary(g => g.Key, g => g.First());

                    foreach (var imported in cleanStaffs)
                    {
                        if (existingLookup.TryGetValue(imported.StaffID, out var current))
                        {
                            imported.VehicleLicenses = current.VehicleLicenses;
                            imported.SkillLicenses = current.SkillLicenses;
                        }
                    }

                    SheetMapper.PopulateWorksheet(cleanStaffs);
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
}
