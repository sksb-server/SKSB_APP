namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Erp;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.ThumbprintSync;
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
                </group>
                <group id='grpThumbprintTerminal' label='Thumbprint Terminal'>
                  <button id='btnPushStaffToTerminal'
                          label='Push ID &amp; Name&#10;to Controller'
                          size='large'
                          imageMso='AccessRefreshAllLists'
                          screentip='Push Staff ID &amp; Name to Controller'
                          supertip='Compares Staff IDs on the thumbprint controller and updates matching users names (preserving fingerprint templates and roles) after confirmation.'
                          onAction='OnPushStaffToTerminalClicked' />
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

            StaffImportResult importResult;
            using (LoadingOverlay.BeginScope("Staff Manager — Importing SQL Payroll", $"Reading {fileNameOnly}..."))
            {
                importResult = Importer.ImportStaffs(selectedFilePath);
            }

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
                using (LoadingOverlay.BeginScope("Staff Manager — Importing SQL Payroll", "Merging staff profiles and licenses..."))
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

    // =========================================================================
    // 4. THUMBPRINT TERMINAL PUSH CALLBACK
    // =========================================================================
    public void OnPushStaffToTerminalClicked(IRibbonControl control)
    {
        const string Title = "Push ID & Name to Controller";
        const int MaxNameLen = 23;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            Dictionary<string, string> staffById;
            try
            {
                var (staffs, _, _) = SheetMapper.ExtractStaffRecordsWithDiagnostics();
                staffById = new Dictionary<string, string>();
                foreach (var s in staffs ?? new List<StaffRecord>())
                {
                    if (s.StaffID <= 0) continue;
                    string name = !string.IsNullOrWhiteSpace(s.StaffDetails?.StaffShortName)
                        ? s.StaffDetails!.StaffShortName.Trim()
                        : (s.StaffDetails?.StaffName ?? "").Trim();
                    if (name.Length == 0) continue;
                    if (name.Length > MaxNameLen) name = name.Substring(0, MaxNameLen);
                    staffById[s.StaffID.ToString()] = name;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed to read staff records for terminal push", ex);
                MessageBox.Show($"Failed to read staff records: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Task.Run(async () =>
            {
                try
                {
                    var terminalUsers = await DeviceSyncManager.FetchTerminalUsersAsync();

                    var updates = new List<TerminalUserUpdate>();
                    int alreadySynced = 0;
                    int unmatched = 0;
                    foreach (var u in terminalUsers)
                    {
                        if (!staffById.TryGetValue(u.EnrollNumber, out var targetName))
                            unmatched++;
                        else if (string.Equals((u.Name ?? "").Trim(), targetName, StringComparison.Ordinal))
                            alreadySynced++;
                        else
                            updates.Add(new TerminalUserUpdate(u.Uid, u.EnrollNumber, targetName));
                    }

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        if (updates.Count == 0)
                        {
                            MessageBox.Show(
                                $"Nothing to update.\n\nAlready synced: {alreadySynced}\nNot in staff list: {unmatched}",
                                Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        var confirm = MessageBox.Show(
                            $"Names to update: {updates.Count}\nAlready synced: {alreadySynced}\nNot in staff list: {unmatched}\n\n" +
                            $"Push the name updates to {DeviceSyncManager.DefaultIp}?\n" +
                            "Fingerprint templates and roles are preserved.",
                            Title, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                        if (confirm != DialogResult.Yes) return;

                        Task.Run(async () =>
                        {
                            try
                            {
                                int updated = await DeviceSyncManager.ApplyTerminalUsersAsync(updates);
                                ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                                    $"Updated {updated} user name(s) on the controller.",
                                    Title, MessageBoxButtons.OK, MessageBoxIcon.Information));
                            }
                            catch (Exception ex)
                            {
                                AppLogger.Error("Terminal name push failed", ex);
                                ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                                    $"Push failed: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error));
                            }
                        });
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Error("Terminal fetch failed", ex);
                    ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                        $"Could not read users from controller: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error));
                }
            });
        });
    }
}
