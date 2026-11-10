namespace SKSB_App.VehicleManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;

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
              <tab id='tabVehicleManager' label='Vehicle Manager'>
                <group id='grpVehiclePersistence' label='Fleet Database Sync'>
                  <button id='btnLoadVehicles' 
                          label='Load from System' 
                          size='large' 
                          imageMso='Refresh' 
                          screentip='Load Vehicle Records'
                          supertip='Fetches all vehicle master records and logs from PostgreSQL or offline storage.'
                          onAction='OnLoadClicked' />
                  <button id='btnSaveVehicles' 
                          label='Save to System' 
                          size='large' 
                          imageMso='FileSave' 
                          screentip='Save Vehicle Records'
                          supertip='Validates and commits fleet specifications and maintenance records.'
                          onAction='OnSaveClicked' />
                </group>

                <!-- PORTAL INTEGRATION GROUP -->
                <group id='grpFuelImport' label='Portal Integration'>
                  <button id='btnImportFuelStatement'
                          label='Import Fuel&#10;Statement'
                          size='large'
                          imageMso='ImportExcel'
                          screentip='Import Petronas / Shell Monthly Statement'
                          supertip='Parses monthly PDF invoices or portal Excel files and appends fuel logs into OdometersTable.'
                          onAction='OnImportFuelStatementClicked' />
                </group>

                <!-- ROW MANAGEMENT GROUP -->
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
  <button id='btnCleanDuplicateRows'
          label='Remove Duplicates'
          size='large'
          imageMso='RemoveDuplicates'
          screentip='Remove Duplicate Rows'
          supertip='Scans all tables on the active sheet and removes identical duplicate rows, keeping the first occurrence.'
          onAction='OnCleanDuplicateRowsClicked' />
  <button id='btnDeleteTableRow'
          label='Delete Row'
          size='large'
          imageMso='TableRowsDelete'
          screentip='Delete Row'
          supertip='Deletes selected row(s) in absolute master tables. Deletion is blocked on additive transaction tables.'
          onAction='OnDeleteTableRowClicked' />
<toggleButton id='btnToggleTotalRow' 
              label='Total Row' 
              size='large' 
              imageMso='AutoSum' 
              screentip='Toggle Total Row'
              supertip='Shows or hides the summary total row at the bottom of the active table.'
              getPressed='GetTotalRowPressed' 
              onAction='OnToggleTotalRowClicked' />
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
        dynamic xlApp = ExcelDnaUtil.Application;
        bool origUpdating = true;
        try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

        try
        {
            var serverData = ServerSyncCommands.FetchVehicleRecords();
            if (serverData != null)
            {
                SheetMapper.PopulateWorksheet(serverData);
            }
        }
        finally
        {
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    public void OnSaveClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            SaveEntry.Save(Events.VehicleSyncDescriptor);
        });
    }

    // =========================================================================
    // FUEL STATEMENT IMPORT CALLBACK
    // =========================================================================
    public void OnImportFuelStatementClicked(IRibbonControl control)
    {
        try
        {
            using var openFileDialog = new OpenFileDialog
            {
                Title = "Select PETRONAS SmartPay or Shell Fleet Hub Statement",
                Filter = "All Statements (*.pdf;*.xlsx;*.xls;*.csv)|*.pdf;*.xlsx;*.xls;*.csv|PDF Invoices (*.pdf)|*.pdf|Excel / CSV (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|All Files (*.*)|*.*",
                FilterIndex = 1,
                RestoreDirectory = true,
                CheckFileExists = true,
                Multiselect = false
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
                return;

            string selectedFilePath = openFileDialog.FileName;
            string fileNameOnly = Path.GetFileName(selectedFilePath);

            // Queue as Excel Macro so COM worksheet manipulation runs safely on the UI thread
            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                var (imported, skipped, totalAmount, provider) = SheetMapper.ImportFuelStatement(selectedFilePath);

                if (imported == 0 && skipped == 0)
                {
                    MessageBox.Show(
                        $"No valid fuel transactions could be parsed from:\n{fileNameOnly}\n\n" +
                        "Ensure the file is an original PDF invoice or portal export from PETRONAS SmartPay or Shell Fleet Hub.",
                        "Import Unsuccessful",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                string summary =
                    $"Provider Detected : {provider}\n" +
                    $"Source File       : {fileNameOnly}\n\n" +
                    $"  • New Logs Appended : {imported}\n" +
                    $"  • Duplicates Skipped: {skipped}\n" +
                    $"  • Net Fuel Expense  : RM {totalAmount:N2}\n\n" +
                    "All records have been written to 'OdometersTable' via TableMapper.";

                MessageBox.Show(summary, "Fuel Statement Import Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Fuel import failed:\n{ex.Message}", "Import Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // =========================================================================
    // ROW MANAGEMENT CALLBACKS
    // =========================================================================
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

    public void OnCleanDuplicateRowsClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                bool origUpdating = true;
                try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }

                int totalDuplicatesRemoved = 0;
                var warnings = new System.Collections.Generic.List<string>();

                try
                {
                    SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                    {
                        foreach (dynamic table in ws.ListObjects)
                        {
                            totalDuplicatesRemoved += TableRowManager.CleanTableDuplicateRows(table, warnings);
                        }
                    }));
                }
                finally
                {
                    try { xlApp.ScreenUpdating = origUpdating; } catch { }
                }

                if (totalDuplicatesRemoved > 0)
                {
                    SheetOperations.ShowStatusConfirmation($"Removed {totalDuplicatesRemoved} duplicate row(s).");
                    MessageBox.Show(
                        $"Successfully removed {totalDuplicatesRemoved} duplicate row(s) from the active sheet.\n\n" +
                        string.Join("\n", warnings),
                        "Duplicates Removed",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
                else
                {
                    SheetOperations.ShowStatusConfirmation("No duplicate rows found on the active sheet.");
                    MessageBox.Show(
                        "No duplicate rows found in any tables on the active sheet.",
                        "Clean Duplicates",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Failed to remove duplicate rows:\n{ex.Message}", "Cleanup Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }

    public void OnDeleteTableRowClicked(IRibbonControl control)
    {
        ExcelAsyncUtil.QueueAsMacro(TableRowManager.DeleteSelectedRowCommand);
    }

    public bool GetTotalRowPressed(IRibbonControl control)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic table = xlApp.ActiveCell?.ListObject;
            if (table != null)
            {
                return (bool)table.ShowTotals;
            }
        }
        catch { }
        return false;
    }

    public void OnToggleTotalRowClicked(IRibbonControl control, bool isPressed)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                dynamic ws = xlApp.ActiveWorkbook?.ActiveSheet;
                if (ws == null) return;

                dynamic table = xlApp.ActiveCell?.ListObject;
                if (table == null && ws.ListObjects.Count > 0)
                {
                    table = ws.ListObjects[1]; // Fallback to first table on sheet
                }
                if (table == null) return;

                SheetProtections.ExecuteWithBypass((object)ws, new Action(() =>
                {
                    table.ShowTotals = isPressed;

                    if (isPressed)
                    {
                        // 1. MUST unlock total row so native header sorting isn't blocked
                        table.TotalsRowRange.Locked = false;

                        // 2. Configure default calculation for Cost columns
                        int colCount = (int)table.ListColumns.Count;
                        for (int c = 1; c <= colCount; c++)
                        {
                            dynamic col = table.ListColumns[c];
                            string header = col.Name?.ToString() ?? string.Empty;

                            if (header.IndexOf("Cost", StringComparison.OrdinalIgnoreCase) >= 0 ||
                                header.IndexOf("Amount", StringComparison.OrdinalIgnoreCase) >= 0)
                            {
                                col.TotalsCalculation = 1; // xlTotalsCalculationSum
                                try { col.Total.NumberFormat = "#,##0.00"; } catch { }
                            }
                        }
                    }
                }));

                _ribbonUi?.InvalidateControl(control.Id);
                SheetOperations.ShowStatusConfirmation(isPressed ? "Total row enabled." : "Total row hidden.");
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed toggling Total Row", ex);
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