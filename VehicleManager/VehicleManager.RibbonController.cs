namespace SKSB_App.VehicleManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ManagerRibbonBase
{
    // =========================================================================
    // 1. MODULE DESCRIPTOR
    // =========================================================================
    protected override ISaveContract Contract => Events.VehicleSyncDescriptor;
    protected override string TabId => "tabVehicleManager";
    protected override string TabLabel => "Vehicle Manager";

    protected override string PersistenceGroupXml => BuildPersistenceGroupXml(
        groupId: "grpVehiclePersistence",
        groupLabel: "Fleet Database Sync",
        loadId: "btnLoadVehicles",
        loadScreentip: "Load Vehicle Records",
        loadSupertip: "Fetches all vehicle master records and logs from PostgreSQL or offline storage.",
        saveId: "btnSaveVehicles",
        saveScreentip: "Save Vehicle Records",
        saveSupertip: "Validates and commits fleet specifications, maintenance records and insurance claims.");

    protected override string GetCustomGroupsXml() => @"
                <group id='grpFuelImport' label='Portal Integration'>
                  <button id='btnImportFuelStatement'
                          label='Import Fuel&#10;Statement'
                          size='large'
                          imageMso='ImportExcel'
                          screentip='Import Petronas / Shell Monthly Statement'
                          supertip='Parses monthly PDF invoices or portal Excel files and appends fuel logs into OdometersTable.'
                          onAction='OnImportFuelStatementClicked' />
                  <button id='btnSortAndHealOdometers'
                          label='Sort &amp; Match&#10;Odometers'
                          size='large'
                          imageMso='SortAscendingExcel'
                          screentip='Auto-Sort &amp; Reconcile OdometersTable'
                          supertip='Sorts all odometer logs chronologically, reads Remarks so edits are never applied twice, normalizes machinery engine hours (0.1 hr), matches cross-card fueling trajectories across the fleet, flags ambiguous or invalid odometers, and recomputes Odometer Before.'
                          onAction='OnSortAndHealOdometersClick' />
                </group>";

    // Original behaviour: an empty fetch still populates (clears) the worksheet.
    protected override bool PopulateWhenEmpty => true;
    protected override bool ShowEmptyLoadDialog => false;

    // =========================================================================
    // 2. EXCEL COMMANDS (must live on a non-generic type)
    // =========================================================================
    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro() => Events.RunDelayedAutoLoad();

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave() => Events.RunHotKeySave();

    // =========================================================================
    // 3. FUEL STATEMENT IMPORT CALLBACK
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
    // 4. SORT & MATCH ODOMETERS CALLBACK
    // =========================================================================
    public void OnSortAndHealOdometersClick(IRibbonControl control)
    {
        // COM work must run on Excel's main thread. The engine restores ScreenUpdating / EnableEvents
        // before returning, so the summary dialogs below are safe to show.
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                var result = SheetMapper.SortAndHealOdometers();

                if (result.RowsSorted == 0)
                {
                    MessageBox.Show("OdometersTable has no rows to sort.", "Sort & Match Odometers",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string summary =
                    $"  • Total Rows Sorted: {result.RowsSorted}\n" +
                    $"  • Vehicles Matched by Odometer Trajectory: {result.VehiclesMatched}\n" +
                    $"  • Ambiguous / Unresolved Plates Flagged for Review: {result.FlaggedPlates}\n" +
                    $"  • Machinery Hour Readings Normalized (÷ 10): {result.HoursNormalized}\n" +
                    $"  • Invalid / Dummy Odometers Zeroed: {result.InvalidZeroed}";

                if (result.FlaggedPlates > 0)
                    summary += "\n\nRows tagged [Unrecognized Plate: ...] or [Ambiguous Odometer ...] need a valid Vehicle Plate before the workbook can be saved.";

                MessageBox.Show(summary, "Sort & Match Odometers Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Sort & Match Odometers failed:\n{ex.Message}", "Odometer Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }
}
