namespace SKSB_App.VehicleManager;

using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.Progress;
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
                <group id='grpFuelImport' getLabel='GetCustomLabel'>
                  <button id='btnImportFuelStatement'
                          getLabel='GetCustomLabel'
                          size='large'
                          imageMso='ImportExcel'
                          getScreentip='GetCustomScreentip'
                          getSupertip='GetCustomSupertip'
                          onAction='OnImportFuelStatementClicked' />
                  <button id='btnSortAndHealOdometers'
                          getLabel='GetCustomLabel'
                          size='large'
                          imageMso='SortAscendingExcel'
                          getScreentip='GetCustomScreentip'
                          getSupertip='GetCustomSupertip'
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
                Title = Loc.T("vehicle.fuelImport.dialogTitle"),
                Filter = $"{Loc.T("vehicle.fuelImport.filterAllStatements")} (*.pdf;*.xlsx;*.xls;*.csv)|*.pdf;*.xlsx;*.xls;*.csv|" +
                         $"{Loc.T("vehicle.fuelImport.filterPdf")} (*.pdf)|*.pdf|" +
                         $"{Loc.T("vehicle.fuelImport.filterExcelCsv")} (*.xlsx;*.xls;*.csv)|*.xlsx;*.xls;*.csv|" +
                         $"{Loc.T("bugReport.filter.all")} (*.*)|*.*",
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
                int imported, skipped;
                decimal totalAmount;
                string provider;
                using (LoadingOverlay.BeginScope(
                    Loc.T("vehicle.fuelImport.overlayTitle", ModuleDisplayName),
                    Loc.T("vehicle.fuelImport.overlayStep", fileNameOnly),
                    operationKey: "Vehicle Manager — Importing Fuel Statement"))
                {
                    (imported, skipped, totalAmount, provider) = SheetMapper.ImportFuelStatement(selectedFilePath);
                }

                if (imported == 0 && skipped == 0)
                {
                    MessageBox.Show(
                        Loc.T("vehicle.fuelImport.noTransactionsBody", fileNameOnly),
                        Loc.T("vehicle.fuelImport.noTransactionsTitle"),
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
                    return;
                }

                string summary = Loc.T(
                    "vehicle.fuelImport.completeBody",
                    provider,
                    fileNameOnly,
                    imported,
                    skipped,
                    totalAmount.ToString("N2"));

                MessageBox.Show(summary, Loc.T("vehicle.fuelImport.completeTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            });
        }
        catch (Exception ex)
        {
            MessageBox.Show(Loc.T("vehicle.fuelImport.errorBody", ex.Message), Loc.T("vehicle.fuelImport.errorTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                SheetMapper.OdometerReconcileResult result;
                using (LoadingOverlay.BeginScope(
                    Loc.T("vehicle.sortOdo.overlayTitle", ModuleDisplayName),
                    Loc.T("vehicle.sortOdo.overlayStep"),
                    operationKey: "Vehicle Manager — Sorting & Matching Odometers"))
                {
                    result = SheetMapper.SortAndHealOdometers();
                }

                if (result.RowsSorted == 0)
                {
                    MessageBox.Show(Loc.T("vehicle.sortOdo.emptyBody"), Loc.T("vehicle.sortOdo.emptyTitle"),
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                string summary = Loc.T(
                    "vehicle.sortOdo.completeBody",
                    result.RowsSorted,
                    result.VehiclesMatched,
                    result.FlaggedPlates,
                    result.HoursNormalized,
                    result.InvalidZeroed);

                if (result.FlaggedPlates > 0)
                    summary += "\n\n" + Loc.T("vehicle.sortOdo.flaggedNote");

                MessageBox.Show(summary, Loc.T("vehicle.sortOdo.completeTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show(Loc.T("vehicle.sortOdo.errorBody", ex.Message), Loc.T("vehicle.sortOdo.errorTitle"),
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });
    }
}
