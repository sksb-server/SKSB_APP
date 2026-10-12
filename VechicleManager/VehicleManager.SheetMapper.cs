namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public static class SheetMapper
{
    // Span from Column A (0) to Column I (8) = 9 columns total
    private const int MasterColumnSpan = 9;

    private static readonly string[] AllVehicleSheetAlias = { "All Vehicle", "All Vehicles", "All_Vehicles", "All_Vehicle" };

    public static List<VehicleRecord> ExtractVehicleRecords()
    {
        var result = new List<VehicleRecord>();
        var validationErrors = new List<string>();

        string masterSheet = WorkbookSession.ResolveSheetName(AllVehicleSheetAlias, "All Vehicles");
        IntPtr masterSheetId = WorkbookSession.GetSheetId(masterSheet);

        // Read columns A (0) to I (8)
        var masterRef = masterSheetId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, MasterColumnSpan - 1, masterSheetId)
            : new ExcelReference(1, 500, 0, MasterColumnSpan - 1);

        if (masterRef.GetValue() is object[,] masterGrid)
        {
            int rows = masterGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                // Column A (Index 0): Vehicle Plate
                string plate = CleanCellString(masterGrid[r, 0]).ToUpperInvariant();
                if (string.IsNullOrWhiteSpace(plate))
                    continue;

                // Column C (Index 2): Vehicle Type
                string typeStr = CleanCellString(masterGrid[r, 2]);
                VehicleType vType = Enum.TryParse(typeStr, true, out VehicleType parsedType)
                    ? parsedType
                    : VehicleType.None;

                // Column I (Index 8): Status (User Override or Explicit Status)
                string statusStr = CleanCellString(masterGrid[r, 8]);
                VehicleStatus vStatus = Enum.TryParse(statusStr, true, out VehicleStatus parsedStatus)
                    ? parsedStatus
                    : VehicleStatus.Active;

                // Dormant / Hidden fields (Col B is Index 1, D-H are indices 3-7)
                // Read them if they happen to have values, otherwise use safe defaults
                string vin = CleanCellString(masterGrid[r, 1]);
                string brand = CleanCellString(masterGrid[r, 3]);
                string model = CleanCellString(masterGrid[r, 4]);

                DateOnly roadTaxExp = ParseExcelDate(masterGrid[r, 5]);
                DateOnly inspExp = ParseExcelDate(masterGrid[r, 6]);
                int.TryParse(CleanCellString(masterGrid[r, 7]), out int odometer);

                // Validation Guard: Ensure Vehicle Type is chosen
                if (vType == VehicleType.None)
                {
                    validationErrors.Add($"• Row {r + 2}: Vehicle '{plate}' does not have a valid Vehicle Type selected.");
                }

                var vehicle = new VehicleRecord
                {
                    VehiclePlate = plate,
                    VehicleDetail = new VehicleDetail
                    {
                        VehiclePlate = plate,
                        VIN = vin,
                        Type = vType,
                        Brand = brand,
                        Model = model,
                        // Set 10-year placeholder if hidden so it never flags "Expired" automatically
                        RoadTaxExpiry = roadTaxExp != default ? roadTaxExp : DateOnly.FromDateTime(DateTime.Today.AddYears(10)),
                        InspectionExpiry = inspExp != default ? inspExp : DateOnly.FromDateTime(DateTime.Today.AddYears(10)),
                        CurrentOdometer = odometer
                    },
                    Maintenances = new List<Maintenance>(),
                    RoadTaxes = new List<RoadTax>(),
                    Inspections = new List<Inspection>()
                };

                // Assign status override
                vehicle.VehicleStatus = vStatus;
                result.Add(vehicle);
            }
        }

        // Pre-Save Validation Gateway
        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save to database due to the following issues:\n");
            foreach (var err in validationErrors) sb.AppendLine(err);
            sb.AppendLine("\nPlease correct the entries and try saving again.");

            MessageBox.Show(sb.ToString(), "Pre-Save Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return new List<VehicleRecord>();
        }

        return result;
    }

    public static void PopulateWorksheet(List<VehicleRecord> vehicleList)
    {
        if (vehicleList == null || vehicleList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(AllVehicleSheetAlias, "All Vehicles");
        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });

            // Clear prior data rows (Row 1 to 500, Cols A to I)
            var clearRef = new ExcelReference(1, 500, 0, MasterColumnSpan - 1);
            clearRef.SetValue(new object[500, MasterColumnSpan]);

            int rowCount = vehicleList.Count;
            object[,] masterOutput = new object[rowCount, MasterColumnSpan];

            for (int i = 0; i < rowCount; i++)
            {
                var v = vehicleList[i];

                // Col A (0): Vehicle Plate
                masterOutput[i, 0] = v.VehiclePlate;
                // Col B (1): VIN (Preserve if exists)
                masterOutput[i, 1] = v.VehicleDetail.VIN;
                // Col C (2): Vehicle Type
                masterOutput[i, 2] = v.VehicleDetail.Type == VehicleType.None ? string.Empty : v.VehicleDetail.Type.ToString();
                // Col D (3): Brand
                masterOutput[i, 3] = v.VehicleDetail.Brand;
                // Col E (4): Model
                masterOutput[i, 4] = v.VehicleDetail.Model;
                // Col F (5): Road Tax Expiry
                masterOutput[i, 5] = v.VehicleDetail.RoadTaxExpiry.ToDateTime(TimeOnly.MinValue).ToOADate();
                // Col G (6): Inspection Expiry
                masterOutput[i, 6] = v.VehicleDetail.InspectionExpiry.ToDateTime(TimeOnly.MinValue).ToOADate();
                // Col H (7): Current Odometer
                masterOutput[i, 7] = v.VehicleDetail.CurrentOdometer;
                // Col I (8): Status
                masterOutput[i, 8] = v.VehicleStatus.ToString();
            }

            var writeRef = new ExcelReference(1, rowCount, 0, MasterColumnSpan - 1);
            writeRef.SetValue(masterOutput);
        }
        finally
        {
            if (initialActiveSheet != null)
            {
                string sheetStr = initialActiveSheet.ToString() ?? string.Empty;
                int bIdx = sheetStr.LastIndexOf(']');
                string restoreName = bIdx >= 0 ? sheetStr.Substring(bIdx + 1) : sheetStr;

                if (!string.IsNullOrEmpty(restoreName))
                {
                    try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { restoreName }); } catch { }
                }
            }

            XlCall.Excel(XlCall.xlcEcho, true);
        }
    }

    private static string CleanCellString(object? cell) =>
        (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            ? string.Empty
            : cell.ToString()?.Trim() ?? string.Empty;

    private static DateOnly ParseExcelDate(object? cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanCellString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return default;
    }
}