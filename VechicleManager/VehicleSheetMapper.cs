namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public static class VehicleSheetMapper
{
    private const int MasterColumnCount = 10;
    private const int MaintenanceColumnCount = 10;

    private static readonly string[] MasterSheetAliases = { "Vehicle_Master", "Vehicle Master", "Master" };
    private static readonly string[] MaintenanceSheetAliases = { "Vehicle_Maintenance", "Maintenance", "Logs" };

    public static List<Vehicle> ExtractVehicleRecords()
    {
        var result = new List<Vehicle>();
        var vehicleLookup = new Dictionary<long, Vehicle>();

        // 1. Ingest Master Records
        string masterSheet = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Vehicle_Master");
        IntPtr masterSheetId = WorkbookSession.GetSheetId(masterSheet);

        var masterRef = masterSheetId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, MasterColumnCount - 1, masterSheetId)
            : new ExcelReference(1, 500, 0, MasterColumnCount - 1);

        if (masterRef.GetValue() is object[,] masterGrid)
        {
            int rows = masterGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = masterGrid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long vehicleId) || vehicleId == 0)
                    continue;

                string regNum = CleanCellString(masterGrid[r, 1]);
                string typeStr = CleanCellString(masterGrid[r, 2]);
                string brand = CleanCellString(masterGrid[r, 3]);
                string model = CleanCellString(masterGrid[r, 4]);

                decimal.TryParse(CleanCellString(masterGrid[r, 5]), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal tonnage);
                DateOnly roadTax = ParseExcelDate(masterGrid[r, 6]);
                DateOnly puspakom = ParseExcelDate(masterGrid[r, 7]);
                int.TryParse(CleanCellString(masterGrid[r, 8]), out int odometer);
                string statusStr = CleanCellString(masterGrid[r, 9]);

                var vehicle = new Vehicle
                {
                    VehicleID = vehicleId,
                    Master = new VehicleMaster
                    {
                        RegistrationNumber = regNum,
                        Type = Enum.TryParse(typeStr, true, out VehicleType vType) ? vType : VehicleType.None,
                        Brand = brand,
                        Model = model,
                        CapacityTonnage = tonnage,
                        RoadTaxExpiry = roadTax,
                        PuspakomExpiry = puspakom,
                        CurrentOdometer = odometer,
                        Status = Enum.TryParse(statusStr, true, out OperationalStatus status) ? status : OperationalStatus.Active
                    },
                    MaintenanceLogs = new List<MaintenanceRecord>()
                };

                vehicleLookup[vehicleId] = vehicle;
                result.Add(vehicle);
            }
        }

        // 2. Ingest Maintenance Logs & Map to Parent Vehicles
        string logSheet = WorkbookSession.ResolveSheetName(MaintenanceSheetAliases, "Vehicle_Maintenance");
        IntPtr logSheetId = WorkbookSession.GetSheetId(logSheet);

        var logRef = logSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, MaintenanceColumnCount - 1, logSheetId)
            : new ExcelReference(1, 1000, 0, MaintenanceColumnCount - 1);

        if (logRef.GetValue() is object[,] logGrid)
        {
            int rows = logGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawMaintId = logGrid[r, 0];
                if (rawMaintId == null || rawMaintId is ExcelEmpty || !long.TryParse(rawMaintId.ToString(), out long maintId) || maintId == 0)
                    continue;

                if (!long.TryParse(CleanCellString(logGrid[r, 1]), out long parentVehId) || !vehicleLookup.TryGetValue(parentVehId, out var targetVehicle))
                    continue;

                string serviceTypeStr = CleanCellString(logGrid[r, 2]);
                DateOnly dateOut = ParseExcelDate(logGrid[r, 3]);
                DateOnly estReturn = ParseExcelDate(logGrid[r, 4]);

                string completedStr = CleanCellString(logGrid[r, 5]);
                DateOnly? dateCompleted = string.IsNullOrWhiteSpace(completedStr) ? null : ParseExcelDate(logGrid[r, 5]);

                int.TryParse(CleanCellString(logGrid[r, 6]), out int odoAtService);
                decimal.TryParse(CleanCellString(logGrid[r, 7]), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal cost);
                string workshop = CleanCellString(logGrid[r, 8]);
                string statusStr = CleanCellString(logGrid[r, 9]);

                var record = new MaintenanceRecord
                {
                    MaintenanceID = maintId,
                    VehicleID = parentVehId,
                    ServiceType = Enum.TryParse(serviceTypeStr, true, out ServiceType sType) ? sType : ServiceType.None,
                    DateOut = dateOut,
                    EstimatedReturn = estReturn,
                    DateCompleted = dateCompleted,
                    OdometerAtService = odoAtService,
                    CostAmount = cost,
                    WorkshopName = workshop,
                    Status = Enum.TryParse(statusStr, true, out MaintenanceStatus mStatus) ? mStatus : MaintenanceStatus.Scheduled
                };

                targetVehicle.MaintenanceLogs.Add(record);
            }
        }

        return result;
    }

    public static void PopulateWorksheet(List<Vehicle> vehicleList)
    {
        if (vehicleList == null || vehicleList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Vehicle_Master");
        string targetLogs = WorkbookSession.ResolveSheetName(MaintenanceSheetAliases, "Vehicle_Maintenance");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            // 1. Populate Master Table
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });
            var clearMasterRef = new ExcelReference(1, 500, 0, MasterColumnCount - 1);
            clearMasterRef.SetValue(new object[500, MasterColumnCount]);

            object[,] masterOutput = new object[vehicleList.Count, MasterColumnCount];
            for (int i = 0; i < vehicleList.Count; i++)
            {
                var v = vehicleList[i];
                masterOutput[i, 0] = v.VehicleID;
                masterOutput[i, 1] = v.Master.RegistrationNumber;
                masterOutput[i, 2] = v.Master.Type == VehicleType.None ? string.Empty : v.Master.Type.ToString();
                masterOutput[i, 3] = v.Master.Brand;
                masterOutput[i, 4] = v.Master.Model;
                masterOutput[i, 5] = v.Master.CapacityTonnage;
                masterOutput[i, 6] = v.Master.RoadTaxExpiry.ToDateTime(TimeOnly.MinValue).ToOADate();
                masterOutput[i, 7] = v.Master.PuspakomExpiry.ToDateTime(TimeOnly.MinValue).ToOADate();
                masterOutput[i, 8] = v.Master.CurrentOdometer;
                masterOutput[i, 9] = v.Master.Status.ToString();
            }

            var writeMasterRef = new ExcelReference(1, vehicleList.Count, 0, MasterColumnCount - 1);
            writeMasterRef.SetValue(masterOutput);

            // 2. Flatten & Populate Maintenance Table
            var allLogs = new List<MaintenanceRecord>();
            foreach (var v in vehicleList)
            {
                if (v.MaintenanceLogs != null) allLogs.AddRange(v.MaintenanceLogs);
            }

            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetLogs });
            var clearLogsRef = new ExcelReference(1, 1000, 0, MaintenanceColumnCount - 1);
            clearLogsRef.SetValue(new object[1000, MaintenanceColumnCount]);

            if (allLogs.Count > 0)
            {
                object[,] logsOutput = new object[allLogs.Count, MaintenanceColumnCount];
                for (int i = 0; i < allLogs.Count; i++)
                {
                    var log = allLogs[i];
                    logsOutput[i, 0] = log.MaintenanceID;
                    logsOutput[i, 1] = log.VehicleID;
                    logsOutput[i, 2] = log.ServiceType == ServiceType.None ? string.Empty : log.ServiceType.ToString();
                    logsOutput[i, 3] = log.DateOut.ToDateTime(TimeOnly.MinValue).ToOADate();
                    logsOutput[i, 4] = log.EstimatedReturn.ToDateTime(TimeOnly.MinValue).ToOADate();
                    logsOutput[i, 5] = log.DateCompleted?.ToDateTime(TimeOnly.MinValue).ToOADate() ?? (object)string.Empty;
                    logsOutput[i, 6] = log.OdometerAtService;
                    logsOutput[i, 7] = log.CostAmount;
                    logsOutput[i, 8] = log.WorkshopName;
                    logsOutput[i, 9] = log.Status.ToString();
                }

                var writeLogsRef = new ExcelReference(1, allLogs.Count, 0, MaintenanceColumnCount - 1);
                writeLogsRef.SetValue(logsOutput);
            }
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

    private static string CleanCellString(object? cell)
    {
        if (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError) return string.Empty;
        return cell.ToString()?.Trim() ?? string.Empty;
    }

    private static DateOnly ParseExcelDate(object cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanCellString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return DateOnly.FromDateTime(DateTime.Today);
    }
}