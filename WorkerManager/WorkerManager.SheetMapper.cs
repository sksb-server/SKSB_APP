namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public static class SheetMapper
{
    // 6 Columns: Worker ID, Worker Name, IC Number, Passport Number, Phone Number, Contract Status
    private const int MasterColumnSpan = 6;
    private const int VehicleLicenseColumnSpan = 7;
    private const int SkillLicenseColumnSpan = 7;

    private static readonly string[] MasterSheetAliases = { "All Workers", "Workers", "Worker_Master", "All Worker", "All_Worker", "Worker_Master" };
    private static readonly string[] VehicleLicenseSheetAliases = { "Vehicle Licenses", "Vehicle_Licenses", "Worker_License" };
    private static readonly string[] SkillLicenseSheetAliases = { "Skill Licenses", "Skill_Licenses", "Worker_Skill" };

    public static List<WorkerRecord> ExtractWorkerRecords()
    {
        var result = new List<WorkerRecord>();
        var workerLookup = new Dictionary<long, WorkerRecord>();
        var validationErrors = new List<string>();

        var allowedLicenseTypes = FetchAllowedSkillLicenseTypes();

        // -------------------------------------------------------------
        // 1. ALL WORKERS (Columns A to F)
        // -------------------------------------------------------------
        string masterSheet = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Workers");
        IntPtr masterSheetId = WorkbookSession.GetSheetId(masterSheet);

        var masterRef = masterSheetId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, MasterColumnSpan - 1, masterSheetId)
            : new ExcelReference(1, 500, 0, MasterColumnSpan - 1);

        if (masterRef.GetValue() is object[,] masterGrid)
        {
            int rows = masterGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = masterGrid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long workerId) || workerId <= 0)
                    continue;

                string workerName = CleanCellString(masterGrid[r, 1]);
                object rawNric = masterGrid[r, 2];
                string passport = CleanCellString(masterGrid[r, 3]);
                string phone = CleanCellString(masterGrid[r, 4]);
                string statusStr = CleanCellString(masterGrid[r, 5]);

                long nric = 0;

                if ( rawNric != null && !(rawNric is ExcelEmpty))
                {
                    string nricStr = CleanCellString(rawNric).Replace("-", string.Empty);
                    long.TryParse(nricStr, out nric);
                }

                if (string.IsNullOrWhiteSpace(workerName))
                {
                    validationErrors.Add($"• Row {r + 2} [{masterSheet}]: Worker ID {workerId} has no Worker Name.");
                    continue;
                }

                ContractStatus status = Enum.TryParse(statusStr, true, out ContractStatus parsedStatus)
                    ? parsedStatus
                    : ContractStatus.Active;

                var worker = new WorkerRecord
                {
                    WorkerID = workerId,
                    WorkerDetails = new WorkerDetails
                    {
                        WorkerID = workerId,
                        WorkerName = workerName,
                        NRIC = nric,
                        Passport = passport,
                        Phone = phone,
                        ContractStatus = status
                    },
                    VehicleLicenses = new List<VehicleLicense>(),
                    SkillLicenses = new List<SkillLicense>()
                };

                workerLookup[workerId] = worker;
                result.Add(worker);
            }
        }

        // -------------------------------------------------------------
        // 2. VEHICLE LICENSES (Columns A to G)
        // -------------------------------------------------------------
        string vLicSheet = WorkbookSession.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        IntPtr vLicSheetId = WorkbookSession.GetSheetId(vLicSheet);

        var vLicRef = vLicSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, VehicleLicenseColumnSpan - 1, vLicSheetId)
            : new ExcelReference(1, 1000, 0, VehicleLicenseColumnSpan - 1);

        if (vLicRef.GetValue() is object[,] vLicGrid)
        {
            int rows = vLicGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawLicId = vLicGrid[r, 0];
                long.TryParse(rawLicId?.ToString(), out long licId);

                object rawWorkerId = vLicGrid[r, 1];
                if (rawWorkerId == null || rawWorkerId is ExcelEmpty || !long.TryParse(rawWorkerId.ToString(), out long workerId) || workerId <= 0)
                    continue;

                if (!workerLookup.TryGetValue(workerId, out var targetWorker))
                {
                    validationErrors.Add($"• Row {r + 2} [Vehicle Licenses]: Worker ID {workerId} does not exist in '{masterSheet}'.");
                    continue;
                }

                string licClassStr = CleanCellString(vLicGrid[r, 3]);
                if (string.IsNullOrWhiteSpace(licClassStr)) continue;

                if (!Enum.TryParse(licClassStr, true, out LicenseClass licClass) || licClass == LicenseClass.None)
                {
                    validationErrors.Add($"• Row {r + 2} [Vehicle Licenses]: Invalid License_Class '{licClassStr}' for Worker ID {workerId}.");
                    continue;
                }

                DateOnly renewDate = ParseExcelDate(vLicGrid[r, 4]);
                DateOnly expiryDate = ParseExcelDate(vLicGrid[r, 5]);
                string statusStr = CleanCellString(vLicGrid[r, 6]);

                if (renewDate == default) renewDate = DateOnly.FromDateTime(DateTime.Today);
                if (expiryDate == default) expiryDate = renewDate.AddYears(1);

                LicenseStatus licStatus = Enum.TryParse(statusStr, true, out LicenseStatus parsedStatus)
                    ? parsedStatus
                    : (DateOnly.FromDateTime(DateTime.Today) > expiryDate ? LicenseStatus.Invalid : LicenseStatus.Valid);

                targetWorker.VehicleLicenses.Add(new VehicleLicense
                {
                    LicenseID = licId,
                    WorkerID = workerId,
                    LicenseClass = licClass,
                    RenewDate = renewDate,
                    ExpiryDate = expiryDate,
                    LicenseStatus = licStatus
                });
            }
        }

        // -------------------------------------------------------------
        // 3. SKILL LICENSES (Columns A to G)
        // -------------------------------------------------------------
        string sLicSheet = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
        IntPtr sLicSheetId = WorkbookSession.GetSheetId(sLicSheet);

        var sLicRef = sLicSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, SkillLicenseColumnSpan - 1, sLicSheetId)
            : new ExcelReference(1, 1000, 0, SkillLicenseColumnSpan - 1);

        if (sLicRef.GetValue() is object[,] sLicGrid)
        {
            int rows = sLicGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawLicId = sLicGrid[r, 0];
                long.TryParse(rawLicId?.ToString(), out long licId);

                object rawWorkerId = sLicGrid[r, 1];
                if (rawWorkerId == null || rawWorkerId is ExcelEmpty || !long.TryParse(rawWorkerId.ToString(), out long workerId) || workerId <= 0)
                    continue;

                if (!workerLookup.TryGetValue(workerId, out var targetWorker))
                {
                    validationErrors.Add($"• Row {r + 2} [Skill Licenses]: Worker ID {workerId} does not exist in '{masterSheet}'.");
                    continue;
                }

                string skillType = CleanCellString(sLicGrid[r, 3]);
                if (string.IsNullOrWhiteSpace(skillType)) continue;

                if (allowedLicenseTypes.Count > 0 && !allowedLicenseTypes.Contains(skillType))
                {
                    validationErrors.Add($"• Row {r + 2} [Skill Licenses]: Unrecognized License Type '{skillType}' for Worker ID {workerId}.");
                    continue;
                }

                DateOnly renewDate = ParseExcelDate(sLicGrid[r, 4]);
                DateOnly expiryDate = ParseExcelDate(sLicGrid[r, 5]);
                string statusStr = CleanCellString(sLicGrid[r, 6]);

                if (renewDate == default) renewDate = DateOnly.FromDateTime(DateTime.Today);
                if (expiryDate == default) expiryDate = renewDate.AddYears(1);

                LicenseStatus licStatus = Enum.TryParse(statusStr, true, out LicenseStatus parsedStatus)
                    ? parsedStatus
                    : (DateOnly.FromDateTime(DateTime.Today) > expiryDate ? LicenseStatus.Invalid : LicenseStatus.Valid);

                targetWorker.SkillLicenses.Add(new SkillLicense
                {
                    LicenseID = licId,
                    WorkerID = workerId,
                    LicenseType = skillType,
                    RenewDate = renewDate,
                    ExpiryDate = expiryDate,
                    LicenseStatus = licStatus
                });
            }
        }

        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save to system due to validation errors:\n");
            foreach (var err in validationErrors) sb.AppendLine(err);
            MessageBox.Show(sb.ToString(), "Validation Errors", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return new List<WorkerRecord>();
        }

        return result;
    }

    public static void PopulateWorksheet(List<WorkerRecord> workerList)
    {
        if (workerList == null || workerList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Workers");
        string targetVLic = WorkbookSession.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string targetSLic = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            // 1. POPULATE WORKERS TABLE (Columns A to F)
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });
            var clearMasterRef = new ExcelReference(1, 500, 0, MasterColumnSpan - 1);
            clearMasterRef.SetValue(new object[500, MasterColumnSpan]);

            int rowCount = workerList.Count;
            object[,] masterOutput = new object[rowCount, MasterColumnSpan];

            for (int i = 0; i < rowCount; i++)
            {
                var s = workerList[i];
                masterOutput[i, 0] = s.WorkerID;
                masterOutput[i, 1] = s.WorkerDetails.WorkerName;
                masterOutput[i, 2] = s.WorkerDetails.NRIC > 0 ? (object)s.WorkerDetails.NRIC : string.Empty;
                masterOutput[i, 3] = s.WorkerDetails.Passport;
                masterOutput[i, 4] = s.WorkerDetails.Phone;
                masterOutput[i, 5] = s.WorkerDetails.ContractStatus.ToString();
            }

            var writeMasterRef = new ExcelReference(1, rowCount, 0, MasterColumnSpan - 1);
            writeMasterRef.SetValue(masterOutput);

            // 2. POPULATE VEHICLE LICENSES (Columns A to G)
            var allVLicenses = new List<(WorkerRecord Worker, VehicleLicense Lic)>();
            foreach (var s in workerList)
            {
                if (s.VehicleLicenses != null)
                {
                    foreach (var lic in s.VehicleLicenses) allVLicenses.Add((s, lic));
                }
            }

            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetVLic });
            var clearVRef = new ExcelReference(1, 1000, 0, VehicleLicenseColumnSpan - 1);
            clearVRef.SetValue(new object[1000, VehicleLicenseColumnSpan]);

            if (allVLicenses.Count > 0)
            {
                object[,] vOutput = new object[allVLicenses.Count, VehicleLicenseColumnSpan];
                for (int i = 0; i < allVLicenses.Count; i++)
                {
                    var (worker, lic) = allVLicenses[i];
                    vOutput[i, 0] = lic.LicenseID > 0 ? (object)lic.LicenseID : string.Empty;
                    vOutput[i, 1] = lic.WorkerID;
                    vOutput[i, 2] = worker.WorkerDetails.WorkerName;
                    vOutput[i, 3] = lic.LicenseClass == LicenseClass.None ? string.Empty : lic.LicenseClass.ToString();
                    vOutput[i, 4] = lic.RenewDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    vOutput[i, 5] = lic.ExpiryDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    vOutput[i, 6] = lic.LicenseStatus.ToString();
                }

                var writeVRef = new ExcelReference(1, allVLicenses.Count, 0, VehicleLicenseColumnSpan - 1);
                writeVRef.SetValue(vOutput);
            }

            // 3. POPULATE SKILL LICENSES (Columns A to G)
            var allSLicenses = new List<(WorkerRecord Worker, SkillLicense Lic)>();
            foreach (var s in workerList)
            {
                if (s.SkillLicenses != null)
                {
                    foreach (var lic in s.SkillLicenses) allSLicenses.Add((s, lic));
                }
            }

            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetSLic });
            var clearSRef = new ExcelReference(1, 1000, 0, SkillLicenseColumnSpan - 1);
            clearSRef.SetValue(new object[1000, SkillLicenseColumnSpan]);

            if (allSLicenses.Count > 0)
            {
                object[,] sOutput = new object[allSLicenses.Count, SkillLicenseColumnSpan];
                for (int i = 0; i < allSLicenses.Count; i++)
                {
                    var (worker, lic) = allSLicenses[i];
                    sOutput[i, 0] = lic.LicenseID > 0 ? (object)lic.LicenseID : string.Empty;
                    sOutput[i, 1] = lic.WorkerID;
                    sOutput[i, 2] = worker.WorkerDetails.WorkerName;
                    sOutput[i, 3] = lic.LicenseType;
                    sOutput[i, 4] = lic.RenewDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    sOutput[i, 5] = lic.ExpiryDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    sOutput[i, 6] = lic.LicenseStatus.ToString();
                }

                var writeSRef = new ExcelReference(1, allSLicenses.Count, 0, SkillLicenseColumnSpan - 1);
                writeSRef.SetValue(sOutput);
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

    private static HashSet<string> FetchAllowedSkillLicenseTypes()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic eval = xlApp.Evaluate("LicenseTypesList");
            if (eval is object[,] arr)
            {
                int rCount = arr.GetLength(0);
                int cCount = arr.GetLength(1);
                for (int r = 1; r <= rCount; r++)
                {
                    for (int c = 1; c <= cCount; c++)
                    {
                        string val = arr[r, c]?.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(val)) set.Add(val);
                    }
                }
            }
            else if (eval != null)
            {
                string single = eval.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(single)) set.Add(single);
            }
        }
        catch { }
        return set;
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