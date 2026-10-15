namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

public static class SheetMapper
{
    private const int MasterColumnSpan = 8;
    private const int VehicleLicenseColumnSpan = 7;
    private const int SkillLicenseColumnSpan = 7;

    private static readonly string[] MasterSheetAliases = { "All Workers", "Workers", "Worker_Master", "All Staff", "All_Staff" };
    private static readonly string[] VehicleLicenseSheetAliases = { "Vehicle Licenses", "Vehicle_Licenses", "Worker_License", "Staff_License" };
    private static readonly string[] SkillLicenseSheetAliases = { "Skill Licenses", "Skill_Licenses", "Worker_Skill", "Staff_Skill" };

    /// <summary>
    /// Performs full extraction and returns records alongside separated lists of Errors and Warnings.
    /// Does not pop up any dialogs directly.
    /// </summary>
    public static (List<WorkerRecord> Records, List<string> Errors, List<string> Warnings) ExtractWorkerRecordsWithDiagnostics()
    {
        var result = new List<WorkerRecord>();
        var workerLookup = new Dictionary<long, WorkerRecord>();
        var validationErrors = new List<string>();
        var validationWarnings = new List<string>();

        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;

        var allowedLicenseTypes = FetchAllowedSkillLicenseTypes();

        string masterSheetName = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Workers");
        string vLicSheetName = WorkbookSession.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string sLicSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        // -------------------------------------------------------------
        // 1. ALL WORKERS (Columns A to H)
        // -------------------------------------------------------------
        try
        {
            dynamic masterWs = wb.Sheets[masterSheetName];
            object[,] masterGrid = masterWs.Range["A2:H1000"].Value2;
            int rows = masterGrid.GetLength(0);

            for (int r = 1; r <= rows; r++)
            {
                int excelRow = r + 1;
                string rawIdStr = CleanCellString(masterGrid[r, 1]);
                string workerName = CleanCellString(masterGrid[r, 2]);
                string rawNric = CleanCellString(masterGrid[r, 3]);
                string passport = CleanCellString(masterGrid[r, 4]);
                string phone = CleanCellString(masterGrid[r, 5]);

                // Completely empty row: skip
                if (string.IsNullOrWhiteSpace(rawIdStr) &&
                    string.IsNullOrWhiteSpace(workerName) &&
                    string.IsNullOrWhiteSpace(rawNric) &&
                    string.IsNullOrWhiteSpace(passport) &&
                    string.IsNullOrWhiteSpace(phone))
                {
                    continue;
                }

                bool hasRowError = false;

                // Rule 1: Worker Name exists but Worker ID is missing
                if (string.IsNullOrWhiteSpace(rawIdStr) && !string.IsNullOrWhiteSpace(workerName))
                {
                    validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Worker '{workerName}' is missing a Worker ID.");
                    hasRowError = true;
                }

                // Rule 2: Invalid or non-numeric Worker ID
                long workerId = 0;
                if (!string.IsNullOrWhiteSpace(rawIdStr))
                {
                    if (!long.TryParse(rawIdStr, out workerId) || workerId <= 0)
                    {
                        validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Invalid Worker ID '{rawIdStr}'. Must be a positive number.");
                        hasRowError = true;
                    }
                }

                // Rule 3: Worker ID exists but Worker Name is missing
                if (string.IsNullOrWhiteSpace(workerName))
                {
                    validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Worker ID {(workerId > 0 ? workerId.ToString() : rawIdStr)} has no Worker Name specified.");
                    hasRowError = true;
                }

                // Rule 4: Duplicate Worker ID check
                if (workerId > 0 && workerLookup.TryGetValue(workerId, out var existingWorker))
                {
                    if (!existingWorker.WorkerDetails.WorkerName.Equals(workerName, StringComparison.OrdinalIgnoreCase))
                    {
                        validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Duplicate Worker ID {workerId} found with conflicting names ('{existingWorker.WorkerDetails.WorkerName}' vs '{workerName}').");
                    }
                    else
                    {
                        validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Duplicate Worker ID {workerId} ('{workerName}') entered multiple times.");
                    }
                    hasRowError = true;
                }

                // Identification warnings
                long nric = 0;
                if (!string.IsNullOrWhiteSpace(rawNric))
                {
                    string cleanNric = System.Text.RegularExpressions.Regex.Replace(rawNric, @"[^\d]", string.Empty);
                    if (!long.TryParse(cleanNric, out nric) || cleanNric.Length != 12)
                    {
                        validationWarnings.Add($"• Row {excelRow} [{masterSheetName}]: Worker ID {workerId} ('{workerName}') NRIC '{rawNric}' is not 12 digits.");
                    }
                }

                if (nric == 0 && string.IsNullOrWhiteSpace(passport))
                {
                    validationWarnings.Add($"• Row {excelRow} [{masterSheetName}]: Worker ID {workerId} ('{workerName}') has neither an NRIC nor a Passport recorded.");
                }

                DateOnly joinDate = ParseExcelDate(masterGrid[r, 6]);
                if (joinDate == default) joinDate = DateOnly.FromDateTime(DateTime.Today);

                DateOnly? resignDate = null;
                DateOnly parsedResign = ParseExcelDate(masterGrid[r, 7]);
                if (parsedResign != default)
                {
                    if (parsedResign < joinDate)
                    {
                        validationErrors.Add($"• Row {excelRow} [{masterSheetName}]: Worker ID {workerId} Resign Date ({parsedResign}) cannot be earlier than Join Date ({joinDate}).");
                        hasRowError = true;
                    }
                    resignDate = parsedResign;
                }

                string statusStr = CleanCellString(masterGrid[r, 8]);
                ContractStatus status = Enum.TryParse(statusStr, true, out ContractStatus parsedStatus)
                    ? parsedStatus
                    : ContractStatus.Active;

                if (!hasRowError && workerId > 0)
                {
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
                            JoinDate = joinDate,
                            ResignDate = resignDate,
                            ContractStatus = status
                        },
                        VehicleLicenses = new List<VehicleLicense>(),
                        SkillLicenses = new List<SkillLicense>()
                    };

                    workerLookup[workerId] = worker;
                    result.Add(worker);
                }
            }
        }
        catch (Exception ex)
        {
            validationErrors.Add($"Could not read master sheet '{masterSheetName}': {ex.Message}");
        }

        // -------------------------------------------------------------
        // PRE-CALCULATE MAXIMUM EXISTING LICENSE IDS FOR AUTO-INCREMENT
        // -------------------------------------------------------------
        long maxVehicleLicId = 0;
        long maxSkillLicId = 0;

        try
        {
            dynamic vWs = wb.Sheets[vLicSheetName];
            object[,] vGrid = vWs.Range["A2:A1000"].Value2;
            int rCount = vGrid.GetLength(0);
            for (int i = 1; i <= rCount; i++)
            {
                if (long.TryParse(vGrid[i, 1]?.ToString(), out long existingId) && existingId > maxVehicleLicId)
                    maxVehicleLicId = existingId;
            }
        }
        catch { }

        try
        {
            dynamic sWs = wb.Sheets[sLicSheetName];
            object[,] sGrid = sWs.Range["A2:A1000"].Value2;
            int rCount = sGrid.GetLength(0);
            for (int i = 1; i <= rCount; i++)
            {
                if (long.TryParse(sGrid[i, 1]?.ToString(), out long existingId) && existingId > maxSkillLicId)
                    maxSkillLicId = existingId;
            }
        }
        catch { }

        var vehicleLicenseIdSet = new HashSet<long>();
        var skillLicenseIdSet = new HashSet<long>();

        // -------------------------------------------------------------
        // 2. VEHICLE LICENSES (Columns A to G)
        // -------------------------------------------------------------
        try
        {
            dynamic vWs = wb.Sheets[vLicSheetName];
            object[,] vLicGrid = vWs.Range["A2:G1000"].Value2;
            int rows = vLicGrid.GetLength(0);

            for (int r = 1; r <= rows; r++)
            {
                int excelRow = r + 1;
                string rawLicIdStr = CleanCellString(vLicGrid[r, 1]);
                string rawWorkerIdStr = CleanCellString(vLicGrid[r, 2]);
                string rawWorkerName = CleanCellString(vLicGrid[r, 3]);
                string licClassStr = CleanCellString(vLicGrid[r, 4]);

                // Completely empty row: skip
                if (string.IsNullOrWhiteSpace(rawLicIdStr) &&
                    string.IsNullOrWhiteSpace(rawWorkerIdStr) &&
                    string.IsNullOrWhiteSpace(rawWorkerName) &&
                    string.IsNullOrWhiteSpace(licClassStr))
                {
                    continue;
                }

                bool hasRowError = false;
                var rowWarnings = new List<string>();

                // A. License ID Check & Auto-Fill
                long licId = 0;
                if (string.IsNullOrWhiteSpace(rawLicIdStr) || !long.TryParse(rawLicIdStr, out licId) || licId <= 0)
                {
                    maxVehicleLicId++;
                    licId = maxVehicleLicId;
                    rowWarnings.Add($"Missing License ID -> Auto-assigned next available ID: {licId}");
                }
                else
                {
                    if (licId > maxVehicleLicId) maxVehicleLicId = licId;
                    if (!vehicleLicenseIdSet.Add(licId))
                    {
                        validationErrors.Add($"• Row {excelRow} [{vLicSheetName}]: Duplicate License ID {licId} detected.");
                        hasRowError = true;
                    }
                }

                // B. Worker ID Check
                long workerId = 0;
                WorkerRecord? targetWorker = null;
                if (string.IsNullOrWhiteSpace(rawWorkerIdStr) || !long.TryParse(rawWorkerIdStr, out workerId) || workerId <= 0)
                {
                    validationErrors.Add($"• Row {excelRow} [{vLicSheetName}]: License ID {licId} is missing a valid numeric Worker ID.");
                    hasRowError = true;
                }
                else if (!workerLookup.TryGetValue(workerId, out targetWorker))
                {
                    validationErrors.Add($"• Row {excelRow} [{vLicSheetName}]: Worker ID {workerId} (License ID {licId}) does not exist in '{masterSheetName}'.");
                    hasRowError = true;
                }

                // C. Worker Name Mismatch Check
                if (targetWorker != null && !string.IsNullOrWhiteSpace(rawWorkerName))
                {
                    string expectedName = targetWorker.WorkerDetails.WorkerName;
                    string expectedShort = targetWorker.WorkerDetails.WorkerShortName;

                    if (!rawWorkerName.Equals(expectedName, StringComparison.OrdinalIgnoreCase) &&
                        !rawWorkerName.Equals(expectedShort, StringComparison.OrdinalIgnoreCase))
                    {
                        rowWarnings.Add($"Worker Name mismatch for ID {workerId} ('{rawWorkerName}' vs master '{expectedName}')");
                    }
                }

                // D. License Class Check
                LicenseClass licClass = LicenseClass.None;
                if (string.IsNullOrWhiteSpace(licClassStr) || !Enum.TryParse(licClassStr, true, out licClass) || licClass == LicenseClass.None)
                {
                    validationErrors.Add($"• Row {excelRow} [{vLicSheetName}]: Invalid License Class '{licClassStr}' for License ID {licId}.");
                    hasRowError = true;
                }

                // E. Date Defaults (Always runs regardless of errors)
                DateOnly parsedRenew = ParseExcelDate(vLicGrid[r, 5]);
                DateOnly parsedExpiry = ParseExcelDate(vLicGrid[r, 6]);
                string statusStr = CleanCellString(vLicGrid[r, 7]);

                DateOnly renewDate;
                DateOnly expiryDate;
                DateOnly today = DateOnly.FromDateTime(DateTime.Today);

                if (parsedRenew == default && parsedExpiry == default)
                {
                    renewDate = (targetWorker != null && targetWorker.WorkerDetails.JoinDate != default)
                        ? targetWorker.WorkerDetails.JoinDate
                        : today;
                    expiryDate = renewDate.AddYears(1);
                    rowWarnings.Add($"Dates empty -> Defaulted Renewal: {renewDate:yyyy-MM-dd}, Expiry: {expiryDate:yyyy-MM-dd}");
                }
                else if (parsedRenew != default && parsedExpiry == default)
                {
                    renewDate = parsedRenew;
                    expiryDate = renewDate.AddYears(1);
                    rowWarnings.Add($"Expiry empty -> Defaulted to 1 year after renewal: {expiryDate:yyyy-MM-dd}");
                }
                else if (parsedRenew == default && parsedExpiry != default)
                {
                    expiryDate = parsedExpiry;
                    renewDate = expiryDate.AddYears(-1);
                    rowWarnings.Add($"Renewal empty -> Defaulted to 1 year prior to expiry: {renewDate:yyyy-MM-dd}");
                }
                else
                {
                    renewDate = parsedRenew;
                    expiryDate = parsedExpiry;
                }

                if (expiryDate <= renewDate)
                {
                    validationErrors.Add($"• Row {excelRow} [{vLicSheetName}]: License ID {licId} Expiry Date ({expiryDate:yyyy-MM-dd}) must be later than Renewal Date ({renewDate:yyyy-MM-dd}).");
                    hasRowError = true;
                }

                // F. Status Defaults
                LicenseStatus licStatus;
                if (!string.IsNullOrWhiteSpace(statusStr) && Enum.TryParse(statusStr, true, out LicenseStatus parsedStatus))
                {
                    licStatus = parsedStatus;
                }
                else
                {
                    licStatus = today > expiryDate ? LicenseStatus.Invalid : LicenseStatus.Valid;
                    rowWarnings.Add($"Status empty -> Defaulted to '{licStatus}' based on Expiry Date ({expiryDate:yyyy-MM-dd})");
                }

                if (rowWarnings.Count > 0)
                {
                    validationWarnings.Add($"• Row {excelRow} [{vLicSheetName}]: License ID {licId}: " + string.Join("; ", rowWarnings));
                }

                if (!hasRowError && targetWorker != null)
                {
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
        }
        catch (Exception ex)
        {
            validationErrors.Add($"Could not read vehicle licenses '{vLicSheetName}': {ex.Message}");
        }

        // -------------------------------------------------------------
        // 3. SKILL LICENSES (Columns A to G)
        // -------------------------------------------------------------
        try
        {
            dynamic sWs = wb.Sheets[sLicSheetName];
            object[,] sLicGrid = sWs.Range["A2:G1000"].Value2;
            int rows = sLicGrid.GetLength(0);

            for (int r = 1; r <= rows; r++)
            {
                int excelRow = r + 1;
                string rawLicIdStr = CleanCellString(sLicGrid[r, 1]);
                string rawWorkerIdStr = CleanCellString(sLicGrid[r, 2]);
                string rawWorkerName = CleanCellString(sLicGrid[r, 3]);
                string skillType = CleanCellString(sLicGrid[r, 4]);

                // Skip completely blank rows
                if (string.IsNullOrWhiteSpace(rawLicIdStr) &&
                    string.IsNullOrWhiteSpace(rawWorkerIdStr) &&
                    string.IsNullOrWhiteSpace(rawWorkerName) &&
                    string.IsNullOrWhiteSpace(skillType))
                {
                    continue;
                }

                bool hasRowError = false;
                var rowWarnings = new List<string>();

                // A. License ID Check & Auto-Fill
                long licId = 0;
                if (string.IsNullOrWhiteSpace(rawLicIdStr) || !long.TryParse(rawLicIdStr, out licId) || licId <= 0)
                {
                    maxSkillLicId++;
                    licId = maxSkillLicId;
                    rowWarnings.Add($"Missing License ID -> Auto-assigned next available ID: {licId}");
                }
                else
                {
                    if (licId > maxSkillLicId) maxSkillLicId = licId;
                    if (!skillLicenseIdSet.Add(licId))
                    {
                        validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: Duplicate Skill License ID {licId} detected.");
                        hasRowError = true;
                    }
                }

                // B. Worker ID Check
                long workerId = 0;
                WorkerRecord? targetWorker = null;
                if (string.IsNullOrWhiteSpace(rawWorkerIdStr) || !long.TryParse(rawWorkerIdStr, out workerId) || workerId <= 0)
                {
                    validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: License ID {licId} is missing a valid numeric Worker ID.");
                    hasRowError = true;
                }
                else if (!workerLookup.TryGetValue(workerId, out targetWorker))
                {
                    validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: Worker ID {workerId} (License ID {licId}) does not exist in '{masterSheetName}'.");
                    hasRowError = true;
                }

                // C. Worker Name Check
                if (targetWorker != null && !string.IsNullOrWhiteSpace(rawWorkerName))
                {
                    string expectedName = targetWorker.WorkerDetails.WorkerName;
                    string expectedShort = targetWorker.WorkerDetails.WorkerShortName;

                    if (!rawWorkerName.Equals(expectedName, StringComparison.OrdinalIgnoreCase) &&
                        !rawWorkerName.Equals(expectedShort, StringComparison.OrdinalIgnoreCase))
                    {
                        rowWarnings.Add($"Worker Name mismatch for ID {workerId} ('{rawWorkerName}' vs master '{expectedName}')");
                    }
                }

                // D. License Type Check
                if (string.IsNullOrWhiteSpace(skillType))
                {
                    validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: License ID {licId} has no License Type specified.");
                    hasRowError = true;
                }
                else if (allowedLicenseTypes.Count > 0 && !allowedLicenseTypes.Contains(skillType))
                {
                    validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: License ID {licId} has invalid License Type '{skillType}'. It is not part of the allowed license types.");
                    hasRowError = true;
                }

                // E. Date Defaults (Always runs regardless of errors)
                DateOnly parsedRenew = ParseExcelDate(sLicGrid[r, 5]);
                DateOnly parsedExpiry = ParseExcelDate(sLicGrid[r, 6]);
                string statusStr = CleanCellString(sLicGrid[r, 7]);

                DateOnly renewDate;
                DateOnly expiryDate;
                DateOnly today = DateOnly.FromDateTime(DateTime.Today);

                if (parsedRenew == default && parsedExpiry == default)
                {
                    renewDate = (targetWorker != null && targetWorker.WorkerDetails.JoinDate != default)
                        ? targetWorker.WorkerDetails.JoinDate
                        : today;
                    expiryDate = renewDate.AddYears(1);
                    rowWarnings.Add($"Dates empty -> Defaulted Renewal: {renewDate:yyyy-MM-dd}, Expiry: {expiryDate:yyyy-MM-dd}");
                }
                else if (parsedRenew != default && parsedExpiry == default)
                {
                    renewDate = parsedRenew;
                    expiryDate = renewDate.AddYears(1);
                    rowWarnings.Add($"Expiry empty -> Defaulted to 1 year after renewal: {expiryDate:yyyy-MM-dd}");
                }
                else if (parsedRenew == default && parsedExpiry != default)
                {
                    expiryDate = parsedExpiry;
                    renewDate = expiryDate.AddYears(-1);
                    rowWarnings.Add($"Renewal empty -> Defaulted to 1 year prior to expiry: {renewDate:yyyy-MM-dd}");
                }
                else
                {
                    renewDate = parsedRenew;
                    expiryDate = parsedExpiry;
                }

                if (expiryDate <= renewDate)
                {
                    validationErrors.Add($"• Row {excelRow} [{sLicSheetName}]: License ID {licId} Expiry Date ({expiryDate:yyyy-MM-dd}) must be later than Renewal Date ({renewDate:yyyy-MM-dd}).");
                    hasRowError = true;
                }

                // F. Status Defaults
                LicenseStatus licStatus;
                if (!string.IsNullOrWhiteSpace(statusStr) && Enum.TryParse(statusStr, true, out LicenseStatus parsedStatus))
                {
                    licStatus = parsedStatus;
                }
                else
                {
                    licStatus = today > expiryDate ? LicenseStatus.Invalid : LicenseStatus.Valid;
                    rowWarnings.Add($"Status empty -> Defaulted to '{licStatus}' based on Expiry Date ({expiryDate:yyyy-MM-dd})");
                }

                if (rowWarnings.Count > 0)
                {
                    validationWarnings.Add($"• Row {excelRow} [{sLicSheetName}]: License ID {licId} ('{skillType}'): " + string.Join("; ", rowWarnings));
                }

                if (!hasRowError && targetWorker != null)
                {
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
        }
        catch (Exception ex)
        {
            validationErrors.Add($"Could not read skill licenses '{sLicSheetName}': {ex.Message}");
        }

        return (result, validationErrors, validationWarnings);
    }

    /// <summary>
    /// Extraction for Ribbon usage: Presents the ValidationSummaryDialog if errors or warnings exist.
    /// Returns empty list if canceled or if blocking errors exist.
    /// </summary>
    public static List<WorkerRecord> ExtractWorkerRecords()
    {
        var (records, errors, warnings) = ExtractWorkerRecordsWithDiagnostics();

        if (errors.Count > 0 || warnings.Count > 0)
        {
            using var dlg = new ValidationSummaryDialog("Worker Manager", errors, warnings);
            dlg.ShowDialog();

            // If there were blocking errors, or user chose Cancel, abort the save
            if (errors.Count > 0 || dlg.Action != ValidationDialogAction.ProceedAnyway)
            {
                return new List<WorkerRecord>();
            }
        }

        return records;
    }

    public static void PopulateWorksheet(List<WorkerRecord> workerList)
    {
        if (workerList == null) return;

        // -----------------------------------------------------------------
        // 1. SORT: Active & Suspended on Top, Resigned at the Bottom
        // -----------------------------------------------------------------
        var sortedList = workerList
            .OrderBy(w => w.WorkerDetails.ContractStatus == ContractStatus.Resigned ? 1 : 0)
            .ThenBy(w => w.WorkerID)
            .ToList();

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Workers");
        string targetVLic = WorkbookSession.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string targetSLic = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            int rowCount = sortedList.Count;

            // -------------------------------------------------------------
            // 2. POPULATE & RESIZE ALL WORKERS (Columns A to H)
            // -------------------------------------------------------------
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });

            var clearMasterRef = new ExcelReference(1, 1000, 0, MasterColumnSpan - 1);
            clearMasterRef.SetValue(new object[1000, MasterColumnSpan]);

            if (rowCount > 0)
            {
                object[,] masterOutput = new object[rowCount, MasterColumnSpan];

                for (int i = 0; i < rowCount; i++)
                {
                    var w = sortedList[i];
                    var details = w.WorkerDetails;

                    masterOutput[i, 0] = w.WorkerID;
                    masterOutput[i, 1] = details.WorkerName;
                    masterOutput[i, 2] = details.NRIC > 0 ? (object)details.NRIC : string.Empty;
                    masterOutput[i, 3] = details.Passport;
                    masterOutput[i, 4] = details.Phone;

                    masterOutput[i, 5] = details.JoinDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    masterOutput[i, 6] = details.ResignDate.HasValue
                        ? (object)details.ResignDate.Value.ToDateTime(TimeOnly.MinValue).ToOADate()
                        : string.Empty;

                    masterOutput[i, 7] = details.ContractStatus.ToString();
                }

                var writeMasterRef = new ExcelReference(1, rowCount, 0, MasterColumnSpan - 1);
                writeMasterRef.SetValue(masterOutput);
            }

            ResizeWorksheetTable(targetMaster, rowCount, MasterColumnSpan);

            // -------------------------------------------------------------
            // 3. POPULATE VEHICLE LICENSES (Columns A to G)
            // -------------------------------------------------------------
            var allVLicenses = new List<(WorkerRecord Worker, VehicleLicense Lic)>();
            foreach (var w in sortedList)
            {
                if (w.VehicleLicenses != null)
                {
                    foreach (var lic in w.VehicleLicenses) allVLicenses.Add((w, lic));
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

            ResizeWorksheetTable(targetVLic, allVLicenses.Count, VehicleLicenseColumnSpan);

            // -------------------------------------------------------------
            // 4. POPULATE SKILL LICENSES (Columns A to G)
            // -------------------------------------------------------------
            var allSLicenses = new List<(WorkerRecord Worker, SkillLicense Lic)>();
            foreach (var w in sortedList)
            {
                if (w.SkillLicenses != null)
                {
                    foreach (var lic in w.SkillLicenses) allSLicenses.Add((w, lic));
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

            ResizeWorksheetTable(targetSLic, allSLicenses.Count, SkillLicenseColumnSpan);

            // -------------------------------------------------------------
            // 5. AUTOMATICALLY RE-HYDRATE SkillLicenseTypesTable (Column I)
            // -------------------------------------------------------------
            PopulateSkillLicenseTypesTable();
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

    public static (bool Success, string Report) PopulateSkillLicenseTypesTable(bool showDialogIfFailed = false)
    {
        var sb = new StringBuilder();
        sb.AppendLine("=== Skill License Types Sync Diagnostics ===");

        var distinctTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        // 1. Database Query Test
        try
        {
            var dbRecords = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
            sb.AppendLine($"[1. Database Query]: Found {dbRecords.Count} row(s) in 'lookup_skill_license_types'.");
            foreach (var r in dbRecords)
            {
                if (!string.IsNullOrWhiteSpace(r.LicenseType))
                {
                    distinctTypes.Add(r.LicenseType.Trim());
                    sb.AppendLine($"    • {r.LicenseType.Trim()}");
                }
            }
        }
        catch (Exception exDb)
        {
            sb.AppendLine($"[1. Database Query FAILED]: {exDb.Message}");
        }

        // 2. Fallback to local snapshot
        if (distinctTypes.Count == 0)
        {
            try
            {
                var localTypes = SKSB_App.Core.FileManager.Services.Load<List<SkillLicenseType>>();
                int localCount = localTypes?.Count ?? 0;
                sb.AppendLine($"[2. Local Snapshot Cache]: Found {localCount} record(s).");
                if (localTypes != null)
                {
                    foreach (var t in localTypes)
                    {
                        if (!string.IsNullOrWhiteSpace(t.LicenseType))
                            distinctTypes.Add(t.LicenseType.Trim());
                    }
                }
            }
            catch (Exception exLoc)
            {
                sb.AppendLine($"[2. Local Snapshot Cache FAILED]: {exLoc.Message}");
            }
        }

        // 3. Fallback: Existing worker records
        if (distinctTypes.Count == 0)
        {
            try
            {
                var localWorkers = SKSB_App.Core.FileManager.Services.Load<List<WorkerRecord>>();
                sb.AppendLine($"[3. Harvest from Workers Cache]: Loaded {localWorkers?.Count ?? 0} worker(s).");
                if (localWorkers != null)
                {
                    foreach (var w in localWorkers)
                    {
                        if (w.SkillLicenses != null)
                        {
                            foreach (var sl in w.SkillLicenses)
                            {
                                if (!string.IsNullOrWhiteSpace(sl.LicenseType))
                                    distinctTypes.Add(sl.LicenseType.Trim());
                            }
                        }
                    }
                }
            }
            catch (Exception exW)
            {
                sb.AppendLine($"[3. Harvest from Workers FAILED]: {exW.Message}");
            }
        }

        if (distinctTypes.Count == 0)
        {
            sb.AppendLine("[Result]: No skill license types found in Database, Local Cache, or Worker records.");
            if (showDialogIfFailed)
            {
                MessageBox.Show(sb.ToString(), "Sync Skill Licenses - Empty", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            return (false, sb.ToString());
        }

        // 4. Excel Table Target Resolution
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;
        dynamic? targetTable = null;
        string foundSheetName = string.Empty;

        try
        {
            foreach (dynamic ws in wb.Sheets)
            {
                try
                {
                    targetTable = ws.ListObjects["SkillLicenseTypesTable"];
                    if (targetTable != null)
                    {
                        foundSheetName = ws.Name?.ToString() ?? string.Empty;
                        break;
                    }
                }
                catch { }
            }

            if (targetTable == null)
            {
                string sLicSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
                try
                {
                    dynamic ws = wb.Sheets[sLicSheetName];
                    foreach (dynamic lo in ws.ListObjects)
                    {
                        string loName = lo.Name?.ToString() ?? string.Empty;
                        if (loName.IndexOf("SkillLicenseType", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            targetTable = lo;
                            foundSheetName = ws.Name?.ToString() ?? string.Empty;
                            break;
                        }
                    }
                }
                catch { }
            }

            if (targetTable == null)
            {
                sb.AppendLine("[4. Excel Table]: FAILED. Table 'SkillLicenseTypesTable' could NOT be found on any worksheet.");
                if (showDialogIfFailed)
                {
                    MessageBox.Show(sb.ToString(), "Sync Skill Licenses - Table Missing", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                return (false, sb.ToString());
            }

            sb.AppendLine($"[4. Excel Table]: Found table '{targetTable.Name}' on sheet '{foundSheetName}'.");

            // 5. Populate and Resize Table
            string[] sortedTypes = distinctTypes.OrderBy(x => x).ToArray();
            int count = sortedTypes.Length;

            dynamic sheet = targetTable.Parent;
            dynamic headerRange = targetTable.HeaderRowRange;
            int headerRow = (int)headerRange.Row;
            int col = (int)headerRange.Column;

            // Clear previous body cells
            try
            {
                if (targetTable.DataBodyRange != null)
                {
                    targetTable.DataBodyRange.ClearContents();
                }
            }
            catch (Exception exClr)
            {
                sb.AppendLine($"[Warning clearing body]: {exClr.Message}");
            }

            // Resize table boundary
            dynamic newRange = sheet.Range[sheet.Cells[headerRow, col], sheet.Cells[headerRow + count, col]];
            targetTable.Resize(newRange);

            // Write values
            object[,] matrix = new object[count, 1];
            for (int i = 0; i < count; i++) matrix[i, 0] = sortedTypes[i];

            dynamic writeRange = sheet.Range[sheet.Cells[headerRow + 1, col], sheet.Cells[headerRow + count, col]];
            writeRange.Value2 = matrix;

            sb.AppendLine($"[5. Writeback]: Successfully written {count} type(s) into {sheet.Name}!R{headerRow + 1}C{col}:R{headerRow + count}C{col}.");

            // Update Name Manager definition
            try
            {
                wb.Names.Add(Name: "LicenseTypesList", RefersTo: $"='{sheet.Name}'!R{headerRow + 1}C{col}:R{headerRow + count}C{col}", Visible: true);
                sb.AppendLine("[6. Name Manager]: 'LicenseTypesList' formula registered.");
            }
            catch (Exception exNm)
            {
                sb.AppendLine($"[6. Name Manager Warning]: {exNm.Message}");
            }

            return (true, sb.ToString());
        }
        catch (Exception exFatal)
        {
            sb.AppendLine($"[FATAL Write Error]: {exFatal.Message}\n{exFatal.StackTrace}");
            if (showDialogIfFailed)
            {
                MessageBox.Show(sb.ToString(), "Sync Skill Licenses - Fatal Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            return (false, sb.ToString());
        }
    }

    /// <summary>
    /// Adjusts the ListObject (Excel Table) boundary to fit the exact row and column count.
    /// If count is 0, shrinks the table down to a single blank template row.
    /// </summary>
    private static void ResizeWorksheetTable(string sheetName, int dataRowCount, int columnSpan)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = xlApp.ActiveWorkbook.Sheets[sheetName];
            if (ws == null || ws.ListObjects.Count == 0) return;

            dynamic table = ws.ListObjects[1]; // First table on the worksheet

            int headerRow = table.HeaderRowRange.Row;
            int startCol = table.HeaderRowRange.Column;
            int endCol = startCol + columnSpan - 1;

            // An Excel table must contain at least 1 data row beneath headers
            int effectiveDataRows = Math.Max(1, dataRowCount);
            int endRow = headerRow + effectiveDataRows;

            dynamic newRange = ws.Range[ws.Cells[headerRow, startCol], ws.Cells[endRow, endCol]];
            table.Resize(newRange);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[ResizeWorksheetTable Warning] Could not resize table on '{sheetName}': {ex.Message}");
        }
    }

    private static HashSet<string> FetchAllowedSkillLicenseTypes()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;

            // 1. Direct Table Lookup by Name across sheets
            dynamic? targetTable = null;
            foreach (dynamic ws in wb.Sheets)
            {
                try
                {
                    targetTable = ws.ListObjects["SkillLicenseTypesTable"];
                    if (targetTable != null) break;
                }
                catch { }
            }

            // Fallback: Check 'Skill Licenses' sheet explicitly
            if (targetTable == null)
            {
                string sLicSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
                try
                {
                    dynamic ws = wb.Sheets[sLicSheetName];
                    foreach (dynamic lo in ws.ListObjects)
                    {
                        string loName = lo.Name?.ToString() ?? string.Empty;
                        if (loName.IndexOf("SkillLicenseType", StringComparison.OrdinalIgnoreCase) >= 0)
                        {
                            targetTable = lo;
                            break;
                        }
                    }
                }
                catch { }
            }

            // Extract values directly from DataBodyRange.Value2 (Safest for COM)
            if (targetTable != null && targetTable.DataBodyRange != null)
            {
                object rawBody = targetTable.DataBodyRange.Value2;
                PopulateSetFromObject(set, rawBody);
            }

            // 2. Fallback: Excel Evaluate structured reference
            if (set.Count == 0)
            {
                try
                {
                    dynamic tableVal = xlApp.Evaluate("SkillLicenseTypesTable[Skill License Types]");
                    PopulateSetFromObject(set, tableVal);
                }
                catch { }
            }

            // 3. Fallback: Name Manager =LicenseTypesList
            if (set.Count == 0)
            {
                try
                {
                    dynamic namedVal = xlApp.Evaluate("LicenseTypesList");
                    PopulateSetFromObject(set, namedVal);
                }
                catch { }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[FetchAllowedSkillLicenseTypes Warning] {ex.Message}");
        }

        return set;
    }

    private static void PopulateSetFromObject(HashSet<string> set, dynamic? obj)
    {
        if (obj == null) return;

        // If Excel returned a COM Range, extract Value2
        try
        {
            if (obj.GetType().Name.Contains("ComObject") || obj.GetType().Name.Contains("Range"))
            {
                obj = obj.Value2;
            }
        }
        catch { }

        if (obj is object[,] arr)
        {
            int rCount = arr.GetLength(0);
            int cCount = arr.GetLength(1);
            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    string val = arr[r, c]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                        set.Add(val);
                }
            }
        }
        else if (obj is object[] vec)
        {
            foreach (var item in vec)
            {
                string val = item?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                    set.Add(val);
            }
        }
        else
        {
            string single = obj.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(single) && !single.StartsWith("#") && !single.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                set.Add(single);
        }
    }

    public static List<SkillLicenseType> ExtractSkillLicenseTypes()
    {
        var distinctTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;

        // 1. Locate SkillLicenseTypesTable
        dynamic? targetTable = null;
        foreach (dynamic ws in wb.Sheets)
        {
            try
            {
                targetTable = ws.ListObjects["SkillLicenseTypesTable"];
                if (targetTable != null) break;
            }
            catch { }
        }

        // Fallback: check on 'Skill Licenses' sheet
        if (targetTable == null)
        {
            string sLicSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
            try
            {
                dynamic ws = wb.Sheets[sLicSheetName];
                foreach (dynamic lo in ws.ListObjects)
                {
                    string loName = lo.Name?.ToString() ?? string.Empty;
                    if (loName.IndexOf("SkillLicenseType", StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        targetTable = lo;
                        break;
                    }
                }
            }
            catch { }
        }

        if (targetTable != null)
        {
            dynamic sheet = targetTable.Parent;
            dynamic headerRange = targetTable.HeaderRowRange;
            int headerRow = (int)headerRange.Row;
            int col = (int)headerRange.Column;

            // CRITICAL FIX: Scan the column directly beneath the header (down to 100 rows)
            // This captures all entries even if Excel hasn't formally expanded DataBodyRange yet!
            try
            {
                dynamic scanRange = sheet.Range[sheet.Cells[headerRow + 1, col], sheet.Cells[headerRow + 100, col]];
                object[,] colGrid = scanRange.Value2;

                if (colGrid != null)
                {
                    int rCount = colGrid.GetLength(0);
                    for (int r = 1; r <= rCount; r++)
                    {
                        string val = CleanCellString(colGrid[r, 1]);
                        if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                        {
                            distinctTypes.Add(val);
                        }
                    }
                }
            }
            catch (Exception exScan)
            {
                System.Diagnostics.Debug.WriteLine($"[ExtractSkillLicenseTypes Column Scan Error] {exScan.Message}");
            }
        }

        // 2. Also harvest from Column D of 'Skill Licenses' to ensure any license assigned to a worker is preserved
        string skillSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
        try
        {
            dynamic ws = wb.Sheets[skillSheetName];
            object[,] grid = ws.Range["D2:D500"].Value2;
            if (grid != null)
            {
                int rCount = grid.GetLength(0);
                for (int r = 1; r <= rCount; r++)
                {
                    string val = CleanCellString(grid[r, 1]);
                    if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                    {
                        distinctTypes.Add(val);
                    }
                }
            }
        }
        catch { }

        // Package into entities
        var result = new List<SkillLicenseType>();
        long id = 1;
        foreach (var type in distinctTypes.OrderBy(x => x))
        {
            result.Add(new SkillLicenseType
            {
                TypeID = id++,
                LicenseType = type
            });
        }

        return result;
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