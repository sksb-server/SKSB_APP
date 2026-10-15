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

            // Clean out previous maximum capacity to ensure no trailing ghost cells
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

                // Write exact used slice (Row 1 down to Row rowCount)
                var writeMasterRef = new ExcelReference(1, rowCount, 0, MasterColumnSpan - 1);
                writeMasterRef.SetValue(masterOutput);
            }

            // 3. Dynamically shrink/expand the native Excel ListObject Table
            ResizeWorksheetTable(targetMaster, rowCount, MasterColumnSpan);

            // -------------------------------------------------------------
            // 4. POPULATE VEHICLE LICENSES (Columns A to G)
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
            // 5. POPULATE SKILL LICENSES (Columns A to G)
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

            // 1. Try reading from DropDown metadata on the Skill Licenses sheet (Column D / Header: License Type)
            string sLicSheetName = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
            try
            {
                dynamic ws = wb.Sheets[sLicSheetName];
                // Check cell D2 metadata configuration
                dynamic cellD2 = ws.Range["D2"];
                var config = SKSB_App.Core.TemplateEditor.DropDown.Metadata.ReadConfig(cellD2);

                if (!string.IsNullOrWhiteSpace(config.Source))
                {
                    dynamic? sourceVal = null;
                    string cleanSource = config.Source.TrimStart('=');

                    if (cleanSource.Contains("[") && cleanSource.Contains("]"))
                    {
                        try { sourceVal = xlApp.Range[cleanSource].Value2; } catch { }
                    }
                    if (sourceVal == null)
                    {
                        sourceVal = SKSB_App.Core.TemplateEditor.DropDown.Metadata.EvaluateFormulaSafely(xlApp, cleanSource);
                    }

                    PopulateSetFromObject(set, sourceVal);
                }
            }
            catch { }

            // 2. Fallback: Check named ranges "SkillLicenseTypes", "LicenseTypesList", or "SkillTypes"
            if (set.Count == 0)
            {
                string[] possibleNames = { "SkillLicenseTypes", "LicenseTypesList", "SkillTypes", "LicenseTypeList" };
                foreach (var name in possibleNames)
                {
                    try
                    {
                        dynamic eval = xlApp.Evaluate(name);
                        PopulateSetFromObject(set, eval);
                        if (set.Count > 0) break;
                    }
                    catch { }
                }
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
        if (obj is object[,] arr)
        {
            int rCount = arr.GetLength(0);
            int cCount = arr.GetLength(1);
            for (int r = 1; r <= rCount; r++)
            {
                for (int c = 1; c <= cCount; c++)
                {
                    string val = arr[r, c]?.ToString()?.Trim() ?? string.Empty;
                    if (!string.IsNullOrEmpty(val) && !val.StartsWith("#"))
                        set.Add(val);
                }
            }
        }
        else if (obj is object[] vec)
        {
            foreach (var item in vec)
            {
                string val = item?.ToString()?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(val) && !val.StartsWith("#"))
                    set.Add(val);
            }
        }
        else if (obj != null)
        {
            string single = obj.ToString()?.Trim() ?? string.Empty;
            if (!string.IsNullOrEmpty(single) && !single.StartsWith("#"))
                set.Add(single);
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