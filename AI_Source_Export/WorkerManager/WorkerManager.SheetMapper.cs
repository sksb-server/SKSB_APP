namespace SKSB_App.WorkerManager;

using ExcelDna.Integration;
using SKSB_App.Core.DialogBoxes;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;
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
    /// Pre-caches validation and dropdown metadata on open.
    /// </summary>
    public static void InitializeSheetRules()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            string masterSheetName = SheetOps.ResolveSheetName(MasterSheetAliases, "All Workers");
            string vLicSheetName = SheetOps.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
            string sLicSheetName = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

            try { CellRulesExtractor.InitializeRulesForSheet((object)wb.Sheets[masterSheetName], headerRow: 1, startCol: 1, columnSpan: MasterColumnSpan); } catch { }
            try { CellRulesExtractor.InitializeRulesForSheet((object)wb.Sheets[vLicSheetName], headerRow: 1, startCol: 1, columnSpan: VehicleLicenseColumnSpan); } catch { }
            try { CellRulesExtractor.InitializeRulesForSheet((object)wb.Sheets[sLicSheetName], headerRow: 1, startCol: 1, columnSpan: SkillLicenseColumnSpan); } catch { }
        }
        catch { }
    }

    public static (List<WorkerRecord> Records, List<string> Errors, List<string> Warnings) ExtractWorkerRecordsWithDiagnostics()
    {
        var result = new List<WorkerRecord>();
        var workerLookup = new Dictionary<long, WorkerRecord>();
        var validationErrors = new List<string>();
        var validationWarnings = new List<string>();

        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;

        string masterSheetName = SheetOps.ResolveSheetName(MasterSheetAliases, "All Workers");
        string vLicSheetName = SheetOps.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string sLicSheetName = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        // -------------------------------------------------------------
        // 1. ALL WORKERS (Evaluated via CellRulesExtractor)
        // -------------------------------------------------------------
        try
        {
            dynamic masterWs = wb.Sheets[masterSheetName];
            var (masterRows, mErrors, mWarnings) = CellRulesExtractor.EvaluateAndExtract(
                (object)masterWs,
                startRow: 2,
                maxRows: 1000,
                startCol: 1,
                columnSpan: MasterColumnSpan,
                driverColIndex: 1
            );

            validationErrors.AddRange(mErrors);
            validationWarnings.AddRange(mWarnings);

            int excelRow = 1;
            foreach (var row in masterRows)
            {
                excelRow++;
                long workerId = DataFormatting.CleanString(GetRowValue(row, "Worker ID", 0)) is string idStr && long.TryParse(idStr, out long parsedId)
                    ? parsedId
                    : 0;
                string workerName = DataFormatting.CleanString(GetRowValue(row, "Worker Name", 1));

                if (workerId > 0 && workerLookup.TryGetValue(workerId, out var existingWorker))
                {
                    if (!existingWorker.WorkerDetails.WorkerName.Equals(workerName, StringComparison.OrdinalIgnoreCase))
                    {
                        validationErrors.Add($"  Row {excelRow} [{masterSheetName}]: Duplicate Worker ID {workerId} with conflicting names ('{existingWorker.WorkerDetails.WorkerName}' vs '{workerName}').");
                    }
                    else
                    {
                        validationErrors.Add($"  Row {excelRow} [{masterSheetName}]: Duplicate Worker ID {workerId} ('{workerName}') entered multiple times.");
                    }
                    continue;
                }

                string rawNric = DataFormatting.CleanString(GetRowValue(row, "NRIC", 2));
                long nric = 0;
                if (!string.IsNullOrWhiteSpace(rawNric))
                {
                    string cleanNric = System.Text.RegularExpressions.Regex.Replace(rawNric, @"[^\d]", string.Empty);
                    long.TryParse(cleanNric, out nric);
                }

                string passport = DataFormatting.CleanString(GetRowValue(row, "Passport", 3));
                string phone = DataFormatting.CleanString(GetRowValue(row, "Phone", 4));
                DateOnly joinDate = DataFormatting.ParseDate(GetRowValue(row, "Join Date", 5));
                if (joinDate == default) joinDate = DateOnly.FromDateTime(DateTime.Today);

                DateOnly? resignDate = null;
                DateOnly parsedResign = DataFormatting.ParseDate(GetRowValue(row, "Resign Date", 6));
                if (parsedResign != default)
                {
                    if (parsedResign < joinDate)
                    {
                        validationErrors.Add($"  Row {excelRow} [{masterSheetName}]: Worker ID {workerId} Resign Date ({parsedResign}) cannot be earlier than Join Date ({joinDate}).");
                        continue;
                    }
                    resignDate = parsedResign;
                }

                string statusStr = DataFormatting.CleanString(GetRowValue(row, "Contract Status", 7));
                ContractStatus status = Enum.TryParse(statusStr, true, out ContractStatus parsedStatus)
                    ? parsedStatus
                    : ContractStatus.Active;

                if (workerId > 0)
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
        // 2. VEHICLE LICENSES (Evaluated via CellRulesExtractor)
        // -------------------------------------------------------------
        try
        {
            dynamic vWs = wb.Sheets[vLicSheetName];
            var (vRows, vErrors, vWarnings) = CellRulesExtractor.EvaluateAndExtract(
                (object)vWs,
                startRow: 2,
                maxRows: 1000,
                startCol: 1,
                columnSpan: VehicleLicenseColumnSpan,
                driverColIndex: 2
            );

            validationErrors.AddRange(vErrors);
            validationWarnings.AddRange(vWarnings);

            var vLicSet = new HashSet<long>();
            long autoLicId = 1000;

            foreach (var row in vRows)
            {
                long.TryParse(DataFormatting.CleanString(GetRowValue(row, "License ID", 0)), out long licId);
                long.TryParse(DataFormatting.CleanString(GetRowValue(row, "Worker ID", 1)), out long workerId);
                string licClassStr = DataFormatting.CleanString(GetRowValue(row, "License Class", 3));

                if (licId <= 0) licId = ++autoLicId;
                if (!vLicSet.Add(licId))
                {
                    validationErrors.Add($"  [{vLicSheetName}]: Duplicate Vehicle License ID {licId}.");
                    continue;
                }

                if (!workerLookup.TryGetValue(workerId, out var targetWorker))
                {
                    validationErrors.Add($"  [{vLicSheetName}]: Worker ID {workerId} (License ID {licId}) does not exist in '{masterSheetName}'.");
                    continue;
                }

                if (!Enum.TryParse(licClassStr, true, out LicenseClass licClass) || licClass == LicenseClass.None)
                {
                    validationErrors.Add($"  [{vLicSheetName}]: Invalid License Class '{licClassStr}' for License ID {licId}.");
                    continue;
                }

                DateOnly renewDate = DataFormatting.ParseDate(GetRowValue(row, "Renew Date", 4));
                DateOnly expiryDate = DataFormatting.ParseDate(GetRowValue(row, "Expiry Date", 5));
                if (renewDate == default) renewDate = targetWorker.WorkerDetails.JoinDate;
                if (expiryDate == default) expiryDate = renewDate.AddYears(1);

                string statusStr = DataFormatting.CleanString(GetRowValue(row, "Status", 6));
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
        catch (Exception ex)
        {
            validationErrors.Add($"Could not read vehicle licenses '{vLicSheetName}': {ex.Message}");
        }

        // -------------------------------------------------------------
        // 3. SKILL LICENSES (Evaluated via CellRulesExtractor)
        // -------------------------------------------------------------
        try
        {
            dynamic sWs = wb.Sheets[sLicSheetName];
            var (sRows, sErrors, sWarnings) = CellRulesExtractor.EvaluateAndExtract(
                (object)sWs,
                startRow: 2,
                maxRows: 1000,
                startCol: 1,
                columnSpan: SkillLicenseColumnSpan,
                driverColIndex: 2
            );

            validationErrors.AddRange(sErrors);
            validationWarnings.AddRange(sWarnings);

            var sLicSet = new HashSet<long>();
            long autoLicId = 2000;

            foreach (var row in sRows)
            {
                long.TryParse(DataFormatting.CleanString(GetRowValue(row, "License ID", 0)), out long licId);
                long.TryParse(DataFormatting.CleanString(GetRowValue(row, "Worker ID", 1)), out long workerId);
                string skillType = DataFormatting.CleanString(GetRowValue(row, "License Type", 3));

                if (licId <= 0) licId = ++autoLicId;
                if (!sLicSet.Add(licId))
                {
                    validationErrors.Add($"  [{sLicSheetName}]: Duplicate Skill License ID {licId}.");
                    continue;
                }

                if (!workerLookup.TryGetValue(workerId, out var targetWorker))
                {
                    validationErrors.Add($"  [{sLicSheetName}]: Worker ID {workerId} (License ID {licId}) does not exist in '{masterSheetName}'.");
                    continue;
                }

                DateOnly renewDate = DataFormatting.ParseDate(GetRowValue(row, "Renew Date", 4));
                DateOnly expiryDate = DataFormatting.ParseDate(GetRowValue(row, "Expiry Date", 5));
                if (renewDate == default) renewDate = targetWorker.WorkerDetails.JoinDate;
                if (expiryDate == default) expiryDate = renewDate.AddYears(1);

                string statusStr = DataFormatting.CleanString(GetRowValue(row, "Status", 6));
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
        catch (Exception ex)
        {
            validationErrors.Add($"Could not read skill licenses '{sLicSheetName}': {ex.Message}");
        }

        return (result, validationErrors, validationWarnings);
    }

    private static object? GetRowValue(Dictionary<string, object?> dict, string key, int fallbackIndex)
    {
        if (dict.TryGetValue(key, out var val)) return val;
        foreach (var entry in dict)
        {
            if (entry.Key.IndexOf(key, StringComparison.OrdinalIgnoreCase) >= 0)
                return entry.Value;
        }
        return dict.Values.ElementAtOrDefault(fallbackIndex);
    }

    public static List<WorkerRecord> ExtractWorkerRecords()
    {
        var (records, errors, warnings) = ExtractWorkerRecordsWithDiagnostics();
        if (errors.Count > 0 || warnings.Count > 0)
        {
            using var dlg = new ValidationSummary("Worker Manager", errors, warnings);
            dlg.ShowDialog();
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
        var sortedList = workerList
            .OrderBy(w => w.WorkerDetails.ContractStatus == ContractStatus.Resigned ? 1 : 0)
            .ThenBy(w => w.WorkerID)
            .ToList();

        string targetMaster = SheetOps.ResolveSheetName(MasterSheetAliases, "All Workers");
        string targetVLic = SheetOps.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string targetSLic = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);
        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);
            int rowCount = sortedList.Count;

            // 1. ALL WORKERS
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
            DataFormatting.ResizeTable(targetMaster, rowCount, MasterColumnSpan);

            // 2. VEHICLE LICENSES
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
            DataFormatting.ResizeTable(targetVLic, allVLicenses.Count, VehicleLicenseColumnSpan);

            // 3. SKILL LICENSES
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
            DataFormatting.ResizeTable(targetSLic, allSLicenses.Count, SkillLicenseColumnSpan);

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

        try
        {
            var dbRecords = ServerSyncCommands.FetchSkillLicenseTypesFromDb();
            sb.AppendLine($"[1. Database Query]: Found {dbRecords.Count} row(s) in 'lookup_skill_license_types'.");
            foreach (var r in dbRecords)
            {
                if (!string.IsNullOrWhiteSpace(r.LicenseType))
                {
                    distinctTypes.Add(r.LicenseType.Trim());
                    sb.AppendLine($"      {r.LicenseType.Trim()}");
                }
            }
        }
        catch (Exception exDb)
        {
            sb.AppendLine($"[1. Database Query FAILED]: {exDb.Message}");
        }

        if (distinctTypes.Count == 0)
        {
            try
            {
                var localTypes = SKSB_App.Core.FileManager.OfflineSaveServices.Load<List<SkillLicenseType>>();
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

        if (distinctTypes.Count == 0)
        {
            try
            {
                var localWorkers = SKSB_App.Core.FileManager.OfflineSaveServices.Load<List<WorkerRecord>>();
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
                string sLicSheetName = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
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
            string[] sortedTypes = distinctTypes.OrderBy(x => x).ToArray();
            int count = sortedTypes.Length;
            dynamic sheet = targetTable.Parent;
            dynamic headerRange = targetTable.HeaderRowRange;
            int headerRow = (int)headerRange.Row;
            int col = (int)headerRange.Column;

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

            dynamic newRange = sheet.Range[sheet.Cells[headerRow, col], sheet.Cells[headerRow + count, col]];
            targetTable.Resize(newRange);
            object[,] matrix = new object[count, 1];
            for (int i = 0; i < count; i++) matrix[i, 0] = sortedTypes[i];
            dynamic writeRange = sheet.Range[sheet.Cells[headerRow + 1, col], sheet.Cells[headerRow + count, col]];
            writeRange.Value2 = matrix;
            sb.AppendLine($"[5. Writeback]: Successfully written {count} type(s) into {sheet.Name}!R{headerRow + 1}C{col}:R{headerRow + count}C{col}.");

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

    public static List<SkillLicenseType> ExtractSkillLicenseTypes()
    {
        var distinctTypes = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        dynamic xlApp = ExcelDnaUtil.Application;
        dynamic wb = xlApp.ActiveWorkbook;
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

        if (targetTable == null)
        {
            string sLicSheetName = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
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
            try
            {
                dynamic scanRange = sheet.Range[sheet.Cells[headerRow + 1, col], sheet.Cells[headerRow + 100, col]];
                object[,] colGrid = scanRange.Value2;
                if (colGrid != null)
                {
                    int rCount = colGrid.GetLength(0);
                    for (int r = 1; r <= rCount; r++)
                    {
                        string val = DataFormatting.CleanString(colGrid[r, 1]);
                        if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                        {
                            distinctTypes.Add(val);
                        }
                    }
                }
            }
            catch { }
        }

        string skillSheetName = SheetOps.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");
        try
        {
            dynamic ws = wb.Sheets[skillSheetName];
            object[,] grid = ws.Range["D2:D500"].Value2;
            if (grid != null)
            {
                int rCount = grid.GetLength(0);
                for (int r = 1; r <= rCount; r++)
                {
                    string val = DataFormatting.CleanString(grid[r, 1]);
                    if (!string.IsNullOrWhiteSpace(val) && !val.StartsWith("#") && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                    {
                        distinctTypes.Add(val);
                    }
                }
            }
        }
        catch { }

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
}