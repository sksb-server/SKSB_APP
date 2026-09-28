namespace SKSB_App.StaffManager;

using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;
using System.Windows.Forms;

public static class StaffSheetMapper
{
    private const int MasterColumnCount = 10;     // Columns A to J (StaffID -> ActiveStatus)
    private const int PreferenceColumnCount = 11; // Columns A to K (StaffID -> DedicatedVehicleID)
    private const int LicenseColumnCount = 5;     // Columns A to E (StaffID, Name, LicenseClass, ExpiryDate, Status)

    private static readonly string[] MasterSheetAliases = { "Staff_Master", "Staff Master", "Master" };
    private static readonly string[] PreferenceSheetAliases = { "Staff_Preference", "Staff Preference", "Preference", "Preferences" };
    private static readonly string[] LicenseSheetAliases = { "Staff_License", "Staff License", "License", "Licenses", "Staff_Licenses" };

    public static List<Staff> ExtractStaffRecords()
    {
        var result = new List<Staff>();
        var staffLookup = new Dictionary<long, Staff>();
        var validationErrors = new List<string>();

        // -----------------------------------------------------------------
        // 1. MASTER RECORDS (Columns A to J)
        // -----------------------------------------------------------------
        string masterSheet = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Staff_Master");
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
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long staffId) || staffId == 0)
                    continue;

                string familyName = CleanCellString(masterGrid[r, 1]);
                string middleName = CleanCellString(masterGrid[r, 2]);
                string givenName = CleanCellString(masterGrid[r, 3]);
                string nric = CleanCellString(masterGrid[r, 4]);
                string deptStr = CleanCellString(masterGrid[r, 5]);
                string payStr = CleanCellString(masterGrid[r, 6]);

                decimal.TryParse(CleanCellString(masterGrid[r, 7]), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal baseRate);
                decimal.TryParse(CleanCellString(masterGrid[r, 8]), NumberStyles.Any, CultureInfo.InvariantCulture, out decimal otRate);
                string statusStr = CleanCellString(masterGrid[r, 9]);

                // HARD VALIDATION: Must have at least a Given or Family name
                if (string.IsNullOrWhiteSpace(givenName) && string.IsNullOrWhiteSpace(familyName))
                {
                    validationErrors.Add($"• Row {r + 2} [Staff_Master]: Staff ID {staffId} has no Family or Given Name.");
                    continue;
                }

                // SMART FALLBACKS
                Department dept = Enum.TryParse(deptStr, true, out Department parsedDept) ? parsedDept : Department.Worker;
                PayType payType = Enum.TryParse(payStr, true, out PayType parsedPay) ? parsedPay : PayType.Hourly;
                ActiveStatus status = Enum.TryParse(statusStr, true, out ActiveStatus parsedStatus) ? parsedStatus : ActiveStatus.OnContract;

                if (baseRate < 0) baseRate = 0;
                if (otRate < 0) otRate = 0;

                var staff = new Staff
                {
                    StaffID = staffId,
                    Master = new Master
                    {
                        FamilyName = familyName,
                        MiddleName = middleName,
                        GivenName = givenName,
                        Name = $"{familyName} {givenName}".Trim(),
                        NRIC = nric,
                        Department = dept,
                        PayType = payType,
                        BaseRate = baseRate,
                        OvertimeRate = otRate,
                        ActiveStatus = status
                    },
                    Licenses = new List<License>()
                };

                staffLookup[staffId] = staff;
                result.Add(staff);
            }
        }

        // -----------------------------------------------------------------
        // 2. PREFERENCES (Columns A to K)
        // -----------------------------------------------------------------
        string prefSheet = WorkbookSession.ResolveSheetName(PreferenceSheetAliases, "Staff_Preference");
        IntPtr prefSheetId = WorkbookSession.GetSheetId(prefSheet);

        var prefRef = prefSheetId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, PreferenceColumnCount - 1, prefSheetId)
            : new ExcelReference(1, 500, 0, PreferenceColumnCount - 1);

        if (prefRef.GetValue() is object[,] prefGrid)
        {
            int rows = prefGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = prefGrid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long staffId) || staffId == 0)
                    continue;

                if (!staffLookup.TryGetValue(staffId, out var targetStaff))
                    continue;

                targetStaff.Preference.PrimaryTeam = CleanCellString(prefGrid[r, 2]);
                targetStaff.Preference.SecondaryTeam = CleanCellString(prefGrid[r, 3]);
                targetStaff.Preference.PrimarySkill = CleanCellString(prefGrid[r, 4]);
                targetStaff.Preference.SecondarySkill = CleanCellString(prefGrid[r, 5]);
                targetStaff.Preference.BlackListTask = CleanCellString(prefGrid[r, 6]);

                string shiftStr = CleanCellString(prefGrid[r, 7]);
                targetStaff.Preference.PreferredShift = Enum.TryParse(shiftStr, true, out Shift s) ? s : Shift.Morning;

                int.TryParse(CleanCellString(prefGrid[r, 8]), out int maxDays);
                targetStaff.Preference.MaxDaysPerWeek = (maxDays > 0 && maxDays <= 7) ? maxDays : 6;

                var validLicenses = targetStaff.Licenses
                    .Where(l => l.LicenseStatus == LicenseStatus.Valid)
                    .Select(l => l.LicenseType)
                    .ToHashSet();

                var prefVehicle = targetStaff.Preference.PreferredVehicleType;

                bool isAuthorized = prefVehicle switch
                {
                    VehicleType.None => true,
                    VehicleType.Van or VehicleType.Pickup => validLicenses.Contains(LicenseClass.D) || validLicenses.Contains(LicenseClass.GDL_D),
                    VehicleType.Lorry_1T => validLicenses.Contains(LicenseClass.GDL_D),
                    VehicleType.Lorry_3T or VehicleType.Lorry_5T => validLicenses.Contains(LicenseClass.GDL_E) || validLicenses.Contains(LicenseClass.GDL_E_Full),
                    VehicleType.Forklift => validLicenses.Contains(LicenseClass.Forklift),
                    _ => false
                };

                if (!isAuthorized)
                {
                    // Auto-reset or log conflict
                    targetStaff.Preference.PreferredVehicleType = VehicleType.None;
                    targetStaff.Preference.DedicatedVehicleID = null;
                }
            }
        }

        // -----------------------------------------------------------------
        // 3. LICENSES (Columns A to E)
        // -----------------------------------------------------------------
        string licenseSheet = WorkbookSession.ResolveSheetName(LicenseSheetAliases, "Staff_License");
        IntPtr licenseSheetId = WorkbookSession.GetSheetId(licenseSheet);

        var licenseRef = licenseSheetId != IntPtr.Zero
            ? new ExcelReference(1, 1000, 0, LicenseColumnCount - 1, licenseSheetId)
            : new ExcelReference(1, 1000, 0, LicenseColumnCount - 1);

        if (licenseRef.GetValue() is object[,] licenseGrid)
        {
            int rows = licenseGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawId = licenseGrid[r, 0];
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long staffId) || staffId == 0)
                    continue;

                if (!staffLookup.TryGetValue(staffId, out var targetStaff))
                {
                    validationErrors.Add($"• Row {r + 2} [Staff_License]: Staff ID {staffId} does not exist in Staff Master.");
                    continue;
                }

                string licStr = CleanCellString(licenseGrid[r, 2]);
                if (string.IsNullOrWhiteSpace(licStr)) continue;

                if (!Enum.TryParse(licStr, true, out LicenseClass licClass) || licClass == LicenseClass.None)
                {
                    validationErrors.Add($"• Row {r + 2} [Staff_License]: Unrecognized License Class '{licStr}' for Staff ID {staffId}.");
                    continue;
                }

                DateOnly expiry = ParseExcelDate(licenseGrid[r, 3]);
                // FALLBACK: If expiry not provided, default to 1 year ahead
                if (expiry == default || expiry <= DateOnly.FromDateTime(DateTime.Today.AddYears(-10)))
                {
                    expiry = DateOnly.FromDateTime(DateTime.Today.AddYears(1));
                }

                string statusStr = CleanCellString(licenseGrid[r, 4]);
                LicenseStatus status = Enum.TryParse(statusStr, true, out LicenseStatus parsedLicStatus)
                    ? parsedLicStatus
                    : (expiry < DateOnly.FromDateTime(DateTime.Today) ? LicenseStatus.Invalid : LicenseStatus.Valid);

                targetStaff.Licenses.Add(new License
                {
                    StaffID = staffId,
                    LicenseType = licClass,
                    ExpiryDate = expiry,
                    LicenseStatus = status
                });
            }
        }

        // -----------------------------------------------------------------
        // 4. PRE-SAVE GATEWAY: ABORT IF HARD ERRORS DETECTED
        // -----------------------------------------------------------------
        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save to database due to the following data validation issues:\n");
            foreach (var err in validationErrors)
            {
                sb.AppendLine(err);
            }
            sb.AppendLine("\nPlease correct these errors on the worksheet and try saving again.");

            MessageBox.Show(sb.ToString(), "Pre-Save Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return new List<Staff>(); // Returns empty list so save safely aborts without overwriting DB
        }

        return result;
    }

    public static void PopulateWorksheet(List<Staff> staffList)
    {
        if (staffList == null || staffList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "Staff_Master");
        string targetPref = WorkbookSession.ResolveSheetName(PreferenceSheetAliases, "Staff_Preference");
        string targetLic = WorkbookSession.ResolveSheetName(LicenseSheetAliases, "Staff_License");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);
            int rowCount = staffList.Count;

            // 1. Populate Master Table
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });
            var clearMasterRef = new ExcelReference(1, 500, 0, MasterColumnCount - 1);
            clearMasterRef.SetValue(new object[500, MasterColumnCount]);

            object[,] masterOutput = new object[rowCount, MasterColumnCount];
            for (int i = 0; i < rowCount; i++)
            {
                var s = staffList[i];
                masterOutput[i, 0] = s.StaffID;
                masterOutput[i, 1] = s.Master.FamilyName;
                masterOutput[i, 2] = s.Master.MiddleName;
                masterOutput[i, 3] = s.Master.GivenName;
                masterOutput[i, 4] = s.Master.NRIC;
                masterOutput[i, 5] = s.Master.Department == Department.None ? "Worker" : s.Master.Department.ToString();
                masterOutput[i, 6] = s.Master.PayType == PayType.None ? "Hourly" : s.Master.PayType.ToString();
                masterOutput[i, 7] = s.Master.BaseRate;
                masterOutput[i, 8] = s.Master.OvertimeRate;
                masterOutput[i, 9] = s.Master.ActiveStatus.ToString();
            }

            var writeMasterRef = new ExcelReference(1, rowCount, 0, MasterColumnCount - 1);
            writeMasterRef.SetValue(masterOutput);

            // 2. Populate Preference Table
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetPref });
            var clearPrefRef = new ExcelReference(1, 500, 0, PreferenceColumnCount - 1);
            clearPrefRef.SetValue(new object[500, PreferenceColumnCount]);

            object[,] prefOutput = new object[rowCount, PreferenceColumnCount];
            for (int i = 0; i < rowCount; i++)
            {
                var s = staffList[i];
                prefOutput[i, 0] = s.StaffID;
                prefOutput[i, 1] = s.Master.Name;
                prefOutput[i, 2] = s.Preference.PrimaryTeam;
                prefOutput[i, 3] = s.Preference.SecondaryTeam;
                prefOutput[i, 4] = s.Preference.PrimarySkill;
                prefOutput[i, 5] = s.Preference.SecondarySkill;
                prefOutput[i, 6] = s.Preference.BlackListTask;
                prefOutput[i, 7] = s.Preference.PreferredShift == Shift.None ? "Morning" : s.Preference.PreferredShift.ToString();
                prefOutput[i, 8] = s.Preference.MaxDaysPerWeek <= 0 ? 6 : s.Preference.MaxDaysPerWeek;
                prefOutput[i, 9] = s.Preference.PreferredVehicleType == VehicleType.None ? string.Empty : s.Preference.PreferredVehicleType.ToString();
                prefOutput[i, 10] = s.Preference.DedicatedVehicleID.HasValue ? s.Preference.DedicatedVehicleID.Value : string.Empty;
            }

            var writePrefRef = new ExcelReference(1, rowCount, 0, PreferenceColumnCount - 1);
            writePrefRef.SetValue(prefOutput);

            // 3. Populate Licenses Table (Columns A to E)
            var allLicenses = new List<(Staff Staff, License License)>();
            foreach (var s in staffList)
            {
                if (s.Licenses != null)
                {
                    foreach (var lic in s.Licenses)
                    {
                        allLicenses.Add((s, lic));
                    }
                }
            }

            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetLic });
            var clearLicRef = new ExcelReference(1, 1000, 0, LicenseColumnCount - 1);
            clearLicRef.SetValue(new object[1000, LicenseColumnCount]);

            if (allLicenses.Count > 0)
            {
                object[,] licOutput = new object[allLicenses.Count, LicenseColumnCount];
                for (int i = 0; i < allLicenses.Count; i++)
                {
                    var (staff, lic) = allLicenses[i];
                    licOutput[i, 0] = lic.StaffID;
                    licOutput[i, 1] = staff.Master.Name;
                    licOutput[i, 2] = lic.LicenseType == LicenseClass.None ? string.Empty : lic.LicenseType.ToString();
                    licOutput[i, 3] = lic.ExpiryDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    licOutput[i, 4] = lic.LicenseStatus.ToString();
                }

                var writeLicRef = new ExcelReference(1, allLicenses.Count, 0, LicenseColumnCount - 1);
                writeLicRef.SetValue(licOutput);
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
        if (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            return string.Empty;

        return cell.ToString()?.Trim() ?? string.Empty;
    }

    private static DateOnly ParseExcelDate(object? cell)
    {
        if (cell is double d) return DateOnly.FromDateTime(DateTime.FromOADate(d));
        if (DateTime.TryParse(CleanCellString(cell), out DateTime dt)) return DateOnly.FromDateTime(dt);
        return default;
    }
}