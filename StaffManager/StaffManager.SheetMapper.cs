namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;

public static class SheetMapper
{
    // All Staff: reading A through D (4 columns)
    private const int MasterColumnSpan = 4;
    // Vehicle Licenses: A through G (7 columns)
    private const int VehicleLicenseColumnSpan = 7;
    // Skill Licenses: A through G (7 columns)
    private const int SkillLicenseColumnSpan = 7;

    private static readonly string[] MasterSheetAliases = { "All Staff", "All_Staff", "Staff_Master" };
    private static readonly string[] VehicleLicenseSheetAliases = { "Vehicle Licenses", "Vehicle_Licenses", "Staff_License" };
    private static readonly string[] SkillLicenseSheetAliases = { "Skill Licenses", "Skill_Licenses", "Staff_Skill" };

    public static List<StaffRecord> ExtractStaffRecords()
    {
        var result = new List<StaffRecord>();
        var staffLookup = new Dictionary<long, StaffRecord>();
        var validationErrors = new List<string>();

        // -------------------------------------------------------------
        // 0. RETRIEVE DYNAMIC LICENSE TYPES LIST FROM EXCEL
        // -------------------------------------------------------------
        var allowedLicenseTypes = FetchAllowedSkillLicenseTypes();

        // -------------------------------------------------------------
        // 1. ALL STAFF (Sheet: "All Staff", Columns A to D)
        // -------------------------------------------------------------
        string masterSheet = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Staff");
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
                if (rawId == null || rawId is ExcelEmpty || !long.TryParse(rawId.ToString(), out long staffId) || staffId <= 0)
                    continue;

                string familyName = CleanCellString(masterGrid[r, 1]);
                string middleName = CleanCellString(masterGrid[r, 2]);
                string givenName = CleanCellString(masterGrid[r, 3]);

                if (string.IsNullOrWhiteSpace(familyName) && string.IsNullOrWhiteSpace(givenName))
                {
                    validationErrors.Add($"• Row {r + 2} [All Staff]: Staff ID {staffId} has no Family or Given Name.");
                    continue;
                }

                var staff = new StaffRecord
                {
                    StaffID = staffId,
                    StaffDetails = new StaffDetails
                    {
                        StaffID = staffId,
                        FamilyName = familyName,
                        MiddleName = middleName,
                        GivenName = givenName,
                        Department = Department.Worker,
                        PayType = PayType.Hourly,
                        BaseRate = 11.00m,
                        ContractStatus = ContractStatus.Active
                    },
                    VehicleLicenses = new List<VehicleLicense>(),
                    SkillLicenses = new List<SkillLicense>()
                };

                staffLookup[staffId] = staff;
                result.Add(staff);
            }
        }

        // -------------------------------------------------------------
        // 2. VEHICLE LICENSES (Sheet: "Vehicle Licenses", Columns A to G)
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

                object rawStaffId = vLicGrid[r, 1];
                if (rawStaffId == null || rawStaffId is ExcelEmpty || !long.TryParse(rawStaffId.ToString(), out long staffId) || staffId <= 0)
                    continue;

                if (!staffLookup.TryGetValue(staffId, out var targetStaff))
                {
                    validationErrors.Add($"• Row {r + 2} [Vehicle Licenses]: Staff ID {staffId} does not exist in 'All Staff'.");
                    continue;
                }

                string licClassStr = CleanCellString(vLicGrid[r, 3]);
                if (string.IsNullOrWhiteSpace(licClassStr))
                    continue;

                if (!Enum.TryParse(licClassStr, true, out LicenseClass licClass) || licClass == LicenseClass.None)
                {
                    validationErrors.Add($"• Row {r + 2} [Vehicle Licenses]: Invalid License_Class '{licClassStr}' for Staff ID {staffId}.");
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

                targetStaff.VehicleLicenses.Add(new VehicleLicense
                {
                    LicenseID = licId,
                    StaffID = staffId,
                    LicenseClass = licClass,
                    RenewDate = renewDate,
                    ExpiryDate = expiryDate,
                    LicenseStatus = licStatus
                });
            }
        }

        // -------------------------------------------------------------
        // 3. SKILL LICENSES (Sheet: "Skill Licenses", Columns A to G)
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

                object rawStaffId = sLicGrid[r, 1];
                if (rawStaffId == null || rawStaffId is ExcelEmpty || !long.TryParse(rawStaffId.ToString(), out long staffId) || staffId <= 0)
                    continue;

                if (!staffLookup.TryGetValue(staffId, out var targetStaff))
                {
                    validationErrors.Add($"• Row {r + 2} [Skill Licenses]: Staff ID {staffId} does not exist in 'All Staff'.");
                    continue;
                }

                string skillType = CleanCellString(sLicGrid[r, 3]);
                if (string.IsNullOrWhiteSpace(skillType))
                    continue;

                // VALIDATION CHECK: License Type must exist in LicenseTypesList
                if (allowedLicenseTypes.Count > 0 && !allowedLicenseTypes.Contains(skillType))
                {
                    string allowedJoined = string.Join(", ", allowedLicenseTypes);
                    validationErrors.Add($"• Row {r + 2} [Skill Licenses]: Unrecognized License Type '{skillType}' for Staff ID {staffId}. Allowed types: [{allowedJoined}].");
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

                targetStaff.SkillLicenses.Add(new SkillLicense
                {
                    LicenseID = licId,
                    StaffID = staffId,
                    LicenseType = skillType,
                    RenewDate = renewDate,
                    ExpiryDate = expiryDate,
                    LicenseStatus = licStatus
                });
            }
        }

        // -------------------------------------------------------------
        // PRE-SAVE VALIDATION GATEWAY: Prevent dirty database saves
        // -------------------------------------------------------------
        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save to system due to validation errors:\n");
            foreach (var err in validationErrors) sb.AppendLine(err);
            sb.AppendLine("\nPlease correct the entries and try saving again.");

            MessageBox.Show(sb.ToString(), "Pre-Save Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return new List<StaffRecord>();
        }

        return result;
    }

    public static void PopulateWorksheet(List<StaffRecord> staffList)
    {
        if (staffList == null || staffList.Count == 0) return;

        string targetMaster = WorkbookSession.ResolveSheetName(MasterSheetAliases, "All Staff");
        string targetVLic = WorkbookSession.ResolveSheetName(VehicleLicenseSheetAliases, "Vehicle Licenses");
        string targetSLic = WorkbookSession.ResolveSheetName(SkillLicenseSheetAliases, "Skill Licenses");

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            // 1. POPULATE ALL STAFF (Columns A to D)
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { targetMaster });
            var clearMasterRef = new ExcelReference(1, 500, 0, MasterColumnSpan - 1);
            clearMasterRef.SetValue(new object[500, MasterColumnSpan]);

            int rowCount = staffList.Count;
            object[,] masterOutput = new object[rowCount, MasterColumnSpan];

            for (int i = 0; i < rowCount; i++)
            {
                var s = staffList[i];
                masterOutput[i, 0] = s.StaffID;
                masterOutput[i, 1] = s.StaffDetails.FamilyName;
                masterOutput[i, 2] = s.StaffDetails.MiddleName;
                masterOutput[i, 3] = s.StaffDetails.GivenName;
            }

            var writeMasterRef = new ExcelReference(1, rowCount, 0, MasterColumnSpan - 1);
            writeMasterRef.SetValue(masterOutput);

            // 2. POPULATE VEHICLE LICENSES (Columns A to G)
            var allVLicenses = new List<(StaffRecord Staff, VehicleLicense Lic)>();
            foreach (var s in staffList)
            {
                if (s.VehicleLicenses != null)
                {
                    foreach (var lic in s.VehicleLicenses)
                    {
                        allVLicenses.Add((s, lic));
                    }
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
                    var (staff, lic) = allVLicenses[i];
                    vOutput[i, 0] = lic.LicenseID > 0 ? (object)lic.LicenseID : string.Empty;
                    vOutput[i, 1] = lic.StaffID;
                    vOutput[i, 2] = staff.StaffDetails.Name;
                    vOutput[i, 3] = lic.LicenseClass == LicenseClass.None ? string.Empty : lic.LicenseClass.ToString();
                    vOutput[i, 4] = lic.RenewDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    vOutput[i, 5] = lic.ExpiryDate.ToDateTime(TimeOnly.MinValue).ToOADate();
                    vOutput[i, 6] = lic.LicenseStatus.ToString();
                }

                var writeVRef = new ExcelReference(1, allVLicenses.Count, 0, VehicleLicenseColumnSpan - 1);
                writeVRef.SetValue(vOutput);
            }

            // 3. POPULATE SKILL LICENSES (Columns A to G)
            var allSLicenses = new List<(StaffRecord Staff, SkillLicense Lic)>();
            foreach (var s in staffList)
            {
                if (s.SkillLicenses != null)
                {
                    foreach (var lic in s.SkillLicenses)
                    {
                        allSLicenses.Add((s, lic));
                    }
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
                    var (staff, lic) = allSLicenses[i];
                    sOutput[i, 0] = lic.LicenseID > 0 ? (object)lic.LicenseID : string.Empty;
                    sOutput[i, 1] = lic.StaffID;
                    sOutput[i, 2] = staff.StaffDetails.Name;
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

    /// <summary>
    /// Evaluates the dynamic =LicenseTypesList defined name in Excel to enforce validation in C#.
    /// </summary>
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