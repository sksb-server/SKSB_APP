namespace SKSB_App.ScheduleManager;

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
    // Sheet Aliases
    private static readonly string[] ScheduleAliases = { "Daily Schedule", "Daily_Schedule", "Schedule" };
    private static readonly string[] AttendanceAliases = { "Daily Attendance", "Daily_Attendance", "Attendance" };
    private static readonly string[] PreferenceAliases = { "Preference", "Worker Preference", "Preferences" };

    public static ScheduleRecord? ExtractScheduleRecord(DateOnly scheduleDate)
    {
        var record = new ScheduleRecord { DailyScheduleDate = scheduleDate };
        var validationErrors = new List<string>();

        // -------------------------------------------------------------
        // 1. INGEST PROJECTS (Sheet: "Daily Schedule")
        // Columns: Project ID (A) | Shift (B) | Workers CSV (C) | Vehicles CSV (D) | Remarks (E)
        // -------------------------------------------------------------
        string schedSheet = WorkbookSession.ResolveSheetName(ScheduleAliases, "Daily Schedule");
        IntPtr schedId = WorkbookSession.GetSheetId(schedSheet);
        var schedRef = schedId != IntPtr.Zero
            ? new ExcelReference(1, 200, 0, 4, schedId)
            : new ExcelReference(1, 200, 0, 4);

        if (schedRef.GetValue() is object[,] schedGrid)
        {
            int rows = schedGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                string projId = CleanCellString(schedGrid[r, 0]);
                if (string.IsNullOrWhiteSpace(projId)) continue;

                string shiftStr = CleanCellString(schedGrid[r, 1]);
                string workersCsv = CleanCellString(schedGrid[r, 2]);
                string vehiclesCsv = CleanCellString(schedGrid[r, 3]);
                string remarks = CleanCellString(schedGrid[r, 4]);

                var workerList = workersCsv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();
                var vehicleList = vehiclesCsv.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries).ToList();

                if (workerList.Count == 0)
                {
                    validationErrors.Add($"• Row {r + 2} [Daily Schedule]: Project '{projId}' has no assigned workers.");
                }

                record.Projects.Add(new DailyProjectSchedule
                {
                    ProjectID = projId,
                    Shift = Enum.TryParse(shiftStr, true, out Shifts s) ? s : Shifts.Morning,
                    Workers = workerList,
                    Vehicles = vehicleList,
                    Remarks = remarks
                });
            }
        }

        // -------------------------------------------------------------
        // 2. INGEST ATTENDANCE (Sheet: "Daily Attendance")
        // Col A: Worker ID | B: Name (Skip) | C: Skill (Skip) | D: Leave (Skip) | E: ClockIn | F: ClockOut | G: Biometric | H: Override | I: Reason | J: AttendanceState | K: IsAllocated
        // -------------------------------------------------------------
        string attSheet = WorkbookSession.ResolveSheetName(AttendanceAliases, "Daily Attendance");
        IntPtr attId = WorkbookSession.GetSheetId(attSheet);
        var attRef = attId != IntPtr.Zero
            ? new ExcelReference(1, 500, 0, 10, attId)
            : new ExcelReference(1, 500, 0, 10);

        if (attRef.GetValue() is object[,] attGrid)
        {
            int rows = attGrid.GetLength(0);
            for (int r = 0; r < rows; r++)
            {
                object rawWorkerId = attGrid[r, 0];
                if (rawWorkerId == null || rawWorkerId is ExcelEmpty || !long.TryParse(rawWorkerId.ToString(), out long workerId) || workerId <= 0)
                    continue;

                TimeOnly clockIn = ParseTime(attGrid[r, 4]);
                TimeOnly clockOut = ParseTime(attGrid[r, 5]);
                string bioStr = CleanCellString(attGrid[r, 6]);
                string overrideStr = CleanCellString(attGrid[r, 7]);
                string reason = CleanCellString(attGrid[r, 8]);
                string attStateStr = CleanCellString(attGrid[r, 9]);
                bool.TryParse(CleanCellString(attGrid[r, 10]), out bool isAllocated);

                AttendanceStatus manualOverride = Enum.TryParse(overrideStr, true, out AttendanceStatus mo) ? mo : AttendanceStatus.Absence;
                if (manualOverride != AttendanceStatus.Absence && string.IsNullOrWhiteSpace(reason))
                {
                    validationErrors.Add($"• Row {r + 2} [Daily Attendance]: Worker ID {workerId} has a manual override without an override reason.");
                }

                record.WorkerAttendance.Add(new DailyWorkerAttendance
                {
                    DailyScheduleDate = scheduleDate,
                    WorkerID = workerId,
                    ClockInTime = clockIn,
                    ClockOutTime = clockOut,
                    BiometricStatus = Enum.TryParse(bioStr, true, out BiometricStatus b) ? b : BiometricStatus.None,
                    ManualOverride = manualOverride,
                    OverrideReason = reason,
                    AttendanceStatus = Enum.TryParse(attStateStr, true, out AttendanceStatus ast) ? ast : AttendanceStatus.Absence,
                    IsAllocated = isAllocated
                });
            }
        }

        // -------------------------------------------------------------
        // PRE-SAVE VALIDATION GATEWAY: Halt persistence if data is dirty
        // -------------------------------------------------------------
        if (validationErrors.Count > 0)
        {
            var sb = new StringBuilder();
            sb.AppendLine("Cannot save schedule to system due to validation errors:\n");
            foreach (var err in validationErrors) sb.AppendLine(err);
            sb.AppendLine("\nPlease correct the entries and try saving again.");

            MessageBox.Show(sb.ToString(), "Pre-Save Validation Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return null;
        }

        return record;
    }

    public static void PopulateScheduleSheets(ScheduleRecord? record)
    {
        if (record == null) return;

        object initialActiveSheet = XlCall.Excel(XlCall.xlfGetDocument, 1);

        try
        {
            XlCall.Excel(XlCall.xlcEcho, false);

            // 1. Populate Daily Schedule Tab
            string schedSheet = WorkbookSession.ResolveSheetName(ScheduleAliases, "Daily Schedule");
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { schedSheet });
            WorkbookSession.ClearActiveSheetDataRows(1, 200, 0, 4);

            if (record.Projects.Count > 0)
            {
                object[,] schedOut = new object[record.Projects.Count, 5];
                for (int i = 0; i < record.Projects.Count; i++)
                {
                    var p = record.Projects[i];
                    schedOut[i, 0] = p.ProjectID;
                    schedOut[i, 1] = p.Shift.ToString();
                    schedOut[i, 2] = string.Join(", ", p.Workers);
                    schedOut[i, 3] = string.Join(", ", p.Vehicles);
                    schedOut[i, 4] = p.Remarks;
                }
                new ExcelReference(1, record.Projects.Count, 0, 4).SetValue(schedOut);
            }

            // 2. Populate Attendance Tab (Formula columns index 1, 2, 3 left blank to let Excel formulas evaluate)
            string attSheet = WorkbookSession.ResolveSheetName(AttendanceAliases, "Daily Attendance");
            XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { attSheet });
            WorkbookSession.ClearActiveSheetDataRows(1, 500, 0, 10);

            if (record.WorkerAttendance.Count > 0)
            {
                object[,] attOut = new object[record.WorkerAttendance.Count, 11];
                for (int i = 0; i < record.WorkerAttendance.Count; i++)
                {
                    var a = record.WorkerAttendance[i];
                    attOut[i, 0] = a.WorkerID;
                    // attOut[i, 1] -> Name (Formula Lookup)
                    // attOut[i, 2] -> Primary Skill (Formula Lookup)
                    // attOut[i, 3] -> Leave Status (Formula Lookup)
                    attOut[i, 4] = a.ClockInTime == TimeOnly.MinValue ? string.Empty : a.ClockInTime.ToString("HH:mm");
                    attOut[i, 5] = a.ClockOutTime == TimeOnly.MaxValue ? string.Empty : a.ClockOutTime.ToString("HH:mm");
                    attOut[i, 6] = a.BiometricStatus.ToString();
                    attOut[i, 7] = a.ManualOverride.ToString();
                    attOut[i, 8] = a.OverrideReason;
                    attOut[i, 9] = a.AttendanceStatus.ToString();
                    attOut[i, 10] = a.IsAllocated;
                }
                new ExcelReference(1, record.WorkerAttendance.Count, 0, 10).SetValue(attOut);
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

    private static string CleanCellString(object? cell) =>
        (cell == null || cell is ExcelEmpty || cell is ExcelMissing || cell is ExcelError)
            ? string.Empty
            : cell.ToString()?.Trim() ?? string.Empty;

    private static TimeOnly ParseTime(object cell)
    {
        string text = CleanCellString(cell);
        return TimeOnly.TryParse(text, out TimeOnly t) ? t : TimeOnly.MinValue;
    }
}