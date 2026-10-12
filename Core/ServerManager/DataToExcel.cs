namespace SKSB_App.Core.ServerManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.WorkbookManager;

public static class DataToExcel
{
    public const string DataSheetName = "Data";

    /// <summary>
    /// Loads entities of type T from disk/database via Services.Load<T>(),
    /// projects the target string/numeric property, and synchronises it into the 'Data' sheet.
    /// Example: DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate);
    /// </summary>
    public static void SyncData<TEntity, TProp>(string namePrefix, Func<TEntity, TProp> propertySelector)
    {
        var entities = Services.Load<List<TEntity>>();
        if (entities == null || entities.Count == 0) return;

        string[] values = entities
            .Select(propertySelector)
            .Where(val => val != null)
            .Select(val => val!.ToString()!.Trim())
            .Where(val => !string.IsNullOrEmpty(val) && val != "0")
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v)
            .ToArray();

        string listName = namePrefix.EndsWith("List", StringComparison.OrdinalIgnoreCase)
            ? namePrefix
            : $"{namePrefix}List";

        SyncValues(listName, values);
    }

    /// <summary>
    /// Writes values to the dedicated 'Data' sheet using smart column placement:
    /// - If column is taken by ANOTHER data list -> Skip.
    /// - If column is taken by the SAME data list -> Wipe and overwrite in-place.
    /// - If column is blank -> Write to this slot.
    /// </summary>
    public static void SyncValues(string listName, string[] values)
    {
        if (string.IsNullOrWhiteSpace(listName) || values == null || values.Length == 0) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                XlCall.Excel(XlCall.xlcEcho, false);

                // 1. Remember original active sheet to restore focus cleanly
                object activeDoc = null;
                try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
                string originalSheet = activeDoc?.ToString() ?? string.Empty;

                // 2. Ensure Data sheet exists and is unhidden
                EnsureDataSheetExists();
                MakeSheetVisibleAndActive(DataSheetName);

                // 3. Auto-organize column inspection
                // Scan columns (0, 2, 4, 6...) looking for matching header or first blank slot
                int targetCol = -1;
                int maxColsToScan = 200;

                for (int c = 0; c < maxColsToScan; c += 2)
                {
                    var headerCell = new ExcelReference(0, 0, c, c);
                    object cellVal = headerCell.GetValue();
                    string existingHeader = cellVal?.ToString()?.Trim() ?? string.Empty;

                    if (string.Equals(existingHeader, listName, StringComparison.OrdinalIgnoreCase))
                    {
                        // Match found: erase and rewrite this column in-place
                        targetCol = c;
                        break;
                    }

                    if (string.IsNullOrEmpty(existingHeader) && targetCol == -1)
                    {
                        // Remember first empty column slot
                        targetCol = c;
                    }
                }

                if (targetCol == -1)
                {
                    targetCol = 0; // Fallback
                }

                // 4. Wipe rows (0 to 2000) for this column to erase stale entries
                var wipeRef = new ExcelReference(0, 2000, targetCol, targetCol);
                object[,] clearGrid = new object[2001, 1];
                for (int r = 0; r <= 2000; r++) clearGrid[r, 0] = string.Empty;
                wipeRef.SetValue(clearGrid);

                // 5. Write Column Header (Row 0)
                var headerRef = new ExcelReference(0, 0, targetCol, targetCol);
                headerRef.SetValue(new object[,] { { listName } });

                // 6. Write Data Rows (Row 1 to values.Length)
                int rowCount = values.Length;
                var valRef = new ExcelReference(1, rowCount, targetCol, targetCol);
                object[,] cellData = new object[rowCount, 1];
                for (int i = 0; i < rowCount; i++)
                {
                    cellData[i, 0] = values[i];
                }
                valRef.SetValue(cellData);

                // 7. Re-register Excel Defined Name: ='Data'!R2C[col]:R[rowCount+1]C[col]
                try { XlCall.Excel(XlCall.xlcDeleteName, listName); } catch { }

                int r1c1StartRow = 2;
                int r1c1EndRow = rowCount + 1;
                int r1c1Col = targetCol + 1; // 1-based R1C1 index
                string refersTo = $"='{DataSheetName}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";

                XlCall.Excel(XlCall.xlcDefineName, listName, refersTo, 3);

                // 8. Move 'Data' sheet behind operational sheets, alongside Lookups
                MoveDataSheetToEnd();

                // 9. Hide 'Data' sheet if NOT in Template Edit Mode
                if (!WorkbookSession.IsTemplateEditMode)
                {
                    try
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { DataSheetName });
                        XlCall.Excel(XlCall.xlcWorkbookHide);
                    }
                    catch { }
                }

                // 10. Restore focus back to caller's original worksheet
                if (!string.IsNullOrWhiteSpace(originalSheet))
                {
                    string cleanOriginal = originalSheet.Contains(']')
                        ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                        : originalSheet;

                    if (!cleanOriginal.Equals(DataSheetName, StringComparison.OrdinalIgnoreCase))
                    {
                        try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[DataToExcel Error] {ex.Message}");
            }
            finally
            {
                XlCall.Excel(XlCall.xlcEcho, true);
            }
        });
    }

    private static void MakeSheetVisibleAndActive(string sheetName)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic ws = xlApp.ActiveWorkbook.Sheets[sheetName];
            if (ws != null)
            {
                ws.Visible = -1; // xlSheetVisible
                ws.Activate();
                return;
            }
        }
        catch { }

        try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { sheetName }); } catch { }
    }

    private static void MoveDataSheetToEnd()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            dynamic sheets = wb.Sheets;
            dynamic dataWs = sheets[DataSheetName];
            if (dataWs != null)
            {
                dataWs.Move(After: sheets[sheets.Count]);
            }
        }
        catch { }
    }

    private static void EnsureDataSheetExists()
    {
        IntPtr existingId = WorkbookSession.GetSheetId(DataSheetName);
        if (existingId != IntPtr.Zero) return;

        try
        {
            object sheetsObj = XlCall.Excel(XlCall.xlfGetWorkbook, 1);
            if (sheetsObj is object[,] sheetArray)
            {
                int dim1 = sheetArray.GetLength(1);
                for (int c = 0; c < dim1; c++)
                {
                    string raw = sheetArray[0, c]?.ToString() ?? string.Empty;
                    int bIdx = raw.LastIndexOf(']');
                    string clean = bIdx >= 0 ? raw.Substring(bIdx + 1) : raw;

                    if (string.Equals(clean, DataSheetName, StringComparison.OrdinalIgnoreCase)) return;
                }
            }
        }
        catch { }

        try
        {
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);

            object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string fullName = docInfo?.ToString() ?? string.Empty;
            string tempName = fullName.Contains(']')
                ? fullName.Substring(fullName.LastIndexOf(']') + 1)
                : fullName;

            XlCall.Excel(XlCall.xlcWorkbookName, tempName, DataSheetName);
        }
        catch { }
    }
}