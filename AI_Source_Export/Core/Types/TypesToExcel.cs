namespace SKSB_App.Core.Services;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookHelper;

public static class TypesToExcel
{
    private const string TargetSheet = SheetOps.LookupsSheetName;

    /// <summary>
    /// Auto-targets and synchronises ALL public Enums found in Core and loaded domain assemblies.
    /// </summary>
    public static void SyncAllCoreEnums(params Type[] excludeTypes)
    {
        var excludeSet = new HashSet<Type>(excludeTypes ?? Array.Empty<Type>());

        // 1. Scan the Core assembly + any loaded SKSB assemblies in the AppDomain
        var targetAssemblies = AppDomain.CurrentDomain.GetAssemblies()
            .Where(a => !a.IsDynamic && (a.FullName?.StartsWith("SKSB_App", StringComparison.OrdinalIgnoreCase) ?? false))
            .Distinct()
            .ToList();

        // Ensure Core is included even if AppDomain filtering misses it
        var coreAssembly = typeof(TypesToExcel).Assembly;
        if (!targetAssemblies.Contains(coreAssembly))
        {
            targetAssemblies.Add(coreAssembly);
        }

        var allEnums = new List<Type>();
        foreach (var asm in targetAssemblies)
        {
            try
            {
                var enums = asm.GetTypes()
                    .Where(t => t.IsEnum && t.IsPublic && !excludeSet.Contains(t));
                allEnums.AddRange(enums);
            }
            catch (ReflectionTypeLoadException ex)
            {
                var loadable = ex.Types.Where(t => t != null && t.IsEnum && t.IsPublic && !excludeSet.Contains(t!));
                allEnums.AddRange(loadable!);
            }
            catch { }
        }

        var uniqueEnums = allEnums
            .GroupBy(t => t.Name)
            .Select(g => g.First())
            .OrderBy(t => t.Name)
            .ToArray();

        if (uniqueEnums.Length == 0)
        {
            System.Diagnostics.Debug.WriteLine("[TypesToExcel] Warning: No public enums discovered.");
            return;
        }

        Sync(uniqueEnums);
    }

    /// <summary>
    /// Synchronises a specific collection of Enum types into defined names on the Lookups sheet.
    /// Named ranges are automatically formatted as [EnumName]List.
    /// </summary>
    public static void Sync(params Type[] enumTypes)
    {
        if (enumTypes == null || enumTypes.Length == 0) return;

        var mappings = new Dictionary<string, string[]>();

        foreach (var type in enumTypes)
        {
            if (!type.IsEnum) continue;

            string listName = $"{type.Name}List";
            string[] values = Enum.GetNames(type)
                                  .Where(n => !string.Equals(n, "None", StringComparison.OrdinalIgnoreCase))
                                  .ToArray();

            mappings[listName] = values;
        }

        SyncDictionary(mappings);
    }

    public static void SyncDictionary(Dictionary<string, string[]> enumMappings)
    {
        if (enumMappings == null || enumMappings.Count == 0) return;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                XlCall.Excel(XlCall.xlcEcho, false);

                // 1. Remember currently active sheet to restore focus cleanly
                object? activeDoc = null;
                try { activeDoc = XlCall.Excel(XlCall.xlfGetDocument, 76); } catch { }
                string originalSheet = activeDoc?.ToString() ?? string.Empty;

                // 2. Ensure Lookups exists
                EnsureLookupsSheetExists();

                // 3. Make Lookups visible and active so ExcelReference.SetValue can write cleanly
                MakeSheetVisibleAndActive(TargetSheet);

                int colZero = 0;

                foreach (var (name, values) in enumMappings)
                {
                    int rowCount = values.Length;

                    // Clear previous rows (0 to 100)
                    var wipeRef = new ExcelReference(0, 100, colZero, colZero);
                    object[,] clearGrid = new object[101, 1];
                    for (int r = 0; r <= 100; r++) clearGrid[r, 0] = string.Empty;
                    wipeRef.SetValue(clearGrid);

                    // Write Column Header
                    var headerRef = new ExcelReference(0, 0, colZero, colZero);
                    headerRef.SetValue(new object[,] { { name } });

                    // Write Data Rows and Define Name
                    if (rowCount > 0)
                    {
                        var valRef = new ExcelReference(1, rowCount, colZero, colZero);
                        object[,] cellData = new object[rowCount, 1];
                        for (int i = 0; i < rowCount; i++) cellData[i, 0] = values[i];
                        valRef.SetValue(cellData);

                        try { XlCall.Excel(XlCall.xlcDeleteName, name); } catch { }

                        int r1c1StartRow = 2;
                        int r1c1EndRow = rowCount + 1;
                        int r1c1Col = colZero + 1;
                        string refersTo = $"='{TargetSheet}'!R{r1c1StartRow}C{r1c1Col}:R{r1c1EndRow}C{r1c1Col}";

                        XlCall.Excel(XlCall.xlcDefineName, name, refersTo, 3);
                    }

                    colZero += 2; // Spacer column
                }

                // 4. Move auxiliary sheets ('Lookups' and 'Data') to the end of the tabs
                SheetOps.MoveAuxiliarySheetsToEnd();

                // 5. Hide Lookups if NOT in template edit mode
                if (!SheetOps.IsTemplateEditMode)
                {
                    try
                    {
                        XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { TargetSheet });
                        XlCall.Excel(XlCall.xlcWorkbookHide);
                    }
                    catch { }
                }

                // 6. Restore original active sheet focus
                if (!string.IsNullOrWhiteSpace(originalSheet))
                {
                    string cleanOriginal = originalSheet.Contains(']')
                        ? originalSheet.Substring(originalSheet.LastIndexOf(']') + 1)
                        : originalSheet;

                    if (!cleanOriginal.Equals(TargetSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        try { XlCall.Excel(XlCall.xlcWorkbookSelect, new object[] { cleanOriginal }); } catch { }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[TypesToExcel Error] {ex.Message}");
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

    private static void EnsureLookupsSheetExists()
    {
        // 1. Check if Lookups already exists across all sheets (visible or hidden) without GetSheetId
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

                    if (string.Equals(clean, TargetSheet, StringComparison.OrdinalIgnoreCase))
                    {
                        SheetOps.MoveAuxiliarySheetsToEnd();
                        return;
                    }
                }
            }
        }
        catch { }

        // 2. Only insert if the sheet does not exist
        try
        {
            XlCall.Excel(XlCall.xlcWorkbookInsert, 1);

            object docInfo = XlCall.Excel(XlCall.xlfGetDocument, 76);
            string fullName = docInfo?.ToString() ?? string.Empty;
            string tempName = fullName.Contains(']')
                ? fullName.Substring(fullName.LastIndexOf(']') + 1)
                : fullName;

            XlCall.Excel(XlCall.xlcWorkbookName, tempName, TargetSheet);
            SheetOps.MoveAuxiliarySheetsToEnd();
        }
        catch { }
    }
}