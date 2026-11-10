namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

internal static partial class SheetMapper
{
    public const string MasterSheetName = "All Vehicles";
    public const string MasterTableName = "AllVehiclesTable";

    private static readonly string[] MasterSheetAliases = { "All Vehicles", "AllVehicles", "Vehicle_Master", "Vehicles" };
    private static readonly string[] MaintenanceAliases = { "Maintenances", "Maintenance", "Vehicle_Maintenance" };
    private static readonly string[] WorkshopAliases = { "Workshops", "Workshop" };
    private static readonly string[] RoadTaxAliases = { "Road Taxes", "RoadTaxes", "Road Tax" };
    private static readonly string[] InspectionAliases = { "Inspections", "Inspection" };
    private static readonly string[] PermitAliases = { "Permits", "Permit" };
    private static readonly string[] InsuranceAliases = { "Insurances", "Insurance" };
    private static readonly string[] OdometerAliases = { "Odometers", "Odometer", "Fuel Logs", "Fuel" };

    // =========================================================================
    // 1. MASTER EXTRACTION PIPELINE
    // =========================================================================
    internal static (List<VehicleRecord>? Data, List<string> Errors, List<string> Warnings) ExtractVehicleRecordsWithDiagnostics()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return (null, new List<string> { "No active workbook found." }, new List<string>());

            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs == null)
            {
                return (null, new List<string> { $"Worksheet '{MasterSheetName}' was not found in workbook." }, new List<string>());
            }

            // 1. Extract Master Vehicle Records
            List<VehicleRecord> vehicles = TableMapper.ExtractFromTable<VehicleRecord>((object)masterWs, MasterTableName);

            // 2. Extract Child Collections
            ExtractChildRecords((object)wb, vehicles);

            // 3. Extract & Validate Auxiliary Tables
            var (auxErrors, auxWarnings) = ExtractAndPersistAuxiliaryTables((object)wb);

            // 4. Validate All Tables Across All Sheets
            var (errors, warnings) = RulesExtractor.ValidateAllWorkbookTables((object)wb);
            errors.AddRange(auxErrors);
            warnings.AddRange(auxWarnings);

            // 5. Primary Key Integrity Check
            var plateGroups = vehicles
                .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                .GroupBy(v => v.VehiclePlate, StringComparer.OrdinalIgnoreCase);

            foreach (var group in plateGroups)
            {
                if (group.Count() > 1)
                {
                    errors.Add($"Duplicate Vehicle Plate '{group.Key}' detected {group.Count()} times in master table.");
                }
            }

            return (vehicles, errors, warnings);
        }
        catch (Exception ex)
        {
            return (null, new List<string> { $"Failed extracting vehicle records: {ex.Message}" }, new List<string>());
        }
    }

    internal static List<VehicleRecord> ExtractVehicleRecords()
    {
        var (data, _, _) = ExtractVehicleRecordsWithDiagnostics();
        return data ?? new List<VehicleRecord>();
    }

    // =========================================================================
    // 2. MASTER POPULATE PIPELINE
    // =========================================================================
    internal static void PopulateWorksheet(List<VehicleRecord> records)
    {
        if (records == null) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        bool origEvents = true;
        bool origUpdating = true;
        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }

        try
        {
            dynamic wb = xlApp.ActiveWorkbook;
            if (wb == null) return;

            // 1. Populate Master Table
            dynamic? masterWs = ResolveSheetSafely((object)wb, MasterSheetAliases, MasterSheetName);
            if (masterWs != null)
            {
                try { masterWs.Unprotect("SKSB_Internal_Guard"); } catch { }

                if (masterWs.Cells[1, 2].Value2 == null)
                    masterWs.Cells[1, 2].Value2 = DateTime.Today.ToOADate();

                if (string.IsNullOrWhiteSpace(masterWs.Cells[1, 3].Value2?.ToString()))
                    masterWs.Cells[1, 3].Value2 = "All_Time";

                TableMapper.PopulateTableSafe<VehicleRecord>((object)masterWs, MasterTableName, records);
                SheetProtections.ApplyStyleDrivenProtection(masterWs);
            }

            // 2. Populate Child Tables
            PopulateChildTables((object)wb, records);

            // 3. Populate Auxiliary Tables
            PopulateAuxiliaryTables((object)wb);

            try
            {
                xlApp.CalculateFull();
            }
            catch (Exception exCalc)
            {
                System.Diagnostics.Debug.WriteLine($"[PopulateWorksheet Calculation Warning]: {exCalc.Message}");
            }

            string[] tableEvalOrder = new[]
        {
            "AllVehiclesTable",
            "WorkshopsTable",
            "InsuranceCompaniesTable",
            "FuelCardsTable",
            "MaintenancesTable",
            "RoadTaxesTable",
            "InspectionsTable",
            "PermitsTable",
            "InsurancesTable",
            "OdometersTable",
            "AllVehiclesTable"
        };

            foreach (string tName in tableEvalOrder)
            {
                foreach (dynamic sheet in wb.Worksheets)
                {
                    try
                    {
                        dynamic tbl = sheet.ListObjects[tName];
                        if (tbl != null)
                        {
                            RuleEvaluator.EvaluateAndHealTable(sheet, tbl);
                            break; // Found the table, move to the next one
                        }
                    }
                    catch { } // Table not on this specific sheet, loop continues
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[VehicleSheetMapper.PopulateWorksheet Error] {ex.Message}");
        }
        finally
        {
            try { xlApp.EnableEvents = origEvents; } catch { }
            try { xlApp.ScreenUpdating = origUpdating; } catch { }
        }
    }

    // =========================================================================
    // 3. SHARED UTILITIES
    // =========================================================================
    internal static dynamic? ResolveSheetSafely(object wbObj, string[] aliases, string defaultName)
    {
        dynamic wb = wbObj;
        string resolvedName = SheetOperations.ResolveSheetName(aliases, defaultName);
        try
        {
            dynamic s = wb.Sheets[resolvedName];
            if (s != null) return s;
        }
        catch { }

        foreach (dynamic ws in wb.Sheets)
        {
            string name = ws.Name?.ToString() ?? string.Empty;
            if (aliases.Any(a => string.Equals(a, name, StringComparison.OrdinalIgnoreCase)))
            {
                return ws;
            }
        }
        return null;
    }

    internal static string NormalizePlate(string plate) =>
        Regex.Replace(plate ?? string.Empty, @"[\s_\-]+", "").ToUpperInvariant();

    internal static string FormatPlate(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        string clean = NormalizePlate(raw);
        var m = Regex.Match(clean, @"^([A-Z]{1,3})(\d{1,4})([A-Z]?)$");
        if (m.Success)
        {
            return $"{m.Groups[1].Value} {m.Groups[2].Value} {m.Groups[3].Value}".Trim();
        }
        return clean;
    }

    internal static T GetProp<T>(object target, params string[] propNames)
    {
        if (target == null) return default!;
        Type type = target.GetType();
        foreach (var name in propNames)
        {
            var p = type.GetProperty(name, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
            if (p != null)
            {
                object? val = p.GetValue(target);
                if (val != null)
                {
                    try { return (T)Convert.ChangeType(val, typeof(T)); }
                    catch { }
                }
            }
        }
        return default!;
    }

    internal static void SetProp(object target, string propName, object? value)
    {
        if (target == null || value == null) return;
        var p = target.GetType().GetProperty(propName, BindingFlags.Public | BindingFlags.Instance | BindingFlags.IgnoreCase);
        if (p != null && p.CanWrite)
        {
            try
            {
                Type uType = Nullable.GetUnderlyingType(p.PropertyType) ?? p.PropertyType;
                if (uType == typeof(DateOnly) && value is DateTime dt)
                    p.SetValue(target, DateOnly.FromDateTime(dt));
                else if (uType == typeof(TimeOnly) && value is string st && TimeOnly.TryParse(st, out var to))
                    p.SetValue(target, to);
                else if (uType.IsEnum)
                {
                    if (Enum.TryParse(uType, value.ToString(), true, out var ev))
                        p.SetValue(target, ev);
                }
                else
                {
                    p.SetValue(target, Convert.ChangeType(value, uType));
                }
            }
            catch { }
        }
    }

#if DEBUG
    private static void ShowPopulateDiagnostics(object wbObj, List<VehicleRecord> records)
    {
        try
        {
            if (!DebugDialog.DebugEnabled) return;

            dynamic wb = wbObj;
            var sb = new System.Text.StringBuilder();
            sb.AppendLine("=== VEHICLE MANAGER POPULATE AUDIT ===");
            sb.AppendLine($"Total Vehicle Profiles Loaded: {records.Count}");
            sb.AppendLine();

            string[] inspectTables =
            {
                "AllVehiclesTable", "MaintenancesTable", "WorkshopsTable", "RoadTaxesTable",
                "InspectionsTable", "PermitsTable", "InsurancesTable", "InsuranceCompaniesTable",
                "OdometersTable", "FuelCardsTable", "FuelCompanyNamesTable"
            };

            foreach (var tName in inspectTables)
            {
                dynamic? tbl = null;
                string sName = "Unknown";
                foreach (dynamic ws in wb.Sheets)
                {
                    try
                    {
                        tbl = ws.ListObjects[tName];
                        if (tbl != null) { sName = ws.Name?.ToString() ?? "Unknown"; break; }
                    }
                    catch { }
                }

                if (tbl != null)
                {
                    int cols = (int)tbl.ListColumns.Count;
                    int rows = tbl.DataBodyRange != null ? (int)tbl.DataBodyRange.Rows.Count : 0;
                    int startCol = (int)tbl.HeaderRowRange.Column;
                    int startRow = (int)tbl.HeaderRowRange.Row;
                    sb.AppendLine($"  [{tName}] on Sheet '{sName}':");
                    sb.AppendLine($"      Dimensions : {rows} data row(s) × {cols} column(s)");
                    sb.AppendLine($"      Coordinates: Row {startRow}, Columns {startCol} to {startCol + cols - 1}");
                    sb.AppendLine();
                }
                else
                {
                    sb.AppendLine($"  [{tName}]: NOT FOUND on any worksheet");
                    sb.AppendLine();
                }
            }

            DebugDialog.ShowLog("Vehicle Manager - Populate Audit", sb.ToString());
        }
        catch { }
    }
#endif
}