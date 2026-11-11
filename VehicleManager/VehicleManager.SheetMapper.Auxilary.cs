namespace SKSB_App.VehicleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Storage;

internal static partial class SheetMapper
{
    private static (List<string> Errors, List<string> Warnings) ExtractAndPersistAuxiliaryTables(object wbObj)
    {
        var errors = new List<string>();
        var warnings = new List<string>();
        dynamic wb = wbObj;

        List<Workshop> workshops = new();
        dynamic? wsWorkshop = SheetMapperHelper.ResolveSheetSafely((object)wb, WorkshopAliases, "Workshops");
        if (wsWorkshop != null)
        {
            workshops = TableMapper.ExtractFromTable<Workshop>((object)wsWorkshop, "WorkshopsTable");
        }

        List<InsuranceCompanyRecord> insuranceCompanies = new();
        dynamic? wsIns = SheetMapperHelper.ResolveSheetSafely((object)wb, InsuranceAliases, "Insurances");
        if (wsIns != null)
        {
            insuranceCompanies = TableMapper.ExtractFromTable<InsuranceCompanyRecord>((object)wsIns, "InsuranceCompaniesTable");
        }

        List<FuelCardRecord> fuelCards = new();
        List<FuelCompanyNameRecord> fuelCompanyRecords = new();
        dynamic? wsOdo = SheetMapperHelper.ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            fuelCards = TableMapper.ExtractFromTable<FuelCardRecord>((object)wsOdo, "FuelCardsTable");
            fuelCompanyRecords = TableMapper.ExtractFromTable<FuelCompanyNameRecord>((object)wsOdo, "FuelCompanyNamesTable", "FuelCompaniesTable");
        }

        var fuelCompanies = fuelCompanyRecords
            .Select(c => c.FuelCompanyName)
            .Where(n => !string.IsNullOrWhiteSpace(n))
            .Select(n => n.Trim())
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        // Relational Check: Warn and auto-restore fuel company if assigned to an active card
        var activeFuelCompanySet = new HashSet<string>(
            fuelCompanies.Where(n => !string.IsNullOrWhiteSpace(n)).Select(n => n.Trim()),
            StringComparer.OrdinalIgnoreCase
        );

        var restoredCompanies = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        foreach (var card in fuelCards)
        {
            if (!string.IsNullOrWhiteSpace(card.FuelCompanyName))
            {
                string compName = card.FuelCompanyName.Trim();
                if (!activeFuelCompanySet.Contains(compName))
                {
                    if (restoredCompanies.Add(compName))
                    {
                        warnings.Add($"Fuel company '{compName}' is still assigned to Fuel Card '{card.FuelCardNumber}'. Deletion was reverted and '{compName}' has been automatically restored.");
                        fuelCompanies.Add(compName);
                        activeFuelCompanySet.Add(compName);
                    }
                }
            }
        }

        // Restore onto sheet if changes occurred
        if (restoredCompanies.Count > 0 && wsOdo != null)
        {
            SheetMapperHelper.SyncReferenceTable((object)wb, "FuelCompanyNamesTable", "", fuelCompanies);
        }

        // Persist Auxiliary Tables
        try
        {
            OfflineSnapshot.Save(workshops, writeToDatabase: false);
            OfflineSnapshot.Save(insuranceCompanies, writeToDatabase: false);
            OfflineSnapshot.Save(fuelCards, writeToDatabase: false);
            OfflineSnapshot.Save(fuelCompanyRecords, writeToDatabase: false);

            ServerDispatcher.SaveAuxiliaryTables(
                workshops,
                insuranceCompanies,
                fuelCards,
                fuelCompanies
            );
        }
        catch (Exception ex)
        {
            AppLogger.Error("Failed persisting auxiliary tables to database", ex);
            errors.Add($"Failed saving auxiliary tables to database: {ex.Message}");
        }

        return (errors, warnings);
    }

    private static void PopulateAuxiliaryTables(object wbObj)
    {
        dynamic wb = wbObj;

        // 1. WorkshopsTable
        dynamic? wsWorkshop = SheetMapperHelper.ResolveSheetSafely((object)wb, WorkshopAliases, "Workshops");
        if (wsWorkshop != null)
        {
            try { wsWorkshop.Unprotect("SKSB_Internal_Guard"); } catch { }
            var workshops = ServerSyncCommands.FetchWorkshops();
            TableMapper.PopulateTableSafe<Workshop>((object)wsWorkshop, "WorkshopsTable", workshops);
            SheetProtections.ApplyStyleDrivenProtection(wsWorkshop);
        }

        // 2. InsuranceCompaniesTable
        dynamic? wsIns = SheetMapperHelper.ResolveSheetSafely((object)wb, InsuranceAliases, "Insurances");
        if (wsIns != null)
        {
            try { wsIns.Unprotect("SKSB_Internal_Guard"); } catch { }
            var insuranceCompanies = ServerSyncCommands.FetchInsuranceCompanies();
            TableMapper.PopulateTableSafe<InsuranceCompanyRecord>((object)wsIns, "InsuranceCompaniesTable", insuranceCompanies);
            SheetProtections.ApplyStyleDrivenProtection(wsIns);
        }

        // 3. FuelCardsTable & FuelCompanyNamesTable
        dynamic? wsOdo = SheetMapperHelper.ResolveSheetSafely((object)wb, OdometerAliases, "Odometers");
        if (wsOdo != null)
        {
            try { wsOdo.Unprotect("SKSB_Internal_Guard"); } catch { }
            var fuelCards = ServerSyncCommands.FetchFuelCards();
            TableMapper.PopulateTableSafe<FuelCardRecord>((object)wsOdo, "FuelCardsTable", fuelCards);

            var fuelCompanies = ServerSyncCommands.FetchFuelCompanies();
            SheetMapperHelper.SyncReferenceTable((object)wb, "FuelCompanyNamesTable", "", fuelCompanies);
            SheetProtections.ApplyStyleDrivenProtection(wsOdo);
        }
    }
}