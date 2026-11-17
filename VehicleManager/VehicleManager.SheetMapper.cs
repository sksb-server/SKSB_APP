namespace SKSB_App.VehicleManager;

using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using System;
using System.Collections.Generic;

internal static partial class SheetMapper
{
    public const string MasterSheetName = "All Vehicles";
    public const string MasterTableName = "AllVehiclesTable";

    internal static readonly string[] AllVehicleAliases = { "All Vehicles", "AllVehicles", "Vehicle_Master", "Vehicles" };
    internal static readonly string[] MaintenanceAliases = { "Maintenances", "Maintenance", "Vehicle_Maintenance" };
    internal static readonly string[] WorkshopAliases = { "Workshops", "Workshop" };
    internal static readonly string[] RoadTaxAliases = { "Road Taxes", "RoadTaxes", "Road Tax" };
    internal static readonly string[] InspectionAliases = { "Inspections", "Inspection" };
    internal static readonly string[] PermitAliases = { "Permits", "Permit" };
    internal static readonly string[] InsuranceAliases = { "Insurances", "Insurance" };
    internal static readonly string[] InsuranceClaimAliases = { "Insurance Claims", "InsuranceClaims", "Insurance Claim", "Claims" };
    internal static readonly string[] OdometerAliases = { "Odometers", "Odometer", "Fuel Logs", "Fuel" };

    // =========================================================================
    // 1. DECLARATIVE WORKBOOK MAPPING (MASTER + CHILD TABLES)
    // =========================================================================
    private static readonly WorkbookMapping<VehicleRecord, string> Mapping = new()
    {
        MasterSheetAliases = AllVehicleAliases,
        MasterSheetName = MasterSheetName,
        MasterTableName = MasterTableName,
        EntityLabel = "vehicle",
        MasterKeyLabel = "Vehicle Plate",
        MasterKeySelector = v => v.VehiclePlate,
        KeyComparer = StringComparer.OrdinalIgnoreCase,
        IsValidKey = plate => !string.IsNullOrWhiteSpace(plate),

        Children = new List<IChildTableBinding<VehicleRecord, string>>
        {
            Child<Maintenance>(MaintenanceAliases, "Maintenances", "MaintenancesTable", v => v.Maintenances, m => m.VehiclePlate, (m, plate) => m.VehiclePlate = plate),
            Child<RoadTax>(RoadTaxAliases, "Road Taxes", "RoadTaxesTable", v => v.RoadTaxes, r => r.VehiclePlate, (r, plate) => r.VehiclePlate = plate),
            Child<Inspection>(InspectionAliases, "Inspections", "InspectionsTable", v => v.Inspections, i => i.VehiclePlate, (i, plate) => i.VehiclePlate = plate),
            Child<VehiclePermit>(PermitAliases, "Permits", "PermitsTable", v => v.Permits, p => p.VehiclePlate, (p, plate) => p.VehiclePlate = plate),
            Child<VehicleInsurance>(InsuranceAliases, "Insurances", "InsurancesTable", v => v.Insurances, i => i.VehiclePlate, (i, plate) => i.VehiclePlate = plate),
            new ChildTableBinding<VehicleRecord, InsuranceClaim, string>
            {
                SheetAliases = InsuranceClaimAliases,
                DefaultSheetName = "Insurance Claims",
                TableName = "InsuranceClaimsTable",
                ForeignKeySelector = c => c.VehiclePlate,
                ParentCollectionSelector = v => v.InsuranceClaims,
                BackfillFromParent = (v, c) =>
                {
                    if (string.IsNullOrWhiteSpace(c.VehiclePlate)) c.VehiclePlate = v.VehiclePlate;
                },
                ChildPrimaryKeySelector = c => c.ClaimID,
                ChildPrimaryKeyLabel = "Claim ID",
                ValidateForeignKeyExists = true
            },
            Child<OdometerLog>(OdometerAliases, "Odometers", "OdometersTable", v => v.OdometerLogs, o => o.VehiclePlate, (o, plate) => o.VehiclePlate = plate)
        },

        // Fuel companies still referenced by a fuel card are restored onto the sheet before rule validation runs.
        PreValidationHook = (wb, masters, errors, warnings) =>
        {
            var (auxErrors, auxWarnings) = RestoreReferencedAuxiliaryRows((object)wb);
            errors.AddRange(auxErrors);
            warnings.AddRange(auxWarnings);
        },

        OnBeforeMasterPopulate = masterWs =>
        {
            if (masterWs.Cells[2, 2].Value2 == null)
                masterWs.Cells[2, 2].Value2 = DateTime.Today.ToOADate();

            if (string.IsNullOrWhiteSpace(masterWs.Cells[1, 3].Value2?.ToString()))
                masterWs.Cells[2, 3].Value2 = "All_Time";
        },

        OnPopulateAuxiliaryTables = (wb, masters) => PopulateAuxiliaryTables((object)wb),

        ApplyMasterCalculations = ApplyVehicleCalculations,

        RecalculateAfterPopulate = true,

        // Claims heal after Maintenances / Workshops (Maintenance ID -> Workshop sync) and Insurances.
        HealTableOrder = new[]
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
            "InsuranceClaimsTable",
            "OdometersTable",
            // Re-heal parents so cross-table rollups see fully synced child tables.
            "WorkshopsTable",
            "InsuranceCompaniesTable",
            "AllVehiclesTable"
        },

        AuditTitle = "Vehicle Manager Populate Audit",
        MasterLabel = "Vehicle Profiles",
        InspectTables = new[]
        {
            "AllVehiclesTable", "MaintenancesTable", "WorkshopsTable", "RoadTaxesTable",
            "InspectionsTable", "PermitsTable", "InsurancesTable", "InsuranceCompaniesTable",
            "InsuranceClaimsTable", "OdometersTable", "FuelCardsTable", "FuelCompanyNamesTable"
        }
    };

    /// <summary>Builds a plate-keyed child binding that backfills a blank Vehicle Plate from its parent.</summary>
    private static IChildTableBinding<VehicleRecord, string> Child<TChild>(
        string[] aliases,
        string defaultSheetName,
        string tableName,
        Func<VehicleRecord, IList<TChild>> collection,
        Func<TChild, string> plateSelector,
        Action<TChild, string> plateSetter) where TChild : new()
    {
        return new ChildTableBinding<VehicleRecord, TChild, string>
        {
            SheetAliases = aliases,
            DefaultSheetName = defaultSheetName,
            TableName = tableName,
            ForeignKeySelector = plateSelector,
            ParentCollectionSelector = collection,
            BackfillFromParent = (v, c) =>
            {
                if (string.IsNullOrWhiteSpace(plateSelector(c))) plateSetter(c, v.VehiclePlate);
            },
            ValidateForeignKeyExists = true
        };
    }

    // =========================================================================
    // 2. MASTER EXTRACTION & POPULATE PIPELINES
    // =========================================================================
    internal static (List<VehicleRecord>? Data, List<string> Errors, List<string> Warnings) ExtractVehicleRecordsWithDiagnostics()
        => WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);

    internal static void PopulateWorksheet(List<VehicleRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 3. SHARED UTILITIES
    // =========================================================================
}
