namespace SKSB_App.VehicleManager;

using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using SKSB_App.Core.WorkbookSave.Validation;
using System;
using System.Collections.Generic;
using System.Linq;

internal static partial class SheetMapper
{
    public const string MasterSheetName = "All Vehicles";
    public const string MasterTableName = "AllVehiclesTable";

    internal static readonly string[] AllVehicleAliases = { "All Vehicles", "AllVehicles", "Vehicle_Master", "Vehicles" };
    internal static readonly string[] MaintenanceAliases = { "Maintenances", "Maintenance", "Vehicle_Maintenance" };
    internal static readonly string[] ServiceAliases = { "Services", "Service", "Vehicle_Services" };
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
            // Date-based child tables are populated newest-date-on-top (not grouped by master plate order).
            Child<Maintenance>(MaintenanceAliases, "Maintenances", "MaintenancesTable", v => v.Maintenances, m => m.VehiclePlate, (m, plate) => m.VehiclePlate = plate,
                list => list.OrderByDescending(m => m.DateIn).ThenByDescending(m => m.MaintenanceID).ToList()),
            Child<VehicleService>(ServiceAliases, "Services", "ServicesTable", v => v.Services, s => s.VehiclePlate, (s, plate) => s.VehiclePlate = plate,
                list => list.OrderByDescending(s => s.ServiceDate).ThenByDescending(s => s.ServiceID).ToList()),
            Child<RoadTax>(RoadTaxAliases, "Road Taxes", "RoadTaxesTable", v => v.RoadTaxes, r => r.VehiclePlate, (r, plate) => r.VehiclePlate = plate,
                list => list.OrderByDescending(r => r.RenewDate).ThenByDescending(r => r.RoadTaxID).ToList()),
            Child<Inspection>(InspectionAliases, "Inspections", "InspectionsTable", v => v.Inspections, i => i.VehiclePlate, (i, plate) => i.VehiclePlate = plate,
                list => list.OrderByDescending(i => i.ScheduleDateIn).ThenByDescending(i => i.InspectionID).ToList()),
            Child<VehiclePermit>(PermitAliases, "Permits", "PermitsTable", v => v.Permits, p => p.VehiclePlate, (p, plate) => p.VehiclePlate = plate,
                list => list.OrderByDescending(p => p.RenewDate).ThenByDescending(p => p.PermitID).ToList()),
            Child<VehicleInsurance>(InsuranceAliases, "Insurances", "InsurancesTable", v => v.Insurances, i => i.VehiclePlate, (i, plate) => i.VehiclePlate = plate,
                list => list.OrderByDescending(i => i.RenewDate).ThenByDescending(i => i.InsuranceID).ToList()),
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
                ValidateForeignKeyExists = true,
                OrderChildRecords = list => list.OrderByDescending(c => c.ClaimDate).ThenByDescending(c => c.ClaimID).ToList()
            },
            Child<OdometerLog>(OdometerAliases, "Odometers", "OdometersTable", v => v.OdometerLogs, o => o.VehiclePlate, (o, plate) => o.VehiclePlate = plate,
                list => list.OrderByDescending(o => o.Date).ThenByDescending(o => o.Time ?? string.Empty).ThenByDescending(o => o.OdometerID).ToList())
        },

        // Fuel companies still referenced by a fuel card are restored onto the sheet before rule validation runs.
        PreValidationHook = (wb, masters, errors, warnings) =>
        {
            var (auxErrors, auxWarnings) = RestoreReferencedAuxiliaryRows((object)wb);
            errors.AddRange(auxErrors);
            warnings.AddRange(auxWarnings);
        },

        OnPopulateAuxiliaryTables = (wb, masters) =>
        {
            PopulateAuxiliaryTables((object)wb);
            InitializeVehicleSummaryContext((object)wb, masters);
        },

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
            "ServicesTable",
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
            "AllVehiclesTable", "MaintenancesTable", "ServicesTable", "WorkshopsTable", "RoadTaxesTable",
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
        Action<TChild, string> plateSetter,
        Func<List<TChild>, List<TChild>>? orderRecords = null) where TChild : new()
    {
        return new ChildTableBinding<VehicleRecord, TChild, string>
        {
            OrderChildRecords = orderRecords,
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
    internal static (List<VehicleRecord>? Data, List<ValidationIssue> Errors, List<ValidationIssue> Warnings) ExtractVehicleRecordsWithDiagnostics()
        => WorkbookMapper.ExtractMasterAndChildrenWithDiagnostics(Mapping);

    internal static void PopulateWorksheet(List<VehicleRecord> records)
        => WorkbookMapper.PopulateMasterAndChildren(Mapping, records);

    // =========================================================================
    // 3. SHARED UTILITIES
    // =========================================================================
}
