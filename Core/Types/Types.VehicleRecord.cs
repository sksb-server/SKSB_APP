namespace SKSB_App.Core.Types;

using SKSB_App.Core.Types.Attributes;
using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;

// =========================================================================
// 1. VEHICLE DOMAIN ENUMS (Generates Lookups Named Ranges)
// =========================================================================

public enum VehicleType
{
    None, Car, Hilux, Van, Lorry, Tipper_Lorry, Roro_Lorry, Crane_Lorry,
    Skylift, Backhoe, Excavator, Roller, Compressor, Generator, Motorcycle, Machinery
}

public enum VehicleStatus
{
    Active, In_Shop, Expired, Reserved, Decommissioned, Suspended, None
}

public enum ServiceType
{
    Regular_Service, Breakdown_Repair, Tyre_Service, Battery_Replacement,
    Aircond_Service, Bodywork_Paint, Puspakom_Prep, Accident_Repair, Other, None
}

public enum MaintenanceStatus
{
    Scheduled, In_Progress, Completed, Cancelled, None
}

public enum VehiclePermitType
{
    None, APAD_Carrier_A, APAD_Carrier_C, APAD_Carrier_KA, APAD_Decontrolled,
    SPAN_Permit, Bomba_Compliance, JKR_JPJ_Special, Other
}

public enum VehiclePermitStatus
{
    Active, Expired, Pending_Renewal, Cancelled, None
}

public enum RoadTaxStatus
{
    Active, Expired, Pending_Renewal, None
}

public enum InspectionStatus
{
    Scheduled, Completed, Expired, Failed, None
}

public enum InsuranceType
{
    Comprehensive, Third_Party_Fire_Theft, Third_Party, None
}

public enum InsuranceStatus
{
    Active, Expired, Cancelled, None
}

public enum InsuranceClaimStatus
{
    Pending, In_Review, Approved, Settled, Rejected, Cancelled, None
}

// =========================================================================
// 2. ROOT VEHICLE RECORD & MASTER TABLE ("AllVehiclesTable")
// =========================================================================

[ExcelTable("AllVehiclesTable")]
public record VehicleRecord
{
    /// <summary>Heavy machinery tracks engine hours (1 decimal place) instead of kilometres.</summary>
    public static bool IsHourMeterType(VehicleType type) =>
        type is VehicleType.Backhoe or VehicleType.Excavator or VehicleType.Roller
             or VehicleType.Compressor or VehicleType.Generator or VehicleType.Machinery;

    [DataList("VehiclePlateList")]
    [ExcelColumn("Vehicle Plate")]
    [ColumnRule(IsRequired = true, Transform = TransformKind.UpperCase)]
    public string VehiclePlate { get; set; } = string.Empty;

    [DataDetail]
    public VehicleDetail VehicleDetail { get; set; } = new();

    [ExcelColumn("Status")]
    [ColumnRule("VehicleStatusList", IsWarningOnly = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Decommissioned", "Active", "Sold Date", DisallowForcedValueWhenFalse = true)]
    [ColumnStateGuard(2, StateConditionKind.DateExpired, "Expired", "Active", "Road Tax Expiry", "Inspection Expiry", "Insurance Expiry", "Permit Expiry")]
    public VehicleStatus VehicleStatus { get; set; } = VehicleStatus.Active;

    public List<Maintenance> Maintenances { get; set; } = new();
    public List<VehiclePermit> Permits { get; set; } = new();
    public List<RoadTax> RoadTaxes { get; set; } = new();
    public List<Inspection> Inspections { get; set; } = new();
    public List<VehicleInsurance> Insurances { get; set; } = new();
    public List<InsuranceClaim> InsuranceClaims { get; set; } = new();
    public List<OdometerLog> OdometerLogs { get; set; } = new();

}

public record VehicleDetail
{
    // ❌ REMOVED: VehiclePlate (Properly exists only as the Primary Key in VehicleRecord now)

    [DataList("VehicleTypeList")]
    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = true)]
    public VehicleType Type { get; set; } = VehicleType.Lorry;

    [ExcelColumn("Brand")]
    public string Brand { get; set; } = string.Empty;

    [ExcelColumn("Model")]
    public string Model { get; set; } = string.Empty;

    [ExcelColumn("VIN")]
    [ColumnRule(IsRequired = false, Transform = TransformKind.UpperCase)]
    public string VIN { get; set; } = string.Empty;

    [ExcelColumn("Bought Date")]
    public DateOnly? BoughtDate { get; set; }

    [ExcelColumn("Bought Price")]
    public decimal? BoughtPrice { get; set; }

    [ExcelColumn("Sold Date")]
    public DateOnly? SoldDate { get; set; }

    [ExcelColumn("Sold Price")]
    public decimal? SoldPrice { get; set; }

    [ColumnRollup(nameof(VehicleRecord.Insurances), RollupKind.Max, nameof(VehicleInsurance.ExpiryDate))]
    [ExcelColumn("Insurance Expiry")]
    public DateOnly? InsuranceExpiry { get; set; }

    [ColumnRollup(nameof(VehicleRecord.Permits), RollupKind.Min, nameof(VehiclePermit.ExpiryDate))]
    [ExcelColumn("Permit Expiry")]
    public DateOnly? PermitExpiry { get; set; }

    [ColumnRollup(nameof(VehicleRecord.RoadTaxes), RollupKind.Max, nameof(RoadTax.ExpiryDate))]
    [ExcelColumn("Road Tax Expiry")]
    public DateOnly? RoadTaxExpiry { get; set; }

    [ColumnRollup(nameof(VehicleRecord.Inspections), RollupKind.Max, nameof(Inspection.ExpiryDate))]
    [ExcelColumn("Inspection Expiry")]
    public DateOnly? InspectionExpiry { get; set; }

    [ColumnRollup(nameof(VehicleRecord.OdometerLogs), RollupKind.Max, nameof(OdometerLog.CurrentOdometer))]
    [ExcelColumn("Current Odometer")]
    public decimal? CurrentOdometer { get; set; }

    [ExcelColumn("Driven By")]
    public string DrivenBy { get; set; } = string.Empty;

    [ExcelColumn("Used At")]
    public string UsedAt { get; set; } = string.Empty;

    // ---- UI-only calculated / rollup columns of 'AllVehiclesTable' (never persisted) ----
    // Cost rollups are All-Time totals; the per-date breakdown lives on the 'Vehicle Summary' sheet.

    [JsonIgnore]
    [ExcelColumn("Years Old")]
    [ColumnCalculate("IF([Bought Date] > 0, YEARS_BETWEEN([Bought Date], IF([Sold Date] > 0, [Sold Date], TODAY())), null)", DecimalPlaces = 0)]
    public int? YearsOld { get; set; }

    [JsonIgnore]
    [ExcelColumn("Return")]
    [ColumnCalculate("IF([Sold Price] > 0, [Sold Price] - [Bought Price], null)")]
    public decimal? ReturnAmount { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.OdometerLogs), RollupKind.Sum, nameof(OdometerLog.Amount), DateProperty = nameof(OdometerLog.Date))]
    [ExcelColumn("Fuel Cost")]
    public decimal? FuelCost { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.Maintenances), RollupKind.Sum, nameof(Maintenance.RepairCost), DateProperty = nameof(Maintenance.DateIn))]
    [ExcelColumn("Maintenance Cost")]
    public decimal? MaintenanceCost { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.Insurances), RollupKind.Sum, nameof(VehicleInsurance.Amount), DateProperty = nameof(VehicleInsurance.RenewDate))]
    [ExcelColumn("Insurance Cost")]
    public decimal? InsuranceCost { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.Permits), RollupKind.Sum, nameof(VehiclePermit.Amount), DateProperty = nameof(VehiclePermit.RenewDate))]
    [ExcelColumn("Permit Cost")]
    public decimal? PermitCost { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.RoadTaxes), RollupKind.Sum, nameof(RoadTax.Amount), DateProperty = nameof(RoadTax.RenewDate))]
    [ColumnRollup(nameof(VehicleRecord.Inspections), RollupKind.Sum, nameof(Inspection.Amount), DateProperty = nameof(Inspection.ScheduleDateIn))]
    [ExcelColumn("Road Tax and Inspection Cost")]
    public decimal? RoadTaxAndInspectionCost { get; set; }

    [JsonIgnore]
    [ColumnRollup(nameof(VehicleRecord.InsuranceClaims), RollupKind.Sum, nameof(InsuranceClaim.ClaimAmount), DateProperty = nameof(InsuranceClaim.ClaimDate))]
    [ExcelColumn("Insurance Claim")]
    public decimal? InsuranceClaimAmount { get; set; }

    [JsonIgnore]
    [ColumnCalculate("[Fuel Cost] + [Maintenance Cost] + [Insurance Cost] + [Permit Cost] + [Road Tax and Inspection Cost] - IF([Insurance Claim] > 0, [Insurance Claim], 0)")]
    [ExcelColumn("Total Cost")]
    public decimal? TotalCost { get; set; }
}

// =========================================================================
// SHARED BASE: every record owned by a vehicle (Vehicle Plate + two-way Vehicle Type lookup)
// =========================================================================

public abstract record VehicleChildRecord
{
    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [JsonIgnore] // UI-only helper derived from Vehicle Plate: never duplicated into DB / JSON snapshots
    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;
}

/// <summary>Shared renewal-cycle columns (insurance, permit, road tax): renewal / expiry dates, cost, remarks.</summary>
public abstract record VehicleRenewalRecord : VehicleChildRecord
{
    [ExcelColumn("Renewal Date", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Renewal Date", offset: 365)]
    [ColumnDateGreaterThan("Renewal Date")]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Cost")]
    public decimal? Amount { get; set; }

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 3. MAINTENANCE TABLE ("MaintenancesTable")
// =========================================================================

[ExcelTable("MaintenancesTable")]
public record Maintenance : VehicleChildRecord
{
    [ExcelColumn("Maintenance ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long MaintenanceID { get; set; }

    [ExcelColumn("Service Type", isRequired: true)]
    [ColumnRule("ServiceTypeList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Regular_Service")]
    public ServiceType ServiceType { get; set; } = ServiceType.Regular_Service;

    [ExcelColumn("Date In", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly DateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Estimated Return")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Date In", offset: 3)]
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(3));

    [ExcelColumn("Date Completed")]
    public DateOnly? DateCompleted { get; set; }

    [ExcelColumn("Odometer In")]
    public decimal OdometerIn { get; set; } = 0m;

    [ExcelColumn("Odometer Out")]
    public decimal OdometerOut { get; set; } = 0m;

    [ExcelColumn("Workshop Name")]
    [ColumnRule("WorkshopsTable[Workshop Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Workshop Address", "WorkshopsTable[Workshop Name]", "WorkshopsTable[Workshop Address]")]
    public string WorkshopName { get; set; } = string.Empty;

    // Persisted to vehicle_maintenance.workshop_address (two-way synced from Workshops).
    [ExcelColumn("Workshop Address")]
    [ColumnRule("WorkshopsTable[Workshop Address]", IsWarningOnly = true, IsRequired = false)]
    [ColumnSync("Workshop Name", "WorkshopsTable[Workshop Address]", "WorkshopsTable[Workshop Name]")]
    public string WorkshopAddress { get; set; } = string.Empty;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("MaintenanceStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "In_Progress")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Completed", "In_Progress", "Date Completed", SkipIfCurrentEquals = "Cancelled", DisallowForcedValueWhenFalse = true)]
    [ColumnStateGuard(2, StateConditionKind.DateExpired, "In_Progress", "Scheduled", "Date in", SkipIfCurrentEquals = "Cancelled")]
    public MaintenanceStatus MaintenanceStatus { get; set; } = MaintenanceStatus.In_Progress;

    [ExcelColumn("Repair Total Cost")]
    [ColumnCalculate("[Tyre] + [Brake] + [Clutch] + [Service] + [Fuel Pump] + [Aircond] + [Wiring] + [Battery] + [Towing] + [Cushion] + [Parts]")]
    public decimal? RepairCost { get; set; }

    [ExcelColumn("Tyre")]
    public decimal? TyreCost { get; set; }

    [ExcelColumn("Brake")]
    public decimal? BrakeCost { get; set; }

    [ExcelColumn("Clutch")]
    public decimal? ClutchCost { get; set; }

    [ExcelColumn("Service")]
    public decimal? ServiceCost { get; set; }

    [ExcelColumn("Fuel Pump")]
    public decimal? FuelPumpCost { get; set; }

    [ExcelColumn("Aircond")]
    public decimal? AircondCost { get; set; }

    [ExcelColumn("Wiring")]
    public decimal? WiringCost { get; set; }

    [ExcelColumn("Battery")]
    public decimal? BatteryCost { get; set; }

    [ExcelColumn("Towing")]
    public decimal? TowingCost { get; set; }

    [ExcelColumn("Cushion")]
    public decimal? CushionCost { get; set; }

    [ExcelColumn("Parts")]
    public decimal? PartsCost { get; set; }

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 4. WORKSHOPS DIRECTORY TABLE ("WorkshopsTable")
// =========================================================================

[ExcelTable("WorkshopsTable")]
public record Workshop
{
    [ExcelColumn("Workshop ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long WorkshopID { get; set; }

    [DataList("WorkshopNameList")]
    [ExcelColumn("Workshop Name", isRequired: true)]
    public string WorkshopName { get; set; } = string.Empty;

    [DataList("WorkshopAddressList")]
    [ExcelColumn("Workshop Address")]
    public string WorkshopAddress { get; set; } = string.Empty;

    [ExcelColumn("Workshop Owner")]
    public string WorkshopOwner { get; set; } = string.Empty;

    [ExcelColumn("Workshop Primary Contact")]
    public string PrimaryContact { get; set; } = string.Empty;

    [ExcelColumn("Workshop Secondary Contact")]
    public string SecondaryContact { get; set; } = string.Empty;

    // ---- UI-only cross-table rollups from 'MaintenancesTable' (never persisted) ----
    [JsonIgnore]
    [ColumnRollup("MaintenancesTable", RollupKind.Count, "Maintenance ID", ForeignKeyHeader = "Workshop Name", LocalKeyHeader = "Workshop Name", DateProperty = "Date In")]
    [ExcelColumn("Number of Vehicle Fixed")]
    public int? NumberOfVehiclesFixed { get; set; }

    [JsonIgnore]
    [ColumnRollup("MaintenancesTable", RollupKind.DistinctJoin, "Vehicle Plate", ForeignKeyHeader = "Workshop Name", LocalKeyHeader = "Workshop Name", DateProperty = "Date In")]
    [ExcelColumn("Fixed Vehicle Plates")]
    public string FixedVehiclePlates { get; set; } = string.Empty;

    [JsonIgnore]
    [ColumnRollup("MaintenancesTable", RollupKind.Sum, "Repair Total Cost", ForeignKeyHeader = "Workshop Name", LocalKeyHeader = "Workshop Name", DateProperty = "Date In")]
    [ExcelColumn("Total Cost")]
    public decimal? TotalCost { get; set; }
}

// =========================================================================
// 5. VEHICLE INSURANCE & INSURANCE COMPANIES ("InsurancesTable", "InsuranceCompaniesTable")
// =========================================================================

[ExcelTable("InsurancesTable")]
public record VehicleInsurance : VehicleRenewalRecord
{
    [ExcelColumn("Insurance Renewal ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long InsuranceID { get; set; }

    // 1. Primary Mapped Column (Stored in Database)
    [ExcelColumn("Insurance Number", isRequired: true)]
    public string InsuranceNumber { get; set; } = string.Empty;

    [ExcelColumn("Insurance Company", isRequired: true)]
    [ColumnRule("InsuranceCompaniesTable[Insurance Company]", IsRequired = true)]
    [ColumnSync("Company Contact Number", "InsuranceCompaniesTable[Insurance Company]", "InsuranceCompaniesTable[Company Contact Number]")]
    [ColumnSync("Insurance Agent", "InsuranceCompaniesTable[Insurance Company]", "InsuranceCompaniesTable[Insurance Agent]")]
    public string InsuranceCompany { get; set; } = string.Empty;

    // 2. Unmapped Helper Column (UI / Display only)
    [JsonIgnore]
    [ExcelColumn("Company Contact Number")]
    [ColumnRule("InsuranceCompaniesTable[Company Contact Number]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Insurance Company", "InsuranceCompaniesTable[Company Contact Number]", "InsuranceCompaniesTable[Insurance Company]")]
    [ColumnSync("Insurance Agent", "InsuranceCompaniesTable[Company Contact Number]", "InsuranceCompaniesTable[Insurance Agent]")]
    public string CompanyContactNumber { get; set; } = string.Empty;

    // 3. Unmapped Helper Column (UI / Display only)
    [JsonIgnore]
    [ExcelColumn("Insurance Agent")]
    [ColumnRule("InsuranceCompaniesTable[Insurance Agent]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Insurance Company", "InsuranceCompaniesTable[Insurance Agent]", "InsuranceCompaniesTable[Insurance Company]")]
    [ColumnSync("Company Contact Number", "InsuranceCompaniesTable[Insurance Agent]", "InsuranceCompaniesTable[Company Contact Number]")]
    public string InsuranceAgent { get; set; } = string.Empty;

    [ExcelColumn("Insurance Type")]
    [ColumnRule("InsuranceTypeList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Comprehensive")]
    public InsuranceType InsuranceType { get; set; } = InsuranceType.Comprehensive;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("InsuranceStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public InsuranceStatus Status { get; set; } = InsuranceStatus.Active;
}

[ExcelTable("InsuranceCompaniesTable")]
public record InsuranceCompanyRecord
{
    [DataList("InsuranceCompanyList")]
    [ExcelColumn("Insurance Company", isRequired: true)]
    public string InsuranceCompany { get; set; } = string.Empty;

    [ExcelColumn("Company Contact Number")]
    public string CompanyContactNumber { get; set; } = string.Empty;

    [ExcelColumn("Company Address")]
    public string CompanyAddress { get; set; } = string.Empty;

    [ExcelColumn("Insurance Agent")]
    public string InsuranceAgent { get; set; } = string.Empty;

    [ExcelColumn("Agent Contact Number")]
    public string AgentContactNumber { get; set; } = string.Empty;

    // ---- UI-only cross-table rollups from 'InsurancesTable' (never persisted) ----
    [JsonIgnore]
    [ColumnRollup("InsurancesTable", RollupKind.Count, "Vehicle Plate", ForeignKeyHeader = "Insurance Agent", LocalKeyHeader = "Insurance Agent")]
    [ExcelColumn("Number of Vehicles Under Agent")]
    public int? NumberOfVehiclesUnderAgent { get; set; }

    [JsonIgnore]
    [ColumnRollup("InsurancesTable", RollupKind.DistinctJoin, "Vehicle Plate", ForeignKeyHeader = "Insurance Agent", LocalKeyHeader = "Insurance Agent")]
    [ExcelColumn("Vehicle Plates Under Agent")]
    public string VehiclePlatesUnderAgent { get; set; } = string.Empty;
}

[ExcelTable("InsuranceClaimsTable")]
public record InsuranceClaim : VehicleChildRecord
{
    [ExcelColumn("Insurance Claim ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long ClaimID { get; set; }

    [ExcelColumn("Claim Date", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly ClaimDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Settlement Date")]
    [ColumnDateGreaterThan("Claim Date")]
    public DateOnly? SettlementDate { get; set; }

    [ExcelColumn("Staff ID")]
    [ColumnRule("StaffIDList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffShortNameList")]
    public long StaffID { get; set; }

    [JsonIgnore] // UI-only helper: derived from Staff ID
    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffShortNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffShortNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Maintenance ID")]
    [ColumnRule("MaintenancesTable[Maintenance ID]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Workshop Name", "MaintenancesTable[Maintenance ID]", "MaintenancesTable[Workshop Name]")]
    [ColumnSync("Workshop Address", "MaintenancesTable[Maintenance ID]", "MaintenancesTable[Workshop Address]")]
    public long MaintenanceID { get; set; } = 0;

    [ExcelColumn("Workshop Name")]
    [ColumnRule("WorkshopsTable[Workshop Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Workshop Address", "WorkshopsTable[Workshop Name]", "WorkshopsTable[Workshop Address]")]
    public string WorkshopName { get; set; } = string.Empty;

    [JsonIgnore] // UI-only helper: derived from Maintenance ID / Workshop Name
    [ExcelColumn("Workshop Address")]
    [ColumnRule("WorkshopsTable[Workshop Address]", IsWarningOnly = true, IsRequired = false)]
    public string WorkshopAddress { get; set; } = string.Empty;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("InsuranceClaimStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Pending")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Settled", "Pending", "Settlement Date", SkipIfCurrentEquals = "Rejected", DisallowForcedValueWhenFalse = true)]
    public InsuranceClaimStatus Status { get; set; } = InsuranceClaimStatus.Pending;

    [ExcelColumn("Claim Amount")]
    public decimal? ClaimAmount { get; set; }

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 6. PERMITS, ROAD TAXES, AND INSPECTIONS TABLES
// =========================================================================

[ExcelTable("PermitsTable")]
public record VehiclePermit : VehicleRenewalRecord
{
    [ExcelColumn("Permit ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long PermitID { get; set; }

    [ExcelColumn("Permit Type")]
    [ColumnRule("VehiclePermitTypeList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "APAD_Carrier_A")]
    public VehiclePermitType PermitType { get; set; } = VehiclePermitType.APAD_Carrier_A;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("VehiclePermitStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public VehiclePermitStatus Status { get; set; } = VehiclePermitStatus.Active;
}

[ExcelTable("RoadTaxesTable")]
public record RoadTax : VehicleRenewalRecord
{
    [ExcelColumn("Road Tax ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long RoadTaxID { get; set; }

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("RoadTaxStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public RoadTaxStatus Status { get; set; } = RoadTaxStatus.Active;
}

[ExcelTable("InspectionsTable")]
public record Inspection : VehicleChildRecord
{
    [ExcelColumn("Inspection ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long InspectionID { get; set; }

    [ExcelColumn("Schedule Date In")]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly ScheduleDateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Date Out", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly DateOut { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnDefault(DefaultKind.DateOffsetMonths, referenceHeader: "Date Out", offset: 6)]
    [ColumnDateGreaterThan("Date Out")]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddMonths(6));

    [ExcelColumn("Cost")]
    public decimal? Amount { get; set; }

    [ExcelColumn("Status")]
    [ColumnRule("InspectionStatusList")]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Scheduled")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Completed", "Expiry Date", SkipIfCurrentEquals = "Failed", DisallowForcedValueWhenFalse = true)]
    [ColumnStateGuard(2, StateConditionKind.DateReached, "Completed", "Scheduled", "Date Out", SkipIfCurrentEquals = "Failed", DisallowForcedValueWhenFalse = true)]
    public InspectionStatus Status { get; set; } = InspectionStatus.Scheduled;

    [ExcelColumn("Address")]
    public string Address { get; set; } = string.Empty;

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 7. ODOMETER LOGS & FUEL CARDS ("OdometersTable", "FuelCardsTable")
// =========================================================================

[ExcelTable("OdometersTable")]
public record OdometerLog : VehicleChildRecord
{
    [ExcelColumn("Fuel Transaction ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long OdometerID { get; set; }

    [ExcelColumn("Date", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    // 1. Stored Primary Foreign Key
    [ExcelColumn("Staff ID")]
    [ColumnRule("StaffIDList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff Name", "StaffIDList", "StaffShortNameList")]
    public long StaffID { get; set; } = 0;

    // 2. Unmapped Helper Column (Two-Way Auto-Evaluated in Excel)
    [JsonIgnore]
    [ExcelColumn("Staff Name")]
    [ColumnRule("StaffShortNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffShortNameList", "StaffIDList")]
    public string StaffName { get; set; } = string.Empty;

    [ExcelColumn("Project")]
    [ColumnRule("ProjectIDList", IsRequired = false, IsWarningOnly = true)]
    public string Project { get; set; } = string.Empty;

    [ColumnRollup("OdometersTable", RollupKind.PriorMax, "Current Odometer", ForeignKeyHeader = "Vehicle Plate", LocalKeyHeader = "Vehicle Plate", DateProperty = "Date", Mode = CalculationMode.DefaultOnly)]
    [ExcelColumn("Odometer Before")]
    public decimal? OdometerBefore { get; set; }

    [ExcelColumn("Current Odometer", isRequired: true)]
    public decimal CurrentOdometer { get; set; } = 0m;

    [JsonIgnore]
    [ExcelColumn("Odometer Difference")]
    [ColumnCalculate("IF([Current Odometer] > [Odometer Before], [Current Odometer] - [Odometer Before], 0)", DecimalPlaces = 1)]
    public decimal? OdometerDifference { get; set; }

    [ExcelColumn("Fuel Card Number")]
    [ColumnRule("FuelCardsTable[Fuel Card Number]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Fuel Company", "FuelCardsTable[Fuel Card Number]", "FuelCardsTable[Fuel Company Name]")]
    [ColumnSync("Card Owner", "FuelCardsTable[Fuel Card Number]", "FuelCardsTable[Card Owner Name]")]
    public string FuelCardNumber { get; set; } = string.Empty;

    [JsonIgnore]
    [ExcelColumn("Card Owner")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true)]
    public string CardOwner { get; set; } = string.Empty;

    [ExcelColumn("Fuel Company")]
    [ColumnRule("FuelCompanyNamesTable[Fuel Company Name]", IsRequired = false, IsWarningOnly = true)]
    public string FuelCompany { get; set; } = string.Empty;

    [ExcelColumn("Station Name")]
    public string StationName { get; set; } = string.Empty;

    [ExcelColumn("Receipt Number")]
    public string ReceiptNumber { get; set; } = string.Empty;

    [ExcelColumn("Fuel Type")]
    public string FuelType { get; set; } = string.Empty;

    [ExcelColumn("Price at Fueling")]
    [ColumnFormat(ColumnFormatKind.PricePerLiter)]
    public decimal PriceAtFueling { get; set; } = 0.00m;

    [ExcelColumn("Fuel Amount")]
    [ColumnFormat(ColumnFormatKind.Liters)]
    public decimal FuelAmount { get; set; } = 0.00m;

    [ExcelColumn("Transaction Amount")]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Time")]
    public string Time { get; set; } = string.Empty;

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;
}

[ExcelTable("FuelCardsTable")]
public record FuelCardRecord
{
    [DataList("FuelCardNumberList")]
    [ExcelColumn("Fuel Card Number", isRequired: true)]
    public string FuelCardNumber { get; set; } = string.Empty;

    [ExcelColumn("Fuel Company Names")]
    [ColumnRule("FuelCompanyNamesTable[Fuel Company Name]", IsRequired = true)]
    public string FuelCompanyName { get; set; } = string.Empty;

    [ExcelColumn("Card Owner ID", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", "StaffIDList", IsRequired = true)]
    [ColumnSync("Card Owner Name", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Brand] + ' ' + AllVehiclesTable[Model]")]
    [ColumnSync("Card Owner Name", "StaffIDList", "StaffShortNameList")]
    public string CardOwnerID { get; set; } = string.Empty;

    [JsonIgnore]
    [ExcelColumn("Card Owner Name")]
    [ColumnRule("AllVehiclesTable[Brand] + ' ' + AllVehiclesTable[Model]", "StaffShortNameList", IsRequired = false, IsWarningOnly = true)]
    // Two-way sync: Picking the name looks up the ID/Plate and fills Card Owner ID
    [ColumnSync("Card Owner ID", "AllVehiclesTable[Brand] + ' ' + AllVehiclesTable[Model]", "AllVehiclesTable[Vehicle Plate]")]
    [ColumnSync("Card Owner ID", "StaffShortNameList", "StaffIDList")]
    public string CardOwnerName { get; set; } = string.Empty;
}

[ExcelTable("FuelCompanyNamesTable")]
public record FuelCompanyNameRecord
{
    [DataList("FuelCompanyNameList")]
    [ExcelColumn("Fuel Company Names", isRequired: true)]
    public string FuelCompanyName { get; set; } = string.Empty;
}

// =========================================================================
// 8. VEHICLE SUMMARY (read-only "VehicleSummaryTable"; never saved)
// =========================================================================

[ExcelTable("VehicleSummaryTable")]
public record VehicleSummaryRow
{
    [ExcelColumn("Date")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public DateOnly Date { get; set; }

    [ExcelColumn("Status")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public VehicleStatus Status { get; set; } = VehicleStatus.Active;

    [ExcelColumn("Driven By")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string DrivenBy { get; set; } = string.Empty;

    [ExcelColumn("Used At")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public string UsedAt { get; set; } = string.Empty;

    [ExcelColumn("Odometer Used")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.OdometerDynamic)]
    public decimal? OdometerUsed { get; set; }

    [ExcelColumn("Fuel Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? FuelCost { get; set; }

    [ExcelColumn("Maintenance Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? MaintenanceCost { get; set; }

    [ExcelColumn("Insurance Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? InsuranceCost { get; set; }

    [ExcelColumn("Permit Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? PermitCost { get; set; }

    [ExcelColumn("Road Tax and Inspection Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? RoadTaxAndInspectionCost { get; set; }

    [ExcelColumn("Insurance Claim")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? InsuranceClaimAmount { get; set; }

    [ExcelColumn("Total Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? TotalCost { get; set; }
}