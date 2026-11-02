namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.Types.Attributes;

// Backward-compatibility aliases
using RoadTaxRecord = RoadTax;
using InspectionRecord = Inspection;
using InsuranceRecord = VehicleInsurance;
using OdometerRecord = OdometerLog;
using FuelCompanyLookup = FuelCompanyNameRecord;

// =========================================================================
// 1. VEHICLE DOMAIN ENUMS (Generates Lookups Named Ranges)
// =========================================================================

public enum VehicleType
{
    None,
    Car,
    Hilux,
    Van,
    Lorry,
    Tipper_Lorry,
    Roro_Lorry,
    Crane_Lorry,
    Skylift,
    Backhoe,
    Excavator,
    Roller,
    Compressor,
    Generator,
    Motorcycle,
    Machinery
}

public enum VehicleStatus
{
    Active,
    In_Shop,
    Expired,
    Reserved,
    Decommissioned,
    None
}

public enum ServiceType
{
    Regular_Service,
    Breakdown_Repair,
    Tyre_Service,
    Battery_Replacement,
    Aircond_Service,
    Bodywork_Paint,
    Puspakom_Prep,
    Accident_Repair,
    Other,
    None
}

public enum MaintenanceStatus
{
    Scheduled,
    In_Progress,
    Done,
    Completed,
    Cancelled,
    None
}

public enum VehiclePermitType
{
    None,
    APAD_Carrier_A,
    APAD_Carrier_C,
    APAD_Carrier_KA,
    APAD_Decontrolled,
    SPAN_Permit,
    Bomba_Compliance,
    JKR_JPJ_Special,
    Other
}

public enum VehiclePermitStatus
{
    Active,
    Expired,
    Pending_Renewal,
    Cancelled,
    None
}

public enum RoadTaxStatus
{
    Active,
    Expired,
    Pending_Renewal,
    None
}

public enum InspectionStatus
{
    Scheduled,
    Completed,
    Expired,
    Failed,
    None
}

public enum InsuranceType
{
    Comprehensive,
    Third_Party_Fire_Theft,
    Third_Party,
    None
}

public enum InsuranceStatus
{
    Active,
    Expired,
    Cancelled,
    None
}

public enum FuelCompany
{
    Petronas,
    Shell,
    Caltex,
    Petron,
    BHPetrol,
    Other,
    None
}

// =========================================================================
// 2. ROOT VEHICLE RECORD & MASTER TABLE ("AllVehiclesTable")
// =========================================================================

[ExcelTable("AllVehiclesTable")]
public record VehicleRecord
{
    [DataList("VehiclePlateList")]
    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule(IsRequired = true, Transform = TransformKind.UpperCase)]
    public string VehiclePlate { get; set; } = string.Empty;

    [DataDetail]
    public VehicleDetail VehicleDetail { get; set; } = new();

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("VehicleStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Decommissioned", "Active", "Sold Date")]
    [ColumnStateGuard(2, StateConditionKind.DateExpired, "Expired", "Active", "Road Tax Expiry", "Inspection Expiry", "Insurance Expiry", "Permit Expiry")]
    public VehicleStatus VehicleStatus { get; set; } = VehicleStatus.Active;

    // Child Collections matching legacy manager names exactly
    public List<Maintenance> Maintenances { get; set; } = new();
    public List<VehiclePermit> Permits { get; set; } = new();
    public List<RoadTax> RoadTaxes { get; set; } = new();
    public List<Inspection> Inspections { get; set; } = new();
    public List<VehicleInsurance> Insurances { get; set; } = new();
    public List<OdometerLog> OdometerLogs { get; set; } = new();

    // Convenience alias for Odometers
    public List<OdometerLog> Odometers
    {
        get => OdometerLogs;
        set => OdometerLogs = value;
    }

    public VehicleStatus EvaluateStatus(DateOnly asOfDate)
    {
        if (VehicleDetail.SoldDate.HasValue && VehicleDetail.SoldDate.Value <= asOfDate)
            return VehicleStatus.Decommissioned;

        bool inWorkshop = Maintenances.Any(m =>
            m.MaintenanceStatus == MaintenanceStatus.In_Progress ||
            (m.DateIn <= asOfDate && (!m.DateCompleted.HasValue || m.DateCompleted.Value > asOfDate)));

        if (inWorkshop)
            return VehicleStatus.In_Shop;

        bool roadTaxExpired = VehicleDetail.RoadTaxExpiry.HasValue && VehicleDetail.RoadTaxExpiry.Value < asOfDate;
        bool inspectionExpired = VehicleDetail.InspectionExpiry.HasValue && VehicleDetail.InspectionExpiry.Value < asOfDate;
        bool insuranceExpired = VehicleDetail.InsuranceExpiry.HasValue && VehicleDetail.InsuranceExpiry.Value < asOfDate;
        bool permitExpired = VehicleDetail.PermitExpiry.HasValue && VehicleDetail.PermitExpiry.Value < asOfDate;

        if (roadTaxExpired || inspectionExpired || insuranceExpired || permitExpired)
            return VehicleStatus.Expired;

        return VehicleStatus == VehicleStatus.Reserved ? VehicleStatus.Reserved : VehicleStatus.Active;
    }
}

public record VehicleDetail
{
    // Restored for backward compatibility with VehicleManager.ServerSyncCommands
    public string VehiclePlate { get; set; } = string.Empty;

    [DataList("VehicleTypeList")]
    [ExcelColumn("Vehicle Type", isRequired: true)]
    [ColumnRule("VehicleTypeList", IsRequired = true)]
    public VehicleType Type { get; set; } = VehicleType.Lorry;

    [ExcelColumn("Vehicle Owner")]
    [ColumnRule("CompanyNameList", IsRequired = false, IsWarningOnly = true)]
    public string VehicleOwner { get; set; } = string.Empty;

    [ExcelColumn("Brand")]
    [ColumnRule(IsRequired = false)]
    public string Brand { get; set; } = string.Empty;

    [ExcelColumn("Model")]
    [ColumnRule(IsRequired = false)]
    public string Model { get; set; } = string.Empty;

    [ExcelColumn("VIN")]
    [ColumnRule(IsRequired = false, Transform = TransformKind.UpperCase)]
    public string VIN { get; set; } = string.Empty;

    [ExcelColumn("Bought Date")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? BoughtDate { get; set; }

    [ExcelColumn("Bought Price")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal BoughtPrice { get; set; } = 0.00m;

    [ExcelColumn("Sold Date")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? SoldDate { get; set; }

    [ExcelColumn("Sold Price")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal SoldPrice { get; set; } = 0.00m;

    [ExcelColumn("Insurance Expiry")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? InsuranceExpiry { get; set; }

    [ExcelColumn("Permit Expiry")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? PermitExpiry { get; set; }

    [ExcelColumn("Road Tax Expiry")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Bought Date", offset: 365)]
    public DateOnly? RoadTaxExpiry { get; set; }

    [ExcelColumn("Inspection Expiry")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true)]
    [ColumnDefault(DefaultKind.DateOffsetMonths, referenceHeader: "Bought Date", offset: 6)]
    public DateOnly? InspectionExpiry { get; set; }

    [ExcelColumn("Current Odometer")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int CurrentOdometer { get; set; } = 0;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 3. MAINTENANCE TABLE ("MaintenancesTable")
// =========================================================================

[ExcelTable("MaintenancesTable")]
public record Maintenance
{
    [ExcelColumn("Maintenance ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long MaintenanceID { get; set; }

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Service Type", isRequired: true)]
    [ColumnRule("ServiceTypeList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Regular_Service")]
    public ServiceType ServiceType { get; set; } = ServiceType.Regular_Service;

    [ExcelColumn("Date In", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly DateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Estimated Return")]
    [ColumnRule(IsRequired = false, IsWarningOnly = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Date In", offset: 3)]
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(3));

    [ExcelColumn("Date Completed")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? DateCompleted { get; set; }

    [ExcelColumn("Odometer In")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int OdometerIn { get; set; } = 0;

    [ExcelColumn("Odometer Out")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int OdometerOut { get; set; } = 0;

    [ExcelColumn("Workshop Name")]
    [ColumnRule("WorkshopsTable[Workshop Name]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Workshop Address", "WorkshopsTable[Workshop Name]", "WorkshopsTable[Workshop Address]")]
    public string WorkshopName { get; set; } = string.Empty;

    [ExcelColumn("Workshop Address")]
    [ColumnRule("WorkshopsTable[Workshop Address]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Workshop Name", "WorkshopsTable[Workshop Address]", "WorkshopsTable[Workshop Name]")]
    public string WorkshopAddress { get; set; } = string.Empty;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("MaintenanceStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "In_Progress")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Completed", "In_Progress", "Date Completed")]
    public MaintenanceStatus MaintenanceStatus { get; set; } = MaintenanceStatus.In_Progress;

    [ExcelColumn("Repair Total Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal RepairCost { get; set; } = 0.00m;

    [ExcelColumn("Tyre Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal TyreCost { get; set; } = 0.00m;

    [ExcelColumn("Brake Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal BrakeCost { get; set; } = 0.00m;

    [ExcelColumn("Clutch Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal ClutchCost { get; set; } = 0.00m;

    [ExcelColumn("Service Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal ServiceCost { get; set; } = 0.00m;

    [ExcelColumn("Fuel Pump Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal FuelPumpCost { get; set; } = 0.00m;

    [ExcelColumn("Aircond Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal AircondCost { get; set; } = 0.00m;

    [ExcelColumn("Wiring Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal WiringCost { get; set; } = 0.00m;

    [ExcelColumn("Battery Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal BatteryCost { get; set; } = 0.00m;

    [ExcelColumn("Towing Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal TowingCost { get; set; } = 0.00m;

    [ExcelColumn("Cushion Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal CushionCost { get; set; } = 0.00m;

    [ExcelColumn("Parts Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal PartsCost { get; set; } = 0.00m;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 4. WORKSHOPS DIRECTORY TABLE ("WorkshopsTable")
// =========================================================================

[ExcelTable("WorkshopsTable")]
public record Workshop
{
    [ExcelColumn("Workshop ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long WorkshopID { get; set; }

    [DataList("WorkshopNameList")]
    [ExcelColumn("Workshop Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string WorkshopName { get; set; } = string.Empty;

    [DataList("WorkshopAddressList")]
    [ExcelColumn("Workshop Address")]
    [ColumnRule(IsRequired = false)]
    public string WorkshopAddress { get; set; } = string.Empty;

    [ExcelColumn("Workshop Owner")]
    [ColumnRule(IsRequired = false)]
    public string WorkshopOwner { get; set; } = string.Empty;

    [ExcelColumn("Primary Contact")]
    [ColumnRule(IsRequired = false)]
    public string PrimaryContact { get; set; } = string.Empty;

    [ExcelColumn("Secondary Contact")]
    [ColumnRule(IsRequired = false)]
    public string SecondaryContact { get; set; } = string.Empty;
}

// =========================================================================
// 5. VEHICLE INSURANCE & INSURANCE COMPANIES ("InsurancesTable", "InsuranceCompaniesTable")
// =========================================================================

[ExcelTable("InsurancesTable")]
public record VehicleInsurance
{
    [ExcelColumn("Insurance Renewal ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long InsuranceID { get; set; }

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Insurance Company", isRequired: true)]
    [ColumnRule("InsuranceCompaniesTable[Insurance Company]", IsRequired = true)]
    public string InsuranceCompany { get; set; } = string.Empty;

    [ExcelColumn("Insurance Type")]
    [ColumnRule("InsuranceTypeList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Comprehensive")]
    public InsuranceType InsuranceType { get; set; } = InsuranceType.Comprehensive;

    [ExcelColumn("Renewal Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Renewal Date", offset: 365)]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Amount")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("InsuranceStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public InsuranceStatus Status { get; set; } = InsuranceStatus.Active;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

[ExcelTable("InsuranceCompaniesTable")]
public record InsuranceCompanyRecord
{
    [DataList("InsuranceCompanyList")]
    [ExcelColumn("Insurance Company", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string InsuranceCompany { get; set; } = string.Empty;

    [ExcelColumn("Company Contact Number")]
    [ColumnRule(IsRequired = false)]
    public string CompanyContactNumber { get; set; } = string.Empty;

    [ExcelColumn("Company Address")]
    [ColumnRule(IsRequired = false)]
    public string CompanyAddress { get; set; } = string.Empty;

    [ExcelColumn("Insurance Agent")]
    [ColumnRule(IsRequired = false)]
    public string InsuranceAgent { get; set; } = string.Empty;

    [ExcelColumn("Agent Contact Number")]
    [ColumnRule(IsRequired = false)]
    public string AgentContactNumber { get; set; } = string.Empty;
}

// =========================================================================
// 6. PERMITS, ROAD TAXES, AND INSPECTIONS TABLES
// =========================================================================

[ExcelTable("PermitsTable")]
public record VehiclePermit
{
    [ExcelColumn("Permit ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long PermitID { get; set; }

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Permit Type")]
    [ColumnRule("VehiclePermitTypeList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "APAD_Carrier_A")]
    public VehiclePermitType PermitType { get; set; } = VehiclePermitType.APAD_Carrier_A;

    [ExcelColumn("Renewal Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Renewal Date", offset: 365)]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Amount")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("VehiclePermitStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public VehiclePermitStatus Status { get; set; } = VehiclePermitStatus.Active;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

[ExcelTable("RoadTaxesTable")]
public record RoadTax
{
    [ExcelColumn("Road Tax ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long RoadTaxID { get; set; }

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Renewal Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly RenewDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Renewal Date", offset: 365)]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Amount")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("RoadTaxStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Active", "Expiry Date")]
    public RoadTaxStatus Status { get; set; } = RoadTaxStatus.Active;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

[ExcelTable("InspectionsTable")]
public record Inspection
{
    [ExcelColumn("Inspection ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long InspectionID { get; set; }

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Schedule Date In")]
    [ColumnRule(IsRequired = false)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly ScheduleDateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Date Out", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly DateOut { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Expiry Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetMonths, referenceHeader: "Date Out", offset: 6)]
    public DateOnly ExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddMonths(6));

    [ExcelColumn("Amount")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Status", isRequired: true)]
    [ColumnRule("InspectionStatusList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Scheduled")]
    [ColumnStateGuard(1, StateConditionKind.DateExpired, "Expired", "Completed", "Expiry Date")]
    public InspectionStatus Status { get; set; } = InspectionStatus.Scheduled;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

// =========================================================================
// 7. ODOMETER LOGS & FUEL CARDS ("OdometersTable", "FuelCardsTable")
// =========================================================================

[ExcelTable("OdometersTable")]
public record OdometerLog
{
    [ExcelColumn("Fuel Transaction ID", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long OdometerID { get; set; }

    [ExcelColumn("Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Vehicle Plate", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", IsRequired = true)]
    [ColumnSync("Vehicle Type", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    public string VehiclePlate { get; set; } = string.Empty;

    [ExcelColumn("Vehicle Type")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Vehicle Plate", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public VehicleType VehicleType { get; set; } = VehicleType.None;

    [ExcelColumn("Driven By")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true, IsMultiSelect = true)]
    public string DrivenBy { get; set; } = string.Empty;

    [ExcelColumn("Project")]
    [ColumnRule("ProjectIDList", IsRequired = false, IsWarningOnly = true)]
    public string Project { get; set; } = string.Empty;

    [ExcelColumn("Odometer Before")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int OdometerBefore { get; set; } = 0;

    [ExcelColumn("Current Odometer")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public int CurrentOdometer { get; set; } = 0;

    [ExcelColumn("Fuel Card Number")]
    [ColumnRule("FuelCardsTable[Fuel Card Number]", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Fuel Company", "FuelCardsTable[Fuel Card Number]", "FuelCardsTable[Fuel Company Name]")]
    public string FuelCardNumber { get; set; } = string.Empty;

    [ExcelColumn("Fuel Company")]
    [ColumnRule("FuelCompanyNamesTable[Fuel Company Name]", IsRequired = false, IsWarningOnly = true)]
    public string FuelCompany { get; set; } = string.Empty;

    [ExcelColumn("Station Name")]
    [ColumnRule(IsRequired = false)]
    public string StationName { get; set; } = string.Empty;

    [ExcelColumn("Receipt Number")]
    [ColumnRule(IsRequired = false)]
    public string ReceiptNumber { get; set; } = string.Empty;

    [ExcelColumn("Amount")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    public decimal Amount { get; set; } = 0.00m;

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;
}

[ExcelTable("FuelCardsTable")]
public record FuelCardRecord
{
    [DataList("FuelCardNumberList")]
    [ExcelColumn("Fuel Card Number", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string FuelCardNumber { get; set; } = string.Empty;

    [ExcelColumn("Fuel Company Name", isRequired: true)]
    [ColumnRule("FuelCompanyNamesTable[Fuel Company Name]", IsRequired = true)]
    public string FuelCompanyName { get; set; } = string.Empty;

    [ExcelColumn("Card Owner ID", isRequired: true)]
    [ColumnRule("AllVehiclesTable[Vehicle Plate]", "StaffIDList", IsRequired = true)]
    [ColumnSync("Card Owner Name", "AllVehiclesTable[Vehicle Plate]", "AllVehiclesTable[Vehicle Type]")]
    [ColumnSync("Card Owner Name", "StaffIDList", "StaffNameList")]
    public string CardOwnerID { get; set; } = string.Empty;

    [ExcelColumn("Card Owner Name", isRequired: true)]
    [ColumnRule("VehicleTypeList", "StaffNameList", IsRequired = true)]
    [ColumnSync("Card Owner ID", "StaffNameList", "StaffIDList")]
    [ColumnSync("Card Owner ID", "AllVehiclesTable[Vehicle Type]", "AllVehiclesTable[Vehicle Plate]")]
    public string CardOwnerName { get; set; } = string.Empty;
}

[ExcelTable("FuelCompanyNamesTable")]
public record FuelCompanyNameRecord
{
    [DataList("FuelCompanyNameList")]
    [ExcelColumn("Fuel Company Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string FuelCompanyName { get; set; } = string.Empty;
}