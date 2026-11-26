namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Text.Json.Serialization;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. PROJECT DOMAIN ENUMS (Auto-discovered into 'Lookups' sheet)
// =========================================================================

public enum ProjectType
{
    Pipe_Repair,
    Pipe_Laying,
    Comm_Pipe_Meter,
    Valve_Hydrant,
    Tapping_Connection,
    Road_Reinstatement,
    Civil_Earthworks,
    Machinery_Rental,
    General_Maintenance,
    Office_Improvement,
    IT_System,
    Admin_Internal,
    Other,
    None
}

/// <summary>Project types selectable on GeneralProjectsTable (dropdown 'GeneralProjectTypeList').</summary>
public enum GeneralProjectType
{
    Pipe_Repair,
    Pipe_Laying,
    Comm_Pipe_Meter,
    Valve_Hydrant,
    Tapping_Connection,
    Road_Reinstatement,
    Civil_Earthworks,
    Machinery_Rental,
    General_Maintenance,
    Other,
    None
}

/// <summary>Project types selectable on InternalProjectsTable (dropdown 'InternalProjectTypeList').</summary>
public enum InternalProjectType
{
    Office_Improvement,
    IT_System,
    Admin_Internal,
    Other,
    None
}

public enum ProjectDistrict
{
    PJ,
    KL,
    Klang,
    Gombak,
    Hulu_Langat,
    Sepang,
    Kuala_Langat,
    Hulu_Selangor,
    Kuala_Selangor,
    Sabak_Bernam,
    Putrajaya,
    Outstation,
    Other,
    None
}

public enum ProjectPriority
{
    Low,
    Standard,
    Urgent,
    Emergency,
    None
}

public enum ProjectStatus
{
    Planned,
    Active,
    Pending,
    On_Hold,
    Completed,
    Expired,
    Cancelled,
    None
}

/// <summary>
/// Project category persisted to <c>project_master.category</c>. iJob / MWO / PVT / Other live on the
/// "General Projects" sheet; Internal (non-operational) projects live on the "Internal Projects" sheet.
/// </summary>
public enum ProjectCategory
{
    iJob,
    MWO,
    PVT,
    Other,
    Internal,
    None
}

/// <summary>Categories selectable on GeneralProjectsTable (dropdown 'GeneralProjectCategoryList').</summary>
public enum GeneralProjectCategory
{
    iJob,
    MWO,
    PVT,
    Other,
    None
}

// =========================================================================
// 2. PROJECT RECORD (GeneralProjectsTable / InternalProjectsTable)
// =========================================================================

[ExcelTable("GeneralProjectsTable")]
[ExcelTable("InternalProjectsTable")]
[ExcelTable("AllProjectsTable")]
public record ProjectRecord
{
    [DataList("ProjectIDList")]
    [ExcelColumn("Project Code", isRequired: true)]
    [ColumnRule(IsRequired = true, Transform = TransformKind.UpperCase)]
    public string ProjectCode { get; set; } = string.Empty;

    /// <summary>'Category' column on GeneralProjectsTable; forced to Internal for InternalProjectsTable rows (no column there).</summary>
    [ExcelColumn("Category", isRequired: false)]
    [ColumnRule("GeneralProjectCategoryList", IsRequired = true)]
    public ProjectCategory Category { get; set; } = ProjectCategory.iJob;

    [DataDetail]
    public ProjectDetails ProjectDetails { get; set; } = new();

    /// <summary>Server <c>updated_at</c> as last loaded (optimistic-concurrency token); never serialized or shown.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

public record ProjectDetails
{
    [ExcelColumn("Project Name")]
    public string ProjectName { get; set; } = string.Empty;

    /// <summary>Whether the schedule board may auto-assign this project; always false for Internal projects.</summary>
    [ExcelColumn("Auto-Schedule")]
    [ColumnDefault(DefaultKind.Literal, literalValue: "TRUE", TableName = "GeneralProjectsTable")]
    public bool AutoSchedule { get; set; } = true;

    [ExcelColumn("Reference No")]
    public string ReferenceNo { get; set; } = string.Empty;

    /// <summary>Multi-select; stored in Excel and PostgreSQL as "Pipe_Repair, Road_Reinstatement".</summary>
    [ExcelColumn("Type", isRequired: true)]
    [ColumnRule("GeneralProjectTypeList", IsRequired = true, IsMultiSelect = true, TableName = "GeneralProjectsTable")]
    [ColumnRule("InternalProjectTypeList", IsRequired = true, IsMultiSelect = true, TableName = "InternalProjectsTable")]
    [ColumnRule("ProjectTypeList", IsRequired = true, IsMultiSelect = true)]
    public List<ProjectType> ProjectType { get; set; } = new();

    [ExcelColumn("Priority", isRequired: true)]
    [ColumnRule("ProjectPriorityList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Standard")]
    public ProjectPriority Priority { get; set; } = ProjectPriority.Standard;

    [ExcelColumn("Status")]
    [ColumnRule("ProjectStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Completed", "Active", "Actual End Date")]
    [ColumnStateGuard(2, StateConditionKind.DateExpired, "Expired", "Active", "Target End Date")]
    public ProjectStatus Status { get; set; } = ProjectStatus.Active;

    // Optional column: required on GeneralProjectsTable, absent on InternalProjectsTable.
    // Category-aware 'SQL Account Project' defaults are applied by ProjectManager's sheet-change observer.
    [ExcelColumn("Company ID", isRequired: false)]
    [ColumnRule("CompaniesTable[Company ID]", IsRequired = true)]
    [ColumnSync("Company Name", "CompaniesTable[Company ID]", "CompaniesTable[Short Company Name]")]
    public long CompanyID { get; set; }

    [ExcelColumn("Company Name")]
    [ColumnRule("CompaniesTable[Short Company Name]", IsRequired = false)]
    [ColumnSync("Company ID", "CompaniesTable[Short Company Name]", "CompaniesTable[Company ID]")]
    public string CompanyName { get; set; } = string.Empty;

    [ExcelColumn("SQL Account Project")]
    [ColumnRule("SqlAccountProjectList", IsRequired = false)]
    public string SqlAccountProjectCode { get; set; } = string.Empty;

    [ExcelColumn("Staff ID")]
    [ColumnRule("StaffIDList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("PIC", "StaffIDList", "StaffShortNameList")]
    public long? StaffID { get; set; }

    [ExcelColumn("PIC")]
    [ColumnRule("StaffShortNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffShortNameList", "StaffIDList")]
    public string PIC { get; set; } = string.Empty;

    [ExcelColumn("District")]
    [ColumnRule("ProjectDistrictList", IsRequired = false)]
    public ProjectDistrict District { get; set; } = ProjectDistrict.None;

    [ExcelColumn("Address")]
    public string Address { get; set; } = string.Empty;

    [ExcelColumn("Start Date", isRequired: true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Target End Date", isRequired: true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Start Date", offset: 365)]
    [ColumnDateGreaterThan("Start Date")]
    public DateOnly TargetEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Actual End Date")]
    [ColumnDateGreaterThan("Start Date", AllowEqual = true)]
    public DateOnly? ActualEndDate { get; set; }

    [ExcelColumn("Estimated Cost")]
    public decimal EstimatedCost { get; set; }

    [ExcelColumn("Promised Return")]
    public decimal PromisedReturn { get; set; }

    [ExcelColumn("Actual Return")]
    public decimal ActualReturn { get; set; }

    [ExcelColumn("Remarks")]
    public string Remarks { get; set; } = string.Empty;

    // ---- Work / cost totals (plain columns; Actual Project Cost and Profit / Margin are calculated below) ----
    [ExcelColumn("Days Worked")]
    public int? DaysWorked { get; set; }

    /// <summary>Sum of Regular Shifts across every daily summary row.</summary>
    [ExcelColumn("Total Regular Shifts")]
    public decimal? TotalRegularShifts { get; set; }

    [ExcelColumn("Total Hours")]
    public decimal? TotalHours { get; set; }

    [ExcelColumn("Overtime Shifts")]
    public int? OvertimeShifts { get; set; }

    [ExcelColumn("Total Overtime Hours")]
    public decimal? TotalOvertimeHours { get; set; }

    [ExcelColumn("Total Odometer")]
    public decimal? TotalOdometer { get; set; }

    [ExcelColumn("Fuel Cost")]
    public decimal? FuelCost { get; set; }

    [ExcelColumn("Labour Cost")]
    public decimal? LabourCost { get; set; }

    [ExcelColumn("Claim Cost")]
    public decimal? ClaimCost { get; set; }

    // ---- Excel formula columns (Z:AA): never written by PopulateTableSafe; mirrored in C# for the database ----
    [ColumnCalculate("[Fuel Cost] + [Labour Cost] + [Claim Cost]")]
    [ExcelColumn("Actual Project Cost")]
    public decimal? ActualProjectCost { get; set; }

    [ColumnCalculate("IF([Actual Return] > 0, [Actual Return], [Promised Return]) - [Actual Project Cost]")]
    [ExcelColumn("Profit / Margin")]
    public decimal? ProfitMargin { get; set; }
}

// =========================================================================
// 3. COMPANY RECORD ("CompaniesTable", "AllCompaniesTable")
// =========================================================================

[ExcelTable("CompaniesTable")]
[ExcelTable("AllCompaniesTable")]
public record CompanyRecord
{
    [DataList("CompanyIDList")]
    [ExcelColumn("Company ID", isRequired: true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    public long CompanyID { get; set; }

    [DataList("LegalCompanyNameList")]
    [ExcelColumn("Legal Company Name", isRequired: true)]
    public string LegalCompanyName { get; set; } = string.Empty;

    [DataList("CompanyNameList")]
    [ExcelColumn("Short Company Name", isRequired: true)]
    public string ShortCompanyName { get; set; } = string.Empty;

    [ExcelColumn("iJob SQL Account Project")]
    [ColumnRule("SqlAccountProjectList", IsRequired = false)]
    public string DefaultSqlAccountProject { get; set; } = string.Empty;

    /// <summary>Server <c>updated_at</c> as last loaded (optimistic-concurrency token); never serialized or shown.</summary>
    [JsonIgnore]
    public DateTime? ServerUpdatedAt { get; set; }
}

// =========================================================================
// 4. PROJECT SUMMARY ROW ("ProjectSummaryTable")
// One row per working day of the project picked in the sheet's selector cells (Project Summary!A2 = project,
// B2 = date filter; see ParameterCells). The table is a derived VIEW: nothing here is persisted to PostgreSQL.
//
// Source wiring is not implemented yet. Columns marked "(derived)" are meant to be filled from the daily
// schedule / attendance tables and the vehicle module, and must never be typed by the user. Only
// "Total Hours" is computed in-row today; the others stay plain properties until their sources exist,
// and the sheet-level lock for them is the template's custom '=FALSE' validation.
// =========================================================================

[ExcelTable("ProjectSummaryTable")]
public record ProjectSummaryRow
{
    [ExcelColumn("Date")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Team Lead")]
    [ColumnRule("StaffNameList", IsRequired = false, IsWarningOnly = true, IsReadOnly = true)]
    public string TeamLead { get; set; } = string.Empty;

    [ExcelColumn("Time Start")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public TimeOnly TimeStart { get; set; } = TimeOnly.MinValue;

    [ExcelColumn("Time End")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public TimeOnly TimeEnd { get; set; } = TimeOnly.MinValue;

    /// <summary>Strictly calculated: worked hours between Time Start and Time End, rounded down to 0.5.</summary>
    [JsonIgnore]
    [ColumnCalculate("FLOOR(HOURS([Time Start], [Time End]), 0.5)")]
    [ExcelColumn("Total Hours")]
    public decimal? TotalHours { get; set; }

    /// <summary>(derived) Count of distinct workers present that day.</summary>
    [ExcelColumn("Workers")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public int Workers { get; set; }

    [ExcelColumn("Regular Shifts")]
    [ColumnRule(IsRequired = false, IsReadOnly = true, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    public decimal RegularShifts { get; set; }

    [ExcelColumn("Overtime Shifts")]
    [ColumnRule(IsRequired = false, IsReadOnly = true, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    public decimal OvertimeShifts { get; set; }

    [ExcelColumn("Overtime Hours")]
    [ColumnRule(IsRequired = false, IsReadOnly = true, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    public decimal OvertimeHours { get; set; }

    [ExcelColumn("Additional OT Hours")]
    [ColumnRule(IsRequired = false, IsReadOnly = true, AllowNonNegativeZero = true, RoundDownStep = 0.5)]
    public decimal AdditionalOTHours { get; set; }

    [ExcelColumn("Is Holiday")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public bool IsHoliday { get; set; }

    /// <summary>(derived) Sum of the day's attendance salaries (DailyAttendanceTable "Total Earned").</summary>
    [ExcelColumn("Labour Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public decimal? LabourCost { get; set; }

    /// <summary>(derived, multi-select, not an input) Plates of every vehicle used that day, e.g. "VAB1234, WXY5678".</summary>
    [ExcelColumn("Vehicles Used")]
    [ColumnRule("VehiclePlateList", IsRequired = false, IsReadOnly = true, IsMultiSelect = true)]
    public List<string> VehiclesUsed { get; set; } = new();

    /// <summary>(derived, multi-select, not an input) Distinct types of the Vehicles Used, e.g. "Lorry, Van".</summary>
    [JsonIgnore]
    [ExcelColumn("Vehicle Types")]
    [ColumnRule("VehicleTypeList", IsRequired = false, IsReadOnly = true, IsMultiSelect = true)]
    public List<VehicleType> VehicleTypes { get; set; } = new();

    /// <summary>(derived) Distance / engine hours of the vehicles used (per-row unit follows the vehicle type).</summary>
    [ExcelColumn("Odometer Used")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public decimal? OdometerUsed { get; set; }

    [ExcelColumn("Fuel Cost")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public decimal? FuelCost { get; set; }

    [ExcelColumn("Claims Made")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public decimal? ClaimsMade { get; set; }

    [ExcelColumn("Pending Claims")]
    [ColumnRule(IsRequired = false, IsReadOnly = true)]
    public decimal? PendingClaims { get; set; }
}
