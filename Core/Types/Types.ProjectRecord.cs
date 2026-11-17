namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
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

/// <summary>Which of the three master sheets a project lives on (Air Selangor iJob / MWO, or Private).</summary>
public enum ProjectCategory
{
    iJob,
    MWO,
    PVT,
    None
}

// =========================================================================
// 2. PROJECT RECORD (iJobProjectsTable / MWOProjectsTable / PVTProjectsTable)
// =========================================================================

[ExcelTable("iJobProjectsTable")]
[ExcelTable("MWOProjectsTable")]
[ExcelTable("PVTProjectsTable")]
[ExcelTable("AllProjectsTable")]
public record ProjectRecord
{
    [DataList("ProjectIDList")]
    [ExcelColumn("Project Code", isRequired: true)]
    [ColumnRule(IsRequired = true, Transform = TransformKind.UpperCase)]
    public string ProjectCode { get; set; } = string.Empty;

    /// <summary>Set from the sheet the row was extracted from; not a table column.</summary>
    public ProjectCategory Category { get; set; } = ProjectCategory.iJob;

    [DataDetail]
    public ProjectDetails ProjectDetails { get; set; } = new();

}

public record ProjectDetails
{
    [ExcelColumn("Reference No")]
    [ColumnRule(IsRequired = false)]
    public string ReferenceNo { get; set; } = string.Empty;

    /// <summary>Multi-select; stored in Excel and PostgreSQL as "Pipe_Repair, Road_Reinstatement".</summary>
    [ExcelColumn("Type", isRequired: true)]
    [ColumnRule("ProjectTypeList", IsRequired = true, IsMultiSelect = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Pipe_Repair")]
    public List<ProjectType> ProjectType { get; set; } = new() { SKSB_App.Core.Types.ProjectType.Pipe_Repair };

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

    [ExcelColumn("Company ID", isRequired: true)]
    [ColumnRule("CompaniesTable[Company ID]", IsRequired = true)]
    [ColumnSync("Company Name", "CompaniesTable[Company ID]", "CompaniesTable[Short Company Name]")]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long CompanyID { get; set; }

    [ExcelColumn("Company Name")]
    [ColumnRule("CompaniesTable[Short Company Name]", IsRequired = false)]
    [ColumnSync("Company ID", "CompaniesTable[Short Company Name]", "CompaniesTable[Company ID]")]
    public string CompanyName { get; set; } = string.Empty;

    [ExcelColumn("Staff ID")]
    [ColumnRule("StaffIDList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("PIC", "StaffIDList", "StaffShortNameList")]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long? StaffID { get; set; }

    [ExcelColumn("PIC")]
    [ColumnRule("StaffShortNameList", IsRequired = false, IsWarningOnly = true)]
    [ColumnSync("Staff ID", "StaffShortNameList", "StaffIDList")]
    public string PIC { get; set; } = string.Empty;

    [ExcelColumn("District")]
    [ColumnRule("ProjectDistrictList", IsRequired = false)]
    public ProjectDistrict District { get; set; } = ProjectDistrict.None;

    [ExcelColumn("Address")]
    [ColumnRule(IsRequired = false)]
    public string Address { get; set; } = string.Empty;

    [ExcelColumn("Start Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    [ColumnFormat(ColumnFormatKind.Date)]
    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Target End Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Start Date", offset: 365)]
    [ColumnDateGreaterThan("Start Date")]
    [ColumnFormat(ColumnFormatKind.Date)]
    public DateOnly TargetEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Actual End Date")]
    [ColumnRule(IsRequired = false)]
    [ColumnFormat(ColumnFormatKind.Date)]
    public DateOnly? ActualEndDate { get; set; }

    [ExcelColumn("Estimated Cost")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal EstimatedCost { get; set; }

    [ExcelColumn("Promised Return")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal PromisedReturn { get; set; }

    [ExcelColumn("Actual Return")]
    [ColumnRule(IsRequired = false, AllowNonNegativeZero = true)]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal ActualReturn { get; set; }

    [ExcelColumn("Remarks")]
    [ColumnRule(IsRequired = false)]
    public string Remarks { get; set; } = string.Empty;

    // ---- Work / cost totals (plain columns; Actual Project Cost and Profit / Margin are calculated below) ----
    [ExcelColumn("Days Worked")]
    [ColumnFormat(ColumnFormatKind.Days)]
    public int? DaysWorked { get; set; }

    /// <summary>Sum of Regular Shifts across every daily summary row.</summary>
    [ExcelColumn("Total Regular Shifts")]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public decimal? TotalRegularShifts { get; set; }

    [ExcelColumn("Total Hours")]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal? TotalHours { get; set; }

    [ExcelColumn("Overtime Shifts")]
    [ColumnFormat(ColumnFormatKind.Shifts)]
    public int? OvertimeShifts { get; set; }

    [ExcelColumn("Total Overtime Hours")]
    [ColumnFormat(ColumnFormatKind.Hours)]
    public decimal? TotalOvertimeHours { get; set; }

    [ExcelColumn("Total Odometer")]
    [ColumnFormat(ColumnFormatKind.Kilometers)]
    public decimal? TotalOdometer { get; set; }

    [ExcelColumn("Fuel Cost")]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? FuelCost { get; set; }

    [ExcelColumn("Labour Cost")]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? LabourCost { get; set; }

    [ExcelColumn("Claim Cost")]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? ClaimCost { get; set; }

    // ---- Excel formula columns (Z:AA): never written by PopulateTableSafe; mirrored in C# for the database ----
    [ColumnCalculate("[Fuel Cost] + [Labour Cost] + [Claim Cost]")]
    [ExcelColumn("Actual Project Cost")]
    [ColumnFormat(ColumnFormatKind.Currency)]
    public decimal? ActualProjectCost { get; set; }

    [ColumnCalculate("IF([Actual Return] > 0, [Actual Return], [Promised Return]) - [Actual Project Cost]")]
    [ExcelColumn("Profit / Margin")]
    [ColumnFormat(ColumnFormatKind.Currency)]
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
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.AutoIncrement)]
    [ColumnFormat(ColumnFormatKind.Id, zero: ZeroBehavior.Blank)]
    public long CompanyID { get; set; }

    [DataList("LegalCompanyNameList")]
    [ExcelColumn("Legal Company Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string LegalCompanyName { get; set; } = string.Empty;

    [DataList("CompanyNameList")]
    [ExcelColumn("Short Company Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string ShortCompanyName { get; set; } = string.Empty;
}
