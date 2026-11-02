namespace SKSB_App.Core.Types;

using System;
using SKSB_App.Core.Types.Attributes;

// =========================================================================
// 1. PROJECT DOMAIN ENUMS (Auto-discovered into 'Lookups' sheet)
// =========================================================================

public enum ProjectType
{
    None,
    Civil_Works,
    Pipe_Laying,
    HDD_Subduct,
    Cable_Pulling,
    Road_Reinstatement,
    Maintenance,
    Telco_Infra,
    Other
}

public enum ProjectStatus
{
    Active,
    Completed,
    On_Hold,
    Cancelled,
    Expired,
    None
}

// =========================================================================
// 2. PROJECT RECORD ("AllProjectsTable", "ProjectsTable")
// =========================================================================

[ExcelTable("AllProjectsTable")]
[ExcelTable("ProjectsTable")]
public record ProjectRecord
{
    [DataList("ProjectIDList")]
    [ExcelColumn("Project ID", isRequired: true)]
    [ColumnRule(IsRequired = true, Transform = TransformKind.UpperCase)]
    public string ProjectID { get; set; } = string.Empty;

    [DataDetail]
    public ProjectDetails ProjectDetails { get; set; } = new();
}

public record ProjectDetails
{
    [DataList("ProjectNameList")]
    [ExcelColumn("Project Name", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    public string ProjectName { get; set; } = string.Empty;

    [ExcelColumn("Project Type", isRequired: true)]
    [ColumnRule("ProjectTypeList", IsRequired = true)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Civil_Works")]
    public ProjectType ProjectType { get; set; } = ProjectType.Civil_Works;

    [ExcelColumn("Company Name", isRequired: true)]
    [ColumnRule("CompaniesTable[Short Company Name]", "CompanyNameList", IsRequired = true)]
    public string CompanyName { get; set; } = string.Empty;

    [ExcelColumn("Start Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.Today)]
    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    [ExcelColumn("Target End Date", isRequired: true)]
    [ColumnRule(IsRequired = true)]
    [ColumnDefault(DefaultKind.DateOffsetDays, referenceHeader: "Start Date", offset: 365)]
    public DateOnly TargetEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today.AddDays(365));

    [ExcelColumn("Actual End Date")]
    [ColumnRule(IsRequired = false)]
    public DateOnly? ActualEndDate { get; set; }

    [ExcelColumn("Status")]
    [ColumnRule("ProjectStatusList", IsRequired = false)]
    [ColumnDefault(DefaultKind.Literal, literalValue: "Active")]
    [ColumnStateGuard(1, StateConditionKind.ColumnNotEmpty, "Completed", "Active", "Actual End Date")]
    [ColumnStateGuard(2, StateConditionKind.DateExpired, "Expired", "Active", "Target End Date")]
    public ProjectStatus Status { get; set; } = ProjectStatus.Active;
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