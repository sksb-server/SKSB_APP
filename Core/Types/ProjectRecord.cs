namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;

public record ProjectRecord
{
    public long ProjectID { get; set; } = 0;
    public ProjectMaster Master { get; set; } = new ProjectMaster();
}

public enum ProjectClassification
{
    iJob,   // Recurring daily operational tasks
    PVT,    // Project / Vehicle Testing / Validation
    SD,     // Site Delivery / Specialized Deployment
    AdHoc,  // Unscheduled or one-off tasks
    None
}

public enum ProjectStatus
{
    Planned,
    Active,
    OnHold,
    Completed,
    Cancelled
}

public enum ProjectPriority
{
    Low,
    Standard,
    Urgent,
    Critical
}

public record ProjectMaster
{
    public ProjectClassification ProjectClass { get; set; } = ProjectClassification.None;
    public string ProjectID { get; set; } = string.Empty;
    public string ProjectName { get; set; } = string.Empty;
    public ProjectPriority Priority { get; set; } = ProjectPriority.Standard;
    public ProjectStatus Status { get; set; } = ProjectStatus.Planned;

    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly TargetEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly ActualEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    // Resource sizing needed by DailyScheduler
    public int RequiredHeadcount { get; set; } = 1;
    public VehicleType RequiredVehicleType { get; set; } = VehicleType.None;
    public string RequiredSkills { get; set;  } = string.Empty;

    public string Address { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;
}