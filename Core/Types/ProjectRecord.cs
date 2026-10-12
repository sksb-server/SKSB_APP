namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;

public record ProjectRecord
{
    public string ProjectID { get; set; } = string.Empty;
    public ProjectDetails ProjectDetails { get; set; } = new();
    public List<DailyProjectSummary> DailyProjectSummary { get; set; } = new();

    /// <summary>
    /// Stored or evaluated aggregate summary of the project.
    /// </summary>
    public ProjectSummary ProjectSummary
    {
        get => _projectSummary ?? ProjectSummary.Create(ProjectDetails, DailyProjectSummary);
        set => _projectSummary = value;
    }
    private ProjectSummary? _projectSummary;
}

public enum CIDBGrade
{
    G1,
    G2,
    G3,
    G4,
    G5,
    G6,
    G7
}

public enum District
{
    PJ,
    Klang,
    Gombak,
    Sepang,
    Project,
    Emergency,
    Private
}

public enum ActiveContractStatus
{
    Committed,
    Available,
    Suspended
}

public record CompanyRecord
{
    public long CompanyID { get; set; } = 0;
    public string LegalCompanyName { get; set; } = string.Empty;
    public string ShortCompanyName { get; set; } = string.Empty;
    public string SSMNumber { get; set; } = string.Empty;
    public string VendorCode { get; set; } = string.Empty;
    public string SPANNumber { get; set; } = string.Empty;
    public DateOnly SPANExpiry { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public CIDBGrade CIDBGrade { get; set; } = CIDBGrade.G1;
    public DateOnly CIDBExpiryDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public District AssignedDistrict { get; set; } = District.PJ;

    public ActiveContractStatus ActiveContractStatus
    {
        get
        {
            if (_activeContractStatus == ActiveContractStatus.Suspended)
                return ActiveContractStatus.Suspended;

            DateOnly today = DateOnly.FromDateTime(DateTime.Today);
            if (today > SPANExpiry || today > CIDBExpiryDate)
                return ActiveContractStatus.Suspended;

            return string.IsNullOrWhiteSpace(ActiveProjectID)
                ? ActiveContractStatus.Available
                : ActiveContractStatus.Committed;
        }
        set => _activeContractStatus = value;
    }
    private ActiveContractStatus _activeContractStatus = ActiveContractStatus.Available;

    public string ActiveProjectID { get; set; } = string.Empty;
}

public enum ProjectType
{
    iJob,
    PVT,
    Private,
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

public record ProjectDetails
{
    public ProjectType ProjectType { get; set; } = ProjectType.None;
    public string ProjectID { get; set; } = string.Empty;
    public string ProjectName { get; set; } = string.Empty;
    public string CompanyName { get; set; } = string.Empty;
    public long CompanyID { get; set; } = 0;
    public ProjectPriority Priority { get; set; } = ProjectPriority.Standard;
    public ProjectStatus Status { get; set; } = ProjectStatus.Planned;
    public string Address { get; set; } = string.Empty;
    public District District { get; set; } = District.PJ;
    public string Remarks { get; set; } = string.Empty;

    public DateOnly StartDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly TargetEndDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? ActualEndDate { get; set; } = null;

    // Financials
    public decimal EstProjectCost { get; set; } = 0.00m;
    public decimal ProjectTotalCost { get; set; } = 0.00m;
    public decimal PromisedReturn { get; set; } = 0.00m;
    public decimal ActualReturn { get; set; } = 0.00m;

    public decimal Profit
    {
        get => _profit != 0.00m
            ? _profit
            : ActualReturn - ProjectTotalCost;
        set => _profit = value;
    }
    private decimal _profit = 0.00m;
}

public record ProjectSummary
{
    // 1. Metadata Rollup (from ProjectDetails)
    public string ProjectID { get; set; } = string.Empty;
    public ProjectStatus Status { get; set; } = ProjectStatus.Planned;
    public string Company { get; set; } = string.Empty;
    public DateOnly DateStart { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public string Address { get; set; } = string.Empty;

    // 2. Timeline & Operational Metrics
    public int ProjectLengthDays { get; set; } = 0;

    // 3. Financial Metrics
    public decimal PromisedReturn { get; set; } = 0.00m;
    public decimal ReturnAmount { get; set; } = 0.00m;
    public decimal EstimatedCost { get; set; } = 0.00m;
    public decimal TotalCost { get; set; } = 0.00m;
    public decimal ClaimMade { get; set; } = 0.00m;
    public decimal Profit => ReturnAmount > 0.00m
        ? ReturnAmount - TotalCost
        : ClaimMade - TotalCost;

    // 4. Resource Usage Sets
    public List<string> VehiclesUsed { get; set; } = new();
    public int OdometerUsed { get; set; } = 0;
    public List<string> InventoryUsed { get; set; } = new();
    public List<string> WorkerPresence { get; set; } = new();

    // 5. Labor Hours
    public decimal TotalWorkerHours { get; set; } = 0.0m;
    public decimal TotalOTHours { get; set; } = 0.0m;
    public decimal AdditionalOTHours { get; set; } = 0.0m;

    public static ProjectSummary Create(ProjectDetails details, IEnumerable<DailyProjectSummary>? dailyLogs)
    {
        var logs = (dailyLogs ?? Enumerable.Empty<DailyProjectSummary>()).ToList();

        var summary = new ProjectSummary
        {
            ProjectID = details.ProjectID,
            Status = details.Status,
            Company = details.CompanyName,
            DateStart = details.StartDate,
            Address = details.Address,
            PromisedReturn = details.PromisedReturn,
            ReturnAmount = details.ActualReturn,
            EstimatedCost = details.EstProjectCost
        };

        if (logs.Count == 0) return summary;

        summary.ProjectLengthDays = logs.Select(l => l.Date).Distinct().Count();
        summary.TotalCost = logs.Sum(l => l.TotalCost);
        summary.ClaimMade = logs.Sum(l => l.ClaimMade);
        summary.OdometerUsed = logs.Sum(l => l.OdometerUsed);

        summary.TotalWorkerHours = logs.Sum(l => l.TotalWorkerHours);
        summary.TotalOTHours = logs.Sum(l => l.OTHours);
        summary.AdditionalOTHours = logs.Sum(l => l.AdditionalOTHours);

        summary.VehiclesUsed = logs
            .SelectMany(l => l.VehiclesUsed)
            .Where(v => !string.IsNullOrWhiteSpace(v))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(v => v)
            .ToList();

        summary.InventoryUsed = logs
            .SelectMany(l => l.InventoryUsed)
            .Where(i => !string.IsNullOrWhiteSpace(i))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(i => i)
            .ToList();

        summary.WorkerPresence = logs
            .SelectMany(l => l.WorkerPresence)
            .Where(w => !string.IsNullOrWhiteSpace(w))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(w => w)
            .ToList();

        return summary;
    }
}

public record DailyProjectSummary
{
    public string ProjectID { get; set; } = string.Empty;
    public DateOnly Date { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public string Company { get; set; } = string.Empty;
    public TimeOnly TimeStart { get; set; } = TimeOnly.FromDateTime(DateTime.Now);
    public TimeOnly TimeEnd { get; set; } = TimeOnly.FromDateTime(DateTime.Now);

    public decimal TotalCost { get; set; } = 0.00m;
    public decimal ClaimMade { get; set; } = 0.00m;

    public List<string> Request { get; set; } = new();
    public List<string> VehiclesUsed { get; set; } = new();
    public int OdometerUsed { get; set; } = 0;
    public List<string> InventoryUsed { get; set; } = new();
    public List<string> WorkerPresence { get; set; } = new();

    public decimal TotalWorkerHours { get; set; } = 0.0m;
    public decimal OTHours { get; set; } = 0.0m;
    public decimal AdditionalOTHours { get; set; } = 0.0m;
}