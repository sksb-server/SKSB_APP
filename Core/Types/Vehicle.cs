namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;

public record Vehicle
{
    public long VehicleID { get; set; } = 0;
    public VehicleMaster Master { get; set; } = new VehicleMaster();
    public List<MaintenanceRecord> MaintenanceLogs { get; set; } = new List<MaintenanceRecord>();
}

public enum VehicleType
{
    Van,
    Lorry_1T,
    Lorry_3T,
    Lorry_5T,
    Pickup,
    Forklift,
    None
}

public enum OperationalStatus
{
    Active,
    In_Shop,
    Reserved,
    Decommissioned
}

public enum ServiceType
{
    Routine_Service,
    Major_Repair,
    Tyre_Replacement,
    Inspection,
    Breakdown,
    None
}

public enum MaintenanceStatus
{
    Scheduled,
    In_Progress,
    Awaiting_Parts,
    Done
}

public record VehicleMaster
{
    public string RegistrationNumber { get; set; } = string.Empty;
    public VehicleType Type { get; set; } = VehicleType.None;
    public string Brand { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public decimal CapacityTonnage { get; set; } = 0m;

    public DateOnly RoadTaxExpiry { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly PuspakomExpiry { get; set; } = DateOnly.FromDateTime(DateTime.Today);

    public int CurrentOdometer { get; set; } = 0;
    public OperationalStatus Status { get; set; } = OperationalStatus.Active;
}

public record MaintenanceRecord
{
    public long MaintenanceID { get; set; } = 0;
    public long VehicleID { get; set; } = 0;
    public ServiceType ServiceType { get; set; } = ServiceType.None;

    public DateOnly DateOut { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? DateCompleted { get; set; }

    public int OdometerAtService { get; set; } = 0;
    public decimal CostAmount { get; set; } = 0m;
    public string WorkshopName { get; set; } = string.Empty;
    public MaintenanceStatus Status { get; set; } = MaintenanceStatus.Scheduled;
}