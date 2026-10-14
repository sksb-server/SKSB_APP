namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;

public record VehicleRecord
{
    public string VehiclePlate { get; set; } = string.Empty;
    public VehicleDetail VehicleDetail { get; set; } = new();
    public List<RoadTax> RoadTaxes { get; set; } = new();
    public List<Inspection> Inspections { get; set; } = new();
    public List<Maintenance> Maintenances { get; set; } = new();

    /// <summary>
    /// Real-time status evaluation based on current date.
    /// </summary>
    public VehicleStatus VehicleStatus
    {
        get => EvaluateStatusOn(DateOnly.FromDateTime(DateTime.Today));
        set => _vehicleStatus = value;
    }
    private VehicleStatus _vehicleStatus = VehicleStatus.Active;

    /// <summary>
    /// Evaluates whether this asset was operational on a specific target date.
    /// Used by DailyScheduler and DataToExcel.
    /// </summary>
    public VehicleStatus EvaluateStatusOn(DateOnly targetDate)
    {
        if (_vehicleStatus == VehicleStatus.Decommissioned)
            return VehicleStatus.Decommissioned;

        if (_vehicleStatus == VehicleStatus.Reserved)
            return VehicleStatus.Reserved;

        // 1. Check if under maintenance during target date
        bool isInShop = Maintenances.Any(m =>
            m.DateIn <= targetDate &&
            (m.DateCompleted == null || targetDate <= m.DateCompleted.Value));

        if (isInShop)
            return VehicleStatus.In_Shop;

        // 2. Regulatory expiration check
        if (targetDate > VehicleDetail.RoadTaxExpiry || targetDate > VehicleDetail.InspectionExpiry)
            return VehicleStatus.Expired;

        return VehicleStatus.Active;
    }

    public bool IsAvailableOn(DateOnly targetDate) => EvaluateStatusOn(targetDate) == VehicleStatus.Active;
}

public enum VehicleType
{
    Car,
    Van,
    Pickup,
    Lorry_1T,
    Lorry_3T,
    Lorry_5T,
    Lorry_Crane,
    Lorry_Tipper,
    Backhoe,
    Mini_Excavator,
    Excavator,
    Water_Tanker,
    Suction_Tanker,
    Jetting_Machine,
    Road_Roller,
    Baby_Roller,
    Forklift,
    Tractor,
    Tow_Truck,
    Motorcycle,
    Trailer_Prime_Mover,
    None
}

public enum VehicleStatus
{
    Active,
    In_Shop,
    Reserved,
    Expired,
    Decommissioned
}

public record VehicleDetail
{
    public string VehiclePlate { get; set; } = string.Empty;
    public string VIN { get; set; } = string.Empty;
    public VehicleType Type { get; set; } = VehicleType.None;
    public string Brand { get; set; } = string.Empty;
    public string Model { get; set; } = string.Empty;
    public DateOnly RoadTaxExpiry { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly InspectionExpiry { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public int CurrentOdometer { get; set; } = 0;
}

public record RoadTax
{
    public long RoadTaxID { get; set; } = 0;
    public string VehiclePlate { get; set; } = string.Empty;
    public DateOnly RenewalDate { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public decimal Cost { get; set; } = 0.00m; // Updated to decimal
}

public enum InspectionStatus
{
    Active,
    Expired
}

public record Inspection
{
    public long InspectionID { get; set; } = 0;
    public string VehiclePlate { get; set; } = string.Empty;
    public DateOnly ScheduleDateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly DateOut { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public int OdometerIn { get; set; } = 0;
    public int OdometerOut { get; set; } = 0;

    public DateOnly ExpiryDate
    {
        get => _expiryDate ?? DateOut.AddMonths(6); // Default 6 months for commercial Puspakom
        set => _expiryDate = value;
    }
    private DateOnly? _expiryDate;

    public decimal Cost { get; set; } = 0.00m; // Updated to decimal
    public string Address { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;

    public InspectionStatus InspectionStatus
    {
        get => DateOnly.FromDateTime(DateTime.Today) > ExpiryDate
            ? InspectionStatus.Expired
            : _inspectionStatus;
        set => _inspectionStatus = value;
    }
    private InspectionStatus _inspectionStatus = InspectionStatus.Active;
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
    InProgress,
    AwaitingParts,
    Done
}

public record Maintenance
{
    public long MaintenanceID { get; set; } = 0;
    public string VehiclePlate { get; set; } = string.Empty;
    public ServiceType ServiceType { get; set; } = ServiceType.None;
    public DateOnly DateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? DateCompleted { get; set; } = null;
    public int OdometerIn { get; set; } = 0;
    public int OdometerOut { get; set; } = 0;
    public decimal RepairCost { get; set; } = 0.00m;
    public string WorkshopName { get; set; } = string.Empty;
    public string WorkshopAddress { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;

    public MaintenanceStatus MaintenanceStatus
    {
        get
        {
            DateOnly today = DateOnly.FromDateTime(DateTime.Today);
            if (DateCompleted.HasValue && today >= DateCompleted.Value)
                return MaintenanceStatus.Done;

            if (today < DateIn)
                return MaintenanceStatus.Scheduled;

            return _maintenanceStatus == MaintenanceStatus.Done
                ? MaintenanceStatus.InProgress
                : _maintenanceStatus;
        }
        set => _maintenanceStatus = value;
    }
    private MaintenanceStatus _maintenanceStatus = MaintenanceStatus.InProgress;
}