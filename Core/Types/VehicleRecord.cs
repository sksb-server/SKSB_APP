namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Linq;

public record VehicleRecord
{
    public string VehiclePlate { get; set; } = string.Empty;
    public VehicleDetail VehicleDetail { get; set; } = new VehicleDetail();
    public List<RoadTax> RoadTaxes { get; set; } = new List<RoadTax>();
    public List<Inspection> Inspections { get; set; } = new List<Inspection>();
    public List<Maintenance> Maintenances { get; set; } = new List<Maintenance>();

    public VehicleStatus VehicleStatus
    {
        get
        {
            if (_vehicleStatus == VehicleStatus.Decommissioned)
                return VehicleStatus.Decommissioned;

            if (_vehicleStatus == VehicleStatus.Reserved)
                return VehicleStatus.Reserved;

            DateOnly today = DateOnly.FromDateTime(DateTime.Today);

            bool isInShop = Maintenances.Any(m =>
                m.MaintenanceStatus == MaintenanceStatus.InProgress ||
                m.MaintenanceStatus == MaintenanceStatus.AwaitingParts ||
                (today >= m.DateIn && (m.DateCompleted == null || today < m.DateCompleted)));

            if (isInShop)
                return VehicleStatus.In_Shop;

            if (today > VehicleDetail.RoadTaxExpiry || today > VehicleDetail.InspectionExpiry)
                return VehicleStatus.Expired;

            return _vehicleStatus;
        }
        set => _vehicleStatus = value;
    }
    private VehicleStatus _vehicleStatus = VehicleStatus.Active;
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
    public int Cost { get; set; } = 0;
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
        get => DateOut.AddMonths(1);
        set => _expiryDate = value;
    }
    private DateOnly _expiryDate = DateOnly.FromDateTime(DateTime.Today);
    public int Cost { get; set; } = 0;
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
    public string VehiclePlate { get; set; } = string.Empty; // <-- Unified to VehiclePlate
    public ServiceType ServiceType { get; set; } = ServiceType.None;
    public DateOnly DateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? DateCompleted { get; set; } = null;
    public int OdometerIn { get; set; } = 0;
    public int OdometerOut { get; set; } = 0;
    public decimal RepairCost { get; set; } = 0m;
    public string WorkshopName { get; set; } = string.Empty;
    public string WorkshopAddress { get; set; } = string.Empty;
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
    public string Remarks { get; set; } = string.Empty;
}