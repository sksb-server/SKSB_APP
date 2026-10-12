namespace SKSB_App.Core.Types;

using System;
using System.Collections.Generic;
using System.Diagnostics.Eventing.Reader;


public enum VehicleStatus
{
    Active,
    In_Shop,
    Reserved,
    Expired,
    Decommissioned
}

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
            // 1. Manual administrative overrides take precedence
            if (_vehicleStatus == VehicleStatus.Decommissioned)
                return VehicleStatus.Decommissioned;

            if (_vehicleStatus == VehicleStatus.Reserved)
                return VehicleStatus.Reserved;

            DateOnly today = DateOnly.FromDateTime(DateTime.Today);

            // 2. Check maintenance / workshop downtime
            bool isInShop = Maintenances.Any(m =>
                m.MaintenanceStatus == MaintenanceStatus.InProgress ||
                m.MaintenanceStatus == MaintenanceStatus.AwaitingParts ||
                (today >= m.DateIn && (m.DateCompleted == null || today < m.DateCompleted)));

            if (isInShop)
                return VehicleStatus.In_Shop;

            // 3. Regulatory check: Road Tax or Puspakom Inspection expired
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
    // --- 1. Light Passenger & Site Transport ---
    Car,                // Site coordination / engineer inspection (e.g. Proton, Perodua, Vios)
    Van,                // Crew transport / mobile tool unit (e.g. Toyota Hiace, Nissan Urvan)
    Pickup,             // 4x4 Site utility / rapid emergency crew (e.g. Hilux, D-Max)

    // --- 2. Rigid Commercial Lorries (Material & Pipe Hauling) ---
    Lorry_1T,           // 1-Tonne light lorry / bonded truck (GDL D)
    Lorry_3T,           // 3-Tonne medium open-top / pipe runner (GDL E)
    Lorry_5T,           // 5-Tonne heavy pipe transporter (GDL E)
    Lorry_Crane,        // Hiab / Lorry with mounted crane (pipe unloading & valve placement)
    Lorry_Tipper,       // Dump truck for excavated soil, crusher run, and asphalt/premix

    // --- 3. Excavation & Earthmoving Machinery ---
    Backhoe,            // Standard loader/excavator (JCB / Case 580) - core pipe repair machine
    Mini_Excavator,     // 1T–3T mini excavator (Kubota/Yanmar) for tight residential lanes / trenches
    Excavator,          // Standard tracked excavator (20-ton class for major trunk mains)

    // --- 4. Specialized Water & Pipe Maintenance Units ---
    Water_Tanker,       // Emergency potable water delivery or site flushing
    Suction_Tanker,     // Sludge / dewatering tanker for flooded pipe trenches
    Jetting_Machine,    // High-pressure pipe and culvert cleaning unit

    // --- 5. Roadworks & Compaction Machinery ---
    Road_Roller,        // Trench compaction roller (tamped asphalt reinstatement)
    Baby_Roller,        // 1-2 Ton walk-behind or ride-on trench roller

    // --- 6. Industrial, Yard & Auxiliary Plant ---
    Forklift,           // Yard/depot pipe and valve loading
    Tractor,            // Yard shunting or rural access
    Tow_Truck,          // Equipment and breakdown recovery
    Motorcycle,         // Leak inspector / site dispatch (B2)
    Trailer_Prime_Mover,// Articulated haulage for large diameter pipes / low-loader transport (GDL E Full)

    None                // Default/unassigned
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
    public DateOnly RenewalDate { get; set; } = DateOnly.FromDateTime(DateTime.Now);
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
    public DateOnly ScheduleDateIn { get; set; } = DateOnly.FromDateTime(DateTime.Now);
    public DateOnly DateOut { get; set; } = DateOnly.FromDateTime(DateTime.Now);
    public int OdometerIn { get; set; } = 0;
    public int OdometerOut { get; set; } = 0;
    public DateOnly ExpiryDate
    {
        get => DateOut.AddMonths(1);
        set => _expiryDate = value;
    } 
    private DateOnly _expiryDate = DateOnly.FromDateTime(DateTime.Now);
    public int Cost { get; set; } = 0;
    public string Address { get; set; } = string.Empty;
    public string Remarks { get; set; } = string.Empty;
    public InspectionStatus InspectionStatus
    {
        get => DateOnly.FromDateTime(DateTime.Now) > ExpiryDate
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
    public long VehicleID { get; set; } = 0;
    public ServiceType ServiceType { get; set; } = ServiceType.None;
    public DateOnly DateIn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly EstimatedReturn { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly? DateCompleted { get; set; } = DateOnly.FromDateTime(DateTime.Today);
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

            // If completed date is set and reached, maintenance is done
            if (DateCompleted.HasValue && today >= DateCompleted.Value)
                return MaintenanceStatus.Done;

            // If scheduled date hasn't arrived yet
            if (today < DateIn)
                return MaintenanceStatus.Scheduled;

            // Otherwise, keep the active status (InProgress or AwaitingParts)
            return _maintenanceStatus == MaintenanceStatus.Done
                ? MaintenanceStatus.InProgress
                : _maintenanceStatus;
        }
        set => _maintenanceStatus = value;
    }
    private MaintenanceStatus _maintenanceStatus = MaintenanceStatus.InProgress;
    public string Remarks { get; set; } = string.Empty;
}