using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.ComponentModel.Design;
using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Numerics;
using System.Security.Permissions;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace SKSB_App.Core.Types;

public record Staff
{
    public long StaffID { get; set; } = 0;
    public Master Master { get; set; } = new Master();
    public Preference Preference { get; set; } = new Preference();
    public List<License> Licenses { get; set; } = new List<License>();
    public List<Leave> Leaves { get; set; } = new List<Leave>();
}


public enum Department
{
    Admin,
    Worker,
    None
}

public enum PayType
{
    Hourly,
    Monthly,
    None
}

public enum ActiveStatus
{
    OnContract,
    Suspended,
    OffContract
}

public record Master
{
    public string FamilyName { get; set; } = string.Empty;
    public string MiddleName { get; set; } = string.Empty;
    public string GivenName { get; set; } = string.Empty;

    /// <summary>
    /// Evaluates to "FamilyName GivenName". Can be set explicitly or generated automatically.
    /// </summary>
    public string Name
    {
        get => string.IsNullOrWhiteSpace(_name)
            ? $"{FamilyName} {GivenName}".Trim()
            : _name;
        set => _name = value;
    }
    private string _name = string.Empty;

    public string NRIC { get; set; } = string.Empty;
    public Department Department { get; set; } = Department.None;
    public PayType PayType { get; set; } = PayType.None;
    public decimal BaseRate { get; set; } = 0;
    public decimal OvertimeRate { get; set; } = 0;
    public ActiveStatus ActiveStatus { get; set; } = ActiveStatus.OffContract;
}

public enum Shift
{
    Morning,
    Afternoon,
    Night,
    None
}

public record Preference
{
    public string PrimaryTeam {  get; set; } = string.Empty;
    public string SecondaryTeam { get; set; } = string.Empty;
    public string PrimarySkill { get; set; } = string.Empty;
    public string SecondarySkill { get; set; } = string.Empty;
    public string BlackListTask { get; set; } = string.Empty;
    public Shift PreferredShift { get; set; } = Shift.None;
    public int MaxDaysPerWeek { get; set; } = 0;
    public DayOfWeek PrimaryRestDay = DayOfWeek.Saturday;
    public DayOfWeek SecondaryRestDay = DayOfWeek.Sunday;
    public VehicleType PreferredVehicleType { get; set; } = VehicleType.None;
    public long? DedicatedVehicleID { get; set; }
}

public enum LicenseClass
{
    B2,         // Motorcycle <= 250cc
    B,          // Motorcycle > 500cc (Full B)
    D,          // Car / Light Vehicle (BTM <= 3500kg)
    DA,         // Automatic Car
    E,          // Heavy Motor Car (Rigid BTM > 7500kg)
    GDL_D,      // Goods Vehicle Driver's Licence (Van / 4WD / Light Pickup)
    GDL_E,      // Goods Vehicle Driver's Licence (Rigid Lorry > 7.5T)
    GDL_E_Full, // Articulated / Trailer / Prime Mover
    Forklift,   // Internal / Industrial Certificate
    None
}
public enum LicenseStatus
{
    Valid,
    Invalid,
    Suspended
}

public record License
{
    public long StaffID { get; set; }
    public LicenseClass LicenseType { get; set; }
    public DateOnly ExpiryDate { get; set; }
    public LicenseStatus LicenseStatus { get; set; }
}

public enum LeaveType
{
    AnnualLeave,
    MedicalLeave,
    EmergencyLeave,
    Unpaid,
    None
}
public enum Session
{
    FullDay,
    Morning,
    Afternoon,
    Night
}
public enum ApprovalStatus
{
    Approve,
    Pending,
    Denied
}

public record Leave
{
    public long LeaveID { get; set; } = 0;
    public long StaffID { get; set; } = 0; // Explicit Foreign Key
    public LeaveType LeaveType { get; set; } = LeaveType.None;
    public DateOnly DateStart { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public DateOnly DateEnd { get; set; } = DateOnly.FromDateTime(DateTime.Today);
    public Session SessionType { get; set; } = Session.FullDay;

    private decimal _totalDays;
    public decimal TotalDays
    {
        get
        {
            if (_totalDays > 0) return _totalDays;

            if (SessionType == Session.FullDay)
            {
                int days = DateEnd.DayNumber - DateStart.DayNumber + 1;
                return days > 0 ? (decimal)days : 0m;
            }

            return 0.5m; // Part-day shifts
        }
        set => _totalDays = value;
    }

    public ApprovalStatus ApprovalStatus { get; set; } = ApprovalStatus.Pending;

}