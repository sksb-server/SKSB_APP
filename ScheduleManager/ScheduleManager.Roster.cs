namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.Storage;

internal sealed record StaffInfo(long Id, string ShortName, string FullName, bool IsActive);
internal sealed record VehicleInfo(string Plate, VehicleType Type, decimal Odometer);
internal sealed record ProjectInfo(string Id, IReadOnlyList<ProjectType> Types);

/// <summary>
/// Read-only reference data the board needs while the user types: staff, vehicles and projects.
/// Server first (one query each), local snapshots when the server is unreachable; cached briefly so a
/// dropdown open never waits on the network.
/// </summary>
internal static class Roster
{
    private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(5);
    private static DateTime _loadedAt = DateTime.MinValue;

    private static List<StaffInfo> _staff = new();
    private static List<VehicleInfo> _vehicles = new();
    private static List<ProjectInfo> _projects = new();
    private static Dictionary<string, StaffInfo> _staffByName = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<long, StaffInfo> _staffById = new();
    private static Dictionary<string, VehicleInfo> _vehicleByPlate = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<string, ProjectInfo> _projectById = new(StringComparer.OrdinalIgnoreCase);

    public static IReadOnlyList<StaffInfo> Staff { get { EnsureLoaded(); return _staff; } }
    public static IReadOnlyList<VehicleInfo> Vehicles { get { EnsureLoaded(); return _vehicles; } }
    public static IReadOnlyList<ProjectInfo> Projects { get { EnsureLoaded(); return _projects; } }

    public static StaffInfo? FindStaff(string? name)
    {
        EnsureLoaded();
        return !string.IsNullOrWhiteSpace(name) && _staffByName.TryGetValue(name.Trim(), out var s) ? s : null;
    }

    public static StaffInfo? FindStaff(long id)
    {
        EnsureLoaded();
        return _staffById.TryGetValue(id, out var s) ? s : null;
    }

    public static VehicleInfo? FindVehicle(string? plate)
    {
        EnsureLoaded();
        return !string.IsNullOrWhiteSpace(plate) && _vehicleByPlate.TryGetValue(plate.Trim(), out var v) ? v : null;
    }

    public static ProjectInfo? FindProject(string? id)
    {
        EnsureLoaded();
        return !string.IsNullOrWhiteSpace(id) && _projectById.TryGetValue(id.Trim(), out var p) ? p : null;
    }

    public static void Invalidate() => _loadedAt = DateTime.MinValue;

    // =========================================================================
    // LOADING
    // =========================================================================
    private static void EnsureLoaded()
    {
        if (DateTime.UtcNow - _loadedAt < CacheLifetime) return;

        try
        {
            if (ServerEnvironment.SimulateOffline) throw new TimeoutException("Simulated offline.");
            var db = new PostgreSqlService();

            _staff = new List<StaffInfo>();
            foreach (var r in db.Query<dynamic>("SELECT staff_id, staff_name, short_name, contract_status FROM staff_master ORDER BY staff_id;"))
            {
                long id = Convert.ToInt64(r.staff_id);
                string? full = r.staff_name;
                string? shortName = r.short_name;
                string? status = r.contract_status;
                _staff.Add(MakeStaff(id, full, shortName, status));
            }

            _vehicles = new List<VehicleInfo>();
            foreach (var r in db.Query<dynamic>("SELECT vehicle_plate, vehicle_type, current_odometer FROM vehicle_master ORDER BY vehicle_plate;"))
            {
                string plate = ((string)r.vehicle_plate).Trim().ToUpperInvariant();
                string? typeText = r.vehicle_type;
                object? odo = r.current_odometer;
                _vehicles.Add(new VehicleInfo(
                    plate,
                    Enum.TryParse<VehicleType>(typeText, true, out var vt) ? vt : VehicleType.None,
                    odo != null ? Convert.ToDecimal(odo) : 0m));
            }

            _projects = new List<ProjectInfo>();
            foreach (var r in db.Query<dynamic>("SELECT project_id, project_type FROM project_master ORDER BY project_id;"))
            {
                string id = ((string)r.project_id).Trim().ToUpperInvariant();
                string? types = r.project_type;
                _projects.Add(new ProjectInfo(id, ParseProjectTypes(types)));
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager.Roster] Server unreachable, using local snapshots: {ex.Message}");
            _staff = (OfflineSnapshot.Load<List<StaffRecord>>() ?? new())
                .Where(s => s.StaffID > 0)
                .Select(s => MakeStaff(s.StaffID, s.StaffDetails.StaffName, s.StaffDetails.StaffShortName, s.StaffDetails.ContractStatus.ToString()))
                .ToList();
            _vehicles = (OfflineSnapshot.Load<List<VehicleRecord>>() ?? new())
                .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
                .Select(v => new VehicleInfo(v.VehiclePlate.Trim().ToUpperInvariant(), v.VehicleDetail.Type, v.VehicleDetail.CurrentOdometer ?? 0m))
                .ToList();
            _projects = (OfflineSnapshot.Load<List<ProjectRecord>>() ?? new())
                .Where(p => !string.IsNullOrWhiteSpace(p.ProjectCode))
                .Select(p => new ProjectInfo(p.ProjectCode.Trim().ToUpperInvariant(), p.ProjectDetails.ProjectType.Where(t => t != ProjectType.None).ToList()))
                .ToList();
        }

        _staffByName = new Dictionary<string, StaffInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in _staff.Where(s => s.ShortName.Length > 0)) _staffByName.TryAdd(s.ShortName, s);
        _staffById = _staff.GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.First());
        _vehicleByPlate = _vehicles.GroupBy(v => v.Plate, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        _projectById = _projects.GroupBy(p => p.Id, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        _loadedAt = DateTime.UtcNow;
    }

    private static StaffInfo MakeStaff(long id, string? fullName, string? shortName, string? contractStatus)
    {
        string full = fullName?.Trim() ?? string.Empty;
        string shortN = !string.IsNullOrWhiteSpace(shortName) ? shortName.Trim() : full;
        bool active = !string.Equals(contractStatus, nameof(ContractStatus.Resigned), StringComparison.OrdinalIgnoreCase);
        return new StaffInfo(id, shortN, full, active);
    }

    internal static List<ProjectType> ParseProjectTypes(string? csv) =>
        (csv ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => Enum.TryParse<ProjectType>(t, true, out var pt) ? pt : ProjectType.None)
            .Where(t => t != ProjectType.None)
            .Distinct()
            .ToList();
}
