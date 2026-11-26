namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.Storage;

/// <summary>An attendance-eligible staff member (Worker, Office or Admin) with the dates needed to decide whether they were employed on a given day.</summary>
internal sealed record StaffInfo(
    long Id, string ShortName, string FullName,
    DateOnly? JoinDate, DateOnly? ResignDate, ContractStatus Status, string Department);

internal sealed record VehicleRenewalSpan(DateOnly? StartDate, DateOnly ExpiryDate, string Status);
internal sealed record VehicleMaintenanceSpan(DateOnly DateIn, DateOnly? DateCompleted, MaintenanceStatus Status);
internal sealed record VehicleInspectionSpan(DateOnly ScheduleDateIn, DateOnly DateOut, InspectionStatus Status);

/// <summary>Date-effective vehicle validity: <see cref="ShortNote"/> for dropdown labels, <see cref="Issues"/> for comments / warnings.</summary>
internal sealed record VehicleEvaluation(bool IsValid, string ShortNote, IReadOnlyList<string> Issues)
{
    internal string FormatComment(DateOnly date) =>
        IsValid ? string.Empty
                : $"⚠ Invalid vehicle on {date.ToString("dd/MM/yyyy", CultureInfo.InvariantCulture)}\n• " + string.Join("\n• ", Issues);
}

internal sealed record VehicleInfo(
    string Plate, VehicleType Type, decimal Odometer, VehicleStatus Status,
    DateOnly? BoughtDate, DateOnly? SoldDate, DateOnly? RoadTaxExpiry, DateOnly? InspectionExpiry,
    IReadOnlyList<VehicleRenewalSpan> RoadTaxes, IReadOnlyList<VehicleInspectionSpan> Inspections,
    IReadOnlyList<VehicleRenewalSpan> Insurances, IReadOnlyList<VehicleRenewalSpan> Permits,
    IReadOnlyList<VehicleMaintenanceSpan> Maintenances);
/// <summary>
/// A project as the board sees it. <see cref="AutoSchedule"/> is false for Internal projects and for General projects
/// with Auto-Schedule = FALSE: those are never auto-listed on the board but can be picked manually.
/// </summary>
internal sealed record ProjectInfo(
    string Id, string Name, IReadOnlyList<ProjectType> Types,
    ProjectStatus Status, ProjectCategory Category, bool AutoSchedule,
    DateOnly StartDate, DateOnly TargetEndDate, DateOnly? ActualEndDate);

/// <summary>
/// Read-only reference data the board needs while the user types: staff, vehicles and projects.
/// <para>
/// <see cref="Staff"/> are FIELD workers only (<see cref="Department.Worker"/>, see <see cref="FieldDepartments"/>); Office and
/// Admin staff never appear there, so the Daily Schedule, preferences and teams exclude them.
/// <see cref="AttendanceStaff"/> holds everyone who appears in attendance: Worker + Office + Admin (see <see cref="AttendanceDepartments"/>).
/// The roster keeps the FULL staff history (resigned and terminated included) so any record can be resolved
/// to a Staff ID / name. Whether a worker may be <i>suggested or newly assigned</i> is a separate, date-effective
/// question answered by <see cref="IsActiveOn(StaffInfo, DateOnly)"/>.
/// </para>
/// Server first (one query each), local snapshots when the server is unreachable; cached briefly so a
/// dropdown open never waits on the network.
/// </summary>
internal static class Roster
{
    /// <summary>
    /// Departments whose staff go on-site (Daily Schedule, Preference, Teams). Office, Admin and None are deliberately absent.
    /// Adjust here if a department stops (or starts) working in the field.
    /// </summary>
    internal static readonly Department[] FieldDepartments = { Department.Worker };

    /// <summary>
    /// Departments that appear in attendance (Daily Attendances / Attendance Entries): field workers, Office staff
    /// (attendance only, never scheduled) and Admin staff (auto-attended on weekdays).
    /// </summary>
    internal static readonly Department[] AttendanceDepartments = { Department.Worker, Department.Office, Department.Admin };

    private static readonly string[] AttendanceDepartmentNames = AttendanceDepartments.Select(d => d.ToString().ToLowerInvariant()).ToArray();
    private static readonly string[] FieldDepartmentNames = FieldDepartments.Select(d => d.ToString().ToLowerInvariant()).ToArray();

    internal static bool IsFieldDepartment(string? department) =>
        !string.IsNullOrWhiteSpace(department) && FieldDepartmentNames.Contains(department.Trim().ToLowerInvariant());

    internal static bool IsAttendanceDepartment(string? department) =>
        !string.IsNullOrWhiteSpace(department) && AttendanceDepartmentNames.Contains(department.Trim().ToLowerInvariant());

    internal static bool IsOfficeDepartment(string? department) =>
        !string.IsNullOrWhiteSpace(department) && string.Equals(department.Trim(), nameof(Department.Office), StringComparison.OrdinalIgnoreCase);

    internal static bool IsAdminDepartment(string? department) =>
        !string.IsNullOrWhiteSpace(department) && string.Equals(department.Trim(), nameof(Department.Admin), StringComparison.OrdinalIgnoreCase);

    /// <summary>Office staff are attendance-eligible but never scheduled; their attendance is booked to the "Office" project.</summary>
    internal static bool IsOfficeStaff(StaffInfo? staff) => staff != null && IsOfficeDepartment(staff.Department);
    internal static bool IsOfficeStaff(long staffId) => IsOfficeStaff(FindAttendanceStaff(staffId));

    /// <summary>Admin staff are attendance-eligible but never scheduled; they are auto checked in / out on weekdays.</summary>
    internal static bool IsAdminStaff(StaffInfo? staff) => staff != null && IsAdminDepartment(staff.Department);
    internal static bool IsAdminStaff(long staffId) => IsAdminStaff(FindAttendanceStaff(staffId));

    private static readonly TimeSpan CacheLifetime = TimeSpan.FromMinutes(5);
    private static DateTime _loadedAt = DateTime.MinValue;

    private static List<StaffInfo> _attendanceStaff = new();
    private static List<StaffInfo> _staff = new();
    private static List<StaffInfo> _adminStaff = new();
    private static Dictionary<string, StaffInfo> _attendanceByName = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<long, StaffInfo> _attendanceById = new();
    private static List<VehicleInfo> _vehicles = new();
    private static List<ProjectInfo> _projects = new();
    private static Dictionary<string, StaffInfo> _staffByName = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<long, StaffInfo> _staffById = new();
    private static Dictionary<string, VehicleInfo> _vehicleByPlate = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<string, ProjectInfo> _projectById = new(StringComparer.OrdinalIgnoreCase);

    /// <summary>Every field worker (Worker), past and present. Use for resolution; use <see cref="ActiveOn"/> for suggestions.</summary>
    public static IReadOnlyList<StaffInfo> Staff { get { EnsureLoaded(); return _staff; } }
    /// <summary>Every Worker, Office and Admin staff member, past and present (attendance resolution).</summary>
    public static IReadOnlyList<StaffInfo> AttendanceStaff { get { EnsureLoaded(); return _attendanceStaff; } }
    /// <summary>Every Admin staff member, past and present.</summary>
    public static IReadOnlyList<StaffInfo> AdminStaff { get { EnsureLoaded(); return _adminStaff; } }
    public static IReadOnlyList<VehicleInfo> Vehicles { get { EnsureLoaded(); return _vehicles; } }
    public static IReadOnlyList<ProjectInfo> Projects { get { EnsureLoaded(); return _projects; } }

    // =========================================================================
    // 1. DATE-EFFECTIVE EMPLOYMENT
    // =========================================================================

    /// <summary>
    /// True when the worker was employed on <paramref name="date"/>: joined on or before it and has not yet resigned.
    /// With a ResignDate the worker counts up to and including that day; without one, a Resigned / Terminated
    /// contract status means not employed.
    /// </summary>
    internal static bool IsActiveOn(StaffInfo s, DateOnly date)
    {
        if (s.JoinDate.HasValue && s.JoinDate.Value > date) return false;
        return s.ResignDate.HasValue
            ? date <= s.ResignDate.Value
            : s.Status is not (ContractStatus.Resigned or ContractStatus.Terminated);
    }

    /// <summary>Field workers only: false for Office / Admin staff and unknown IDs.</summary>
    internal static bool IsActiveOn(long staffId, DateOnly date) => FindStaff(staffId) is { } s && IsActiveOn(s, date);

    /// <summary>Field workers employed on <paramref name="date"/> - the only ones suggested or newly assignable that day.</summary>
    public static IReadOnlyList<StaffInfo> ActiveOn(DateOnly date) => Staff.Where(s => IsActiveOn(s, date)).ToList();

    /// <summary>Worker, Office and Admin staff employed on <paramref name="date"/> - the people who appear in attendance that day.</summary>
    public static IReadOnlyList<StaffInfo> ActiveAttendanceOn(DateOnly date) => AttendanceStaff.Where(s => IsActiveOn(s, date)).ToList();

    /// <summary>Admin staff employed on <paramref name="date"/>.</summary>
    public static IReadOnlyList<StaffInfo> ActiveAdminOn(DateOnly date) => AdminStaff.Where(s => IsActiveOn(s, date)).ToList();

    /// <summary>Why <paramref name="s"/> cannot work on <paramref name="date"/> (for validation messages).</summary>
    internal static string DescribeInactive(StaffInfo s, DateOnly date)
    {
        if (s.JoinDate.HasValue && s.JoinDate.Value > date) return $"joined on {s.JoinDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}";
        if (s.ResignDate.HasValue) return $"resigned on {s.ResignDate.Value.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture)}";
        return $"is {s.Status}";
    }

    // =========================================================================
    // 2. RESOLUTION (FindStaff: field workers; FindAttendanceStaff: Worker + Office + Admin; any date)
    // =========================================================================
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

    public static StaffInfo? FindAttendanceStaff(string? name)
    {
        EnsureLoaded();
        return !string.IsNullOrWhiteSpace(name) && _attendanceByName.TryGetValue(name.Trim(), out var s) ? s : null;
    }

    public static StaffInfo? FindAttendanceStaff(long id)
    {
        EnsureLoaded();
        return _attendanceById.TryGetValue(id, out var s) ? s : null;
    }

    /// <summary>Display name for a Staff ID; resolves former workers, Office and Admin staff too so past dates keep their names.</summary>
    public static string NameOf(long id, string fallback = "") =>
        FindAttendanceStaff(id)?.ShortName is { Length: > 0 } n ? n : fallback;

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

    /// <summary>True when <paramref name="projectId"/> exists in the roster and includes <see cref="ProjectType.Road_Reinstatement"/>.</summary>
    internal static bool SupportsPremixHoles(string? projectId) =>
        FindProject(projectId)?.Types.Contains(ProjectType.Road_Reinstatement) == true;

    private static bool IsActiveWindow(ProjectInfo p, DateOnly date) =>
        p.Status == ProjectStatus.Active && p.StartDate <= date && date <= (p.ActualEndDate ?? p.TargetEndDate);

    /// <summary>
    /// Auto-scheduled projects that are Active on <paramref name="date"/> (started, and not past the actual / target end
    /// date), ordered iJob -> MWO -> PVT -> Other, then by project ID. Internal projects and projects with
    /// Auto-Schedule = FALSE are excluded (see <see cref="ActiveManualProjectsOn"/>) but stay in <see cref="Projects"/>.
    /// </summary>
    public static IReadOnlyList<ProjectInfo> ActiveProjectsOn(DateOnly date) =>
        Projects
            .Where(p => p.Category != ProjectCategory.Internal && p.AutoSchedule && IsActiveWindow(p, date))
            .OrderBy(p => p.Category switch
            {
                ProjectCategory.iJob => 0, ProjectCategory.MWO => 1, ProjectCategory.PVT => 2, ProjectCategory.Other => 3, _ => 4
            })
            .ThenBy(p => p.Id, StringComparer.OrdinalIgnoreCase)
            .ToList();

    /// <summary>
    /// Active projects on <paramref name="date"/> that are only ever placed on the board by hand: Internal projects and
    /// General projects with Auto-Schedule = FALSE. General projects first, then Internal, then by project ID.
    /// </summary>
    public static IReadOnlyList<ProjectInfo> ActiveManualProjectsOn(DateOnly date) =>
        Projects
            .Where(p => (p.Category == ProjectCategory.Internal || !p.AutoSchedule) && IsActiveWindow(p, date))
            .OrderBy(p => p.Category == ProjectCategory.Internal ? 1 : 0)
            .ThenBy(p => p.Id, StringComparer.OrdinalIgnoreCase)
            .ToList();

    public static void Invalidate() => _loadedAt = DateTime.MinValue;

    // =========================================================================
    // 3. LOADING
    // =========================================================================
    private static void EnsureLoaded()
    {
        if (DateTime.UtcNow - _loadedAt < CacheLifetime) return;

        try
        {
            if (ServerEnvironment.IsOffline) throw new TimeoutException("Offline Mode: roster comes from the local snapshot.");
            var db = new PostgreSqlService();

            _attendanceStaff = new List<StaffInfo>();
            foreach (var r in db.Query<dynamic>(
                @"SELECT staff_id, staff_name, short_name, contract_status, join_date, resign_date, department FROM staff_master
                  WHERE LOWER(department) = ANY(@Departments)
                  ORDER BY staff_id;",
                new { Departments = AttendanceDepartmentNames }))
            {
                long id = Convert.ToInt64(r.staff_id);
                string? full = r.staff_name;
                string? shortName = r.short_name;
                string? status = r.contract_status;
                object? join = r.join_date;
                object? resign = r.resign_date;
                string? department = r.department;
                _attendanceStaff.Add(MakeStaff(id, full, shortName, status, ToDate(join), ToDate(resign), department));
            }

            _vehicles = new List<VehicleInfo>();
            var roadTaxes = LoadRenewals(db, "SELECT vehicle_plate, renew_date, expiry_date, status FROM vehicle_road_taxes;", "renew_date", "vehicle_road_taxes");
            var inspections = new Dictionary<string, List<VehicleInspectionSpan>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var r in db.Query<dynamic>("SELECT vehicle_plate, schedule_date_in, date_out, status FROM vehicle_inspections;"))
                {
                    string key = ((string)r.vehicle_plate).Trim().ToUpperInvariant();
                    object? sin = r.schedule_date_in;
                    object? dout = r.date_out;
                    string? st = r.status;
                    var dateIn = NullIfUnset(ToDate(sin));
                    var dateOut = NullIfUnset(ToDate(dout));
                    if (dateIn == null || dateOut == null) continue;
                    if (!inspections.TryGetValue(key, out var list)) inspections[key] = list = new();
                    list.Add(new VehicleInspectionSpan(dateIn.Value, dateOut.Value,
                        Enum.TryParse<InspectionStatus>(st, true, out var ist) ? ist : InspectionStatus.Scheduled));
                }
            }
            catch (Exception ex) { AppLogger.Warn($"[ScheduleManager.Roster] vehicle_inspections load skipped: {ex.Message}"); }
            var insurances = LoadRenewals(db, "SELECT vehicle_plate, renew_date, expiry_date, status FROM vehicle_insurances;", "renew_date", "vehicle_insurances");
            var permits = LoadRenewals(db, "SELECT vehicle_plate, renew_date, expiry_date, status FROM vehicle_permits;", "renew_date", "vehicle_permits");
            var maintenances = new Dictionary<string, List<VehicleMaintenanceSpan>>(StringComparer.OrdinalIgnoreCase);
            try
            {
                foreach (var r in db.Query<dynamic>("SELECT vehicle_plate, date_in, date_completed, status FROM vehicle_maintenance;"))
                {
                    string key = ((string)r.vehicle_plate).Trim().ToUpperInvariant();
                    object? din = r.date_in;
                    object? dcomp = r.date_completed;
                    string? st = r.status;
                    var dateIn = NullIfUnset(ToDate(din));
                    if (dateIn == null) continue;
                    if (!maintenances.TryGetValue(key, out var list)) maintenances[key] = list = new();
                    list.Add(new VehicleMaintenanceSpan(dateIn.Value, NullIfUnset(ToDate(dcomp)),
                        Enum.TryParse<MaintenanceStatus>(st, true, out var ms) ? ms : MaintenanceStatus.None));
                }
            }
            catch (Exception ex) { AppLogger.Warn($"[ScheduleManager.Roster] vehicle_maintenance load skipped: {ex.Message}"); }

            foreach (var r in db.Query<dynamic>("SELECT vehicle_plate, vehicle_type, current_odometer, status, bought_date, sold_date, road_tax_expiry, inspection_expiry FROM vehicle_master ORDER BY vehicle_plate;"))
            {
                string plate = ((string)r.vehicle_plate).Trim().ToUpperInvariant();
                string? typeText = r.vehicle_type;
                object? odo = r.current_odometer;
                string? statusText = r.status;
                object? bought = r.bought_date;
                object? sold = r.sold_date;
                object? rtExp = r.road_tax_expiry;
                object? inExp = r.inspection_expiry;
                _vehicles.Add(new VehicleInfo(
                    plate,
                    Enum.TryParse<VehicleType>(typeText, true, out var vt) ? vt : VehicleType.None,
                    odo != null ? Convert.ToDecimal(odo) : 0m,
                    Enum.TryParse<VehicleStatus>(statusText, true, out var vs) ? vs : VehicleStatus.None,
                    NullIfUnset(ToDate(bought)), NullIfUnset(ToDate(sold)), NullIfUnset(ToDate(rtExp)), NullIfUnset(ToDate(inExp)),
                    roadTaxes.GetValueOrDefault(plate) ?? new(), inspections.GetValueOrDefault(plate) ?? new(),
                    insurances.GetValueOrDefault(plate) ?? new(), permits.GetValueOrDefault(plate) ?? new(),
                    maintenances.GetValueOrDefault(plate) ?? new()));
            }

            // project_name / auto_schedule come from the General / Internal Projects migration.
            ServerDispatcher.EnsureProjectSchema();
            _projects = new List<ProjectInfo>();
            foreach (var r in db.Query<dynamic>("SELECT project_id, project_name, project_type, status, category, auto_schedule, start_date, target_end_date, actual_end_date FROM project_master ORDER BY project_id;"))
            {
                string id = ((string)r.project_id).Trim().ToUpperInvariant();
                string? name = r.project_name;
                string? types = r.project_type;
                string? statusText = r.status;
                string? categoryText = r.category;
                object? autoRaw = r.auto_schedule;
                object? start = r.start_date;
                object? target = r.target_end_date;
                object? actual = r.actual_end_date;
                var pc = Enum.TryParse<ProjectCategory>(categoryText, true, out var parsedCategory) ? parsedCategory : ProjectCategory.None;
                bool auto = autoRaw is bool b ? b : pc != ProjectCategory.Internal;
                _projects.Add(new ProjectInfo(
                    id, name?.Trim() ?? string.Empty, ParseProjectTypes(types),
                    Enum.TryParse<ProjectStatus>(statusText, true, out var ps) ? ps : ProjectStatus.None,
                    pc, pc != ProjectCategory.Internal && auto,
                    ToDate(start) ?? DateOnly.MinValue, ToDate(target) ?? DateOnly.MaxValue, ToDate(actual)));
            }
        }
        catch (Exception ex)
        {
            if (ServerEnvironment.IsOffline)
            {
                AppLogger.Warn($"[ScheduleManager.Roster] OFFLINE MODE, using local snapshots: {ex.Message}");
                LoadRosterFromSnapshots();
            }
            else
            {
                // Online + SQL error: never mix in offline data.
                AppLogger.Error("[ScheduleManager.Roster] Server query failed while online; roster left empty", ex);
                _attendanceStaff = new List<StaffInfo>();
                _vehicles = new List<VehicleInfo>();
                _projects = new List<ProjectInfo>();
            }
        }

        _staff = _attendanceStaff.Where(s => IsFieldDepartment(s.Department)).ToList();
        _adminStaff = _attendanceStaff.Where(s => IsAdminDepartment(s.Department)).ToList();
        _attendanceByName = new Dictionary<string, StaffInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in _attendanceStaff.Where(s => s.ShortName.Length > 0)) _attendanceByName.TryAdd(s.ShortName, s);
        _attendanceById = _attendanceStaff.GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.First());
        _staffByName = new Dictionary<string, StaffInfo>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in _staff.Where(s => s.ShortName.Length > 0)) _staffByName.TryAdd(s.ShortName, s);
        _staffById = _staff.GroupBy(s => s.Id).ToDictionary(g => g.Key, g => g.First());
        _vehicleByPlate = _vehicles.GroupBy(v => v.Plate, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        _projectById = _projects.GroupBy(p => p.Id, StringComparer.OrdinalIgnoreCase).ToDictionary(g => g.Key, g => g.First(), StringComparer.OrdinalIgnoreCase);
        _loadedAt = DateTime.UtcNow;
    }

    private static void LoadRosterFromSnapshots()
    {
        _attendanceStaff = (OfflineSnapshot.Load<List<StaffRecord>>() ?? new())
            .Where(s => s.StaffID > 0 && IsAttendanceDepartment(s.StaffDetails.Department))
            .Select(s => MakeStaff(
                s.StaffID, s.StaffDetails.StaffName, s.StaffDetails.StaffShortName, s.StaffDetails.ContractStatus.ToString(),
                s.StaffDetails.JoinDate, s.StaffDetails.ResignDate, s.StaffDetails.Department))
            .ToList();
        _vehicles = (OfflineSnapshot.Load<List<VehicleRecord>>() ?? new())
            .Where(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
            .Select(v => new VehicleInfo(
                v.VehiclePlate.Trim().ToUpperInvariant(), v.VehicleDetail.Type, v.VehicleDetail.CurrentOdometer ?? 0m, v.VehicleStatus,
                NullIfUnset(v.VehicleDetail.BoughtDate), NullIfUnset(v.VehicleDetail.SoldDate),
                NullIfUnset(v.VehicleDetail.RoadTaxExpiry), NullIfUnset(v.VehicleDetail.InspectionExpiry),
                v.RoadTaxes.Select(x => new VehicleRenewalSpan(NullIfUnset(x.RenewDate), x.ExpiryDate, x.Status.ToString())).ToList(),
                v.Inspections.Where(x => NullIfUnset(x.ScheduleDateIn) != null && NullIfUnset(x.DateOut) != null)
                    .Select(x => new VehicleInspectionSpan(x.ScheduleDateIn, x.DateOut, x.Status)).ToList(),
                v.Insurances.Select(x => new VehicleRenewalSpan(NullIfUnset(x.RenewDate), x.ExpiryDate, x.Status.ToString())).ToList(),
                v.Permits.Select(x => new VehicleRenewalSpan(NullIfUnset(x.RenewDate), x.ExpiryDate, x.Status.ToString())).ToList(),
                v.Maintenances.Select(x => new VehicleMaintenanceSpan(x.DateIn, NullIfUnset(x.DateCompleted), x.MaintenanceStatus)).ToList()))
            .ToList();
        _projects = (OfflineSnapshot.Load<List<ProjectRecord>>() ?? new())
            .Where(p => !string.IsNullOrWhiteSpace(p.ProjectCode))
            .Select(p => new ProjectInfo(
                p.ProjectCode.Trim().ToUpperInvariant(),
                (p.ProjectDetails.ProjectName ?? string.Empty).Trim(),
                p.ProjectDetails.ProjectType.Where(t => t != ProjectType.None).ToList(),
                p.ProjectDetails.Status, p.Category,
                p.Category != ProjectCategory.Internal && p.ProjectDetails.AutoSchedule,
                p.ProjectDetails.StartDate, p.ProjectDetails.TargetEndDate, p.ProjectDetails.ActualEndDate))
            .ToList();
    }

    private static DateOnly? NullIfUnset(DateOnly? d) => d is { Year: <= 1900 } ? null : d;

    private static Dictionary<string, List<VehicleRenewalSpan>> LoadRenewals(PostgreSqlService db, string sql, string startColumn, string table)
    {
        var map = new Dictionary<string, List<VehicleRenewalSpan>>(StringComparer.OrdinalIgnoreCase);
        try
        {
            foreach (IDictionary<string, object?> r in db.Query<dynamic>(sql))
            {
                string key = Convert.ToString(r["vehicle_plate"])?.Trim().ToUpperInvariant() ?? string.Empty;
                var expiry = NullIfUnset(ToDate(r["expiry_date"]));
                if (key.Length == 0 || expiry == null) continue;
                var start = NullIfUnset(ToDate(r[startColumn]));
                if (!map.TryGetValue(key, out var list)) map[key] = list = new();
                list.Add(new VehicleRenewalSpan(start, expiry.Value, Convert.ToString(r["status"]) ?? string.Empty));
            }
        }
        catch (Exception ex) { AppLogger.Warn($"[ScheduleManager.Roster] {table} load skipped: {ex.Message}"); }
        return map;
    }

    /// <summary>Is the vehicle usable on <paramref name="date"/>? Checks status, ownership, maintenance and document validity.</summary>
    internal static VehicleEvaluation EvaluateVehicleOn(VehicleInfo v, DateOnly date)
    {
        var issues = new List<(string Short, string Detail)>();
        const string F = "dd/MM/yyyy";
        var ic = CultureInfo.InvariantCulture;
        var today = DateOnly.FromDateTime(DateTime.Today);

        if (v.BoughtDate.HasValue && v.BoughtDate.Value > date)
            issues.Add(("not yet owned", $"Not yet owned (bought {v.BoughtDate.Value.ToString(F, ic)})"));

        if (v.SoldDate.HasValue && date > v.SoldDate.Value)
            issues.Add(("decommissioned", $"Decommissioned (sold on {v.SoldDate.Value.ToString(F, ic)})"));
        else if (!v.SoldDate.HasValue && v.Status == VehicleStatus.Decommissioned)
            issues.Add(("decommissioned", "Vehicle is Decommissioned"));

        if (v.Status == VehicleStatus.Suspended) issues.Add(("suspended", "Vehicle is Suspended"));
        if (v.Status == VehicleStatus.Reserved) issues.Add(("reserved", "Vehicle is Reserved"));

        var shop = v.Maintenances.FirstOrDefault(m => m.Status != MaintenanceStatus.Cancelled &&
            ((m.DateCompleted.HasValue && m.DateIn <= date && date <= m.DateCompleted.Value) ||
             (!m.DateCompleted.HasValue && m.Status == MaintenanceStatus.In_Progress && m.DateIn <= date)));
        if (shop != null)
            issues.Add(("in shop", shop.DateCompleted.HasValue
                ? $"In shop ({shop.DateIn.ToString(F, ic)} – {shop.DateCompleted.Value.ToString(F, ic)})"
                : $"In shop since {shop.DateIn.ToString(F, ic)}"));
        else if (v.Status == VehicleStatus.In_Shop && date >= today)
            issues.Add(("in shop", "Vehicle is In Shop"));

        int docIssues = 0;
        void CheckDocument(string docName, IReadOnlyList<VehicleRenewalSpan> spans, DateOnly? masterFallbackExpiry)
        {
            var valid = spans.Where(s => !s.Status.Equals("Cancelled", StringComparison.OrdinalIgnoreCase)
                                      && !s.Status.Equals("Failed", StringComparison.OrdinalIgnoreCase)).ToList();
            if (valid.Count > 0)
            {
                if (valid.Any(s => (s.StartDate == null || s.StartDate.Value <= date) && date <= s.ExpiryDate)) return;
                var past = valid.Where(s => s.ExpiryDate < date).ToList();
                if (past.Count > 0)
                {
                    var latest = past.MaxBy(s => s.ExpiryDate)!;
                    issues.Add(($"{docName} expired", $"{docName} expired on {latest.ExpiryDate.ToString(F, ic)}"));
                    docIssues++;
                }
                else
                {
                    var future = valid.Where(s => s.StartDate.HasValue && s.StartDate.Value > date).ToList();
                    if (future.Count > 0)
                    {
                        var earliest = future.MinBy(s => s.StartDate!.Value)!;
                        issues.Add(($"{docName} not yet valid", $"{docName} not valid until {earliest.StartDate!.Value.ToString(F, ic)}"));
                        docIssues++;
                    }
                }
            }
            else if (masterFallbackExpiry.HasValue && masterFallbackExpiry.Value < date)
            {
                issues.Add(($"{docName} expired", $"{docName} expired on {masterFallbackExpiry.Value.ToString(F, ic)}"));
                docIssues++;
            }
        }
        CheckDocument("Road Tax", v.RoadTaxes, v.RoadTaxExpiry);
        CheckDocument("Insurance", v.Insurances, null);
        CheckDocument("Permit", v.Permits, null);

        // Expired inspection is OK; only the inspection window itself makes the vehicle unavailable.
        var insp = v.Inspections.FirstOrDefault(i => i.ScheduleDateIn <= date && date <= i.DateOut);
        if (insp != null)
            issues.Add(("in inspection", insp.ScheduleDateIn == insp.DateOut
                ? $"In inspection ({insp.ScheduleDateIn.ToString(F, ic)})"
                : $"In inspection ({insp.ScheduleDateIn.ToString(F, ic)} – {insp.DateOut.ToString(F, ic)})"));

        if (docIssues == 0 && v.Status == VehicleStatus.Expired && date >= today && !(v.InspectionExpiry.HasValue && v.InspectionExpiry.Value < today))
            issues.Add(("expired", "Vehicle status is Expired"));

        string shortNote = issues.Count == 0 ? string.Empty
            : issues.Count == 1 ? issues[0].Short
            : $"{issues[0].Short}, +{issues.Count - 1} more";
        return new VehicleEvaluation(issues.Count == 0, shortNote, issues.Select(i => i.Detail).ToList());
    }

    private static DateOnly? ToDate(object? raw) =>
        raw == null ? null : DateOnly.FromDateTime(Convert.ToDateTime(raw));

    private static StaffInfo MakeStaff(long id, string? fullName, string? shortName, string? contractStatus, DateOnly? join, DateOnly? resign, string? department)
    {
        string full = fullName?.Trim() ?? string.Empty;
        string shortN = !string.IsNullOrWhiteSpace(shortName) ? shortName.Trim() : full;
        var status = Enum.TryParse<ContractStatus>(contractStatus, true, out var parsed) ? parsed : ContractStatus.None;
        // Year <= 1900 is the "unset" sentinel used elsewhere in the app.
        if (join is { Year: <= 1900 }) join = null;
        if (resign is { Year: <= 1900 }) resign = null;
        return new StaffInfo(id, shortN, full, join, resign, status, department?.Trim() ?? string.Empty);
    }

    internal static List<ProjectType> ParseProjectTypes(string? csv) =>
        (csv ?? string.Empty)
            .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
            .Select(t => Enum.TryParse<ProjectType>(t, true, out var pt) ? pt : ProjectType.None)
            .Where(t => t != ProjectType.None)
            .Distinct()
            .ToList();
}
