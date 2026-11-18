namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.SheetMapper;

/// <summary>
/// Workbook-resident reference data the board needs on every edit (preferences, teams, project-type requirements,
/// icon legend, late cut-off) plus the weekly-load figures from the database. Reloaded on the Excel macro thread
/// only (<see cref="Reload"/>); dropdowns and observers read the cached copy.
/// </summary>
internal static class BoardContext
{
    // =========================================================================
    // 1. BUILT-IN DEFAULTS (used until / unless the editable Lookups tables say otherwise)
    // =========================================================================
    private static readonly Dictionary<string, string> DefaultIcons = new(StringComparer.OrdinalIgnoreCase)
    {
        ["Driver"] = "🚗", ["Excavator Operator"] = "🚜", ["Pipe Fitter"] = "🔧", ["Foreman"] = "👷", ["General Worker"] = "🧰",
        ["Car"] = "🚗", ["Hilux"] = "🚙", ["Van"] = "🚐", ["Lorry"] = "🚚", ["Tipper_Lorry"] = "🚛", ["Roro_Lorry"] = "🚛",
        ["Crane_Lorry"] = "🏗", ["Skylift"] = "🏗", ["Backhoe"] = "🚜", ["Excavator"] = "🚜", ["Roller"] = "🚧",
        ["Compressor"] = "🔨", ["Generator"] = "⚡", ["Motorcycle"] = "🏍", ["Machinery"] = "⚙",
        ["Pipe_Repair"] = "🔧", ["Pipe_Laying"] = "🚰", ["Comm_Pipe_Meter"] = "💧", ["Valve_Hydrant"] = "🚒", ["Tapping_Connection"] = "🔗",
        ["Road_Reinstatement"] = "🛣", ["Civil_Earthworks"] = "⛏", ["Machinery_Rental"] = "⚙", ["General_Maintenance"] = "🧰", ["Other"] = "📋"
    };

    private static readonly Dictionary<ProjectType, (string[] Skills, VehicleType[] Vehicles)> DefaultRequirements = new()
    {
        [ProjectType.Pipe_Repair] = (new[] { "Pipe Fitter", "General Worker" }, new[] { VehicleType.Hilux, VehicleType.Backhoe }),
        [ProjectType.Pipe_Laying] = (new[] { "Pipe Fitter", "Excavator Operator", "General Worker" }, new[] { VehicleType.Excavator, VehicleType.Tipper_Lorry }),
        [ProjectType.Comm_Pipe_Meter] = (new[] { "Pipe Fitter" }, new[] { VehicleType.Hilux }),
        [ProjectType.Valve_Hydrant] = (new[] { "Pipe Fitter" }, new[] { VehicleType.Hilux }),
        [ProjectType.Tapping_Connection] = (new[] { "Pipe Fitter" }, new[] { VehicleType.Hilux }),
        [ProjectType.Road_Reinstatement] = (new[] { "General Worker", "Driver" }, new[] { VehicleType.Roller, VehicleType.Tipper_Lorry }),
        [ProjectType.Civil_Earthworks] = (new[] { "Excavator Operator", "Driver" }, new[] { VehicleType.Excavator, VehicleType.Tipper_Lorry }),
        [ProjectType.Machinery_Rental] = (new[] { "Driver" }, new[] { VehicleType.Machinery }),
        [ProjectType.General_Maintenance] = (new[] { "General Worker" }, new[] { VehicleType.Hilux })
    };

    // =========================================================================
    // 2. CACHED STATE
    // =========================================================================
    private static List<WorkerPreference> _prefs = new();
    private static List<TeamRecord> _teams = new();
    private static Dictionary<ProjectType, ProjectTypeRequirement> _requirements = new();
    private static Dictionary<string, string> _icons = new(DefaultIcons, StringComparer.OrdinalIgnoreCase);
    public static readonly TimeOnly DefaultLateCutOff = new(11, 0);
    public static readonly TimeOnly DefaultOtStartTime = new(18, 0);

    private static TimeOnly _lateCutOff = DefaultLateCutOff;
    private static TimeOnly _otStartTime = DefaultOtStartTime;

    private static DateOnly _weekKey;
    private static DateTime _weekLoadedAt = DateTime.MinValue;
    private static Dictionary<long, int> _daysWorkedBefore = new();

    public static IReadOnlyList<WorkerPreference> Preferences => _prefs;
    public static IReadOnlyList<TeamRecord> Teams => _teams;
    public static TimeOnly LateCutOff => _lateCutOff;
    public static TimeOnly OTStartTime => _otStartTime;

    public static WorkerPreference? PreferenceFor(long staffId) => _prefs.FirstOrDefault(p => p.StaffID == staffId);

    public static void SetPreferences(IEnumerable<WorkerPreference> prefs, IEnumerable<TeamRecord> teams)
    {
        _prefs = prefs.ToList();
        _teams = teams.ToList();
    }

    // =========================================================================
    // 3. WORKBOOK RELOAD (Excel macro thread only)
    // =========================================================================
    public static void Reload(object wbObj)
    {
        try
        {
            dynamic? prefSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetMapper.PreferenceAliases, "Preference");
            if (prefSheet != null)
            {
                _prefs = TableMapper.ExtractFromTable<WorkerPreference>((object)prefSheet, "WorkerPreferencesTable");
                _teams = TableMapper.ExtractFromTable<TeamRecord>((object)prefSheet, "TeamTable");
            }

            dynamic? lookups = SheetMapperHelper.ResolveSheetSafely(wbObj, new[] { "Lookups" }, "Lookups");
            if (lookups != null)
            {
                var reqs = TableMapper.ExtractFromTable<ProjectTypeRequirement>((object)lookups, "ProjectTypeRequirementsTable");
                _requirements = reqs.Where(r => r.ProjectType != ProjectType.None).GroupBy(r => r.ProjectType).ToDictionary(g => g.Key, g => g.First());

                var icons = new Dictionary<string, string>(DefaultIcons, StringComparer.OrdinalIgnoreCase);
                foreach (var i in TableMapper.ExtractFromTable<SkillIcon>((object)lookups, "SkillIconsTable"))
                    if (!string.IsNullOrWhiteSpace(i.Key) && !string.IsNullOrWhiteSpace(i.Icon)) icons[i.Key.Trim()] = i.Icon.Trim();
                _icons = icons;
            }

            dynamic? attSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetMapper.MasterSheetAliases, "Daily Attendances");
            if (attSheet != null)
            {
                _lateCutOff = ReadTimeCell(attSheet, SheetMapper.LateCutOffCell, DefaultLateCutOff);
                _otStartTime = ReadTimeCell(attSheet, SheetMapper.OtStartCell, DefaultOtStartTime);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ScheduleManager.BoardContext] Reload warning: {ex.Message}");
        }
    }

    /// <summary>Reads a time-of-day cell (Excel serial fraction or text); blank / unreadable => <paramref name="fallback"/>.</summary>
    public static TimeOnly ReadTimeCell(dynamic attendanceSheet, string cellAddress, TimeOnly fallback)
    {
        try
        {
            object? raw = attendanceSheet.Range[cellAddress].Value2;
            if (raw is double d && d >= 0)
                return TimeOnly.FromTimeSpan(TimeSpan.FromSeconds(Math.Round((d - Math.Floor(d)) * 86400.0) % 86400));
            if (raw != null && TimeOnly.TryParse(raw.ToString(), CultureInfo.InvariantCulture, DateTimeStyles.None, out var t)) return t;
        }
        catch { }
        return fallback;
    }

    // =========================================================================
    // 4. REQUIREMENTS & ICONS
    // =========================================================================
    public static (IReadOnlyList<string> Skills, IReadOnlyList<VehicleType> Vehicles) RequirementsFor(IEnumerable<ProjectType> types)
    {
        var skills = new List<string>();
        var vehicles = new List<VehicleType>();
        foreach (var type in types)
        {
            if (_requirements.TryGetValue(type, out var r))
            {
                skills.AddRange(r.RequiredSkills);
                vehicles.AddRange(r.RequiredVehicleTypes);
            }
            else if (DefaultRequirements.TryGetValue(type, out var d))
            {
                skills.AddRange(d.Skills);
                vehicles.AddRange(d.Vehicles);
            }
        }
        return (skills.Where(s => !string.IsNullOrWhiteSpace(s)).Distinct(StringComparer.OrdinalIgnoreCase).ToList(),
                vehicles.Where(v => v != VehicleType.None).Distinct().ToList());
    }

    /// <summary>Icon for a skill / vehicle type; unmapped keys fall back to a short text tag so nothing is hidden.</summary>
    public static string IconFor(string key)
    {
        if (string.IsNullOrWhiteSpace(key)) return string.Empty;
        string k = key.Trim();
        if (_icons.TryGetValue(k, out var icon) || _icons.TryGetValue(k.Replace(' ', '_'), out icon)) return icon;
        return k.Length <= 4 ? k : k[..4];
    }

    public static string IconFor(VehicleType type) => type == VehicleType.None ? string.Empty : IconFor(type.ToString());

    public static string IconFor(ProjectType type) => type == ProjectType.None ? string.Empty : IconFor(type.ToString());

    // =========================================================================
    // 5. WEEKLY LOAD (database, cached; offline => unknown => treated as 0)
    // =========================================================================
    /// <summary>Days each worker was present on the Monday..(date-1) stretch of <paramref name="date"/>'s week.</summary>
    public static int DaysWorkedEarlierThisWeek(long staffId, DateOnly date)
    {
        DateOnly weekStart = date.AddDays(-(((int)date.DayOfWeek + 6) % 7));
        bool stale = _weekKey != date || DateTime.UtcNow - _weekLoadedAt > TimeSpan.FromMinutes(2);
        if (stale)
        {
            _weekKey = date;
            _weekLoadedAt = DateTime.UtcNow;
            _daysWorkedBefore = new Dictionary<long, int>();
            try
            {
                if (ServerEnvironment.IsOnline && date > weekStart)
                {
                    ServerDispatcher.EnsureScheduleSchema();
                    var rows = new PostgreSqlService().Query<dynamic>(
                        @"SELECT staff_id, COUNT(*) AS days FROM daily_attendances
                          WHERE schedule_date >= @From AND schedule_date < @To
                            AND attendance_status IN ('Presence','Late','NightShift')
                          GROUP BY staff_id;",
                        new { From = weekStart.ToDateTime(TimeOnly.MinValue), To = date.ToDateTime(TimeOnly.MinValue) });
                    foreach (var r in rows) _daysWorkedBefore[Convert.ToInt64(r.staff_id)] = Convert.ToInt32(r.days);
                }
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[ScheduleManager.BoardContext] Weekly load unavailable: {ex.Message}");
            }
        }
        return _daysWorkedBefore.TryGetValue(staffId, out int n) ? n : 0;
    }
}
