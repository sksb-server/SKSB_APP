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
        ["Driver"] = "🚗", ["Excavator Operator"] = "🚜", [SkillCatalog.GeneralWorker] = SkillCatalog.GeneralWorkerIcon,
        ["Car"] = "🚗", ["Hilux"] = "🚙", ["Van"] = "🚐", ["Lorry"] = "🚚", ["Tipper_Lorry"] = "🚛", ["Roro_Lorry"] = "🚛",
        ["Crane_Lorry"] = "🏗", ["Skylift"] = "🏗", ["Backhoe"] = "🚜", ["Excavator"] = "🚜", ["Roller"] = "🚧",
        ["Compressor"] = "🔨", ["Generator"] = "⚡", ["Motorcycle"] = "🏍", ["Machinery"] = "⚙",
        ["Pipe_Repair"] = "🔧", ["Pipe_Laying"] = "🚰", ["Comm_Pipe_Meter"] = "💧", ["Valve_Hydrant"] = "🚒", ["Tapping_Connection"] = "🔗",
        ["Road_Reinstatement"] = "🛣", ["Civil_Earthworks"] = "⛏", ["Machinery_Rental"] = "⚙", ["General_Maintenance"] = "🧰", ["Other"] = "📋"
    };

    // =========================================================================
    // 2. CACHED STATE
    // =========================================================================
    private static List<WorkerPreference> _prefs = new();
    private static List<TeamRecord> _teams = new();
    private static Dictionary<ProjectType, (List<string> Skills, List<VehicleType> Vehicles)> _projectSkillRequirements = new();
    private static Dictionary<string, string> _icons = BuildBaseIcons();
    private static List<ScheduleSkillRecord> _skills = new();
    private static List<SkillLicenseType> _skillLicenseTypes = new();
    /// <summary>Last-read Lookups[SkillIconsTable] overlay, reused when icons are rebuilt without the sheet.</summary>
    private static Dictionary<string, string> _legendIcons = new(StringComparer.OrdinalIgnoreCase);
    public static readonly TimeOnly DefaultDayStartTime = new(6, 0);
    public static readonly TimeOnly DefaultLateCutOff = new(11, 0);
    public static readonly TimeOnly DefaultOtStartTime = new(18, 0);

    private static TimeOnly _dayStartTime = DefaultDayStartTime;
    private static TimeOnly _lateCutOff = DefaultLateCutOff;
    private static TimeOnly _otStartTime = DefaultOtStartTime;

    private static DateOnly _weekKey;
    private static DateTime _weekLoadedAt = DateTime.MinValue;
    private static Dictionary<long, int> _daysWorkedBefore = new();

    public static IReadOnlyList<WorkerPreference> Preferences => _prefs;
    public static IReadOnlyList<TeamRecord> Teams => _teams;
    public static IReadOnlyList<ScheduleSkillRecord> Skills => _skills;
    public static IReadOnlyList<SkillLicenseType> SkillLicenseTypes => _skillLicenseTypes;
    public static TimeOnly DayStartTime => _dayStartTime;
    public static TimeOnly LateCutOff => _lateCutOff;
    public static TimeOnly OTStartTime => _otStartTime;

    private static List<VehicleLicense> _vehicleLicenses = new();
    private static List<SkillLicense> _skillLicenses = new();
    private static bool _licensesLoaded;

    public static IReadOnlyList<VehicleLicense> VehicleLicenses => _vehicleLicenses;
    public static IReadOnlyList<SkillLicense> SkillLicenses => _skillLicenses;
    public static bool HasLicensesLoaded => _licensesLoaded;

    /// <summary>StaffManager licences cached so Preference[Licensed Skills] can be re-evaluated live.</summary>
    public static void SetLicenses(IEnumerable<VehicleLicense>? vehicleLicenses, IEnumerable<SkillLicense>? skillLicenses)
    {
        _vehicleLicenses = (vehicleLicenses ?? Enumerable.Empty<VehicleLicense>()).ToList();
        _skillLicenses = (skillLicenses ?? Enumerable.Empty<SkillLicense>()).ToList();
        _licensesLoaded = true;
    }

    public static WorkerPreference? PreferenceFor(long staffId) => _prefs.FirstOrDefault(p => p.StaffID == staffId);

    public static void SetPreferences(IEnumerable<WorkerPreference> prefs, IEnumerable<TeamRecord> teams)
    {
        _prefs = prefs.ToList();
        _teams = teams.ToList();
    }

    /// <summary>Preference[SkillsTable] rows (source 3). Rebuilds the icon map.</summary>
    public static void SetSkills(IEnumerable<ScheduleSkillRecord>? skills)
    {
        _skills = (skills ?? Enumerable.Empty<ScheduleSkillRecord>()).ToList();
        RebuildIcons();
    }

    /// <summary>Preference[ProjectRequirementsTable] rows.</summary>
    public static void SetProjectSkillRequirements(IEnumerable<ProjectSkillRequirement>? reqs)
    {
        _projectSkillRequirements = (reqs ?? Enumerable.Empty<ProjectSkillRequirement>())
            .Where(r => r.ProjectType != ProjectType.None)
            .GroupBy(r => r.ProjectType)
            .ToDictionary(
                g => g.Key,
                g => (
                    g.First().Skills
                        .Where(s => !string.IsNullOrWhiteSpace(s))
                        .Select(s => s.Trim())
                        .Distinct(StringComparer.OrdinalIgnoreCase)
                        .ToList(),
                    (g.First().Vehicles ?? new List<VehicleType>())
                        .Where(v => v != VehicleType.None)
                        .Distinct()
                        .ToList()));
    }

    /// <summary>StaffManager's lookup_skill_license_types rows (source 1). Rebuilds the icon map.</summary>
    public static void SetSkillLicenseTypes(IEnumerable<SkillLicenseType>? types)
    {
        _skillLicenseTypes = (types ?? Enumerable.Empty<SkillLicenseType>()).ToList();
        RebuildIcons();
    }

    /// <summary>
    /// Distinct (case-insensitive, first spelling wins) union of every configured skill:
    /// SkillsTable -> SkillLicenseTypesTable -> VehicleLicenseClass role skills -> General Worker.
    /// </summary>
    public static List<string> AllConfiguredSkills()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        void Add(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            string n = name.Trim();
            if (seen.Add(n)) result.Add(n);
        }

        foreach (var s in _skills) Add(s.SkillName);
        foreach (var t in _skillLicenseTypes) Add(t.SkillLicenseName);
        foreach (var r in SkillCatalog.VehicleLicenseRoleSkills) Add(r);
        Add(SkillCatalog.GeneralWorker);
        return result;
    }

    /// <summary>Every licensed skill name: SkillLicenseTypes plus vehicle-license role skills (case-insensitive).</summary>
    public static HashSet<string> LicensedSkillNames()
    {
        var set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var t in _skillLicenseTypes)
            if (!string.IsNullOrWhiteSpace(t.SkillLicenseName)) set.Add(t.SkillLicenseName.Trim());
        foreach (var r in SkillCatalog.VehicleLicenseRoleSkills) set.Add(r);
        return set;
    }

    /// <summary>True when the skill is licensed (known licence skill, or a "License {class}" entry from ApplyLicensedSkills).</summary>
    public static bool IsLicensedSkill(string? skill)
    {
        if (string.IsNullOrWhiteSpace(skill)) return false;
        string s = skill.Trim();
        return LicensedSkillNames().Contains(s) || s.StartsWith("License ", StringComparison.OrdinalIgnoreCase);
    }

    /// <summary>
    /// The worker's Primary Skill when it counts: unlicensed skills always count; a licensed Primary Skill only counts
    /// while it is among the worker's active LicensedSkills. Otherwise null.
    /// </summary>
    public static string? EffectivePrimarySkill(WorkerPreference? pref)
    {
        if (pref == null || string.IsNullOrWhiteSpace(pref.PrimarySkill)) return null;
        string primary = pref.PrimarySkill.Trim();
        if (!IsLicensedSkill(primary)) return primary;
        return pref.LicensedSkills.Any(s => s.Trim().Equals(primary, StringComparison.OrdinalIgnoreCase)) ? primary : null;
    }

    /// <summary>
    /// When the worker's licensed Primary Skill is inactive on <paramref name="scheduleDate"/> but a historical /
    /// suspended / future license record exists, returns a short dropdown note and a detailed comment text.
    /// </summary>
    public static (string ShortNote, string Detail)? DescribeInactivePrimarySkillLicense(WorkerPreference? pref, DateOnly scheduleDate)
    {
        if (pref == null || pref.StaffID <= 0 || string.IsNullOrWhiteSpace(pref.PrimarySkill)) return null;
        string primary = pref.PrimarySkill.Trim();
        if (!IsLicensedSkill(primary) || EffectivePrimarySkill(pref) != null) return null;

        var records = new List<(bool Suspended, DateOnly RenewDate, DateOnly ExpiryDate)>();
        records.AddRange(_vehicleLicenses
            .Where(l => l.StaffID == pref.StaffID && l.VehicleLicenseClass != VehicleLicenseClass.None &&
                        (SkillCatalog.RoleSkillFor(l.VehicleLicenseClass) ?? SkillCatalog.Driver).Equals(primary, StringComparison.OrdinalIgnoreCase))
            .Select(l => (l.VehicleLicenseStatus == VehicleLicenseStatus.Suspended, l.RenewDate, l.ExpiryDate)));
        records.AddRange(_skillLicenses
            .Where(l => l.StaffID == pref.StaffID && !string.IsNullOrWhiteSpace(l.SkillLicenseType) &&
                        l.SkillLicenseType.Trim().Equals(primary, StringComparison.OrdinalIgnoreCase))
            .Select(l => (l.SkillLicenseStatus == SkillLicenseStatus.Suspended, l.RenewDate, l.ExpiryDate)));
        if (records.Count == 0) return null;

        if (records.All(r => r.Suspended))
            return ($"{primary} license suspended", $"{primary} license is suspended");

        var live = records.Where(r => !r.Suspended).ToList();
        var expired = live.Where(r => r.ExpiryDate < scheduleDate).OrderByDescending(r => r.ExpiryDate).Cast<(bool Suspended, DateOnly RenewDate, DateOnly ExpiryDate)?>().FirstOrDefault();
        if (expired != null)
            return ($"{primary} license expired", $"{primary} license expired on {expired.Value.ExpiryDate:dd/MM/yyyy}");

        var future = live.Where(r => r.RenewDate > scheduleDate).OrderBy(r => r.RenewDate).Cast<(bool Suspended, DateOnly RenewDate, DateOnly ExpiryDate)?>().FirstOrDefault();
        if (future != null)
            return ($"{primary} license not yet valid", $"{primary} license not valid until {future.Value.RenewDate:dd/MM/yyyy}");

        return ($"{primary} license inactive", $"{primary} license is not valid on {scheduleDate:dd/MM/yyyy}");
    }

    /// <summary>Distinct (first spelling wins) union of SkillsTable and General Worker, excluding licensed skills.</summary>
    public static List<string> UnlicensedSkills()
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var result = new List<string>();
        void Add(string? name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            string n = name.Trim();
            if (IsLicensedSkill(n)) return;
            if (seen.Add(n)) result.Add(n);
        }

        foreach (var s in _skills) Add(s.SkillName);
        Add(SkillCatalog.GeneralWorker);
        return result;
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
                _skills = TableMapper.ExtractFromTable<ScheduleSkillRecord>((object)prefSheet, "SkillsTable");
                SetProjectSkillRequirements(TableMapper.ExtractFromTable<ProjectSkillRequirement>((object)prefSheet, "ProjectRequirementsTable"));
            }

            dynamic? lookups = SheetMapperHelper.ResolveSheetSafely(wbObj, new[] { "Lookups" }, "Lookups");
            RebuildIcons((object?)lookups);

            dynamic? attSheet = SheetMapperHelper.ResolveSheetSafely(wbObj, SheetMapper.MasterSheetAliases, "Daily Attendances");
            if (attSheet != null)
            {
                _dayStartTime = ReadTimeCell(attSheet, SheetMapper.DayStartCell, DefaultDayStartTime);
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
    private static Dictionary<string, string> BuildBaseIcons()
    {
        var icons = new Dictionary<string, string>(DefaultIcons, StringComparer.OrdinalIgnoreCase);
        foreach (var kv in SkillCatalog.VehicleLicenseRoleIcons) icons[kv.Key] = kv.Value;
        return icons;
    }

    /// <summary>
    /// Rebuilds the icon map in precedence order: built-in defaults + vehicle-licence role icons ->
    /// SkillLicenseTypesTable -> SkillsTable -> Lookups[SkillIconsTable] (for extra unmapped keys).
    /// </summary>
    public static void RebuildIcons(object? lookupsSheet = null)
    {
        if (lookupsSheet != null)
        {
            try
            {
                var legend = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                foreach (var i in TableMapper.ExtractFromTable<SkillIcon>(lookupsSheet, "SkillIconsTable"))
                    if (!string.IsNullOrWhiteSpace(i.Key) && !string.IsNullOrWhiteSpace(i.Icon)) legend[i.Key.Trim()] = i.Icon.Trim();
                _legendIcons = legend;
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[ScheduleManager.BoardContext] SkillIconsTable read warning: {ex.Message}");
            }
        }

        var icons = BuildBaseIcons();
        foreach (var t in _skillLicenseTypes)
            if (!string.IsNullOrWhiteSpace(t.SkillLicenseName) && !string.IsNullOrWhiteSpace(t.DisplayIcon))
                icons[t.SkillLicenseName.Trim()] = t.DisplayIcon.Trim();
        foreach (var s in _skills)
            if (!string.IsNullOrWhiteSpace(s.SkillName) && !string.IsNullOrWhiteSpace(s.DisplayIcon))
                icons[s.SkillName.Trim()] = s.DisplayIcon.Trim();
        foreach (var kv in _legendIcons) icons.TryAdd(kv.Key, kv.Value);
        _icons = icons;
    }
    public static (IReadOnlyList<string> Skills, IReadOnlyList<VehicleType> Vehicles) RequirementsFor(IEnumerable<ProjectType> types)
    {
        var skills = new List<string>();
        var vehicles = new List<VehicleType>();
        foreach (var type in types)
        {
            // Skills and vehicles come exclusively from Preference[ProjectRequirementsTable].
            if (_projectSkillRequirements.TryGetValue(type, out var req))
            {
                skills.AddRange(req.Skills);
                vehicles.AddRange(req.Vehicles);
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
