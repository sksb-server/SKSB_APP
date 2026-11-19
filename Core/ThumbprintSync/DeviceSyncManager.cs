namespace SKSB_App.Core.ThumbprintSync;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Types;

/// <summary>One raw punch read from the attendance terminal.</summary>
public sealed record TerminalPunch(string EnrollNumber, DateTime Timestamp);

/// <summary>One user stored on the terminal.</summary>
public sealed record TerminalUser(int Uid, string EnrollNumber, string Name, int Role, uint Card);

/// <summary>Rewrite of a terminal user's EnrollNumber (and optionally Name) by Uid.</summary>
public sealed record TerminalUserUpdate(int Uid, string EnrollNumber, string? Name = null);

/// <summary>
/// Biometric terminal sync. Talks to the terminal through the external <c>ZkBridge.exe</c> helper placed next to the
/// add-in and converts raw punches into append-only <see cref="AttendanceEntry"/> rows
/// (<see cref="AttendanceSource.Thumbprint"/>). Contains no COM access: the caller appends the result to the
/// workbook on the Excel macro thread.
/// </summary>
public static class DeviceSyncManager
{
    public const string DefaultIp = "192.168.0.110";
    public const int DefaultPort = 4370;

    /// <summary>Punches of the same worker closer than this are one physical scan (double tap).</summary>
    private static readonly TimeSpan DoubleScanWindow = TimeSpan.FromMinutes(1);

    /// <summary>Default working-day boundary (Daily Attendances C2): punches before 06:00 belong to the previous working day.</summary>
    public static readonly TimeOnly DefaultDayStartTime = new(6, 0);

    /// <summary>Working date a punch belongs to: before <paramref name="dayStartTime"/> it rolls back to the previous day.</summary>
    public static DateOnly WorkingDate(DateOnly date, TimeOnly time, TimeOnly dayStartTime) =>
        time < dayStartTime ? date.AddDays(-1) : date;

    /// <summary>True when the punch (real calendar date + time) belongs to <paramref name="workingDate"/>.</summary>
    public static bool IsInWorkingDay(DateOnly punchDate, TimeOnly punchTime, DateOnly workingDate, TimeOnly dayStartTime) =>
        WorkingDate(punchDate, punchTime, dayStartTime) == workingDate;

    // =========================================================================
    // 1. TERMINAL READ (background thread only - never touches Excel)
    // =========================================================================

    /// <summary>
    /// Runs <c>ZkBridge.exe &lt;ip&gt; &lt;port&gt; &lt;commKey&gt; attlog [yyyy-MM-dd]</c>, which must print a JSON array of
    /// <c>{"EnrollNumber":"12","Timestamp":"2026-11-17T08:05:12"}</c> objects on stdout.
    /// </summary>
    public static Task<List<TerminalPunch>> FetchPunchesAsync(DateOnly? since = null, string ip = DefaultIp, int port = DefaultPort, int commKey = 0, string? addinDirectory = null)
    {
        return Task.Run(() =>
        {
            string output = RunBridge($"{ip} {port} {commKey} attlog" + (since.HasValue ? $" {since.Value:yyyy-MM-dd}" : string.Empty), null, $"{ip}:{port}", addinDirectory);
            return Deserialize<List<TerminalPunch>>(output) ?? new List<TerminalPunch>();
        });
    }

    /// <summary>Runs <c>ZkBridge.exe &lt;ip&gt; &lt;port&gt; &lt;commKey&gt; users</c> and returns every user stored on the terminal.</summary>
    public static Task<List<TerminalUser>> FetchTerminalUsersAsync(string ip = DefaultIp, int port = DefaultPort, int commKey = 0, string? addinDirectory = null)
    {
        return Task.Run(() =>
        {
            string output = RunBridge($"{ip} {port} {commKey} users", null, $"{ip}:{port}", addinDirectory);
            return Deserialize<List<TerminalUser>>(output) ?? new List<TerminalUser>();
        });
    }

    /// <summary>
    /// Runs <c>ZkBridge.exe ... applyusers</c> with the updates as a JSON array on stdin. Rewrites EnrollNumber (and Name when
    /// non-null) by Uid, preserving templates and all other user data. Returns the number of users updated.
    /// </summary>
    public static Task<int> ApplyTerminalUsersAsync(IReadOnlyCollection<TerminalUserUpdate> updates, string ip = DefaultIp, int port = DefaultPort, int commKey = 0, string? addinDirectory = null)
    {
        return Task.Run(() =>
        {
            string json = JsonSerializer.Serialize(updates);
            string output = RunBridge($"{ip} {port} {commKey} applyusers", json, $"{ip}:{port}", addinDirectory);
            using var doc = JsonDocument.Parse(output);
            return doc.RootElement.TryGetProperty("Updated", out var n) ? n.GetInt32() : 0;
        });
    }

    private static T? Deserialize<T>(string output) =>
        string.IsNullOrWhiteSpace(output)
            ? default
            : JsonSerializer.Deserialize<T>(output, new JsonSerializerOptions { PropertyNameCaseInsensitive = true });

    /// <summary>Launches ZkBridge.exe, optionally feeds stdin, enforces a 30 s timeout and returns stdout (throws on any error).</summary>
    private static string RunBridge(string arguments, string? stdin, string target, string? addinDirectory = null)
    {
        if (string.IsNullOrWhiteSpace(addinDirectory))
        {
            try { addinDirectory = Path.GetDirectoryName(ExcelDnaUtil.XllPath); } catch { }
        }
        if (string.IsNullOrWhiteSpace(addinDirectory)) addinDirectory = AppContext.BaseDirectory;
        string? parentDirectory = null;
        try { parentDirectory = Path.GetDirectoryName(addinDirectory.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)); } catch { }
        string? rootDirectory = null;
        try { rootDirectory = FolderRegistry.SKSB_App.GetPath(); } catch { }

        string? bridgePath = null;
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var searchDirs = new List<string?> { addinDirectory, parentDirectory, rootDirectory };
        if (!string.IsNullOrWhiteSpace(rootDirectory))
        {
            searchDirs.Add(Path.Combine(rootDirectory, "AddIns"));
#if DEBUG
            searchDirs.Add(Path.Combine(rootDirectory, "ZkBridge", "bin", "Debug", "net8.0-windows"));
#endif
        }
        searchDirs.Add(AppDomain.CurrentDomain.BaseDirectory);
        foreach (string? dir in searchDirs)
        {
            if (string.IsNullOrWhiteSpace(dir) || !seen.Add(dir)) continue;
            string candidate = Path.Combine(dir, "ZkBridge.exe");
            if (File.Exists(candidate)) { bridgePath = candidate; break; }
        }
        if (bridgePath == null)
        {
            string where = !string.IsNullOrWhiteSpace(rootDirectory) && !string.Equals(rootDirectory, addinDirectory, StringComparison.OrdinalIgnoreCase)
                ? $"{addinDirectory} or {rootDirectory}"
                : addinDirectory;
            throw new FileNotFoundException($"Cannot locate ZkBridge.exe next to the add-in ({where}).");
        }

        var startInfo = new ProcessStartInfo
        {
            FileName = bridgePath,
            Arguments = arguments,
            UseShellExecute = false,
            RedirectStandardInput = stdin != null,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            CreateNoWindow = true
        };

        using var process = Process.Start(startInfo) ?? throw new InvalidOperationException("Could not start ZkBridge.exe.");
        Task<string> stdout = process.StandardOutput.ReadToEndAsync();
        Task<string> stderr = process.StandardError.ReadToEndAsync();
        if (stdin != null)
        {
            process.StandardInput.Write(stdin);
            process.StandardInput.Close();
        }
        if (!process.WaitForExit(30000))
        {
            try { process.Kill(); } catch { }
            throw new TimeoutException($"Terminal {target} did not answer within 30 seconds.");
        }

        string error = stderr.GetAwaiter().GetResult();
        if (process.ExitCode != 0 || !string.IsNullOrWhiteSpace(error))
            throw new InvalidOperationException($"ZkBridge error (code {process.ExitCode}): {error}");

        return stdout.GetAwaiter().GetResult();
    }

    // =========================================================================
    // 2. PUNCH -> ATTENDANCE ENTRY CONVERSION (pure, de-duplicating)
    // =========================================================================

    /// <summary>
    /// Removes duplicate rows already present (order preserved). Thumbprint rows within <see cref="DoubleScanWindow"/> of the
    /// previous kept row for the same (staff, working day) are dropped and the kept timeline's in/out alternation is normalized;
    /// other sources drop rows repeating an earlier (date, staff, type, source, hh:mm). The working day is derived from
    /// <paramref name="dayStartTime"/> (default <see cref="DefaultDayStartTime"/>) so night shifts crossing midnight stay together.
    /// </summary>
    public static (List<AttendanceEntry> CleanEntries, int RemovedDuplicates) DeduplicateExisting(IReadOnlyCollection<AttendanceEntry> existing, TimeOnly? dayStartTime = null)
    {
        TimeOnly dayStart = dayStartTime ?? DefaultDayStartTime;
        var drop = new HashSet<AttendanceEntry>(ReferenceEqualityComparer.Instance);

        foreach (var g in existing.Where(e => e.Source == AttendanceSource.Thumbprint).GroupBy(e => (e.StaffID, WorkingDate(e.Date, e.Time, dayStart))))
        {
            var kept = new List<AttendanceEntry>();
            foreach (var e in g.OrderBy(e => e.Date).ThenBy(e => e.Time))   // stable sort
            {
                if (kept.Count > 0 && Math.Abs((e.Date.ToDateTime(e.Time) - kept[^1].Date.ToDateTime(kept[^1].Time)).TotalSeconds) < DoubleScanWindow.TotalSeconds)
                    drop.Add(e);
                else
                    kept.Add(e);
            }
            for (int i = 0; i < kept.Count; i++)
                kept[i].EntryType = i % 2 == 0 ? AttendanceEntryType.CheckIn : AttendanceEntryType.CheckOut;
        }

        var seen = new HashSet<(DateOnly, long, AttendanceEntryType, AttendanceSource, int, int)>();
        var clean = new List<AttendanceEntry>();
        foreach (var e in existing)
        {
            if (drop.Contains(e)) continue;
            if (e.Source != AttendanceSource.Thumbprint && !seen.Add((e.Date, e.StaffID, e.EntryType, e.Source, e.Time.Hour, e.Time.Minute))) continue;
            clean.Add(e);
        }
        return (clean, existing.Count - clean.Count);
    }

    /// <summary>
    /// Converts raw punches into new <see cref="AttendanceSource.Thumbprint"/> entries.
    /// The terminal's enroll number is the Staff ID. Punches are dropped when the worker is unknown, when the same
    /// instant already exists in <paramref name="existing"/> (re-sync), or when it is a double scan.
    /// Entry type alternates per worker and working day (see <see cref="WorkingDate"/>): the first punch is a CheckIn,
    /// then CheckOut, CheckIn... A night-shift check-out after midnight stays in the previous working day's timeline.
    /// </summary>
    public static (List<AttendanceEntry> NewEntries, int Duplicates, int UnknownStaff) BuildEntries(
        IEnumerable<TerminalPunch> punches,
        IReadOnlyCollection<AttendanceEntry> existing,
        Func<long, bool> isKnownStaff,
        TimeOnly? dayStartTime = null)
    {
        TimeOnly dayStart = dayStartTime ?? DefaultDayStartTime;
        var result = new List<AttendanceEntry>();
        int duplicates = 0;
        int unknown = 0;

        // Existing thumbprint rows per (staff, working day), chronological: they seed the in/out alternation and the de-dup check.
        var history = existing
            .Where(e => e.Source == AttendanceSource.Thumbprint)
            .GroupBy(e => (e.StaffID, WorkingDate(e.Date, e.Time, dayStart)))
            .ToDictionary(g => g.Key, g => g.OrderBy(e => e.Date).ThenBy(e => e.Time).ToList());

        var parsed = punches
            .Select(p => (Ok: long.TryParse(p.EnrollNumber?.Trim(), out long id), Id: id, p.Timestamp))
            .ToList();

        foreach (var staffDay in parsed.Where(p => p.Ok).GroupBy(p => (p.Id, WorkDate: WorkingDate(DateOnly.FromDateTime(p.Timestamp), TimeOnly.FromDateTime(p.Timestamp), dayStart))))
        {
            if (!isKnownStaff(staffDay.Key.Id)) { unknown += staffDay.Count(); continue; }

            if (!history.TryGetValue((staffDay.Key.Id, staffDay.Key.WorkDate), out var timeline))
                timeline = new List<AttendanceEntry>();

            foreach (var punch in staffDay.OrderBy(p => p.Timestamp))
            {
                var t = TimeOnly.FromDateTime(punch.Timestamp);
                var time = new TimeOnly(t.Hour, t.Minute, t.Second); // drop sub-second noise
                var date = DateOnly.FromDateTime(punch.Timestamp);  // keep the real calendar date
                DateTime punchDt = date.ToDateTime(time);

                bool isRepeat = timeline.Any(e => Math.Abs((e.Date.ToDateTime(e.Time) - punchDt).TotalSeconds) < DoubleScanWindow.TotalSeconds);
                if (isRepeat) { duplicates++; continue; }

                var previous = timeline.LastOrDefault(e => e.Date.ToDateTime(e.Time) < punchDt);
                var entry = new AttendanceEntry
                {
                    Date = date,
                    Time = time,
                    StaffID = staffDay.Key.Id,
                    EntryType = previous != null && previous.EntryType == AttendanceEntryType.CheckIn
                        ? AttendanceEntryType.CheckOut
                        : AttendanceEntryType.CheckIn,
                    Source = AttendanceSource.Thumbprint,
                    IsApproved = true
                };
                timeline.Add(entry);
                timeline.Sort((a, b) => a.Date.ToDateTime(a.Time).CompareTo(b.Date.ToDateTime(b.Time)));
                result.Add(entry);
            }
        }

        unknown += parsed.Count(p => !p.Ok);
        AppLogger.Info($"[DeviceSyncManager] BuildEntries: new={result.Count}, duplicates={duplicates}, unknownStaff={unknown}");
        return (result, duplicates, unknown);
    }
}
