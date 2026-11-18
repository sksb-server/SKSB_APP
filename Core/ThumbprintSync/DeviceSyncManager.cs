namespace SKSB_App.Core.ThumbprintSync;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ExcelDna.Integration;
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
        string bridgePath = Path.Combine(addinDirectory, "ZkBridge.exe");
        if (!File.Exists(bridgePath)) bridgePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "ZkBridge.exe");
        if (!File.Exists(bridgePath))
            throw new FileNotFoundException($"Cannot locate ZkBridge.exe next to the add-in ({addinDirectory}).");

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
    /// Converts raw punches into new <see cref="AttendanceSource.Thumbprint"/> entries.
    /// The terminal's enroll number is the Staff ID. Punches are dropped when the worker is unknown, when the same
    /// instant already exists in <paramref name="existing"/> (re-sync), or when it is a double scan.
    /// Entry type alternates per worker and day: the first punch is a CheckIn, then CheckOut, CheckIn...
    /// </summary>
    public static (List<AttendanceEntry> NewEntries, int Duplicates, int UnknownStaff) BuildEntries(
        IEnumerable<TerminalPunch> punches,
        IReadOnlyCollection<AttendanceEntry> existing,
        Func<long, bool> isKnownStaff)
    {
        var result = new List<AttendanceEntry>();
        int duplicates = 0;
        int unknown = 0;

        // Existing thumbprint rows per (staff, date), time-ordered: they seed the in/out alternation and the de-dup check.
        var history = existing
            .Where(e => e.Source == AttendanceSource.Thumbprint)
            .GroupBy(e => (e.StaffID, e.Date))
            .ToDictionary(g => g.Key, g => g.OrderBy(e => e.Time).ToList());

        var parsed = punches
            .Select(p => (Ok: long.TryParse(p.EnrollNumber?.Trim(), out long id), Id: id, p.Timestamp))
            .ToList();

        foreach (var staffDay in parsed.Where(p => p.Ok).GroupBy(p => (p.Id, Date: DateOnly.FromDateTime(p.Timestamp))))
        {
            if (!isKnownStaff(staffDay.Key.Id)) { unknown += staffDay.Count(); continue; }

            if (!history.TryGetValue((staffDay.Key.Id, staffDay.Key.Date), out var timeline))
                timeline = new List<AttendanceEntry>();

            foreach (var punch in staffDay.OrderBy(p => p.Timestamp))
            {
                var time = TimeOnly.FromDateTime(punch.Timestamp);
                time = new TimeOnly(time.Hour, time.Minute, time.Second); // drop sub-second noise

                bool isRepeat = timeline.Any(e => Math.Abs((e.Time.ToTimeSpan() - time.ToTimeSpan()).TotalSeconds) < DoubleScanWindow.TotalSeconds);
                if (isRepeat) { duplicates++; continue; }

                var previous = timeline.LastOrDefault(e => e.Time < time);
                var entry = new AttendanceEntry
                {
                    Date = staffDay.Key.Date,
                    Time = time,
                    StaffID = staffDay.Key.Id,
                    EntryType = previous != null && previous.EntryType == AttendanceEntryType.CheckIn
                        ? AttendanceEntryType.CheckOut
                        : AttendanceEntryType.CheckIn,
                    Source = AttendanceSource.Thumbprint,
                    IsApproved = true
                };
                timeline.Add(entry);
                timeline.Sort((a, b) => a.Time.CompareTo(b.Time));
                result.Add(entry);
            }
        }

        unknown += parsed.Count(p => !p.Ok);
        AppLogger.Info($"[DeviceSyncManager] BuildEntries: new={result.Count}, duplicates={duplicates}, unknownStaff={unknown}");
        return (result, duplicates, unknown);
    }
}
