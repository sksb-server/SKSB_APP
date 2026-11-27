namespace SKSB_App.Core.RuntimeUI.Progress;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Npgsql;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.BugReport;

internal readonly record struct LoadingEstimate(double? LocalMs, double? ServerAvgMs)
{
    public double? EffectiveMs => LocalMs ?? ServerAvgMs;
    public bool IsServerFallback => !LocalMs.HasValue && ServerAvgMs.HasValue;
}

/// <summary>
/// Learns how long each loading operation takes (per PC + server-wide) and serves in-memory estimates
/// to <see cref="LoadingOverlay"/>. Never blocks on the database and never flips the app Offline.
/// </summary>
internal static class LoadingTimeEstimator
{
    // =========================================================================
    // 1. CONSTANTS & STATE
    // =========================================================================
    public const string GlobalMachineKey = "__global__";
#if DEBUG
    public const string CacheFileName = "loading_timings_cache.debug.json";
#else
    public const string CacheFileName = "loading_timings_cache.json";
#endif
    public const double MinValidMs = 150.0;
    public const double MaxValidMs = 300_000.0;
    public const double EmaAlpha = 0.30;
    public const int EstimateWindowDays = 90;
    public const int MinSamplesForTier = 3;
    public const int MaxFallbackPreviousVersions = 3;
    public const double MinValidStepMs = 20.0;

    private static readonly object _lock = new();
#if !DEBUG
    private static readonly object _schemaLock = new();
#endif
    private static readonly object _fileLock = new();
    private static readonly string _machine = Environment.MachineName;

    private static readonly Lazy<string> _appVersion = new(() =>
    {
        try
        {
            string? v = BugReportRepository.ResolveAppVersion();
            return string.IsNullOrWhiteSpace(v) ? "unknown" : v.Trim();
        }
        catch
        {
            return "unknown";
        }
    });

    public static string CurrentAppVersion => _appVersion.Value;

    private static Dictionary<string, double> _localAvgMs = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<string, double> _globalAvgMs = new(StringComparer.OrdinalIgnoreCase);
    private static readonly HashSet<string> _unsyncedLocalKeys = new(StringComparer.OrdinalIgnoreCase);
    private static bool _loaded;
#if !DEBUG
    private static volatile bool _schemaBootstrapped;
#endif
    private static volatile bool _syncedFromServer;
    private static int _refreshRunning;

    private static readonly Regex _longDate = new(@"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b", RegexOptions.Compiled);
    private static readonly Regex _isoDate = new(@"\b\d{4}-\d{2}-\d{2}\b", RegexOptions.Compiled);
    private static readonly Regex _rowCountParens = new(@"\(\s*\d+\s+row\(s\)\s*\)", RegexOptions.Compiled | RegexOptions.IgnoreCase);
    private static readonly Regex _spaces = new(@"\s{2,}", RegexOptions.Compiled);

    private sealed class CacheDto
    {
        public Dictionary<string, double> LocalAvgMs { get; set; } = new();
        public Dictionary<string, double> GlobalAvgMs { get; set; } = new();
    }

    private readonly record struct VersionBucket(
        string AppVersion,
        bool InWindow,
        int Count,
        double SumMs,
        DateTime LatestAt);

#if !DEBUG
    private const string SchemaDdl = @"
CREATE SCHEMA IF NOT EXISTS app;
CREATE TABLE IF NOT EXISTS app.operation_timings (
    operation_key TEXT NOT NULL,
    machine_name TEXT NOT NULL,
    avg_ms DOUBLE PRECISION NOT NULL,
    min_ms DOUBLE PRECISION NOT NULL,
    max_ms DOUBLE PRECISION NOT NULL,
    last_ms DOUBLE PRECISION NOT NULL,
    sample_count INTEGER NOT NULL DEFAULT 1,
    updated_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP,
    PRIMARY KEY (operation_key, machine_name)
);
CREATE OR REPLACE VIEW app.pc_speed_ranking AS
SELECT
    p.machine_name,
    COUNT(*) AS operations_measured,
    SUM(p.sample_count) AS total_samples,
    ROUND(AVG(p.avg_ms / NULLIF(g.avg_ms, 0))::numeric, 2) AS relative_duration_ratio,
    ROUND(AVG(p.avg_ms)::numeric, 0) AS avg_ms,
    MAX(p.updated_at) AS last_seen_at
FROM app.operation_timings p
JOIN app.operation_timings g
  ON g.operation_key = p.operation_key AND g.machine_name = '__global__'
WHERE p.machine_name <> '__global__'
GROUP BY p.machine_name
ORDER BY relative_duration_ratio ASC;

CREATE TABLE IF NOT EXISTS app.operation_timing_samples (
    id BIGSERIAL PRIMARY KEY,
    operation_key TEXT NOT NULL,
    machine_name TEXT NOT NULL,
    app_version TEXT NOT NULL DEFAULT 'unknown',
    duration_ms DOUBLE PRECISION NOT NULL,
    recorded_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_operation_timing_samples_op_rec
    ON app.operation_timing_samples (operation_key, recorded_at DESC);
CREATE INDEX IF NOT EXISTS ix_operation_timing_samples_pc_rec
    ON app.operation_timing_samples (machine_name, recorded_at DESC);

CREATE TABLE IF NOT EXISTS app.operation_step_samples (
    id BIGSERIAL PRIMARY KEY,
    operation_key TEXT NOT NULL,
    step_key TEXT NOT NULL,
    machine_name TEXT NOT NULL,
    app_version TEXT NOT NULL DEFAULT 'unknown',
    duration_ms DOUBLE PRECISION NOT NULL,
    recorded_at TIMESTAMPTZ NOT NULL DEFAULT CURRENT_TIMESTAMP
);
CREATE INDEX IF NOT EXISTS ix_operation_step_samples_op_rec
    ON app.operation_step_samples (operation_key, recorded_at DESC);
CREATE INDEX IF NOT EXISTS ix_operation_step_samples_step_rec
    ON app.operation_step_samples (step_key, recorded_at DESC);";

    private const string InsertSampleSql = @"
INSERT INTO app.operation_timing_samples (operation_key, machine_name, app_version, duration_ms)
VALUES (@Key, @Machine, @Version, @Ms);";

    private const string InsertStepSampleSql = @"
INSERT INTO app.operation_step_samples (operation_key, step_key, machine_name, app_version, duration_ms)
VALUES (@Key, @StepKey, @Machine, @Version, @Ms);";

    private const string UpsertSql = @"
INSERT INTO app.operation_timings (operation_key, machine_name, avg_ms, min_ms, max_ms, last_ms, sample_count, updated_at)
VALUES (@Key, @Machine, @Ms, @Ms, @Ms, @Ms, 1, CURRENT_TIMESTAMP)
ON CONFLICT (operation_key, machine_name) DO UPDATE SET
    avg_ms = CASE WHEN app.operation_timings.sample_count < 3
                  THEN ((app.operation_timings.avg_ms * app.operation_timings.sample_count) + EXCLUDED.last_ms) / (app.operation_timings.sample_count + 1)
                  ELSE (app.operation_timings.avg_ms * 0.7) + (EXCLUDED.last_ms * 0.3) END,
    min_ms = LEAST(app.operation_timings.min_ms, EXCLUDED.last_ms),
    max_ms = GREATEST(app.operation_timings.max_ms, EXCLUDED.last_ms),
    last_ms = EXCLUDED.last_ms,
    sample_count = app.operation_timings.sample_count + 1,
    updated_at = CURRENT_TIMESTAMP
RETURNING avg_ms;";
#endif

    // =========================================================================
    // 2. KEY NORMALIZATION
    // =========================================================================
    public static string NormalizeKey(string? title, string? explicitKey = null)
    {
        if (!string.IsNullOrWhiteSpace(explicitKey)) return explicitKey.Trim();
        if (string.IsNullOrWhiteSpace(title)) return string.Empty;

        string s = _longDate.Replace(title, "{date}");
        s = _isoDate.Replace(s, "{date}");
        return _spaces.Replace(s, " ").Trim();
    }

    public static string NormalizeStepKey(string? step)
    {
        if (string.IsNullOrWhiteSpace(step)) return string.Empty;

        string s = _longDate.Replace(step, "{date}");
        s = _isoDate.Replace(s, "{date}");
        s = _rowCountParens.Replace(s, "({n} row(s))");
        s = _spaces.Replace(s, " ").Trim();
        if (s.Length > 160) s = s.Substring(0, 160).TrimEnd();
        return s;
    }

    // =========================================================================
    // 3. PUBLIC API
    // =========================================================================
    public static void WarmUp() => EnsureLoaded();

    public static LoadingEstimate GetEstimate(string operationKey)
    {
        if (string.IsNullOrWhiteSpace(operationKey)) return default;
        EnsureLoaded();
        lock (_lock)
        {
            double? local = _localAvgMs.TryGetValue(operationKey, out double l) ? l : null;
            double? global = _globalAvgMs.TryGetValue(operationKey, out double g) ? g : null;
            return new LoadingEstimate(local, global);
        }
    }

    public static double? GetEstimateMs(string operationKey) => GetEstimate(operationKey).EffectiveMs;

    public static void RecordDuration(
        string operationKey,
        double elapsedMs,
        IReadOnlyList<(string StepKey, double ElapsedMs)>? steps = null)
    {
        if (string.IsNullOrWhiteSpace(operationKey)) return;
        if (double.IsNaN(elapsedMs) || double.IsInfinity(elapsedMs)) return;
        if (elapsedMs < MinValidMs || elapsedMs > MaxValidMs) return;

        EnsureLoaded();
        lock (_lock)
        {
            _localAvgMs[operationKey] = _localAvgMs.TryGetValue(operationKey, out double old)
                ? old * (1 - EmaAlpha) + elapsedMs * EmaAlpha
                : elapsedMs;
            _unsyncedLocalKeys.Add(operationKey);
        }
        SaveCache();
#if !DEBUG
        _ = Task.Run(() => SyncSampleToServerSafe(operationKey, elapsedMs, steps));
#endif
    }

    // =========================================================================
    // 4. LOCAL CACHE
    // =========================================================================
    private static void EnsureLoaded()
    {
        lock (_lock)
        {
            if (_loaded) return;
            _loaded = true;
        }

        try
        {
            string path = ResolveCacheFilePath();
            if (File.Exists(path))
            {
                var dto = JsonSerializer.Deserialize<CacheDto>(File.ReadAllText(path));
                if (dto != null)
                {
                    lock (_lock)
                    {
                        // Merge so samples recorded during load are not lost.
                        foreach (var kv in dto.LocalAvgMs ?? new()) _localAvgMs.TryAdd(kv.Key, kv.Value);
                        foreach (var kv in dto.GlobalAvgMs ?? new()) _globalAvgMs.TryAdd(kv.Key, kv.Value);
                    }
                }
            }
        }
        catch { /* corrupt/missing cache is non-fatal */ }

        try { ServerEnvironment.StatusChanged += OnStatusChanged; } catch { }
        _ = Task.Run(RefreshFromServerSafe);
    }

    private static string ResolveCacheFilePath()
    {
        return Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "SKSB_App", "Data", CacheFileName);
    }

    private static void OnStatusChanged()
    {
        if (ServerEnvironment.IsOnline && !_syncedFromServer)
            _ = Task.Run(RefreshFromServerSafe);
    }

    private static void SaveCache()
    {
        try
        {
            CacheDto dto;
            lock (_lock)
                dto = new CacheDto { LocalAvgMs = new(_localAvgMs), GlobalAvgMs = new(_globalAvgMs) };

            lock (_fileLock)
            {
                string path = ResolveCacheFilePath();
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                File.WriteAllText(path, JsonSerializer.Serialize(dto));
            }
        }
        catch { /* IO errors ignored */ }
    }

    // =========================================================================
    // 5. SERVER SYNC (background only; never touches Excel COM)
    // =========================================================================
    private static NpgsqlConnection OpenTelemetryConnection()
    {
        string cs = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString)
        {
            Timeout = 3,
            CommandTimeout = 5
        }.ConnectionString;
        var conn = new NpgsqlConnection(cs);
        try { conn.Open(); }
        catch { conn.Dispose(); throw; }
        return conn;
    }

#if !DEBUG
    private static void EnsureSchemaBootstrapped(NpgsqlConnection conn)
    {
        if (_schemaBootstrapped) return;
        lock (_schemaLock)
        {
            if (_schemaBootstrapped) return;
            using var cmd = new NpgsqlCommand(SchemaDdl, conn);
            cmd.ExecuteNonQuery();
            _schemaBootstrapped = true;
        }
    }
#endif

    private static void RefreshFromServerSafe()
    {
        if (ServerEnvironment.IsOffline) return;
        if (Interlocked.Exchange(ref _refreshRunning, 1) != 0) return;
        try
        {
            using var conn = OpenTelemetryConnection();
#if !DEBUG
            EnsureSchemaBootstrapped(conn);
#endif

            string currentVersion = CurrentAppVersion;
            var localBuckets = new Dictionary<string, List<VersionBucket>>(StringComparer.OrdinalIgnoreCase);
            var globalBuckets = new Dictionary<string, List<VersionBucket>>(StringComparer.OrdinalIgnoreCase);

            const string samplesSql = @"
SELECT operation_key,
       (machine_name = @Machine) AS is_local,
       app_version,
       (recorded_at >= CURRENT_TIMESTAMP - (@WindowDays * INTERVAL '1 day')) AS in_window,
       COUNT(*)::int AS cnt,
       SUM(duration_ms) AS sum_ms,
       MAX(recorded_at) AS latest_at
FROM app.operation_timing_samples
GROUP BY operation_key, (machine_name = @Machine), app_version, (recorded_at >= CURRENT_TIMESTAMP - (@WindowDays * INTERVAL '1 day'));";

            try
            {
                using var cmd = new NpgsqlCommand(samplesSql, conn);
                cmd.Parameters.AddWithValue("Machine", _machine);
                cmd.Parameters.AddWithValue("WindowDays", EstimateWindowDays);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string opKey = r.GetString(0);
                    bool isLocal = r.GetBoolean(1);
                    string version = r.IsDBNull(2) ? "unknown" : r.GetString(2);
                    bool inWindow = r.GetBoolean(3);
                    int cnt = r.GetInt32(4);
                    double sumMs = r.GetDouble(5);
                    DateTime latestAt = r.GetDateTime(6);

                    var bucket = new VersionBucket(version, inWindow, cnt, sumMs, latestAt);

                    if (!globalBuckets.TryGetValue(opKey, out var gList))
                    {
                        gList = new List<VersionBucket>();
                        globalBuckets[opKey] = gList;
                    }
                    gList.Add(bucket);

                    if (isLocal)
                    {
                        if (!localBuckets.TryGetValue(opKey, out var lList))
                        {
                            lList = new List<VersionBucket>();
                            localBuckets[opKey] = lList;
                        }
                        lList.Add(bucket);
                    }
                }
            }
            catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
            {
                // operation_timing_samples not created yet; fall back to legacy operation_timings below.
            }

            // Tier 4 fallback: legacy aggregated app.operation_timings
            var legacyLocal = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var legacyGlobal = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            try
            {
                using var cmd = new NpgsqlCommand(
                    "SELECT operation_key, machine_name, avg_ms FROM app.operation_timings WHERE machine_name IN (@Machine, '__global__')", conn);
                cmd.Parameters.AddWithValue("Machine", _machine);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string key = r.GetString(0);
                    string machine = r.GetString(1);
                    double avg = r.GetDouble(2);
                    if (machine == GlobalMachineKey) legacyGlobal[key] = avg;
                    else legacyLocal[key] = avg;
                }
            }
            catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
            {
                if (globalBuckets.Count == 0 && localBuckets.Count == 0)
                    return;
            }

            var local = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var global = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);

            var allLocalKeys = new HashSet<string>(localBuckets.Keys, StringComparer.OrdinalIgnoreCase);
            allLocalKeys.UnionWith(legacyLocal.Keys);
            foreach (string opKey in allLocalKeys)
            {
                double? resolved = localBuckets.TryGetValue(opKey, out var lb)
                    ? ResolveTieredAverage(lb, currentVersion)
                    : null;
                if (resolved.HasValue) local[opKey] = resolved.Value;
                else if (legacyLocal.TryGetValue(opKey, out double leg)) local[opKey] = leg;
            }

            var allGlobalKeys = new HashSet<string>(globalBuckets.Keys, StringComparer.OrdinalIgnoreCase);
            allGlobalKeys.UnionWith(legacyGlobal.Keys);
            foreach (string opKey in allGlobalKeys)
            {
                double? resolved = globalBuckets.TryGetValue(opKey, out var gb)
                    ? ResolveTieredAverage(gb, currentVersion)
                    : null;
                if (resolved.HasValue) global[opKey] = resolved.Value;
                else if (legacyGlobal.TryGetValue(opKey, out double leg)) global[opKey] = leg;
            }

            lock (_lock)
            {
                _globalAvgMs.Clear();
                foreach (var kv in global) _globalAvgMs[kv.Key] = kv.Value;

                var staleKeys = new List<string>();
                foreach (string k in _localAvgMs.Keys)
                {
                    if (!local.ContainsKey(k) && !_unsyncedLocalKeys.Contains(k))
                        staleKeys.Add(k);
                }
                foreach (string k in staleKeys) _localAvgMs.Remove(k);
                foreach (var kv in local) _localAvgMs[kv.Key] = kv.Value;
            }
            _syncedFromServer = true;
            SaveCache();
        }
        catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
        {
            // Table not yet created on server (DEBUG never runs DDL); ignore quietly.
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[LoadingTimeEstimator] Refresh from server failed: {ex.Message}");
        }
        finally
        {
            Interlocked.Exchange(ref _refreshRunning, 0);
        }
    }

    private static double? ResolveTieredAverage(List<VersionBucket> buckets, string currentVersion)
    {
        if (buckets.Count == 0) return null;

        // Tier 1: Current app_version within the rolling window (>= MinSamplesForTier)
        int t1Cnt = 0;
        double t1Sum = 0;
        foreach (var b in buckets)
        {
            if (b.InWindow && string.Equals(b.AppVersion, currentVersion, StringComparison.OrdinalIgnoreCase))
            {
                t1Cnt += b.Count;
                t1Sum += b.SumMs;
            }
        }
        if (t1Cnt >= MinSamplesForTier)
            return t1Sum / t1Cnt;

        // Tier 2: Current app_version + up to MaxFallbackPreviousVersions recent versions within window (>= MinSamplesForTier)
        var t2Allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { currentVersion };
        foreach (string v in buckets
            .Where(b => b.InWindow && !string.Equals(b.AppVersion, currentVersion, StringComparison.OrdinalIgnoreCase))
            .GroupBy(b => b.AppVersion, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(g => g.Max(x => x.LatestAt))
            .Take(MaxFallbackPreviousVersions)
            .Select(g => g.Key))
        {
            t2Allowed.Add(v);
        }

        int t2Cnt = 0;
        double t2Sum = 0;
        foreach (var b in buckets)
        {
            if (b.InWindow && t2Allowed.Contains(b.AppVersion))
            {
                t2Cnt += b.Count;
                t2Sum += b.SumMs;
            }
        }
        if (t2Cnt >= MinSamplesForTier)
            return t2Sum / t2Cnt;

        // Tier 3: Outside the window (all-time) for current + up to MaxFallbackPreviousVersions recent versions
        var t3Allowed = new HashSet<string>(StringComparer.OrdinalIgnoreCase) { currentVersion };
        foreach (string v in buckets
            .Where(b => !string.Equals(b.AppVersion, currentVersion, StringComparison.OrdinalIgnoreCase))
            .GroupBy(b => b.AppVersion, StringComparer.OrdinalIgnoreCase)
            .OrderByDescending(g => g.Max(x => x.LatestAt))
            .Take(MaxFallbackPreviousVersions)
            .Select(g => g.Key))
        {
            t3Allowed.Add(v);
        }

        int t3Cnt = 0;
        double t3Sum = 0;
        foreach (var b in buckets)
        {
            if (t3Allowed.Contains(b.AppVersion))
            {
                t3Cnt += b.Count;
                t3Sum += b.SumMs;
            }
        }
        if (t3Cnt >= MinSamplesForTier)
            return t3Sum / t3Cnt;

        // Fallback within Tier 3: all available samples across all versions if still < MinSamplesForTier
        int allCnt = 0;
        double allSum = 0;
        foreach (var b in buckets)
        {
            allCnt += b.Count;
            allSum += b.SumMs;
        }
        return allCnt > 0 ? allSum / allCnt : null;
    }

#if !DEBUG
    private static void SyncSampleToServerSafe(
        string operationKey,
        double elapsedMs,
        IReadOnlyList<(string StepKey, double ElapsedMs)>? steps)
    {
        if (ServerEnvironment.IsOffline) return;
        try
        {
            using var conn = OpenTelemetryConnection();
            EnsureSchemaBootstrapped(conn);

            string version = CurrentAppVersion;
            double localAvg, globalAvg;
            using (var tx = conn.BeginTransaction())
            {
                using (var cmd = new NpgsqlCommand(InsertSampleSql, conn, tx))
                {
                    cmd.Parameters.AddWithValue("Key", operationKey);
                    cmd.Parameters.AddWithValue("Machine", _machine);
                    cmd.Parameters.AddWithValue("Version", version);
                    cmd.Parameters.AddWithValue("Ms", elapsedMs);
                    cmd.ExecuteNonQuery();
                }

                if (steps != null && steps.Count > 0)
                {
                    using var stepCmd = new NpgsqlCommand(InsertStepSampleSql, conn, tx);
                    var pKey = stepCmd.Parameters.Add("Key", NpgsqlTypes.NpgsqlDbType.Text);
                    var pStep = stepCmd.Parameters.Add("StepKey", NpgsqlTypes.NpgsqlDbType.Text);
                    var pMachine = stepCmd.Parameters.Add("Machine", NpgsqlTypes.NpgsqlDbType.Text);
                    var pVersion = stepCmd.Parameters.Add("Version", NpgsqlTypes.NpgsqlDbType.Text);
                    var pMs = stepCmd.Parameters.Add("Ms", NpgsqlTypes.NpgsqlDbType.Double);

                    pKey.Value = operationKey;
                    pMachine.Value = _machine;
                    pVersion.Value = version;

                    foreach (var (stepKey, stepMs) in steps)
                    {
                        if (string.IsNullOrWhiteSpace(stepKey)) continue;
                        if (double.IsNaN(stepMs) || double.IsInfinity(stepMs)) continue;
                        if (stepMs < MinValidStepMs || stepMs > MaxValidMs) continue;

                        pStep.Value = stepKey;
                        pMs.Value = stepMs;
                        stepCmd.ExecuteNonQuery();
                    }
                }

                localAvg = Upsert(conn, tx, operationKey, _machine, elapsedMs);
                globalAvg = Upsert(conn, tx, operationKey, GlobalMachineKey, elapsedMs);
                tx.Commit();
            }

            lock (_lock)
            {
                _localAvgMs[operationKey] = localAvg;
                _globalAvgMs[operationKey] = globalAvg;
                _unsyncedLocalKeys.Remove(operationKey);
            }
            SaveCache();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[LoadingTimeEstimator] Sync sample failed: {ex.Message}");
        }
    }

    private static double Upsert(NpgsqlConnection conn, NpgsqlTransaction tx, string key, string machine, double ms)
    {
        using var cmd = new NpgsqlCommand(UpsertSql, conn, tx);
        cmd.Parameters.AddWithValue("Key", key);
        cmd.Parameters.AddWithValue("Machine", machine);
        cmd.Parameters.AddWithValue("Ms", ms);
        return Convert.ToDouble(cmd.ExecuteScalar());
    }
#endif
}
