namespace SKSB_App.Core.RuntimeUI.Progress;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;
using Npgsql;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;

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
    public const string CacheFileName = "loading_timings_cache.json";
    public const double MinValidMs = 150.0;
    public const double MaxValidMs = 300_000.0;
    public const double EmaAlpha = 0.30;

    private static readonly object _lock = new();
    private static readonly object _schemaLock = new();
    private static readonly object _fileLock = new();
    private static readonly string _machine = Environment.MachineName;

    private static Dictionary<string, double> _localAvgMs = new(StringComparer.OrdinalIgnoreCase);
    private static Dictionary<string, double> _globalAvgMs = new(StringComparer.OrdinalIgnoreCase);
    private static bool _loaded;
    private static volatile bool _schemaBootstrapped;
    private static volatile bool _syncedFromServer;
    private static int _refreshRunning;

    private static readonly Regex _longDate = new(@"\b\d{1,2}\s+[A-Za-z]{3,9}\s+\d{4}\b", RegexOptions.Compiled);
    private static readonly Regex _isoDate = new(@"\b\d{4}-\d{2}-\d{2}\b", RegexOptions.Compiled);
    private static readonly Regex _spaces = new(@"\s{2,}", RegexOptions.Compiled);

    private sealed class CacheDto
    {
        public Dictionary<string, double> LocalAvgMs { get; set; } = new();
        public Dictionary<string, double> GlobalAvgMs { get; set; } = new();
    }

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
ORDER BY relative_duration_ratio ASC;";

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

    // =========================================================================
    // 3. PUBLIC API
    // =========================================================================
    public static double? GetEstimateMs(string operationKey)
    {
        if (string.IsNullOrWhiteSpace(operationKey)) return null;
        EnsureLoaded();
        lock (_lock)
        {
            if (_localAvgMs.TryGetValue(operationKey, out double l)) return l;
            if (_globalAvgMs.TryGetValue(operationKey, out double g)) return g;
        }
        return null;
    }

    public static void RecordDuration(string operationKey, double elapsedMs)
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
        }
        SaveCache();
        _ = Task.Run(() => SyncSampleToServerSafe(operationKey, elapsedMs));
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

    private static void RefreshFromServerSafe()
    {
        if (ServerEnvironment.IsOffline) return;
        if (Interlocked.Exchange(ref _refreshRunning, 1) != 0) return;
        try
        {
            using var conn = OpenTelemetryConnection();
            EnsureSchemaBootstrapped(conn);

            var local = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            var global = new Dictionary<string, double>(StringComparer.OrdinalIgnoreCase);
            using (var cmd = new NpgsqlCommand(
                "SELECT operation_key, machine_name, avg_ms FROM app.operation_timings WHERE machine_name IN (@Machine, '__global__')", conn))
            {
                cmd.Parameters.AddWithValue("Machine", _machine);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string key = r.GetString(0);
                    string machine = r.GetString(1);
                    double avg = r.GetDouble(2);
                    if (machine == GlobalMachineKey) global[key] = avg;
                    else local[key] = avg;
                }
            }

            lock (_lock)
            {
                foreach (var kv in local) _localAvgMs[kv.Key] = kv.Value;
                foreach (var kv in global) _globalAvgMs[kv.Key] = kv.Value;
            }
            _syncedFromServer = true;
            SaveCache();
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

    private static void SyncSampleToServerSafe(string operationKey, double elapsedMs)
    {
        if (ServerEnvironment.IsOffline) return;
        try
        {
            using var conn = OpenTelemetryConnection();
            EnsureSchemaBootstrapped(conn);

            double localAvg, globalAvg;
            using (var tx = conn.BeginTransaction())
            {
                localAvg = Upsert(conn, tx, operationKey, _machine, elapsedMs);
                globalAvg = Upsert(conn, tx, operationKey, GlobalMachineKey, elapsedMs);
                tx.Commit();
            }

            lock (_lock)
            {
                _localAvgMs[operationKey] = localAvg;
                _globalAvgMs[operationKey] = globalAvg;
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
}
