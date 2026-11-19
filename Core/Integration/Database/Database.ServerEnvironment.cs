namespace SKSB_App.Core.Integration.Database;

using Npgsql;
using SKSB_App.Core.Diagnostics;
using System;
using System.IO;
using System.Net.Sockets;

public static class ServerEnvironment
{
#if DEBUG
    private static bool _simulateOffline = false;
#else
    private static readonly bool _simulateOffline = false;
#endif

    private static bool _isServerOffline;
    private static string _lastOfflineReason = string.Empty;

    /// <summary>Raised whenever <see cref="IsOffline"/> flips between Online and Offline Mode.</summary>
    public static event Action? StatusChanged;

    public static bool SimulateOffline
    {
        get => _simulateOffline;
        set
        {
#if DEBUG
            bool before = IsOffline;
            _simulateOffline = value;
            if (before != IsOffline) RaiseStatusChanged();
#else
            throw new InvalidOperationException("Offline simulation is disabled in Release builds.");
#endif

        }
    }

    public static bool IsOffline => _simulateOffline || _isServerOffline;
    public static bool IsOnline => !IsOffline;
    public static string LastOfflineReason => _simulateOffline ? "Simulated offline mode active." : _lastOfflineReason;

    public static void EnterOfflineMode(string? reason = null)
    {
        bool before = IsOffline;
        _isServerOffline = true;
        _lastOfflineReason = reason ?? string.Empty;
        if (!before)
        {
            AppLogger.Warn($"[ServerEnvironment] Entering OFFLINE MODE: {reason}");
            RaiseStatusChanged();
        }
    }

    public static void SetOnline()
    {
        bool before = IsOffline;
        _isServerOffline = false;
        _lastOfflineReason = string.Empty;
        if (before && !IsOffline)
        {
            AppLogger.Info("[ServerEnvironment] Back ONLINE.");
            RaiseStatusChanged();
        }
    }

    private static void RaiseStatusChanged()
    {
        try { StatusChanged?.Invoke(); }
        catch (Exception ex) { AppLogger.Error("[ServerEnvironment] StatusChanged handler failed", ex); }
    }

    /// <summary>Opens a short-lived connection and runs SELECT 1; updates Online/Offline Mode from the outcome.</summary>
    public static bool PingServer(out string errorMessage)
    {
        errorMessage = string.Empty;
        if (SimulateOffline)
        {
            errorMessage = "Simulated offline mode active.";
            EnterOfflineMode(errorMessage);
            return false;
        }

        try
        {
            var csb = new NpgsqlConnectionStringBuilder(PostgreSqlService.DefaultConnectionString)
            {
                Timeout = 3,
                CommandTimeout = 3,
                Pooling = false
            };
            using var conn = new NpgsqlConnection(csb.ConnectionString);
            conn.Open();
            using var cmd = new NpgsqlCommand("SELECT 1;", conn);
            cmd.ExecuteScalar();
            SetOnline();
#if DEBUG
            try { EnsureDebugSchemaBootstrapped(); }
            catch (Exception ex) { AppLogger.Warn($"[Debug Server] Schema bootstrap warning: {ex.Message}"); }
#endif
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            EnterOfflineMode(ex.Message);
            return false;
        }
    }

    /// <summary>True for network-level failures. A PostgresException means the server answered (it is online).</summary>
    public static bool IsConnectionException(Exception? ex)
    {
        for (var e = ex; e != null; e = e.InnerException)
        {
            if (e is PostgresException) return false;
            if (e is TimeoutException || e is SocketException || e is IOException || e is NpgsqlException) return true;
        }
        return false;
    }

    public static void ToggleSimulatedOffline(ExcelDna.Integration.CustomUI.IRibbonUI? ribbonUi, string controlId, bool pressed)
    {
#if DEBUG
        SimulateOffline = pressed;
        if (!pressed) PingServer(out _);
        ribbonUi?.Invalidate();
        AppLogger.Warn(pressed
            ? "[Debug Server Mode] Server is now SIMULATED OFFLINE: all PostgreSQL operations fail immediately to test offline local stashing."
            : "[Debug Server Mode] Simulation disabled: server re-pinged.");
#endif
    }

#if DEBUG
    // =========================================================================
    // DEBUG SCHEMA SANDBOX (debug.* mirrors public.*)
    // =========================================================================
    private static bool _debugSchemaBootstrapped;
    private static readonly object _mirrorLock = new();

    public static void EnsureDebugSchemaBootstrapped()
    {
        if (_debugSchemaBootstrapped || IsOffline) return;
        lock (_mirrorLock)
        {
            if (_debugSchemaBootstrapped) return;
            var csb = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString) { Timeout = 5, CommandTimeout = 30 };
            using var conn = new NpgsqlConnection(csb.ConnectionString);
            conn.Open();
            using (var check = new NpgsqlCommand("SELECT COUNT(*) FROM information_schema.tables WHERE table_schema = 'debug' AND table_name IN ('staff_master', 'vehicle_master', 'project_master');", conn))
            {
                long count = Convert.ToInt64(check.ExecuteScalar());
                if (count < 3)
                {
                    MirrorReleaseToDebugCore(conn, out int tableCount, out int rowCount);
                    AppLogger.Info($"[Debug Server] Bootstrapped 'debug' schema from 'public' ({tableCount} tables, {rowCount} rows mirrored).");
                }
            }
            _debugSchemaBootstrapped = true;
        }
    }

    public static bool MirrorReleaseToDebug(out int tableCount, out int totalRows, out string errorMessage)
    {
        tableCount = 0;
        totalRows = 0;
        errorMessage = string.Empty;
        if (IsOffline)
        {
            errorMessage = LastOfflineReason;
            return false;
        }
        lock (_mirrorLock)
        {
            try
            {
                var csb = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString) { Timeout = 5, CommandTimeout = 60 };
                using var conn = new NpgsqlConnection(csb.ConnectionString);
                conn.Open();
                MirrorReleaseToDebugCore(conn, out tableCount, out totalRows);
                _debugSchemaBootstrapped = true;
                return true;
            }
            catch (Exception ex)
            {
                AppLogger.Error("[Debug Server] Mirror failed", ex);
                errorMessage = ex.Message;
                return false;
            }
        }
    }

    private static void MirrorReleaseToDebugCore(NpgsqlConnection conn, out int tableCount, out int totalRows)
    {
        totalRows = 0;
        using var trans = conn.BeginTransaction();

        void Exec(string sql, out int affected)
        {
            using var cmd = new NpgsqlCommand(sql, conn, trans);
            affected = cmd.ExecuteNonQuery();
        }

        Exec("DROP SCHEMA IF EXISTS debug CASCADE; CREATE SCHEMA debug;", out _);

        var tables = new System.Collections.Generic.List<string>();
        using (var cmd = new NpgsqlCommand("SELECT table_name FROM information_schema.tables WHERE table_schema = 'public' AND table_type = 'BASE TABLE' ORDER BY table_name;", conn, trans))
        using (var r = cmd.ExecuteReader())
            while (r.Read()) tables.Add(r.GetString(0));

        var seqs = new System.Collections.Generic.List<(string Seq, string Table, string Col)>();
        const string seqSql = @"SELECT c.relname AS seq_name, t.relname AS table_name, a.attname AS col_name
FROM pg_class c
JOIN pg_namespace n ON n.oid = c.relnamespace AND n.nspname = 'public'
JOIN pg_depend d ON d.objid = c.oid
JOIN pg_class t ON d.refobjid = t.oid
JOIN pg_attribute a ON a.attrelid = t.oid AND a.attnum = d.refobjsubid
WHERE c.relkind = 'S';";
        using (var cmd = new NpgsqlCommand(seqSql, conn, trans))
        using (var r = cmd.ExecuteReader())
            while (r.Read()) seqs.Add((r.GetString(0), r.GetString(1), r.GetString(2)));

        foreach (var (seqName, _, _) in seqs)
            Exec($"CREATE SEQUENCE debug.\"{seqName}\";", out _);

        foreach (var tbl in tables)
        {
            Exec($"CREATE TABLE debug.\"{tbl}\" (LIKE public.\"{tbl}\" INCLUDING DEFAULTS INCLUDING CONSTRAINTS INCLUDING INDEXES);", out _);
            Exec($"INSERT INTO debug.\"{tbl}\" SELECT * FROM public.\"{tbl}\";", out int n);
            totalRows += n;
        }

        foreach (var (seqName, tblName, colName) in seqs)
        {
            Exec($"ALTER TABLE debug.\"{tblName}\" ALTER COLUMN \"{colName}\" SET DEFAULT nextval('debug.\"{seqName}\"'::regclass);", out _);
            Exec($"ALTER SEQUENCE debug.\"{seqName}\" OWNED BY debug.\"{tblName}\".\"{colName}\";", out _);
            using var cmd = new NpgsqlCommand($"SELECT setval('debug.\"{seqName}\"', COALESCE((SELECT MAX(\"{colName}\") FROM debug.\"{tblName}\"), 0) + 1, false);", conn, trans);
            cmd.ExecuteScalar();
        }

        trans.Commit();
        tableCount = tables.Count;
    }
#endif

    public static void GuardConnection()
    {
        if (IsOffline)
        {
            throw new TimeoutException(_simulateOffline
                ? "[Simulated Offline Mode Active] Connection to PostgreSQL server blocked by user switch."
                : $"[Offline Mode] PostgreSQL server unreachable. {_lastOfflineReason}".TrimEnd());
        }
    }
}
