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
