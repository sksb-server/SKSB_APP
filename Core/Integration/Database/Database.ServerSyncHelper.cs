namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.WorkbookSave.Storage;

/// <summary>
/// Shared fetch / save plumbing for every module's ServerSyncCommands: offline simulation guard,
/// local JSON snapshot caching, DEBUG download audit, error logging and offline fallback.
/// </summary>
public static class ServerSyncHelper
{
    // =========================================================================
    // 1. FETCH (ONLINE = SERVER ONLY, OFFLINE = LOCAL SNAPSHOT ONLY)
    // =========================================================================

    /// <summary>
    /// Online Mode: runs <paramref name="dbFetcher"/> against PostgreSQL and returns the result; a null result or a
    /// non-connection failure yields null (never snapshot data). Offline Mode (or a connection failure, which flips
    /// the app into Offline Mode): returns <see cref="OfflineSnapshot.Load{T}"/>. Fetching never writes the snapshot.
    /// </summary>
    public static T? FetchWithOfflineFallback<T>(
        string scopeName,
        Func<PostgreSqlService, T?> dbFetcher,
        bool saveSnapshot = true,
        Func<T, string>? debugFormatter = null)
    {
        using var _ = AppLogger.Scope(scopeName);

        if (ServerEnvironment.IsOnline)
        {
            try
            {
                T? result = dbFetcher(new PostgreSqlService());
#if DEBUG
                if (result != null && debugFormatter != null)
                {
                    AppLogger.Info($"[PostgreSQL Download - {scopeName}]\n{debugFormatter(result)}");
                }
#endif
                return result;
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[{scopeName}] Server fetch failed", ex);
                if (ServerEnvironment.IsConnectionException(ex)) ServerEnvironment.EnterOfflineMode(ex.Message);
                else return default; // Online + SQL error: never mix in offline data.
            }
        }

        T? fallback = OfflineSnapshot.Load<T>();
        AppLogger.Warn($"[{scopeName}] OFFLINE MODE: {(fallback == null ? "no local snapshot found" : "loaded local snapshot")}.");
        return fallback;
    }

    // =========================================================================
    // 1b. SERVER-AUTHORITATIVE LIST FETCH
    // =========================================================================

    /// <summary>
    /// Online Mode: the server result is authoritative even when empty (never mixed with the snapshot).
    /// Offline Mode (or a connection failure that flips the app offline): the local snapshot is used.
    /// A non-connection error while online is logged and yields an empty list. Never writes the snapshot.
    /// </summary>
    public static List<T> FetchAuthoritativeList<T>(
        string module,
        string tableLabel,
        Func<PostgreSqlService, List<T>> dbFetcher,
        Func<List<T>, string>? debugFormatter = null)
    {
        using var _ = AppLogger.Scope($"{module}.Fetch {tableLabel}");

        if (ServerEnvironment.IsOnline)
        {
            try
            {
                var list = dbFetcher(new PostgreSqlService());
                OfflineSnapshot.LogDataSource(module, tableLabel, "SERVER", list.Count == 0 ? "server reachable but empty" : $"{list.Count} row(s)");
#if DEBUG
                if (debugFormatter != null)
                {
                    AppLogger.Info($"[PostgreSQL Download - {tableLabel}]\n{debugFormatter(list)}");
                }
#endif
                return list;
            }
            catch (Exception ex)
            {
                AppLogger.Error($"[{module}] {tableLabel} server query failed", ex);
                if (ServerEnvironment.IsConnectionException(ex)) ServerEnvironment.EnterOfflineMode(ex.Message);
                else return new List<T>(); // Online + SQL error: never mix in offline data.
            }
        }

        OfflineSnapshot.LogDataSource(module, tableLabel, "OFFLINE CACHE", ServerEnvironment.LastOfflineReason);
        var fallback = OfflineSnapshot.Load<List<T>>() ?? new List<T>();
        AppLogger.Warn($"[{module}] OFFLINE MODE: loaded {fallback.Count} {tableLabel} row(s) from the local snapshot.");
        return fallback;
    }

    /// <summary>Server-authoritative fetch of a lookup / auxiliary table with a plain Dapper-mapped query.</summary>
    public static List<T> FetchAuxiliaryTable<T>(string module, string tableLabel, string sql) =>
        FetchAuthoritativeList<T>(module, tableLabel, db => db.Query<T>(sql).ToList());

    // =========================================================================
    // 1c. DYNAMIC DAPPER ROW HELPERS
    // =========================================================================

    /// <summary>Reads a column that may not exist yet on an un-migrated database (DapperRow is an IDictionary); missing => null.</summary>
    public static object? Col(object row, string column) =>
        row is IDictionary<string, object?> d && d.TryGetValue(column, out var value) ? value : null;

    /// <summary>Text form of <see cref="Col"/>; missing / NULL => empty string.</summary>
    public static string ColumnText(object row, string column) => Col(row, column)?.ToString() ?? string.Empty;

    /// <summary>Nullable DB date; NULL and sentinel dates (Year &lt;= 1900) become null instead of DateOnly.MinValue.</summary>
    public static DateOnly? ValidDate(object? raw)
    {
        if (raw == null) return null;
        try
        {
            var d = DateOnly.FromDateTime(Convert.ToDateTime(raw));
            return d.Year <= 1900 ? null : d;
        }
        catch { return null; }
    }

    /// <summary>DB date or <paramref name="fallback"/> when NULL.</summary>
    public static DateOnly DateOrDefault(object? raw, DateOnly fallback) =>
        raw != null ? DateOnly.FromDateTime(Convert.ToDateTime(raw)) : fallback;

    public static decimal Dec(object? raw) => raw != null ? Convert.ToDecimal(raw) : 0m;

    public static decimal? DecN(object? raw) => raw != null ? Convert.ToDecimal(raw) : null;

    // =========================================================================
    // 2. SILENT SAVE (LOCAL SNAPSHOT + SERVER DISPATCH)
    // =========================================================================

    /// <summary>
    /// Runs <paramref name="preSaveAuxiliary"/> (reference / lookup tables), then persists <paramref name="data"/>
    /// to PostgreSQL via <see cref="ServerDispatcher.Dispatch{T}"/> and only then refreshes the local snapshot. Any failure is captured into
    /// <paramref name="errorMessage"/> instead of being thrown. An empty collection is treated as
    /// "nothing to persist" and succeeds without touching the snapshot or the server.
    /// </summary>
    public static bool SaveSilent<T>(T data, out string errorMessage, Action? preSaveAuxiliary = null)
    {
        errorMessage = string.Empty;
        try
        {
            preSaveAuxiliary?.Invoke();

            if (data == null) return true;
            if (data is ICollection { Count: 0 }) return true;

            // Database first: a failed save must never overwrite the local snapshot.
            ServerDispatcher.Dispatch(data);
            OfflineSnapshot.Save(data, writeToDatabase: false);
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error($"[ServerSyncHelper] SaveSilent<{typeof(T).Name}> failed", ex);
            return false;
        }
    }
}
