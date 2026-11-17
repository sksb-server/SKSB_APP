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
    // 1. FETCH WITH OFFLINE FALLBACK
    // =========================================================================

    /// <summary>
    /// Runs <paramref name="dbFetcher"/> against PostgreSQL and caches the result to the local JSON snapshot.
    /// Falls back to <see cref="OfflineSnapshot.Load{T}"/> when the server is (simulated) offline, the query
    /// throws, or the fetcher returns null (the fetcher's way of saying "nothing usable on the server").
    /// </summary>
    public static T? FetchWithOfflineFallback<T>(
        string scopeName,
        Func<PostgreSqlService, T?> dbFetcher,
        bool saveSnapshot = true,
        Func<T, string>? debugFormatter = null)
    {
        using var _ = AppLogger.Scope(scopeName);

        try
        {
            if (!ServerEnvironment.SimulateOffline)
            {
                T? result = dbFetcher(new PostgreSqlService());
                if (result != null)
                {
                    if (saveSnapshot)
                    {
                        // A cache write failure must never discard a successful server read.
                        try { OfflineSnapshot.Save(result, writeToDatabase: false); }
                        catch (Exception exCache) { AppLogger.Warn($"[{scopeName}] Snapshot cache write failed: {exCache.Message}"); }
                    }

#if DEBUG
                    if (debugFormatter != null)
                    {
                        AppLogger.Info($"[PostgreSQL Download - {scopeName}]\n{debugFormatter(result)}");
                    }
#endif
                    return result;
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{scopeName}] Server fetch failed; falling back to offline snapshot", ex);
        }

        T? fallback = OfflineSnapshot.Load<T>();
        AppLogger.Warn($"[{scopeName}] Offline fallback: {(fallback == null ? "no local snapshot found" : "loaded local snapshot")}.");
        return fallback;
    }

    // =========================================================================
    // 1b. SERVER-AUTHORITATIVE LIST FETCH
    // =========================================================================

    /// <summary>
    /// Runs <paramref name="dbFetcher"/> and treats the result as authoritative even when empty: the local snapshot
    /// is refreshed (so a stale cache can never resurrect wiped data) and the list is returned as-is.
    /// The offline snapshot is used only when the server is (simulated) offline or the query throws.
    /// </summary>
    public static List<T> FetchAuthoritativeList<T>(
        string module,
        string tableLabel,
        Func<PostgreSqlService, List<T>> dbFetcher,
        Func<List<T>, string>? debugFormatter = null)
    {
        using var _ = AppLogger.Scope($"{module}.Fetch {tableLabel}");

        if (ServerEnvironment.SimulateOffline)
        {
            OfflineSnapshot.LogDataSource(module, tableLabel, "OFFLINE CACHE", "Simulate Server Offline is ON");
            return OfflineSnapshot.Load<List<T>>() ?? new List<T>();
        }

        try
        {
            var list = dbFetcher(new PostgreSqlService());

            // A cache write failure must never discard a successful server read.
            try { OfflineSnapshot.Save(list, writeToDatabase: false); }
            catch (Exception exCache) { AppLogger.Warn($"[{module}] {tableLabel} snapshot cache write failed: {exCache.Message}"); }

            OfflineSnapshot.LogDataSource(module, tableLabel, "SERVER", list.Count == 0 ? "server reachable but empty; local cache reset (no offline fallback)" : $"{list.Count} row(s)");
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
            OfflineSnapshot.LogDataSource(module, tableLabel, "OFFLINE CACHE", $"server query failed: {ex.GetType().Name}: {ex.Message}");
            var fallback = OfflineSnapshot.Load<List<T>>() ?? new List<T>();
            AppLogger.Warn($"[{module}] Offline fallback: loaded {fallback.Count} {tableLabel} row(s) from the local snapshot.");
            return fallback;
        }
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
    /// through <see cref="OfflineSnapshot.Save{T}"/> with server dispatch. Any failure is captured into
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

            string snapshotPath = OfflineSnapshot.Save(data, writeToDatabase: true);
            return !string.IsNullOrEmpty(snapshotPath);
        }
        catch (Exception ex)
        {
            errorMessage = ex.Message;
            AppLogger.Error($"[ServerSyncHelper] SaveSilent<{typeof(T).Name}> failed", ex);
            return false;
        }
    }
}
