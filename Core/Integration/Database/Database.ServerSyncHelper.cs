namespace SKSB_App.Core.Integration.Database;

using System;
using System.Collections;
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
                    if (debugFormatter != null && DebugDialog.DebugEnabled)
                    {
                        DebugDialog.ShowLog($"[DEBUG] PostgreSQL Download - {scopeName}", debugFormatter(result));
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
