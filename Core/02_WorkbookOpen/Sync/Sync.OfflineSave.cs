namespace SKSB_App.Core.WorkbookOpen.Sync;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.Progress;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.IO;
using System.Windows.Forms;

public static class OfflineSave
{
    public static void HandleStartup<T>(SaveContract<T> contract)
    {
        using var _ = AppLogger.Scope($"OfflineSave.HandleStartup<{typeof(T).Name}>({contract.ModuleName})");
        try
        {
            if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode)
            {
                AppLogger.Info("Design Mode / Master Template File detected: Auto-load bypassed.");
                SheetOperations.ShowStatusConfirmation("Design Mode: Auto-load bypassed.");
                return;
            }

            using var progress = LoadingOverlay.BeginScope($"{contract.ModuleName} — Starting Up", "Checking PostgreSQL server connection...");

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = xlApp.ActiveWorkbook;
            dynamic? activeSheetBefore = null;
            try { activeSheetBefore = xlApp.ActiveSheet; } catch { }

            // Make Online/Offline authoritative BEFORE any lookup sync or fetch decides where to read from.
            bool online = ServerEnvironment.PingServer(out string pingError);
            AppLogger.Info($"Startup ping: {(online ? "ONLINE" : $"OFFLINE ({pingError})")}");
            contract.UpdateWindowCaption();

            AppLogger.Info("Calling contract.SyncLookups...");
            LoadingOverlay.Report("Synchronizing lookup lists & dropdowns...");
            contract.SyncLookups?.Invoke(false);

            bool hasPending = OfflineSnapshot.HasPendingSnapshot<T>();
            AppLogger.Info($"HasPendingSnapshot<{typeof(T).Name}> = {hasPending}");

            // Offline Mode: never try to upload; the pending file stays queued until the server is reachable.
            if (ServerEnvironment.IsOffline)
            {
                LoadOfflineSnapshotIntoWorkbook(contract, xlApp, activeSheetBefore, hasPending);
                return;
            }

            if (hasPending)
            {
                string pendingFilePath = OfflineSnapshot.GetPendingUploadPath<T>();
                var pendingData = OfflineSnapshot.LoadPendingSnapshot<T>();
                bool isValid = pendingData != null && (contract.ValidateOfflinePayload?.Invoke(pendingData) ?? true);
                AppLogger.Info($"Pending snapshot loaded. Valid = {isValid}");

                if (!isValid)
                {
                    AppLogger.Warn("Found a corrupt or invalid offline pending file. Aborting upload.");
                    using (LoadingOverlay.Suspend())
                    {
                        MessageBox.Show(
                            "Found a corrupt or invalid offline pending file. The upload has been aborted to protect the database.",
                            "Upload Rejected",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Warning
                        );
                    }
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                }
                else
                {
                    DateTime localOfflineUtc = File.GetLastWriteTimeUtc(pendingFilePath);
                    DateTime? serverLastUpdateRaw = string.IsNullOrEmpty(contract.ServerTableName)
                        ? null
                        : OfflineSnapshot.GetServerLatestTimestamp(contract.ServerTableName);

                    DateTime? serverLastUpdateUtc = serverLastUpdateRaw.HasValue
                        ? (serverLastUpdateRaw.Value.Kind == DateTimeKind.Utc
                            ? serverLastUpdateRaw.Value
                            : serverLastUpdateRaw.Value.ToUniversalTime())
                        : null;

                    AppLogger.Info($"LocalPendingUtc={localOfflineUtc:yyyy-MM-dd HH:mm:ss}, ServerLatestUtc={serverLastUpdateUtc:yyyy-MM-dd HH:mm:ss}");

                    bool canCommit = true;
                    if (serverLastUpdateUtc.HasValue && serverLastUpdateUtc.Value > localOfflineUtc)
                    {
                        var serverData = contract.FetchFromServer != null ? contract.FetchFromServer() : default;
                        string diffMessage = contract.FormatConflictSummary != null && serverData != null
                            ? contract.FormatConflictSummary(pendingData!, serverData)
                            : $"  Your Offline Save: {localOfflineUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}\n  Current Server:     {serverLastUpdateUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}\n";

                        DialogResult dialogChoice;
                        using (LoadingOverlay.Suspend())
                        {
                            dialogChoice = MessageBox.Show(
                                $"=== DATA VERSION CONFLICT DETECTED ===\n" +
                                $"The server was modified after your last offline save.\n\n" +
                                $"{diffMessage}\n\n" +
                                $"Would you like to OVERWRITE the server with your local offline version?\n" +
                                $"  Yes: Overwrite server database with your local offline records.\n" +
                                $"  No:  Discard your offline file and load the latest server copy.",
                                "Synchronization Conflict",
                                MessageBoxButtons.YesNo,
                                MessageBoxIcon.Question
                            );
                        }

                        canCommit = (dialogChoice == DialogResult.Yes);
                        AppLogger.Info($"Offline sync conflict decision: {(canCommit ? "overwrite server with local offline save" : "discard offline save")}.");
                        if (!canCommit) OfflineSnapshot.ClearPendingSnapshot<T>();
                    }
                    if (canCommit)
                    {
                        AppLogger.Info("Committing offline pending save to server...");
                        LoadingOverlay.Report("Uploading pending offline changes to server...");
                        var (success, dbError) = contract.SaveToServer(pendingData!);
                        if (success)
                        {
                            OfflineSnapshot.ClearPendingSnapshot<T>();
                            AppLogger.Info("Committed pending snapshot to server successfully. Calling PopulateWorksheet...");

                            // FREEZE SCREEN DURING POPULATION
                            bool origUpdating = true;
                            bool origEvents = true;
                            int origCalculation = -4105;
                            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
                            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                            try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }
                            try
                            {
                                LoadingOverlay.Report($"Fetching {contract.ModuleName} records from server...");
                                var authoritative = contract.FetchFromServer != null ? contract.FetchFromServer() : default;
                                LoadingOverlay.Report($"Populating {contract.ModuleName} worksheets...");
                                contract.PopulateWorksheet?.Invoke(authoritative ?? pendingData!);
                                if (activeSheetBefore != null)
                                {
                                    try { activeSheetBefore.Activate(); } catch { }
                                }
                            }
                            finally
                            {
                                try { xlApp.Calculation = origCalculation; } catch { }
                                try { xlApp.EnableEvents = origEvents; } catch { }
                                try { xlApp.ScreenUpdating = origUpdating; } catch { }
                            }

                            SaveNotificationDialog.ShowSuccess(
                                contract.ModuleName,
                                "Offline changes uploaded successfully!\n\n  Committed records to PostgreSQL server.\n  Local pending upload stash has been cleared."
                            );
                            return;
                        }
                        else
                        {
                            AppLogger.Error($"Could not commit offline file to database: {dbError}");
                            using (LoadingOverlay.Suspend())
                            {
                                MessageBox.Show(
                                    $"Could not commit offline file to database:\n\n{dbError}\n\nThe offline file has been retained.",
                                    "Server Upload Failed",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Error
                                );
                            }
                        }
                    }
                }
            }

            if (contract.FetchFromServer != null)
            {
                LoadingOverlay.Report($"Fetching {contract.ModuleName} records from server...");
                AppLogger.Info("Calling contract.FetchFromServer()...");
                var serverData = contract.FetchFromServer();
                bool isEmpty = contract.IsDatasetEmpty(serverData);
                AppLogger.Info($"FetchFromServer returned: DataIsNull={serverData == null}, IsEmpty={isEmpty}");

                if (serverData != null && (!isEmpty || contract.PopulateStartupWhenEmpty))
                {
                    AppLogger.Info("Calling contract.PopulateWorksheet with server records...");

                    // TOP-LEVEL FREEZE FOR THE ENTIRE POPULATE PASS
                    bool origUpdating = true;
                    bool origEvents = true;
                    int origCalculation = -4105; // xlCalculationAutomatic
                    try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
                    try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                    try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }

                    try
                    {
                        LoadingOverlay.Report($"Populating {contract.ModuleName} worksheets...");
                        contract.PopulateWorksheet?.Invoke(serverData);

                        if (activeSheetBefore != null)
                        {
                            try { activeSheetBefore.Activate(); } catch { }
                        }
                    }
                    finally
                    {
                        try { xlApp.Calculation = origCalculation; } catch { }
                        try { xlApp.EnableEvents = origEvents; } catch { }
                        try { xlApp.ScreenUpdating = origUpdating; } catch { }
                    }

                    SheetOperations.ShowStatusConfirmation(ServerEnvironment.IsOffline
                        ? $"[OFFLINE MODE] {contract.ModuleName} loaded from local offline snapshot (server became unreachable)."
                        : $"{contract.ModuleName} records loaded from server.");
                    contract.UpdateWindowCaption();
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("OfflineSave.HandleStartup failed", ex);
        }
    }

    /// <summary>Offline Mode startup: loads the local snapshot (and auxiliary snapshots) into the workbook.</summary>
    private static void LoadOfflineSnapshotIntoWorkbook<T>(SaveContract<T> contract, dynamic xlApp, dynamic? activeSheetBefore, bool hasPending)
    {
        AppLogger.Info("OFFLINE MODE: loading local offline snapshot (no upload attempted).");
        LoadingOverlay.Report("Loading local offline snapshot...");
        var data = contract.FetchFromServer != null ? contract.FetchFromServer() : OfflineSnapshot.Load<T>();
        bool isEmpty = contract.IsDatasetEmpty(data);

        if (data != null && (!isEmpty || contract.PopulateStartupWhenEmpty))
        {
            bool origUpdating = true;
            bool origEvents = true;
            int origCalculation = -4105; // xlCalculationAutomatic
            try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
            try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
            try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }
            try
            {
                LoadingOverlay.Report("Populating worksheets from offline snapshot...");
                contract.PopulateWorksheet?.Invoke(data);
                if (activeSheetBefore != null)
                {
                    try { activeSheetBefore.Activate(); } catch { }
                }
            }
            finally
            {
                try { xlApp.Calculation = origCalculation; } catch { }
                try { xlApp.EnableEvents = origEvents; } catch { }
                try { xlApp.ScreenUpdating = origUpdating; } catch { }
            }
        }

        SheetOperations.ShowStatusConfirmation(hasPending
            ? $"[OFFLINE MODE] {contract.ModuleName} loaded from local offline snapshot (pending upload queued)."
            : $"[OFFLINE MODE] {contract.ModuleName} loaded from local offline snapshot.");
    }
}