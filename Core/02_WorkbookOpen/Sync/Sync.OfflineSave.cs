namespace SKSB_App.Core.WorkbookOpen.Sync;

using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
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

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp?.ActiveWorkbook;
            dynamic? activeSheetBefore = null;
            try { activeSheetBefore = xlApp?.ActiveSheet; } catch { }

            AppLogger.Info("Calling contract.SyncLookups...");
            contract.SyncLookups?.Invoke(false);

            bool hasPending = OfflineSnapshot.HasPendingSnapshot<T>();
            AppLogger.Info($"HasPendingSnapshot<{typeof(T).Name}> = {hasPending}");

            if (hasPending)
            {
                string pendingFilePath = OfflineSnapshot.GetPendingUploadPath<T>();
                var pendingData = OfflineSnapshot.LoadPendingSnapshot<T>();
                bool isValid = pendingData != null && (contract.ValidateOfflinePayload?.Invoke(pendingData) ?? true);
                AppLogger.Info($"Pending snapshot loaded. Valid = {isValid}");

                if (!isValid)
                {
                    AppLogger.Warn("Found a corrupt or invalid offline pending file. Aborting upload.");
                    MessageBox.Show(
                        "Found a corrupt or invalid offline pending file. The upload has been aborted to protect the database.",
                        "Upload Rejected",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
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
#if DEBUG
                    var serverDataDebug = contract.FetchFromServer != null ? contract.FetchFromServer() : default;
                    string diffMessage = contract.FormatConflictSummary != null && serverDataDebug != null
                        ? contract.FormatConflictSummary(pendingData!, serverDataDebug)
                        : string.Empty;

                    string promptMessage =
                        $"[DEBUG MODE] PENDING OFFLINE SAVE DETECTED\n" +
                        $"Module: {contract.ModuleName}\n\n" +
                        $"  Local Offline Save: {localOfflineUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}\n" +
                        $"  Last Server Commit: {(serverLastUpdateUtc.HasValue ? serverLastUpdateUtc.Value.ToLocalTime().ToString("yyyy-MM-dd HH:mm:ss") : "Unknown / Server Offline")}\n\n" +
                        (string.IsNullOrWhiteSpace(diffMessage) ? string.Empty : $"Changes:\n{diffMessage}\n\n") +
                        $"Choose action:\n" +
                        $"  [Yes]    Commit Offline Save to PostgreSQL Server\n" +
                        $"  [No]     Delete Offline Save (Discard local and load Server)\n" +
                        $"  [Cancel] Keep Offline Save on disk (Skip without uploading)";

                    var debugChoice = MessageBox.Show(
                        promptMessage,
                        "[DEBUG] Offline Sync Decision",
                        MessageBoxButtons.YesNoCancel,
                        MessageBoxIcon.Question
                    );

                    if (debugChoice == DialogResult.Cancel) canCommit = false;
                    else if (debugChoice == DialogResult.No)
                    {
                        OfflineSnapshot.ClearPendingSnapshot<T>();
                        canCommit = false;
                        AppLogger.Info("User discarded offline pending file.");
                    }
                    else canCommit = true;
#else
                    if (serverLastUpdateUtc.HasValue && serverLastUpdateUtc.Value > localOfflineUtc)
                    {
                        var serverData = contract.FetchFromServer != null ? contract.FetchFromServer() : default;
                        string diffMessage = contract.FormatConflictSummary != null && serverData != null
                            ? contract.FormatConflictSummary(pendingData!, serverData)
                            : $"  Your Offline Save: {localOfflineUtc.ToLocalTime():yyyy-MM-dd HH:mm:ss}\n  Current Server:     {serverLastUpdateUtc.Value.ToLocalTime():yyyy-MM-dd HH:mm:ss}\n";

                        var dialogChoice = MessageBox.Show(
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

                        canCommit = (dialogChoice == DialogResult.Yes);
                        if (!canCommit) OfflineSnapshot.ClearPendingSnapshot<T>();
                    }
#endif
                    if (canCommit)
                    {
                        AppLogger.Info("Committing offline pending save to server...");
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
                                contract.PopulateWorksheet?.Invoke(pendingData!);
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

            if (contract.FetchFromServer != null)
            {
                AppLogger.Info("Calling contract.FetchFromServer()...");
                var serverData = contract.FetchFromServer();
                bool isEmpty = contract.IsDatasetEmpty(serverData);
                AppLogger.Info($"FetchFromServer returned: DataIsNull={serverData == null}, IsEmpty={isEmpty}");

                if (serverData != null && !isEmpty)
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

                    SheetOperations.ShowStatusConfirmation($"{contract.ModuleName} records loaded from server.");
                }
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("OfflineSave.HandleStartup failed", ex);
        }
    }
}