namespace SKSB_App.Core.WorkbookSave.Entry;

using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

public static class SaveEntry
{
    public static void Save<T>(SaveContract<T> contract)
    {
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;

        while (true)
        {
            var (data, validationErrors, validationWarnings) = contract.ExtractAndValidate();
            if (validationErrors.Count > 0 || validationWarnings.Count > 0)
            {
                using var valDialog = new SummaryDialog(contract.ModuleName, validationErrors, validationWarnings);
                valDialog.ShowDialog();
                if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                {
                    return;
                }
            }
#if DEBUG
            using var debugPrompt = new ExitConflictDialog(
                contract.ModuleName,
                string.Empty,
                allowSaveOffline: true,
                isClosing: false,
                isDebugPrompt: true
            );
            debugPrompt.ShowDialog();
            switch (debugPrompt.Resolution)
            {
                case ExitResolution.Retry:
                    var (dbSuccess, dbErr) = contract.SaveToServer(data!);
                    if (dbSuccess)
                    {
                        OfflineSnapshot.ClearPendingSnapshot<T>();
                        contract.PopulateWorksheet?.Invoke(data!);
                        try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                        SaveNotificationDialog.ShowSuccess(contract.ModuleName, "Debug: Data committed to PostgreSQL server successfully.");
                        return;
                    }
                    else
                    {
                        MessageBox.Show($"Server Save Failed:\n\n{dbErr}", "Debug Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        continue;
                    }
                case ExitResolution.SaveLocallyAndQuit:
                    OfflineSnapshot.SaveOfflineSnapshot(data);
                    contract.PopulateWorksheet?.Invoke(data!);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    SaveNotificationDialog.ShowOffline(contract.ModuleName, "Debug: Records stashed to offline JSON storage.");
                    return;
                case ExitResolution.QuitAnyway:
                case ExitResolution.Cancel:
                default:
                    return;
            }
#else
            if (data == null || contract.IsDatasetEmpty(data))
            {
                SaveNotificationDialog.ShowSuccess(contract.ModuleName, "The worksheet contains no records to save.");
                return;
            }

            var (success, dbErrRelease) = contract.SaveToServer(data);
            if (success)
            {
                OfflineSnapshot.ClearPendingSnapshot<T>();
                contract.PopulateWorksheet?.Invoke(data);
                try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                SaveNotificationDialog.ShowSuccess(contract.ModuleName, "Save completed successfully. All records persisted to database.");
                return;
            }

            using var conflictDialog = new ExitConflictDialog(
                contract.ModuleName,
                string.IsNullOrWhiteSpace(dbErrRelease) ? "Could not establish a connection to PostgreSQL." : dbErrRelease,
                allowSaveOffline: true,
                isClosing: false,
                isDebugPrompt: false
            );
            conflictDialog.ShowDialog();
            switch (conflictDialog.Resolution)
            {
                case ExitResolution.Retry:
                    continue;
                case ExitResolution.SaveLocallyAndQuit:
                    OfflineSnapshot.SaveOfflineSnapshot(data);
                    contract.PopulateWorksheet?.Invoke(data);
                    try { ((dynamic)ExcelDnaUtil.Application).ActiveWorkbook.Saved = true; } catch { }
                    SaveNotificationDialog.ShowOffline(contract.ModuleName, "Saved records to local offline storage.");
                    return;
                case ExitResolution.Cancel:
                default:
                    return;
            }
#endif
        }
    }

    public static void InterceptClose<T>(
        dynamic wb,
        ref bool cancel,
        SaveContract<T> contract,
        ref bool isExiting,
        Action unEventHooks)
    {
        if (isExiting)
        {
            cancel = false;
            return;
        }

        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;

        try
        {
            while (true)
            {
                var (data, validationErrors, validationWarnings) = contract.ExtractAndValidate();
                if (validationErrors.Count > 0 || validationWarnings.Count > 0)
                {
                    using var valDialog = new SummaryDialog(contract.ModuleName, validationErrors, validationWarnings);
                    valDialog.ShowDialog();
#if DEBUG
                    if (valDialog.Action == ValidationDialogAction.DebugQuitAnyway)
                    {
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    }
#endif
                    if (validationErrors.Count > 0 || valDialog.Action != ValidationDialogAction.ProceedAnyway)
                    {
                        cancel = true;
                        return;
                    }
                }
#if DEBUG
                using var closePrompt = new ExitConflictDialog(
                    contract.ModuleName,
                    string.Empty,
                    allowSaveOffline: true,
                    isClosing: true,
                    isDebugPrompt: true
                );
                closePrompt.ShowDialog();
                switch (closePrompt.Resolution)
                {
                    case ExitResolution.Retry:
                        var (dbSuccess, dbErr) = contract.SaveToServer(data!);
                        if (dbSuccess)
                        {
                            OfflineSnapshot.ClearPendingSnapshot<T>();
                            isExiting = true;
                            unEventHooks();
                            wb.Saved = true;
                            cancel = false;
                            QuitExcelIfLastWorkbook(wb);
                            return;
                        }
                        else
                        {
                            MessageBox.Show($"Server Save Failed:\n\n{dbErr}", "Debug Server Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            continue;
                        }
                    case ExitResolution.SaveLocallyAndQuit:
                        OfflineSnapshot.SaveOfflineSnapshot(data);
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    case ExitResolution.QuitAnyway:
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    case ExitResolution.Cancel:
                    default:
                        cancel = true;
                        return;
                }
#else
                if (data == null || contract.IsDatasetEmpty(data))
                {
                    isExiting = true;
                    wb.Saved = true;
                    cancel = false;
                    QuitExcelIfLastWorkbook(wb);
                    return;
                }

                var (success, dbError) = contract.SaveToServer(data);
                if (success)
                {
                    OfflineSnapshot.ClearPendingSnapshot<T>();
                    isExiting = true;
                    unEventHooks();
                    wb.Saved = true;
                    cancel = false;
                    QuitExcelIfLastWorkbook(wb);
                    return;
                }

                using var dialog = new ExitConflictDialog(
                    contract.ModuleName,
                    string.IsNullOrWhiteSpace(dbError) ? "Could not persist records to the PostgreSQL server." : dbError,
                    allowSaveOffline: true,
                    isClosing: true,
                    isDebugPrompt: false
                );
                dialog.ShowDialog();
                switch (dialog.Resolution)
                {
                    case ExitResolution.Retry:
                        continue;
                    case ExitResolution.SaveLocallyAndQuit:
                        OfflineSnapshot.SaveOfflineSnapshot(data);
                        isExiting = true;
                        unEventHooks();
                        wb.Saved = true;
                        cancel = false;
                        QuitExcelIfLastWorkbook(wb);
                        return;
                    case ExitResolution.Cancel:
                    default:
                        cancel = true;
                        return;
                }
#endif
            }
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Error encountered during exit validation:\n\n{ex.Message}\n\n{ex.StackTrace}",
                "Exit Validation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            cancel = true;
        }
    }

    private static void QuitExcelIfLastWorkbook(dynamic wb)
    {
        try
        {
            dynamic app = wb.Application;
            if (app != null && (int)app.Workbooks.Count <= 1)
            {
                app.Quit();
            }
        }
        catch { }
    }
}