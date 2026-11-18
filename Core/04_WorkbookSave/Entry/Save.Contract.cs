namespace SKSB_App.Core.WorkbookSave.Entry;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.WorkbookSave.Storage;

/// <summary>
/// Non-generic view of a <see cref="SaveContract{T}"/>. Lets COM-facing classes (Ribbon controllers)
/// drive a module without carrying a generic type argument: the CLR does not expose a COM class
/// interface for any type with a generic class in its inheritance chain, which would make every
/// Ribbon callback unreachable from Excel.
/// </summary>
public interface ISaveContract
{
    string ModuleName { get; }
    string[] MasterSheetAliases { get; }
    string MasterTableName { get; }
    Action<bool>? SyncLookups { get; }

    /// <summary>Runs the full validate + persist pipeline (<see cref="SaveEntry.Save{T}"/>).</summary>
    void Save();

    /// <summary>
    /// Fetches the dataset and returns whether it is empty, plus an action that populates the worksheet
    /// with it (null when nothing was fetched or the module has no populate routine).
    /// </summary>
    (bool IsEmpty, Action? Populate) FetchForPopulate();

    /// <summary>True while an offline save is queued for upload (<c>*_PendingUpload.json</c> exists).</summary>
    bool HasPendingUpload { get; }

    /// <summary>Deletes the queued offline upload.</summary>
    void ClearPendingUpload();

    /// <summary>Applies the module window caption, with an <c>[OFFLINE MODE]</c> suffix while offline.</summary>
    void UpdateWindowCaption();
}

public class SaveContract<T> : ISaveContract
{
    public string ModuleName { get; set; } = "Manager";
    public string ServerTableName { get; set; } = string.Empty;

    // =========================================================================
    // 1. WORKBOOK IDENTIFICATION & STARTUP CONFIGURATION
    // =========================================================================
    public string WindowCaption { get; set; } = string.Empty;
    public string[] MasterSheetAliases { get; set; } = Array.Empty<string>();
    public string MasterTableName { get; set; } = string.Empty;
    public Type[] EnumsToSync { get; set; } = Array.Empty<Type>();

    // =========================================================================
    // 2. SAVE / LOAD PIPELINE DELEGATES
    // =========================================================================
    public Func<(T? Data, List<string> Errors, List<string> Warnings)> ExtractAndValidate { get; set; } = null!;
    public Predicate<T?> IsDatasetEmpty { get; set; } = data => data == null;

    /// <summary>
    /// When true, an empty (but non-null) server dataset still runs <see cref="PopulateWorksheet"/> on startup,
    /// so auxiliary/reference tables (e.g. fuel companies, workshops) load even before any master record exists.
    /// </summary>
    public bool PopulateStartupWhenEmpty { get; set; }
    public Func<T, (bool Success, string ErrorMessage)> SaveToServer { get; set; } = null!;

    /// <summary>
    /// Offline-save hook: modules with auxiliary tables override it to snapshot those tables too. It must also
    /// persist the master payload (default: <see cref="OfflineSnapshot.SaveOfflineSnapshot{T}"/>).
    /// </summary>
    public Action<T>? SaveToOffline { get; set; }

    /// <summary>Offline save: module hook (or default) writes <c>*_data.json</c> + <c>*_PendingUpload.json</c>.</summary>
    public void SaveOffline(T data)
    {
        if (SaveToOffline != null) SaveToOffline(data);
        else OfflineSnapshot.SaveOfflineSnapshot(data);
    }

    public Action<T>? PopulateWorksheet { get; set; }
    public Func<T?>? FetchFromServer { get; set; }
    public Action<bool>? SyncLookups { get; set; }
    public Predicate<T>? ValidateOfflinePayload { get; set; }
    public Func<T, T, string>? FormatConflictSummary { get; set; }

    /// <summary>
    /// Last chance to stop a commit: runs after validation (and the empty-dataset check) and before anything is
    /// written. Return false to cancel the save / keep the workbook open on close. Typically wired to
    /// <c>PastRecordGuard</c>. Not used for the silent offline-upload on startup.
    /// </summary>
    public Func<T, bool>? ConfirmBeforeSave { get; set; }

    // =========================================================================
    // 3. NON-GENERIC COMMAND SURFACE (ISaveContract)
    // =========================================================================
    void ISaveContract.Save() => SaveEntry.Save(this);

    public bool HasPendingUpload => OfflineSnapshot.HasPendingSnapshot<T>();

    public void ClearPendingUpload() => OfflineSnapshot.ClearPendingSnapshot<T>();

    public void UpdateWindowCaption()
    {
        if (string.IsNullOrWhiteSpace(WindowCaption)) return;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp?.ActiveWindow == null) return;
            xlApp.ActiveWindow.Caption = ServerEnvironment.IsOffline ? $"{WindowCaption} [OFFLINE MODE]" : WindowCaption;
        }
        catch { }
    }

    (bool IsEmpty, Action? Populate) ISaveContract.FetchForPopulate()
    {
        if (FetchFromServer == null) return (true, null);

        T? data = FetchFromServer();
        bool isEmpty = IsDatasetEmpty(data);
        var populate = PopulateWorksheet;

        return data != null && populate != null
            ? (isEmpty, () => populate(data))
            : (isEmpty, null);
    }
}
