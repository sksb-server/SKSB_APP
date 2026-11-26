namespace SKSB_App.Core.WorkbookSave.Entry;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.WorkbookSave.Storage;
using SKSB_App.Core.WorkbookSave.Validation;

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

    /// <summary>Runs extraction + validation only and shows the results in the Validation Task Pane (<see cref="SaveEntry.Validate{T}"/>).</summary>
    void Validate();

    /// <summary>
    /// Fetches the dataset and returns whether it is empty, plus an action that populates the worksheet
    /// with it (null when nothing was fetched or the module has no populate routine).
    /// </summary>
    (bool IsEmpty, Action? Populate) FetchForPopulate();

    /// <summary>True while an offline save is queued for upload (<c>*_PendingUpload.json</c> exists).</summary>
    bool HasPendingUpload { get; }

    /// <summary>Deletes the queued offline upload.</summary>
    void ClearPendingUpload();

    /// <summary>True while a "Save Draft &amp; Close" draft (<c>*_Draft.json</c>) with unfixed errors exists.</summary>
    bool HasLocalDraft { get; }

    /// <summary>Deletes the local draft (never uploaded to the server).</summary>
    void ClearLocalDraft();

    /// <summary>Applies the module window caption, with an <c>[OFFLINE MODE]</c> / <c>[LOCAL DRAFT]</c> suffix when applicable.</summary>
    void UpdateWindowCaption();

    /// <summary>
    /// Runs the module's post-load hook (e.g. an advisory dialog) on the macro thread after the worksheets were
    /// populated and Excel's ScreenUpdating / EnableEvents / Calculation state was restored.
    /// </summary>
    void NotifyAfterLoad();
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
    public Func<(T? Data, List<ValidationIssue> Errors, List<ValidationIssue> Warnings)> ExtractAndValidate { get; set; } = null!;
    public Predicate<T?> IsDatasetEmpty { get; set; } = data => data == null;

    /// <summary>
    /// When true, an empty (but non-null) server dataset still runs <see cref="PopulateWorksheet"/> on startup,
    /// so auxiliary/reference tables (e.g. fuel companies, workshops) load even before any master record exists.
    /// </summary>
    public bool PopulateStartupWhenEmpty { get; set; }
    public Func<T, (bool Success, string ErrorMessage)> SaveToServer { get; set; } = null!;

    /// <summary>
    /// Optional delta save with optimistic concurrency. Arg 2 is <c>forceOverwrite</c>. When the save is rolled back
    /// because other users changed the same rows, return <c>Success = false</c> with a non-empty <c>Conflicts</c>
    /// list; <see cref="SaveEntry"/> then lists the field-level diffs in the Validation Task Pane (Conflict Mode) with
    /// "Keep Server Values" / "Overwrite Server &amp; Save". Takes precedence over <see cref="SaveToServer"/> for
    /// interactive saves.
    /// </summary>
    public Func<T, bool, (bool Success, string ErrorMessage, IReadOnlyList<ConcurrencyConflict>? Conflicts)>? SaveToServerWithConflictCheck { get; set; }

    /// <summary>
    /// Runs after "Keep Server Values" wrote the server's values into the conflicting cells: the module refreshes the
    /// delta baselines of ONLY those rows to the server's latest fingerprint / <c>updated_at</c>, so the next save
    /// treats them as unchanged while every other edited row keeps its local edits and original baseline.
    /// </summary>
    public Action<IReadOnlyList<ConcurrencyConflict>>? OnAcceptServerValues { get; set; }

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

    /// <summary>
    /// Optional lightweight post-save refresh. When set, <see cref="SaveEntry"/> runs it after a successful save
    /// instead of the full <see cref="PopulateWorksheet"/> (e.g. to avoid rewriting large append-only tables).
    /// </summary>
    public Action<T>? RefreshAfterSave { get; set; }
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

    /// <summary>
    /// Optional hook that runs after a startup load or a Ribbon "Load from System" finished populating the
    /// workbook (Excel state already restored). Typically shows an advisory / warnings dialog.
    /// </summary>
    public Action? OnAfterLoad { get; set; }

    public void NotifyAfterLoad() => OnAfterLoad?.Invoke();

    // =========================================================================
    // 3. NON-GENERIC COMMAND SURFACE (ISaveContract)
    // =========================================================================
    void ISaveContract.Save() => SaveEntry.Save(this);

    void ISaveContract.Validate() => SaveEntry.Validate(this);

    public bool HasPendingUpload => OfflineSnapshot.HasPendingSnapshot<T>();

    public void ClearPendingUpload() => OfflineSnapshot.ClearPendingSnapshot<T>();

    public bool HasLocalDraft => WorkbookDraft.HasDraft<T>();

    public void ClearLocalDraft() => WorkbookDraft.Clear<T>();

    public void UpdateWindowCaption()
    {
        if (string.IsNullOrWhiteSpace(WindowCaption)) return;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp?.ActiveWindow == null) return;
            if (HasLocalDraft)
            {
#if DEBUG
                xlApp.ActiveWindow.Caption = ServerEnvironment.IsOffline ? $"{WindowCaption} [DEBUG - OFFLINE / LOCAL DRAFT]" : $"{WindowCaption} [DEBUG - LOCAL DRAFT]";
#else
                xlApp.ActiveWindow.Caption = ServerEnvironment.IsOffline ? $"{WindowCaption} [OFFLINE - LOCAL DRAFT]" : $"{WindowCaption} [LOCAL DRAFT - FIX ERRORS]";
#endif
                return;
            }
#if DEBUG
            xlApp.ActiveWindow.Caption = ServerEnvironment.IsOffline ? $"{WindowCaption} [DEBUG - OFFLINE MODE]" : $"{WindowCaption} [DEBUG SERVER]";
#else
            xlApp.ActiveWindow.Caption = ServerEnvironment.IsOffline ? $"{WindowCaption} [OFFLINE MODE]" : WindowCaption;
#endif
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
