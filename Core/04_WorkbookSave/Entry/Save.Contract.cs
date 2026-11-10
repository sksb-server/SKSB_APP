namespace SKSB_App.Core.WorkbookSave.Entry;

using System;
using System.Collections.Generic;

/// <summary>
/// Non-generic view of a <see cref="SaveContract{T}"/>. Lets COM-facing classes (Ribbon controllers)
/// drive a module without carrying a generic type argument: the CLR does not expose a COM class
/// interface for any type with a generic class in its inheritance chain, which would make every
/// Ribbon callback unreachable from Excel.
/// </summary>
public interface ISaveContract
{
    string ModuleName { get; }
    string WindowCaption { get; }
    string[] MasterSheetAliases { get; }
    string MasterTableName { get; }
    Type[] EnumsToSync { get; }
    Action<bool>? SyncLookups { get; }

    /// <summary>Runs the full validate + persist pipeline (<see cref="SaveEntry.Save{T}"/>).</summary>
    void Save();

    /// <summary>
    /// Fetches the dataset and returns whether it is empty, plus an action that populates the worksheet
    /// with it (null when nothing was fetched or the module has no populate routine).
    /// </summary>
    (bool IsEmpty, Action? Populate) FetchForPopulate();
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
    public Action<dynamic>? OnAfterDelayedLoad { get; set; }

    // =========================================================================
    // 2. SAVE / LOAD PIPELINE DELEGATES
    // =========================================================================
    public Func<(T? Data, List<string> Errors, List<string> Warnings)> ExtractAndValidate { get; set; } = null!;
    public Predicate<T?> IsDatasetEmpty { get; set; } = data => data == null;
    public Func<T, (bool Success, string ErrorMessage)> SaveToServer { get; set; } = null!;

    public Action<T>? PopulateWorksheet { get; set; }
    public Func<T?>? FetchFromServer { get; set; }
    public Action<bool>? SyncLookups { get; set; }
    public Predicate<T>? ValidateOfflinePayload { get; set; }
    public Func<T, T, string>? FormatConflictSummary { get; set; }

    // =========================================================================
    // 3. NON-GENERIC COMMAND SURFACE (ISaveContract)
    // =========================================================================
    void ISaveContract.Save() => SaveEntry.Save(this);

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
