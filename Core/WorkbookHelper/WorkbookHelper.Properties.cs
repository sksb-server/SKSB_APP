namespace SKSB_App.Core.WorkbookHelper;

using System;
using System.Collections.Generic;

/// <summary>
/// Defines how a specific manager workbook extracts, validates, saves, and populates its domain data.
/// </summary>
public class SyncDescriptor<T>
{
    public string ModuleName { get; set; } = "Manager";
    public string ServerTableName { get; set; } = string.Empty;

    public Func<(T? Data, List<string> Errors, List<string> Warnings)> ExtractAndValidate { get; set; } = null!;
    public Predicate<T?> IsDatasetEmpty { get; set; } = data => data == null;
    public Func<T, (bool Success, string ErrorMessage)> SaveToServer { get; set; } = null!;
    public Action<T>? PopulateWorksheet { get; set; }
    public Func<T?>? FetchFromServer { get; set; }
    public Action<bool>? SyncLookups { get; set; }
    public Predicate<T>? ValidateOfflinePayload { get; set; }
    public Func<T, T, string>? FormatConflictSummary { get; set; }
}