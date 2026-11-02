namespace SKSB_App.Core.RuntimeUI.DateFilter;

/// <summary>
/// Universal temporal filter scope for operational queries and worksheet views.
/// Automatically discovered by Sync.EnumList.SyncAllCoreEnums() into '=DateFilterScopeList'.
/// </summary>
public enum DateFilterScope
{
    All_Time,
    Annual,
    Month,
    Day,
    None
}