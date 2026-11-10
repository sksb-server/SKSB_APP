namespace SKSB_App.Core.Types.Attributes;

using System;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false)]
public sealed class ColumnFilterAttribute : Attribute
{
    public string HeaderName { get; }
    public string SourceList { get; }

    /// <summary>
    /// Creates a UI-only filter column that does not map to database storage.
    /// Used to narrow down options in a primary mapped column (e.g., picking a Vehicle Type to filter Vehicle Plates).
    /// </summary>
    /// <param name="headerName">The column header name in Excel.</param>
    /// <param name="sourceList">The Named Range to populate the dropdown (e.g., "VehicleTypeList").</param>
    public ColumnFilterAttribute(string headerName, string sourceList)
    {
        HeaderName = headerName;
        SourceList = sourceList;
    }
}