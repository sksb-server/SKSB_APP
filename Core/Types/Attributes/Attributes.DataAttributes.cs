namespace SKSB_App.Core.Types.Attributes;

using System;

/// <summary>
/// Marks a model property to be exposed as an Excel Defined Name list (e.g. [PropertyName]List) on the 'Data' tab.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class DataListAttribute : Attribute
{
    public string? CustomListName { get; }

    public DataListAttribute(string? customListName = null)
    {
        CustomListName = customListName;
    }
}

/// <summary>
/// Marks a child record (e.g. StaffDetails, ProjectDetails, VehicleDetail)
/// so the reflection scanner inspects its inner [DataList] properties.
/// </summary>
[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class DataDetailAttribute : Attribute
{
}