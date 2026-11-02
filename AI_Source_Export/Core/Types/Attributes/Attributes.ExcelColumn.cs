namespace SKSB_App.Core.Types.Attributes;

using System;

[AttributeUsage(AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
public sealed class ExcelColumnAttribute : Attribute
{
    public string HeaderName { get; }
    public bool IsRequired { get; }

    public ExcelColumnAttribute(string headerName, bool isRequired = false)
    {
        HeaderName = headerName;
        IsRequired = isRequired;
    }
}