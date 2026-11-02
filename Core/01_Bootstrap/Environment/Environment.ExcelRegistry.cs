namespace SKSB_App.Core.Bootstrap.Environment;

using System;

public static class ExcelRegistry
{
    private static string _excelPath;
    private static Bitness _officeBitness;

    static ExcelRegistry()
    {
        _excelPath = ExcelPathFinder.ResolveExcelExecutable();
        _officeBitness = BinaryInspector.GetExcelBitness(_excelPath);
    }

    public static string ExcelPath
    {
        get => _excelPath;
        set
        {
            if (string.IsNullOrWhiteSpace(value))
                throw new ArgumentException("Excel path cannot be null or empty.", nameof(value));

            _excelPath = value;
            _officeBitness = BinaryInspector.GetExcelBitness(value);
        }
    }

    public static Bitness OfficeBitness => _officeBitness;
}