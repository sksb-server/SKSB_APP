namespace SKSB_App.Core.Bootstrap.Environment;


public static class ExcelRegistry
{
    private static string _excelPath;
    private static Bitness _officeBitness;

    static ExcelRegistry()
    {
        _excelPath = ExcelPathFinder.ResolveExcelExecutable();
        _officeBitness = BinaryInspector.GetExcelBitness(_excelPath);
    }

    public static string ExcelPath => _excelPath;

    public static Bitness OfficeBitness => _officeBitness;
}