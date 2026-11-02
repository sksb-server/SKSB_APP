namespace SKSB_App.Core.Bootstrap.Environment;

using System.IO;
using Microsoft.Win32;

public static class ExcelPathFinder
{
    /// <summary>
    /// Locates the EXCEL.EXE binary via App Paths registry key or standard Office install directories.
    /// </summary>
    public static string ResolveExcelExecutable()
    {
        using var key = Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe");
        if (key?.GetValue(null) is string appPath && File.Exists(appPath))
        {
            return appPath;
        }

        string[] standardLocations =
        {
            @"C:\Program Files\Microsoft Office\root\Office16\EXCEL.EXE",
            @"C:\Program Files (x86)\Microsoft Office\root\Office16\EXCEL.EXE",
            @"C:\Program Files\Microsoft Office\Office16\EXCEL.EXE",
            @"C:\Program Files\Microsoft Office\Office15\EXCEL.EXE"
        };

        foreach (var path in standardLocations)
        {
            if (File.Exists(path)) return path;
        }

        return "excel.exe";
    }
}
