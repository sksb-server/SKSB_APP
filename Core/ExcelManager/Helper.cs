using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace SKSB_App.Core.ExcelManager;

public enum Bitness
{
    Unknown,
    Win32,
    Win64
}

public static class ExcelHelper
{
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

public static class BinaryInspector
{
    private const ushort PE_SIGNATURE_OFFSET = 0x3C;
    private const uint PE_SIGNATURE = 0x00004550; // "PE\0\0"
    private const ushort IMAGE_FILE_MACHINE_I386 = 0x014C;
    private const ushort IMAGE_FILE_MACHINE_AMD64 = 0x8664;

    /// <summary>
    /// Inspects the Portable Executable (PE) header of EXCEL.EXE directly.
    /// Eliminates registry lookup inaccuracies.
    /// </summary>
    public static Bitness GetExcelBitness(string excelExePath)
    {
        try
        {
            using var stream = new FileStream(excelExePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            using var reader = new BinaryReader(stream);

            // Read DOS Header to locate the PE Header pointer
            if (stream.Length < 64) return Bitness.Unknown;
            stream.Seek(PE_SIGNATURE_OFFSET, SeekOrigin.Begin);
            int peHeaderOffset = reader.ReadInt32();

            if (peHeaderOffset + 24 > stream.Length) return Bitness.Unknown;

            // Verify PE Signature
            stream.Seek(peHeaderOffset, SeekOrigin.Begin);
            uint signature = reader.ReadUInt32();
            if (signature != PE_SIGNATURE) return Bitness.Unknown;

            // Read Machine field from COFF File Header (immediately after signature)
            ushort machine = reader.ReadUInt16();

            return machine switch
            {
                IMAGE_FILE_MACHINE_AMD64 => Bitness.Win64,
                IMAGE_FILE_MACHINE_I386 => Bitness.Win32,
                _ => Bitness.Unknown
            };
        }
        catch
        {
            return Bitness.Unknown;
        }
    }
}

