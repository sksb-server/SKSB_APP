namespace SKSB_App.Core.Bootstrap.Environment;

using System;
using System.IO;

public enum Bitness
{
    Unknown,
    Win32,
    Win64
}

public static class BinaryInspector
{
    private const ushort PE_SIGNATURE_OFFSET = 0x3C;
    private const uint PE_SIGNATURE = 0x00004550; // "PE\0\0"
    private const ushort IMAGE_FILE_MACHINE_I386 = 0x014C;
    private const ushort IMAGE_FILE_MACHINE_AMD64 = 0x8664;

    /// <summary>
    /// Inspects the Portable Executable (PE) header of EXCEL.EXE directly.
    /// Eliminates registry lookup inaccuracies and side-by-side Click-to-Run false positives.
    /// </summary>
    public static Bitness GetExcelBitness(string excelExePath)
    {
        try
        {
            if (string.IsNullOrWhiteSpace(excelExePath) || !File.Exists(excelExePath))
            {
                return Bitness.Unknown;
            }

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