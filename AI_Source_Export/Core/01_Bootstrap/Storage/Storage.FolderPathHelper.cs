namespace SKSB_App.Core.Bootstrap.Storage;

using System;
using System.IO;
using SKSB_App.Core.Bootstrap.Environment;

public sealed class AppFolder
{
    private readonly string _path;

    public AppFolder(string path)
    {
        _path = path ?? throw new ArgumentNullException(nameof(path));
    }

    public string GetPath() => _path;
    public bool Exists() => Directory.Exists(_path);

    public string EnsureExists()
    {
        if (!Directory.Exists(_path))
        {
            Directory.CreateDirectory(_path);
        }
        return _path;
    }

    public string GetFilePath(string fileName, bool mustExist = false)
    {
        if (string.IsNullOrWhiteSpace(fileName))
            throw new ArgumentException("Filename cannot be empty.", nameof(fileName));

        string fullPath = Path.Combine(_path, fileName);
        if (mustExist && !File.Exists(fullPath))
        {
            throw new FileNotFoundException($"Required operational file was not found at '{fullPath}'.", fullPath);
        }
        return fullPath;
    }

    public override string ToString() => _path;
    public static implicit operator string(AppFolder folder) => folder?._path ?? string.Empty;
}

public sealed class TemplateAddInPair
{
    private readonly string _template;
    private readonly string _addIn;

    public TemplateAddInPair(string template, string addIn)
    {
        _template = template ?? throw new ArgumentNullException(nameof(template));
        _addIn = addIn ?? throw new ArgumentNullException(nameof(addIn));
    }

    public string Template() => _template;

    public string AddIn()
    {
        return ExcelRegistry.OfficeBitness.Equals(Bitness.Win32)
            ? $"{_addIn}.xll"
            : $"{_addIn}64.xll";
    }

    public static string ResolveAddInPath(string addInFileName)
    {
#if DEBUG
        // 1. Current executing project base directory
        string localBinPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, addInFileName);
        if (File.Exists(localBinPath)) return localBinPath;

        // 2. Sibling module build directory: codes\SKSB_App\<Module>\bin\Debug\net8.0-windows\<addInFileName>
        string moduleName = Path.GetFileNameWithoutExtension(addInFileName)
                                .Replace("-AddIn64", "")
                                .Replace("-AddIn", "");

        string siblingPath = Path.Combine(
            FolderRegistry.SKSB_App.GetPath(),
            moduleName,
            "bin", "Debug", "net8.0-windows",
            addInFileName
        );
        if (File.Exists(siblingPath)) return siblingPath;

        // 3. Packed fallback: publish\<addIn>-packed.xll
        string baseName = Path.GetFileNameWithoutExtension(addInFileName);
        string packedPath = Path.Combine(
            FolderRegistry.SKSB_App.GetPath(),
            moduleName,
            "bin", "Debug", "net8.0-windows",
            "publish",
            $"{baseName}-packed.xll"
        );
        if (File.Exists(packedPath)) return packedPath;

        return siblingPath;
#else
        // In Release: Strictly use the deployed Documents\SKSB_App\AddIns folder
        return Path.Combine(FolderRegistry.AddIns.GetPath(), addInFileName);
#endif
    }

    public static implicit operator string(TemplateAddInPair pair) => nameof(pair) ?? string.Empty;
}