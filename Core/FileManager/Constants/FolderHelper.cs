using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using SKSB_App.Core.ExcelManager;

namespace SKSB_App.Core.FileManager.Constants;

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

    /// <summary>
    /// Resolves file path. If 'mustExist' is true, enforces that the file is on disk.
    /// </summary>
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
        return Context.OfficeBitness.Equals(Bitness.Win32)
            ? $"{_addIn}.xll"
            : $"{_addIn}64.xll";
    }
    public static implicit operator string(TemplateAddInPair pair) => nameof(pair) ?? string.Empty;
}