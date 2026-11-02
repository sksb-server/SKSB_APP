namespace SKSB_App.Core.Bootstrap.Storage;

using System;
using System.IO;
using System.Text.RegularExpressions;

public static class FolderRegistry
{
    // Canonical deployed runtime path for clients / Release
    private static readonly string ReleaseRootPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "SKSB_App"
    );

    // Development source repository root for Debug
    private static readonly string CodesRootPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "codes",
        "SKSB_App"
    );

#if DEBUG
    // In Debug: Root is strictly the active solution codebase root (Documents\codes\SKSB_App)
    private static readonly string SolutionRoot = ResolveDebugSolutionRoot();
    public static AppFolder SKSB_App { get; } = new AppFolder(SolutionRoot);
    public static AppFolder Templates { get; } = new AppFolder(ResolveFolder("Templates"));
    public static AppFolder AddIns { get; } = new AppFolder(ResolveFolder("AddIns"));
    public static AppFolder Data { get; } = new AppFolder(ResolveFolder("Data"));
#else
    // In Release: Root is strictly the client deployment directory (Documents\SKSB_App)
    public static AppFolder SKSB_App { get; } = new AppFolder(ReleaseRootPath);
    public static AppFolder AddIns { get; } = new AppFolder(Path.Combine(ReleaseRootPath, "AddIns"));
    public static AppFolder Templates { get; } = new AppFolder(Path.Combine(ReleaseRootPath, "Templates"));
    public static AppFolder Data { get; } = new AppFolder(Path.Combine(ReleaseRootPath, "Data"));
#endif

    private static string ResolveFolder(string subFolder)
    {
#if DEBUG
        if (!string.IsNullOrEmpty(SolutionRoot) && Directory.Exists(SolutionRoot))
        {
            string slnPath = Path.Combine(SolutionRoot, subFolder);
            if (Directory.Exists(slnPath)) return slnPath;
        }

        string directCodesPath = Path.Combine(CodesRootPath, subFolder);
        if (Directory.Exists(directCodesPath)) return directCodesPath;

        return Path.Combine(AppContext.BaseDirectory, subFolder);
#else
        return Path.Combine(ReleaseRootPath, subFolder);
#endif
    }

#if DEBUG
    private static string ResolveDebugSolutionRoot()
    {
        try
        {
            string? startDir = AppContext.BaseDirectory;
            if (!string.IsNullOrWhiteSpace(startDir))
            {
                var dir = new DirectoryInfo(startDir);
                while (dir != null)
                {
                    if (dir.GetFiles("*.sln").Length > 0 || dir.GetDirectories(".git").Length > 0)
                    {
                        return dir.FullName;
                    }
                    dir = dir.Parent;
                }
            }
        }
        catch { }

        // Hard fallback to codes repository root
        if (Directory.Exists(CodesRootPath))
        {
            return CodesRootPath;
        }

        return AppContext.BaseDirectory;
    }
#endif
}

public static class ModuleRegistry
{
    public static TemplateAddInPair DailyScheduler { get; } = CreateDefault("DailyScheduler");
    public static TemplateAddInPair StaffManager { get; } = CreateDefault("StaffManager");
    public static TemplateAddInPair ProjectManager { get; } = CreateDefault("ProjectManager");
    public static TemplateAddInPair VehicleManager { get; } = CreateDefault("VehicleManager");

    public static TemplateAddInPair CreateDefault(string moduleName)
    {
        string snake = Regex.Replace(moduleName, "([a-z])([A-Z])", "$1_$2");
        return new TemplateAddInPair(
            template: $"{snake}_Template.xltx",
            addIn: $"{moduleName}-AddIn"
        );
    }
}