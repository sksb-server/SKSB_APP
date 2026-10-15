namespace SKSB_App.Core.FileManager.Constants;

using System;
using System.IO;

public static class Folder
{
    // The canonical user directory in Documents
    private static readonly string DocumentsAppPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
        "SKSB_App"
    );

#if DEBUG
    // 1. Dynamic resolver function instead of hardcoded relative path
    private static readonly string SolutionRoot = ResolveDebugSolutionRoot();

    // Prefer Documents if present, otherwise fall back to the Solution Root
    private static readonly string RootPath = Directory.Exists(DocumentsAppPath)
        ? DocumentsAppPath
        : SolutionRoot;

    public static AppFolder SKSB_App { get; } = new AppFolder(RootPath);
    public static AppFolder Templates { get; } = new AppFolder(ResolveFolder("Templates"));
    public static AppFolder AddIns { get; } = new AppFolder(ResolveFolder("AddIns"));
    public static AppFolder Data { get; } = new AppFolder(ResolveFolder("Data"));

#else
    // Production / Release
    private static readonly string LocalAppPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SKSB_App"
    );

    public static AppFolder SKSB_App { get; } = new AppFolder(LocalAppPath);
    public static AppFolder AddIns { get; } = new AppFolder(Path.Combine(LocalAppPath, "AddIns"));
    public static AppFolder Templates { get; } = new AppFolder(Path.Combine(DocumentsAppPath, "Templates"));
    public static AppFolder Data { get; } = new AppFolder(Path.Combine(DocumentsAppPath, "Data"));
#endif

    private static string ResolveFolder(string subFolder)
    {
        // Check Documents\SKSB_App\<subFolder> first
        string docPath = Path.Combine(DocumentsAppPath, subFolder);
        if (Directory.Exists(docPath))
        {
            return docPath;
        }

        // Check <SolutionRoot>\<subFolder> second
        string slnPath = Path.Combine(SolutionRoot, subFolder);
        if (Directory.Exists(slnPath))
        {
            return slnPath;
        }

        // Default to Solution directory
        return slnPath;
    }

    private static string ResolveDebugSolutionRoot()
    {
        string? startDir = Path.GetDirectoryName(typeof(Folder).Assembly.Location);
        if (string.IsNullOrEmpty(startDir) || startDir.Contains("Office", StringComparison.OrdinalIgnoreCase))
        {
            startDir = AppContext.BaseDirectory;
        }

        var dir = new DirectoryInfo(startDir);

        while (dir != null)
        {
            if (dir.GetFiles("*.sln").Length > 0 || dir.GetDirectories(".git").Length > 0)
            {
                return dir.FullName;
            }
            dir = dir.Parent;
        }

        return DocumentsAppPath;
    }
}

public static class Window
{
    // Define with new or unified project names
    public static TemplateAddInPair DailyScheduler { get; } = CreateDefault("DailyScheduler");
    public static TemplateAddInPair WorkerManager { get; } = CreateDefault("WorkerManager");
    public static TemplateAddInPair ProjectManager { get; } = CreateDefault("ProjectManager");
    public static TemplateAddInPair VehicleManager { get; } = CreateDefault("VehicleManager");

    /// <summary>
    /// Dynamically constructs the template and add-in filenames from the base module name.
    /// Example: "WorkerManager" -> Template: "Worker_Manager_Template.xltx", AddIn: "WorkerManager-AddIn"
    /// </summary>
    public static TemplateAddInPair CreateDefault(string moduleName)
    {
        // Converts PascalCase to Snake_Case (e.g. WorkerManager -> Worker_Manager)
        string snake = System.Text.RegularExpressions.Regex.Replace(moduleName, "([a-z])([A-Z])", "$1_$2");
        return new TemplateAddInPair(
            template: $"{snake}_Template.xltx",
            addIn: $"{moduleName}-AddIn"
        );
    }
}