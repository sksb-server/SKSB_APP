namespace SKSB_App.Portal;

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using SKSB_App.Core.RuntimeUI.BugReport;
#if DEBUG
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using SKSB_App.Core.Localization;
#endif

static class Program
{
    private const string PortalMutexName = "Global\\SKSB_App_Portal_SingleInstance_Mutex";

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    [STAThread]
    static int Main(string[] args)
    {
#if DEBUG
        if (args.Length >= 2 && string.Equals(args[0], "--export-dict-state", StringComparison.OrdinalIgnoreCase))
        {
            return RunExportDictStateCli(args[1]);
        }

        if (args.Length >= 2 && string.Equals(args[0], "--submit-dict-suggestions", StringComparison.OrdinalIgnoreCase))
        {
            string? outJsonPath = args.Length >= 3 ? args[2] : null;
            return RunSubmitDictSuggestionsCli(args[1], outJsonPath);
        }

        if (TryRelaunchFromShadow()) return 0;
#endif
#if DEBUG
        string mutexName = GetPortalMutexName();
#else
        string mutexName = PortalMutexName;
#endif
        using var mutex = new Mutex(true, mutexName, out bool isInitialInstance);
        if (!isInitialInstance)
        {
            FocusExistingPortal();
            return 0;
        }

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        AutoBugReporter.InstallForPortal();
        Application.Run(new PortalForm());
        return 0;
    }

#if DEBUG
    private static readonly UTF8Encoding Utf8NoBom = new(encoderShouldEmitUTF8Identifier: false);

    private static void WriteCliOutputFile(string outJsonPath, string content)
    {
        string? dir = Path.GetDirectoryName(outJsonPath);
        if (!string.IsNullOrEmpty(dir))
            Directory.CreateDirectory(dir);
        File.WriteAllText(outJsonPath, content, Utf8NoBom);
    }

    private static int RunExportDictStateCli(string outJsonPath)
    {
        try
        {
            var export = CollectiveDictionaryRepository.ExportStateForReview();
            string json = JsonSerializer.Serialize(export, new JsonSerializerOptions { WriteIndented = true });
            WriteCliOutputFile(outJsonPath, json);
            return 0;
        }
        catch (Exception ex)
        {
            string errJson = JsonSerializer.Serialize(new { error = ex.Message });
            try { WriteCliOutputFile(outJsonPath, errJson); } catch { }
            return 1;
        }
    }

    private static int RunSubmitDictSuggestionsCli(string jsonPath, string? outJsonPath)
    {
        try
        {
            string json = File.ReadAllText(jsonPath, Encoding.UTF8);
            var options = new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true,
                ReadCommentHandling = JsonCommentHandling.Skip,
                AllowTrailingCommas = true
            };
            var items = JsonSerializer.Deserialize<List<CollectiveSuggestionInput>>(json, options)
                ?? new List<CollectiveSuggestionInput>();
            var result = CollectiveDictionaryRepository.SubmitAiSuggestions(items);
            string outputJson = JsonSerializer.Serialize(result);
            Console.WriteLine(outputJson);
            if (!string.IsNullOrWhiteSpace(outJsonPath))
                WriteCliOutputFile(outJsonPath, outputJson);
            return 0;
        }
        catch (Exception ex)
        {
            string errJson = JsonSerializer.Serialize(new { error = ex.Message });
            Console.WriteLine(errJson);
            if (!string.IsNullOrWhiteSpace(outJsonPath))
            {
                try { WriteCliOutputFile(outJsonPath, errJson); } catch { }
            }
            return 1;
        }
    }
    private static string GetPortalMutexName()
    {
        try
        {
            string root = SKSB_App.Core.Bootstrap.Storage.FolderRegistry.SKSB_App.GetPath();
            string normalized = Path.GetFullPath(root).TrimEnd('\\', '/').ToLowerInvariant();
            byte[] hash = System.Security.Cryptography.SHA256.HashData(System.Text.Encoding.UTF8.GetBytes(normalized));
            string shortHash = Convert.ToHexString(hash)[..12];
            return $"Global\\SKSB_App_Portal_Debug_{shortHash}_Mutex";
        }
        catch
        {
            return PortalMutexName;
        }
    }

    private static bool TryRelaunchFromShadow()
    {
        try
        {
            if (Debugger.IsAttached) return false;

            string shadowRoot = Path.Combine(Path.GetTempPath(), "SKSB_Debug_Shadow");
            string baseDir = AppContext.BaseDirectory;
            if (baseDir.StartsWith(shadowRoot, StringComparison.OrdinalIgnoreCase)) return false;

            string solutionRoot = SKSB_App.Core.Bootstrap.Storage.FolderRegistry.SKSB_App.GetPath();

            PortalRouter.CleanupUnusedShadowDirectories(shadowRoot, "Portal_*");

            string targetDir = Path.Combine(shadowRoot, $"Portal_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{Environment.ProcessId}");
            Directory.CreateDirectory(targetDir);

            foreach (string file in Directory.GetFiles(baseDir))
            {
                File.Copy(file, Path.Combine(targetDir, Path.GetFileName(file)), true);
            }

            File.WriteAllText(Path.Combine(targetDir, ".sksb_solution_root"), solutionRoot);
            Environment.SetEnvironmentVariable("SKSB_SOLUTION_ROOT", solutionRoot);

            var psi = new ProcessStartInfo
            {
                FileName = Path.Combine(targetDir, "Portal.exe"),
                WorkingDirectory = targetDir,
                UseShellExecute = false
            };
            return Process.Start(psi) != null;
        }
        catch
        {
            return false;
        }
    }
#endif

    private static void FocusExistingPortal()
    {
        Process current = Process.GetCurrentProcess();
        foreach (Process process in Process.GetProcessesByName(current.ProcessName))
        {
            if (process.Id != current.Id && process.MainWindowHandle != IntPtr.Zero)
            {
                ShowWindow(process.MainWindowHandle, SW_RESTORE);
                SetForegroundWindow(process.MainWindowHandle);
                break;
            }
        }
    }
}