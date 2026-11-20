namespace SKSB_App.Portal;

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using SKSB_App.Core.RuntimeUI.BugReport;

static class Program
{
    private const string PortalMutexName = "Global\\SKSB_App_Portal_SingleInstance_Mutex";

    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    private const int SW_RESTORE = 9;

    [STAThread]
    static void Main()
    {
#if DEBUG
        if (TryRelaunchFromShadow()) return;
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
            return;
        }

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);
        AutoBugReporter.InstallForPortal();
        Application.Run(new PortalForm());
    }

#if DEBUG
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