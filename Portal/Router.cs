namespace SKSB_App.Portal;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Win32;
using SKSB_App.Core.Bootstrap.Environment;
using SKSB_App.Core.Bootstrap.Storage;

public static class PortalRouter
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("oleaut32.dll", PreserveSig = false)]
    private static extern void GetActiveObject(
        ref Guid rclsid,
        IntPtr pvReserved,
        [MarshalAs(UnmanagedType.IUnknown)] out object ppunk
    );

    [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = false)]
    private static extern void CLSIDFromProgID(
        [MarshalAs(UnmanagedType.LPWStr)] string progId,
        out Guid clsid
    );

    private delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool EnumWindows(EnumWindowsProc lpEnumFunc, IntPtr lParam);

    [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
    private static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, int nMaxCount);

    [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
    private static extern int GetWindowTextLength(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsIconic(IntPtr hWnd);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    private const int SW_RESTORE = 9;
    private static readonly TimeSpan LaunchGracePeriod = TimeSpan.FromSeconds(15);

    private static readonly object _launchLock = new();
    private static readonly Dictionary<string, (Process Process, DateTime LaunchedAtUtc)> _activeLaunches
        = new(StringComparer.OrdinalIgnoreCase);

    public static void Launch(TemplateAddInPair pair)
    {
        string excelPath = ExcelRegistry.ExcelPath;
        if (string.IsNullOrEmpty(excelPath) || !File.Exists(excelPath))
        {
            MessageBox.Show("Excel installation could not be detected on this workstation.",
                            "Missing Excel", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        try
        {
            lock (_launchLock)
            {
            string templateFileName = pair.Template();
            string templateBaseName = Path.GetFileNameWithoutExtension(templateFileName);

            // 1. Check if an Excel window with this specific workbook is already open
            if (TryFocusExistingWorkbook(templateBaseName))
            {
                return; // Module already active and brought to front
            }

            // 1b. Guard against a launch still booting (workbook window not yet visible)
            if (_activeLaunches.TryGetValue(templateBaseName, out var entry))
            {
                bool exited;
                try { exited = entry.Process.HasExited; }
                catch { exited = true; }

                if (!exited && (DateTime.UtcNow - entry.LaunchedAtUtc) < LaunchGracePeriod)
                {
                    try
                    {
                        entry.Process.Refresh();
                        if (entry.Process.MainWindowHandle != IntPtr.Zero)
                            FocusWindow(entry.Process.MainWindowHandle);
                    }
                    catch { }
                    return;
                }

                try { entry.Process.Dispose(); } catch { }
                _activeLaunches.Remove(templateBaseName);
            }

            // 2. Resolve Template & Target Add-In paths
            string templatePath = ResolveTemplatePath(templateFileName);
            string addInPath = TemplateAddInPair.ResolveAddInPath(pair.AddIn());

            // 3. Resolve Core Add-In path so it loads alongside every manager
            string coreAddInName = ExcelRegistry.OfficeBitness.Equals(Bitness.Win32)
                ? "Core-AddIn.xll"
                : "Core-AddIn64.xll";
            string coreAddInPath = TemplateAddInPair.ResolveAddInPath(coreAddInName);

            if (!File.Exists(templatePath))
                throw new FileNotFoundException($"Template file not found at:\n{templatePath}", templatePath);
            if (!File.Exists(addInPath))
                throw new FileNotFoundException($"Module Add-in not found at:\n{addInPath}", addInPath);
            if (!File.Exists(coreAddInPath))
                throw new FileNotFoundException($"Core Add-in not found at:\n{coreAddInPath}", coreAddInPath);

#if DEBUG
            (coreAddInPath, addInPath) = StageDebugShadowAddIns(coreAddInPath, addInPath, templateBaseName);
#endif

            // 4. Use /x to force an isolated Excel process, preventing add-in cross-contamination
            var startInfo = new ProcessStartInfo
            {
                FileName = excelPath,
                Arguments = $"/x /t \"{templatePath}\" \"{coreAddInPath}\" \"{addInPath}\"",
                UseShellExecute = true
            };

            Process? started = Process.Start(startInfo);
            if (started != null)
                _activeLaunches[templateBaseName] = (started, DateTime.UtcNow);
            }
        }
        catch (FileNotFoundException fnf)
        {
            MessageBox.Show($"File missing:\n{fnf.FileName}",
                            "Launch Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"Failed to dispatch module:\n{ex.Message}",
                            "Execution Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

#if DEBUG
    private static (string CoreAddInPath, string ModuleAddInPath) StageDebugShadowAddIns(
        string coreAddInPath, string moduleAddInPath, string moduleBaseName)
    {
        try
        {
            string shadowRoot = Path.Combine(Path.GetTempPath(), "SKSB_Debug_Shadow");
            string solutionRoot = FolderRegistry.SKSB_App.GetPath();
            Environment.SetEnvironmentVariable("SKSB_SOLUTION_ROOT", solutionRoot);
            Directory.CreateDirectory(shadowRoot);

            EnsureDebugTrustedLocations(shadowRoot, solutionRoot);

            // Best-effort sweep: folders in use by running Excel/Portal are locked and skipped.
            CleanupUnusedShadowDirectories(shadowRoot);

            string sessionDir = Path.Combine(shadowRoot,
                $"{moduleBaseName}_{DateTime.UtcNow:yyyyMMdd_HHmmss_fff}_{Guid.NewGuid().ToString("N")[..6]}");
            Directory.CreateDirectory(sessionDir);

            foreach (string srcDir in new[] { Path.GetDirectoryName(coreAddInPath)!, Path.GetDirectoryName(moduleAddInPath)! })
            {
                foreach (string file in Directory.GetFiles(srcDir))
                {
                    File.Copy(file, Path.Combine(sessionDir, Path.GetFileName(file)), overwrite: true);
                }
            }

            File.WriteAllText(Path.Combine(sessionDir, ".sksb_solution_root"), solutionRoot);

            return (Path.Combine(sessionDir, Path.GetFileName(coreAddInPath)),
                    Path.Combine(sessionDir, Path.GetFileName(moduleAddInPath)));
        }
        catch
        {
            return (coreAddInPath, moduleAddInPath);
        }
    }

    internal static void CleanupUnusedShadowDirectories(string shadowRoot, string searchPattern = "*")
    {
        if (!Directory.Exists(shadowRoot)) return;

        string[] dirs;
        try { dirs = Directory.GetDirectories(shadowRoot, searchPattern); }
        catch { return; }

        foreach (string dir in dirs)
        {
            try
            {
                // Deletion is not atomic: skip the whole directory if any file is in use.
                foreach (string file in Directory.EnumerateFiles(dir, "*", SearchOption.AllDirectories))
                {
                    using (File.Open(file, FileMode.Open, FileAccess.Read, FileShare.None)) { }
                }

                string trashDir = dir.EndsWith(".deleting", StringComparison.OrdinalIgnoreCase) ? dir : $"{dir}.deleting";
                if (!string.Equals(dir, trashDir, StringComparison.OrdinalIgnoreCase)) Directory.Move(dir, trashDir);
                Directory.Delete(trashDir, recursive: true);
            }
            catch { }
        }
    }

    private static void EnsureDebugTrustedLocations(string shadowRoot, string solutionRoot)
    {
        foreach (string ver in new[] { "16.0", "15.0" })
        {
            try
            {
                if (ver == "15.0")
                {
                    using var probe = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{ver}\Excel");
                    if (probe == null) continue;
                }

                string basePath = $@"Software\Microsoft\Office\{ver}\Excel\Security\Trusted Locations";
                RegisterTrustedLocation(basePath, "SKSB_App_DebugShadow", shadowRoot, "SKSB_App Debug Shadow");
                RegisterTrustedLocation(basePath, "SKSB_App_DebugRepo", solutionRoot, "SKSB_App Debug Repo");
            }
            catch { }
        }
    }

    private static void RegisterTrustedLocation(string basePath, string name, string path, string description)
    {
        using var key = Registry.CurrentUser.CreateSubKey($@"{basePath}\{name}");
        key?.SetValue("Path", path, RegistryValueKind.String);
        key?.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
        key?.SetValue("Description", description, RegistryValueKind.String);
    }
#endif

    private static string ResolveTemplatePath(string templateFileName)
    {
#if DEBUG
        // 1. Check Codes Root: Documents\codes\SKSB_App\Templates\<file>
        string codesTemplates = Path.Combine(FolderRegistry.SKSB_App.GetPath(), "Templates", templateFileName);
        if (File.Exists(codesTemplates)) return codesTemplates;

        // 2. Check explicit MyDocuments\codes\SKSB_App\Templates
        string directCodes = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
            "codes", "SKSB_App", "Templates", templateFileName);
        if (File.Exists(directCodes)) return directCodes;

        // 3. Check AppContext.BaseDirectory\Templates
        string binTemplates = Path.Combine(AppContext.BaseDirectory, "Templates", templateFileName);
        if (File.Exists(binTemplates)) return binTemplates;

        return codesTemplates;
#else
        // Release: Always Documents\SKSB_App\Templates\<file>
        return FolderRegistry.Templates.GetFilePath(templateFileName, mustExist: true);
#endif
    }

    private static bool TryFocusExistingWorkbook(string targetWorkbookName)
    {
        try
        {
            object? xlAppObj = GetActiveObjectHelper("Excel.Application");
            if (xlAppObj != null)
            {
                dynamic xlApp = xlAppObj;
                foreach (dynamic wb in xlApp.Workbooks)
                {
                    string wbName = wb.Name?.ToString() ?? string.Empty;
                    if (wbName.IndexOf(targetWorkbookName, StringComparison.OrdinalIgnoreCase) >= 0)
                    {
                        wb.Activate();
                        xlApp.Visible = true;
                        IntPtr hwnd = new IntPtr((int)xlApp.Hwnd);
                        if (hwnd != IntPtr.Zero)
                        {
                            FocusWindow(hwnd);
                            return true;
                        }
                    }
                }
            }
        }
        catch { }

        IntPtr foundHwnd = IntPtr.Zero;
        EnumWindows((hWnd, lParam) =>
        {
            if (!IsWindowVisible(hWnd)) return true;
            int length = GetWindowTextLength(hWnd);
            if (length > 0)
            {
                var sb = new StringBuilder(length + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                string title = sb.ToString();

                if (title.IndexOf(targetWorkbookName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    bool isExcel = false;
                    try
                    {
                        GetWindowThreadProcessId(hWnd, out uint pid);
                        isExcel = Process.GetProcessById((int)pid).ProcessName
                            .Equals("EXCEL", StringComparison.OrdinalIgnoreCase);
                    }
                    catch { }

                    if (isExcel)
                    {
                        foundHwnd = hWnd;
                        return false;
                    }
                }
            }
            return true;
        }, IntPtr.Zero);

        if (foundHwnd != IntPtr.Zero)
        {
            FocusWindow(foundHwnd);
            return true;
        }

        return false;
    }

    private static void FocusWindow(IntPtr hWnd)
    {
        if (hWnd == IntPtr.Zero) return;
        if (IsIconic(hWnd)) ShowWindow(hWnd, SW_RESTORE);
        SetForegroundWindow(hWnd);
    }

    private static object? GetActiveObjectHelper(string progId)
    {
        try
        {
            CLSIDFromProgID(progId, out Guid clsid);
            GetActiveObject(ref clsid, IntPtr.Zero, out object obj);
            return obj;
        }
        catch
        {
            return null;
        }
    }
}