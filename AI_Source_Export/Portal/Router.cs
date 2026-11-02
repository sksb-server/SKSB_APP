namespace SKSB_App.Portal;

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
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

    private const int SW_RESTORE = 9;

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
            string templateFileName = pair.Template();
            string templateBaseName = Path.GetFileNameWithoutExtension(templateFileName);

            // 1. Bring existing window to front if already running
            if (TryFocusExistingWorkbook(templateBaseName))
            {
                return;
            }

            // 2. Resolve Template & Add-In paths
            string templatePath = ResolveTemplatePath(templateFileName);
            string addInPath = TemplateAddInPair.ResolveAddInPath(pair.AddIn());

            if (!File.Exists(templatePath))
            {
                throw new FileNotFoundException($"Template file not found at:\n{templatePath}", templatePath);
            }

            if (!File.Exists(addInPath))
            {
                throw new FileNotFoundException($"Add-In file not found at:\n{addInPath}", addInPath);
            }

            var startInfo = new ProcessStartInfo
            {
                FileName = excelPath,
                Arguments = $"/t \"{templatePath}\" \"{addInPath}\"",
                UseShellExecute = true
            };

            Process.Start(startInfo);
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
                            ShowWindow(hwnd, SW_RESTORE);
                            SetForegroundWindow(hwnd);
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
            int length = GetWindowTextLength(hWnd);
            if (length > 0)
            {
                var sb = new StringBuilder(length + 1);
                GetWindowText(hWnd, sb, sb.Capacity);
                string title = sb.ToString();

                if (title.IndexOf("Excel", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    title.IndexOf(targetWorkbookName, StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    foundHwnd = hWnd;
                    return false;
                }
            }
            return true;
        }, IntPtr.Zero);

        if (foundHwnd != IntPtr.Zero)
        {
            ShowWindow(foundHwnd, SW_RESTORE);
            SetForegroundWindow(foundHwnd);
            return true;
        }

        return false;
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