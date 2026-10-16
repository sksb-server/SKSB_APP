namespace SKSB_App.Portal;

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using SKSB_App.Core.ExcelBitrateManager;
using SKSB_App.Core.FileManager.Structure;

public static class PortalRouter
{
    [DllImport("user32.dll")]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    // .NET 8 P/Invoke replacement for Marshal.GetActiveObject
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

    private const int SW_RESTORE = 9;

    public static void Launch(TemplateAddInPair pair)
    {
        string excelPath = Context.ExcelPath;
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

            // 1. Check if an Excel window with this specific workbook is already open
            if (TryFocusExistingWorkbook(templateBaseName))
            {
                return; // Module already active and brought to front
            }

            // 2. Not already open: dispatch Excel with the requested template and companion add-in
            string templatePath = Folder.Templates.GetFilePath(templateFileName, mustExist: true);
            string addInPath = Folder.AddIns.GetFilePath(pair.AddIn(), mustExist: true);

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

    /// <summary>
    /// Searches active Excel instances for a workbook whose name matches the target template.
    /// If found, activates the workbook and brings its window to the front.
    /// </summary>
    private static bool TryFocusExistingWorkbook(string targetWorkbookName)
    {
        try
        {
            // Query running Excel COM instance via OLEAUT32 in .NET 8
            object? xlAppObj = GetActiveObjectHelper("Excel.Application");
            if (xlAppObj == null) return false;

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
                    }

                    return true;
                }
            }
        }
        catch
        {
            // Throws if Excel is not running or COM ROT is unavailable
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