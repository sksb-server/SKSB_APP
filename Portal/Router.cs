using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using SKSB_App.Core.ExcelManager;
using SKSB_App.Core.FileManager.Constants;

namespace SKSB_App.Portal;

public static class PortalRouter
{
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
            string templatePath = Folder.Templates.GetFilePath(pair.Template(), mustExist: true);
            string addInPath = Folder.AddIns.GetFilePath(pair.AddIn(), mustExist: true);

            var startInfo = new ProcessStartInfo
            {
                FileName = excelPath,
                // /r opens read-only template instance and loads companion .xll
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
}