namespace SKSB_App.Portal;

using System;
using System.Windows.Forms;

static class Program
{
    [STAThread]
    static void Main()
    {
        // 1. Critical: Tell Windows to scale WinForms per display DPI
        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        // 2. Launch the portal
        Application.Run(new PortalForm());
    }
}