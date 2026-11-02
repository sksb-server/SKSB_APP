namespace SKSB_App.Core.Integration.Database;

using System;

public static class ServerEnvironment
{
#if DEBUG
    private static bool _simulateOffline = false;
#else
    private static readonly bool _simulateOffline = false;
#endif

    public static bool SimulateOffline
    {
        get => _simulateOffline;
        set
        {
#if DEBUG
            _simulateOffline = value;
#else
            throw new InvalidOperationException("Offline simulation is disabled in Release builds.");
#endif

        }
    }

    public static bool IsSimulatedOffline()
    {
#if DEBUG
        return SimulateOffline;
#else
        return false;
#endif
    }

    public static void ToggleSimulatedOffline(ExcelDna.Integration.CustomUI.IRibbonUI? ribbonUi, string controlId, bool pressed)
    {
#if DEBUG
        SimulateOffline = pressed;
        ribbonUi?.InvalidateControl(controlId);
        System.Windows.Forms.MessageBox.Show(
            pressed
                ? "Server is now SIMULATED OFFLINE.\nAll PostgreSQL database operations will fail immediately to test offline local stashing."
                : "Server is now ONLINE.\nNormal PostgreSQL database connectivity restored.",
            "Debug Server Mode",
            System.Windows.Forms.MessageBoxButtons.OK,
            pressed ? System.Windows.Forms.MessageBoxIcon.Warning : System.Windows.Forms.MessageBoxIcon.Information
        );
#endif
    }

    public static void GuardConnection()
    {
        if (SimulateOffline)
        {
            throw new TimeoutException("[Simulated Offline Mode Active] Connection to PostgreSQL server blocked by user switch.");
        }
    }
}