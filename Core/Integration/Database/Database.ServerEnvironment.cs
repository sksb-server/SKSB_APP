namespace SKSB_App.Core.Integration.Database;

using SKSB_App.Core.Diagnostics;
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

    public static void ToggleSimulatedOffline(ExcelDna.Integration.CustomUI.IRibbonUI? ribbonUi, string controlId, bool pressed)
    {
#if DEBUG
        SimulateOffline = pressed;
        ribbonUi?.InvalidateControl(controlId);
        AppLogger.Warn(pressed
            ? "[Debug Server Mode] Server is now SIMULATED OFFLINE: all PostgreSQL operations fail immediately to test offline local stashing."
            : "[Debug Server Mode] Server is now ONLINE: normal PostgreSQL connectivity restored.");
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