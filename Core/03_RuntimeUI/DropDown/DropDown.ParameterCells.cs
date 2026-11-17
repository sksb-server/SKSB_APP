namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using SKSB_App.Core.Diagnostics;

/// <summary>
/// Sheet-change observers for selector cells that live OUTSIDE any ListObject (e.g. Project Summary!A2 / B2).
/// </summary>
public static class ParameterCells
{
    private static readonly List<Action<dynamic, dynamic>> _observers = new();
    private static readonly object _lock = new();

    // =========================================================================
    // 1. OBSERVER REGISTRATION
    // =========================================================================

    /// <summary>Registers a callback invoked (on the Excel event thread) after any sheet change. Idempotent per delegate.</summary>
    public static void RegisterChangeObserver(Action<dynamic, dynamic> observer)
    {
        lock (_lock)
        {
            if (!_observers.Contains(observer)) _observers.Add(observer);
        }
    }

    public static void UnregisterChangeObserver(Action<dynamic, dynamic> observer)
    {
        lock (_lock) { _observers.Remove(observer); }
    }

    // =========================================================================
    // 2. CHANGE NOTIFICATION
    // =========================================================================

    /// <summary>Fans a SheetChange out to every observer. Observers must queue their own COM mutations.</summary>
    public static void NotifyChanged(dynamic sheet, dynamic target)
    {
        Action<dynamic, dynamic>[] snapshot;
        lock (_lock) { snapshot = _observers.ToArray(); }

        foreach (var observer in snapshot)
        {
            try { observer(sheet, target); }
            catch (Exception ex) { AppLogger.Error("[ParameterCells] Change observer failed", ex); }
        }
    }
}
