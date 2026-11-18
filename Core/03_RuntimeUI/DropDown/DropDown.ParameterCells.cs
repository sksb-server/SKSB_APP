namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;

/// <summary>
/// Sheet-change observers for selector cells that live OUTSIDE any ListObject (e.g. Project Summary!A2 / B2).
/// The observer list lives in <see cref="AppDomain"/> data (BCL types only) so every .xll - each loaded in its own
/// AssemblyLoadContext with its own copy of Core.dll - shares one list.
/// </summary>
public static class ParameterCells
{
    private const string SharedKey = "SKSB_PARAMETER_CELL_OBSERVERS";

    /// <summary>Identifies this Core.dll copy (one per .xll ALC).</summary>
    private static readonly Guid ScopeId = Guid.NewGuid();

    private static List<(Guid ScopeId, Action<object, object> Observer)> Shared()
    {
        lock (AppDomain.CurrentDomain)
        {
            if (AppDomain.CurrentDomain.GetData(SharedKey) is List<(Guid, Action<object, object>)> existing) return existing;
            var created = new List<(Guid ScopeId, Action<object, object> Observer)>();
            AppDomain.CurrentDomain.SetData(SharedKey, created);
            return created;
        }
    }

    // =========================================================================
    // 1. OBSERVER REGISTRATION
    // =========================================================================

    /// <summary>Registers a callback invoked (on the Excel event thread) after any sheet change. Idempotent per (scope, delegate).</summary>
    public static void RegisterChangeObserver(Action<dynamic, dynamic> observer)
    {
        Action<object, object> o = observer;
        var list = Shared();
        lock (list)
        {
            if (!list.Any(e => e.ScopeId == ScopeId && e.Observer.Equals(o))) list.Add((ScopeId, o));
        }
    }

    public static void UnregisterChangeObserver(Action<dynamic, dynamic> observer)
    {
        Action<object, object> o = observer;
        var list = Shared();
        lock (list) { list.RemoveAll(e => e.ScopeId == ScopeId && e.Observer.Equals(o)); }
    }

    // =========================================================================
    // 2. CHANGE NOTIFICATION
    // =========================================================================

    /// <summary>Fans a SheetChange out to every observer from every .xll. Observers must queue their own COM mutations.</summary>
    public static void NotifyChanged(dynamic sheet, dynamic target)
    {
        var list = Shared();
        Action<object, object>[] snapshot;
        lock (list) { snapshot = list.Select(e => e.Observer).ToArray(); }

        foreach (var observer in snapshot)
        {
            try { observer((object)sheet, (object)target); }
            catch (Exception ex) { AppLogger.Error("[ParameterCells] Change observer failed", ex); }
        }
    }
}
