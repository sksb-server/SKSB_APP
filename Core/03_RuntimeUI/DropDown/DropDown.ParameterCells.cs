namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;

/// <summary>
/// Code-registered dropdown cells that live OUTSIDE any ListObject (e.g. the selector cells at the top of a
/// sheet), plus sheet-change observers. Replaces per-cell "CFG_" defined names for modules that need
/// dynamic item lists and change callbacks (Project Summary!A2 / B2).
/// </summary>
public static class ParameterCells
{
    private sealed record Binding(string SheetName, string Address, Func<dynamic, IEnumerable<string>> ItemsProvider);

    private static readonly List<Binding> _bindings = new();
    private static readonly List<Action<dynamic, dynamic>> _observers = new();
    private static readonly object _lock = new();

    // =========================================================================
    // 1. REGISTRATION
    // =========================================================================

    /// <summary>Registers a dropdown on a single cell. <paramref name="address"/> uses A1 style, e.g. "A2" or "$A$2".</summary>
    public static void RegisterDropdown(string sheetName, string address, Func<dynamic, IEnumerable<string>> itemsProvider)
    {
        string normalized = NormalizeAddress(address);
        lock (_lock)
        {
            _bindings.RemoveAll(b =>
                b.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase) &&
                b.Address.Equals(normalized, StringComparison.OrdinalIgnoreCase));
            _bindings.Add(new Binding(sheetName, normalized, itemsProvider));
        }
    }

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
    // 2. LOOKUP
    // =========================================================================

    public static bool IsParameterCell(dynamic singleCell) => Find(singleCell) != null;

    /// <summary>Returns the registered items for <paramref name="singleCell"/>, or null if it is not a parameter cell.</summary>
    public static List<string>? GetItems(dynamic singleCell, dynamic wb)
    {
        var binding = Find(singleCell);
        if (binding == null) return null;

        try
        {
            IEnumerable<string> raw = binding.ItemsProvider((object)wb);
            return raw
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[ParameterCells] Item provider for {binding.SheetName}!{binding.Address} failed: {ex.Message}");
            return new List<string>();
        }
    }

    // =========================================================================
    // 3. CHANGE NOTIFICATION
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

    // =========================================================================
    // 4. HELPERS
    // =========================================================================

    private static Binding? Find(dynamic singleCell)
    {
        if (singleCell == null) return null;

        Binding[] snapshot;
        lock (_lock) { snapshot = _bindings.ToArray(); }
        if (snapshot.Length == 0) return null;

        try
        {
            string sheetName = singleCell.Worksheet.Name?.ToString() ?? string.Empty;
            string address = NormalizeAddress(singleCell.Address?.ToString() ?? string.Empty);
            return snapshot.FirstOrDefault(b =>
                b.SheetName.Equals(sheetName, StringComparison.OrdinalIgnoreCase) &&
                b.Address.Equals(address, StringComparison.OrdinalIgnoreCase));
        }
        catch
        {
            return null;
        }
    }

    private static string NormalizeAddress(string address) => (address ?? string.Empty).Replace("$", string.Empty).Trim();
}
