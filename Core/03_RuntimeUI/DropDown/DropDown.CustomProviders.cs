namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;

/// <summary>What a module-supplied dropdown shows for one cell and how a pick is written back.</summary>
/// <param name="Options">Items in display order (use <see cref="DropdownOption.Section"/> as the rank so the order is preserved).</param>
/// <param name="Commit">Writes the chosen option into the cell. Runs with EnableEvents already suppressed.</param>
public sealed record CustomDropdownResult(IReadOnlyList<DropdownOption> Options, Action<dynamic, DropdownOption> Commit);

/// <summary>
/// Extension point for dropdowns that are not driven by a ListObject column rule (e.g. the Schedule board,
/// whose staff cells are ranked by match-% at open time). Providers are consulted before the attribute rule engine.
/// </summary>
public static class CustomDropdowns
{
    private sealed record Provider(Func<dynamic, bool> AppliesTo, Func<dynamic, CustomDropdownResult?> Build);

    private static readonly List<Provider> _providers = new();
    private static readonly object _lock = new();

    public static void Register(Func<dynamic, bool> appliesTo, Func<dynamic, CustomDropdownResult?> build)
    {
        lock (_lock) { _providers.Add(new Provider(appliesTo, build)); }
    }

    public static void Clear()
    {
        lock (_lock) { _providers.Clear(); }
    }

    public static bool AppliesTo(dynamic singleCell)
    {
        Provider[] snapshot;
        lock (_lock) { snapshot = _providers.ToArray(); }
        foreach (var p in snapshot)
        {
            try { if (p.AppliesTo(singleCell)) return true; }
            catch (Exception ex) { AppLogger.Warn($"[CustomDropdowns] AppliesTo failed: {ex.Message}"); }
        }
        return false;
    }

    public static CustomDropdownResult? TryBuild(dynamic singleCell)
    {
        Provider[] snapshot;
        lock (_lock) { snapshot = _providers.ToArray(); }
        foreach (var p in snapshot)
        {
            try
            {
                if (!p.AppliesTo(singleCell)) continue;
                var result = p.Build(singleCell);
                if (result != null) return result;
            }
            catch (Exception ex) { AppLogger.Error("[CustomDropdowns] Build failed", ex); }
        }
        return null;
    }
}
