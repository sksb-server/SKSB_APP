namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using System.Linq;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;

/// <summary>What a module-supplied dropdown shows for one cell and how a pick is written back.</summary>
/// <param name="Options">Items in display order (use <see cref="DropdownOption.Section"/> as the rank so the order is preserved).</param>
/// <param name="Commit">Writes the chosen option into the cell. Runs with EnableEvents already suppressed.</param>
public sealed record CustomDropdownResult(IReadOnlyList<DropdownOption> Options, Action<dynamic, DropdownOption> Commit);

/// <summary>
/// Extension point for dropdowns that are not driven by a ListObject column rule (e.g. the Schedule board,
/// whose staff cells are ranked by match-% at open time). Providers are consulted before the attribute rule engine.
/// The provider list lives in <see cref="AppDomain"/> data and only holds BCL types, because every .xll loads its own
/// copy of Core.dll (distinct <see cref="CustomDropdownResult"/> / <see cref="DropdownOption"/> types per ALC).
/// </summary>
public static class CustomDropdowns
{
    private const string SharedKey = "SKSB_CUSTOM_DROPDOWN_PROVIDERS";

    /// <summary>Identifies this Core.dll copy (one per .xll ALC).</summary>
    private static readonly Guid ScopeId = Guid.NewGuid();

    private static List<(Guid ScopeId, Func<object, bool> AppliesTo,
        Func<object, (List<(string Display, string Raw, int Section, bool IsValid)> Options, Action<object, string, string> Commit)?> Build)> Shared()
    {
        lock (AppDomain.CurrentDomain)
        {
            if (AppDomain.CurrentDomain.GetData(SharedKey) is List<(Guid, Func<object, bool>,
                Func<object, (List<(string, string, int, bool)>, Action<object, string, string>)?>)> existing) return existing;

            var created = new List<(Guid ScopeId, Func<object, bool> AppliesTo,
                Func<object, (List<(string Display, string Raw, int Section, bool IsValid)> Options, Action<object, string, string> Commit)?> Build)>();
            AppDomain.CurrentDomain.SetData(SharedKey, created);
            return created;
        }
    }

    public static void Register(Func<dynamic, bool> appliesTo, Func<dynamic, CustomDropdownResult?> build)
    {
        Func<object, bool> bclApplies = cell => appliesTo(cell);

        Func<object, (List<(string Display, string Raw, int Section, bool IsValid)> Options, Action<object, string, string> Commit)?> bclBuild = cell =>
        {
            CustomDropdownResult? result = build(cell);
            if (result == null) return null;

            var options = result.Options.Select(o => (o.Display, o.Raw, o.Section, o.IsValid)).ToList();
            Action<object, string, string> commit = (c, raw, display) =>
            {
                // Prefer the original option (keeps Section / IsValid); fall back to a minimal one.
                var match = result.Options.FirstOrDefault(o => o.Raw == raw && o.Display == display)
                            ?? new DropdownOption { Display = display, Raw = raw };
                result.Commit(c, match);
            };
            return (options, commit);
        };

        var list = Shared();
        lock (list) { list.Add((ScopeId, bclApplies, bclBuild)); }
    }

    /// <summary>Removes only the providers registered by this .xll; other add-ins keep theirs.</summary>
    public static void Clear()
    {
        var list = Shared();
        lock (list) { list.RemoveAll(p => p.ScopeId == ScopeId); }
    }

    private static (Func<object, bool> AppliesTo,
        Func<object, (List<(string Display, string Raw, int Section, bool IsValid)> Options, Action<object, string, string> Commit)?> Build)[] Snapshot()
    {
        var list = Shared();
        lock (list) { return list.Select(p => (p.AppliesTo, p.Build)).ToArray(); }
    }

    /// <summary>Providers registered by every add-in in this process (diagnostics).</summary>
    public static int ProviderCount => Snapshot().Length;

    public static bool AppliesTo(dynamic singleCell)
    {
        object cell = singleCell;
        foreach (var p in Snapshot())
        {
            try { if (p.AppliesTo(cell)) return true; }
            catch (Exception ex) { AppLogger.Warn($"[CustomDropdowns] AppliesTo failed: {ex.GetType().Name}: {ex.Message}"); }
        }
        return false;
    }

    public static CustomDropdownResult? TryBuild(dynamic singleCell)
    {
        object cell = singleCell;
        foreach (var p in Snapshot())
        {
            try
            {
                if (!p.AppliesTo(cell)) continue;
                var built = p.Build(cell);
                if (built == null) continue;

                var (options, commit) = built.Value;
                var local = options.Select(o => new DropdownOption { Display = o.Display, Raw = o.Raw, Section = o.Section, IsValid = o.IsValid }).ToList();
                return new CustomDropdownResult(local, (c, opt) => commit((object)c, opt.Raw, opt.Display));
            }
            catch (Exception ex) { AppLogger.Error("[CustomDropdowns] Build failed", ex); }
        }
        return null;
    }
}
