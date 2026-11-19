namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using SKSB_App.Core.Types;

/// <summary>What a slot needs, used to rank workers.</summary>
internal sealed record SlotContext(
    DateOnly Date,
    Shifts Band,
    IReadOnlyList<string> RequiredSkills,
    IReadOnlyList<VehicleType> RequiredVehicleTypes,
    IReadOnlyList<string> SlotVehiclePlates,
    IReadOnlyList<VehicleType> SlotVehicleTypes,
    IReadOnlyDictionary<long, int> SameShiftAssignmentCounts,
    IReadOnlyDictionary<long, int> OtherShiftAssignmentCounts);

/// <summary>Everything the engine knows about one worker for scoring.</summary>
internal sealed record WorkerProfile(
    StaffInfo Staff,
    WorkerPreference? Preference,
    int DaysWorkedThisWeek,
    bool IsUnavailableToday);

internal sealed record MatchResult(int Percent, bool Blocked, string Reason);

/// <summary>
/// Scores workers (and whole teams) against a slot. Weights add up to 100:
/// primary skill 30, licensed/other skills 20, vehicle fit 15, shift 10, rest day 10, weekly cap 10, availability 5.
/// Workers already on other project slots the same day are penalized by inverse-proportion decay:
/// load = sameShift + 0.5 * otherShift, and the positive score is divided by 1 + 0.25 * load + 0.25 * load^2.
/// A blacklisted required skill subtracts 40 and blocks the worker (listed last, flagged).
/// </summary>
internal static class ScheduleMatchEngine
{
    private static readonly Regex PercentPrefix = new(@"^\s*\[\d{1,3}%\]\s*", RegexOptions.Compiled);

    // =========================================================================
    // 1. SCORING
    // =========================================================================
    public static MatchResult Score(WorkerProfile worker, SlotContext slot)
    {
        var pref = worker.Preference;
        var required = slot.RequiredSkills.Where(s => !string.IsNullOrWhiteSpace(s)).ToList();
        var has = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        string? effectivePrimary = BoardContext.EffectivePrimarySkill(pref);
        if (pref != null)
        {
            if (effectivePrimary != null) has.Add(effectivePrimary);
            foreach (var s in pref.LicensedSkills.Concat(pref.OtherSkills)) if (!string.IsNullOrWhiteSpace(s)) has.Add(s.Trim());
        }

        double score = 0;
        var notes = new List<string>();

        // Primary skill (30)
        if (required.Count == 0) score += 15;
        else if (!string.IsNullOrEmpty(effectivePrimary) &&
                 required.Contains(effectivePrimary, StringComparer.OrdinalIgnoreCase)) score += 30;
        if (BoardContext.DescribeInactivePrimarySkillLicense(pref, slot.Date) is { } inactive)
            notes.Add(inactive.ShortNote);

        // Licensed + other skills coverage (20)
        if (required.Count == 0) score += 10;
        else score += 20.0 * required.Count(r => has.Contains(r)) / required.Count;

        // Vehicle fit (15)
        bool hasSlotVehicles = slot.SlotVehiclePlates.Count > 0 || slot.RequiredVehicleTypes.Count > 0;
        if (!hasSlotVehicles) score += 7;
        else if (pref != null)
        {
            bool dedicated = !string.IsNullOrWhiteSpace(pref.DedicatedVehiclePlate) &&
                             slot.SlotVehiclePlates.Contains(pref.DedicatedVehiclePlate, StringComparer.OrdinalIgnoreCase);
            bool typeFit = pref.PreferredVehicleType != VehicleType.None &&
                           (slot.SlotVehicleTypes.Contains(pref.PreferredVehicleType) || slot.RequiredVehicleTypes.Contains(pref.PreferredVehicleType));
            score += dedicated ? 15 : typeFit ? 10 : 0;
        }

        // Preferred shift (10)
        if (pref == null || pref.PreferredShift == Shifts.None) score += 6;
        else if (BandOf(pref.PreferredShift) == slot.Band) score += 10;
        else notes.Add("prefers other shift");

        // Rest day (10)
        if (pref != null && pref.RestDayPreference != Days.None && ToDays(slot.Date.DayOfWeek) == pref.RestDayPreference)
            notes.Add("rest day");
        else score += 10;

        // Weekly cap (10)
        if (pref != null && pref.MaxDaysPerWeek > 0 && worker.DaysWorkedThisWeek >= pref.MaxDaysPerWeek) notes.Add("weekly cap reached");
        else score += 10;

        // Availability (5)
        slot.SameShiftAssignmentCounts.TryGetValue(worker.Staff.Id, out int sameShiftCount);
        slot.OtherShiftAssignmentCounts.TryGetValue(worker.Staff.Id, out int otherShiftCount);
        if (worker.IsUnavailableToday) notes.Add("unavailable");
        else if (sameShiftCount > 0)
            notes.Add(sameShiftCount == 1 ? "already on 1 project this shift" : $"already on {sameShiftCount} projects this shift");
        else score += 5;
        if (otherShiftCount > 0)
            notes.Add(otherShiftCount == 1 ? "already on 1 project other shift" : $"already on {otherShiftCount} projects other shift");

        // Multi-project decay
        double load = sameShiftCount + (0.5 * otherShiftCount);
        if (load > 0)
            score /= 1.0 + (0.25 * load) + (0.25 * load * load);

        // Blacklist
        bool blacklisted = pref != null && required.Any(r => pref.BlacklistedSkills.Contains(r, StringComparer.OrdinalIgnoreCase));
        if (blacklisted) { score -= 40; notes.Add("blacklisted skill"); }

        int percent = (int)Math.Round(Math.Clamp(score, 0, 100));
        return new MatchResult(percent, blacklisted || worker.IsUnavailableToday || !Roster.IsActiveOn(worker.Staff, slot.Date), string.Join(", ", notes));
    }

    /// <summary>A team's score is the mean of its members' scores.</summary>
    public static MatchResult ScoreTeam(IReadOnlyList<WorkerProfile> members, SlotContext slot)
    {
        if (members.Count == 0) return new MatchResult(0, true, "no members");
        var results = members.Select(m => Score(m, slot)).ToList();
        return new MatchResult((int)Math.Round(results.Average(r => r.Percent)), results.All(r => r.Blocked), string.Empty);
    }

    // =========================================================================
    // 2. DISPLAY PREFIX
    // =========================================================================
    public static string WithPercent(string name, int percent) => $"[{percent}%] {name}";

    /// <summary>Removes a leading "[95%] " rank prefix; text without one is returned trimmed.</summary>
    public static string StripPercent(string? text) => PercentPrefix.Replace(text ?? string.Empty, string.Empty).Trim();

    // =========================================================================
    // 3. HELPERS
    // =========================================================================
    internal static Shifts BandOf(Shifts s) => s == Shifts.Night ? Shifts.Night : Shifts.Morning;

    internal static Days ToDays(DayOfWeek d) => d switch
    {
        DayOfWeek.Monday => Days.Monday,
        DayOfWeek.Tuesday => Days.Tuesday,
        DayOfWeek.Wednesday => Days.Wednesday,
        DayOfWeek.Thursday => Days.Thursday,
        DayOfWeek.Friday => Days.Friday,
        DayOfWeek.Saturday => Days.Saturday,
        _ => Days.Sunday
    };
}
