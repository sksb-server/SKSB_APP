namespace SKSB_App.ScheduleManager;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
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
    public static MatchResult Score(WorkerProfile worker, SlotContext slot) => Evaluate(worker, slot, false).Result;

    /// <summary>Multi-line Best Match breakdown for a worker cell comment (same math as <see cref="Score"/>).</summary>
    public static string FormatBreakdown(WorkerProfile worker, SlotContext slot) => Evaluate(worker, slot, true).Breakdown ?? string.Empty;

    private static (MatchResult Result, string? Breakdown) Evaluate(WorkerProfile worker, SlotContext slot, bool buildBreakdown)
    {
        var inv = CultureInfo.InvariantCulture;
        var lines = buildBreakdown ? new List<string>() : null;
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
        var inactive = BoardContext.DescribeInactivePrimarySkillLicense(pref, slot.Date);
        if (required.Count == 0)
        {
            score += 15;
            lines?.Add($"Primary Skill: +15 / 30 (no skills req{(inactive != null ? "; ⚠ " + inactive.Value.Detail : "")})");
        }
        else if (!string.IsNullOrEmpty(effectivePrimary) &&
                 required.Contains(effectivePrimary, StringComparer.OrdinalIgnoreCase))
        {
            score += 30;
            lines?.Add($"Primary Skill: +30 / 30 ({effectivePrimary})");
        }
        else
        {
            string primaryName = string.IsNullOrWhiteSpace(pref?.PrimarySkill) ? "none" : pref!.PrimarySkill.Trim();
            lines?.Add(inactive != null ? $"Primary Skill: +0 / 30 (⚠ {inactive.Value.Detail})" : $"Primary Skill: +0 / 30 ({primaryName})");
        }
        if (inactive != null) notes.Add(inactive.Value.ShortNote);

        // Licensed + other skills coverage (20)
        if (required.Count == 0)
        {
            score += 10;
            lines?.Add("Skill Coverage: +10 / 20 (no skills req)");
        }
        else
        {
            int matched = required.Count(r => has.Contains(r));
            double covPts = 20.0 * matched / required.Count;
            score += covPts;
            lines?.Add($"Skill Coverage: +{covPts.ToString("0.#", inv)} / 20 ({matched}/{required.Count} required)");
        }

        // Vehicle fit (15)
        bool hasSlotVehicles = slot.SlotVehiclePlates.Count > 0 || slot.RequiredVehicleTypes.Count > 0;
        if (!hasSlotVehicles)
        {
            score += 7;
            lines?.Add("Vehicle Fit: +7 / 15 (no vehicles req)");
        }
        else if (pref == null) lines?.Add("Vehicle Fit: +0 / 15");
        else
        {
            bool dedicated = !string.IsNullOrWhiteSpace(pref.DedicatedVehiclePlate) &&
                             slot.SlotVehiclePlates.Contains(pref.DedicatedVehiclePlate, StringComparer.OrdinalIgnoreCase);
            bool typeFit = pref.PreferredVehicleType != VehicleType.None &&
                           (slot.SlotVehicleTypes.Contains(pref.PreferredVehicleType) || slot.RequiredVehicleTypes.Contains(pref.PreferredVehicleType));
            score += dedicated ? 15 : typeFit ? 10 : 0;
            lines?.Add(dedicated ? $"Vehicle Fit: +15 / 15 (dedicated: {pref.DedicatedVehiclePlate})"
                : typeFit ? $"Vehicle Fit: +10 / 15 (preferred: {pref.PreferredVehicleType.ToString().Replace('_', ' ')})"
                : "Vehicle Fit: +0 / 15");
        }

        // Preferred shift (10)
        if (pref == null || pref.PreferredShift == Shifts.None)
        {
            score += 6;
            lines?.Add("Preferred Shift: +6 / 10 (no preference)");
        }
        else if (BandOf(pref.PreferredShift) == slot.Band)
        {
            score += 10;
            lines?.Add($"Preferred Shift: +10 / 10 ({slot.Band})");
        }
        else
        {
            notes.Add("prefers other shift");
            lines?.Add($"Preferred Shift: +0 / 10 (prefers {BandOf(pref.PreferredShift)})");
        }

        // Rest day (10)
        if (pref != null && pref.RestDayPreference != Days.None && ToDays(slot.Date.DayOfWeek) == pref.RestDayPreference)
        {
            notes.Add("rest day");
            lines?.Add($"Rest Day: +0 / 10 (rest day: {pref.RestDayPreference})");
        }
        else
        {
            score += 10;
            lines?.Add("Rest Day: +10 / 10");
        }

        // Weekly cap (10)
        bool hasCap = pref != null && pref.MaxDaysPerWeek > 0;
        if (hasCap && worker.DaysWorkedThisWeek >= pref!.MaxDaysPerWeek)
        {
            notes.Add("weekly cap reached");
            lines?.Add($"Weekly Cap: +0 / 10 ({worker.DaysWorkedThisWeek}/{pref.MaxDaysPerWeek} days)");
        }
        else
        {
            score += 10;
            lines?.Add(hasCap ? $"Weekly Cap: +10 / 10 ({worker.DaysWorkedThisWeek}/{pref!.MaxDaysPerWeek} days)" : "Weekly Cap: +10 / 10 (no cap)");
        }

        // Availability (5)
        slot.SameShiftAssignmentCounts.TryGetValue(worker.Staff.Id, out int sameShiftCount);
        slot.OtherShiftAssignmentCounts.TryGetValue(worker.Staff.Id, out int otherShiftCount);
        if (worker.IsUnavailableToday)
        {
            notes.Add("unavailable");
            lines?.Add("Availability: +0 / 5 (unavailable)");
        }
        else if (sameShiftCount > 0)
        {
            notes.Add(sameShiftCount == 1 ? "already on 1 project this shift" : $"already on {sameShiftCount} projects this shift");
            lines?.Add($"Availability: +0 / 5 ({sameShiftCount} other project(s) this shift)");
        }
        else
        {
            score += 5;
            lines?.Add("Availability: +5 / 5");
        }
        if (otherShiftCount > 0)
            notes.Add(otherShiftCount == 1 ? "already on 1 project other shift" : $"already on {otherShiftCount} projects other shift");

        // Multi-project decay
        double load = sameShiftCount + (0.5 * otherShiftCount);
        if (load > 0)
        {
            double divisor = 1.0 + (0.25 * load) + (0.25 * load * load);
            score /= divisor;
            lines?.Add($"Multi-Project Decay: ÷{divisor.ToString("0.##", inv)} ({sameShiftCount} same, {otherShiftCount} other shift) -> {score.ToString("0.#", inv)}");
        }

        // Blacklist
        var blacklistedSkills = pref == null
            ? new List<string>()
            : required.Where(r => pref.BlacklistedSkills.Contains(r, StringComparer.OrdinalIgnoreCase)).ToList();
        bool blacklisted = blacklistedSkills.Count > 0;
        if (blacklisted)
        {
            score -= 40;
            notes.Add("blacklisted skill");
            lines?.Add($"Blacklist Penalty: -40 ({string.Join(", ", blacklistedSkills)})");
        }

        bool activeOn = Roster.IsActiveOn(worker.Staff, slot.Date);
        if (!activeOn) lines?.Add($"⚠ Inactive staff on {slot.Date.ToString("dd/MM/yyyy", inv)}");

        int percent = (int)Math.Round(Math.Clamp(score, 0, 100));
        var result = new MatchResult(percent, blacklisted || worker.IsUnavailableToday || !activeOn, string.Join(", ", notes));

        string? breakdown = null;
        if (lines != null)
        {
            var sb = new StringBuilder();
            sb.Append("Best Match: ").Append(percent.ToString(inv)).Append('%');
            if (result.Blocked) sb.Append(" (Blocked)");
            foreach (var l in lines) sb.Append('\n').Append(l);
            breakdown = sb.ToString();
        }
        return (result, breakdown);
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
