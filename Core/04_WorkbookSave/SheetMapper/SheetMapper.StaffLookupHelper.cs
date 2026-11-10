namespace SKSB_App.Core.WorkbookSave.SheetMapper;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.Storage;

/// <summary>
/// Cross-domain staff resolution: lets any Manager module turn a free-text name (driver on a fuel
/// statement, worker on a schedule) into a StaffID using the cached <see cref="StaffRecord"/> directory.
/// </summary>
public static class StaffLookupHelper
{
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    /// <summary>Trims and collapses internal whitespace so names compare reliably.</summary>
    public static string NormalizeName(string? name)
    {
        if (string.IsNullOrWhiteSpace(name)) return string.Empty;
        return WhitespaceRegex.Replace(name.Trim(), " ");
    }

    /// <summary>Loads the staff directory from the local snapshot refreshed by Staff Manager.</summary>
    public static List<StaffRecord> LoadCachedStaff() =>
        OfflineSnapshot.Load<List<StaffRecord>>() ?? new List<StaffRecord>();

    /// <summary>
    /// Builds the exact-match index (normalised ShortName / FullName -> StaffID).
    /// A short name always wins over another staff member's full name.
    /// </summary>
    public static Dictionary<string, long> BuildExactNameMap(IEnumerable<StaffRecord> staffList)
    {
        var map = new Dictionary<string, long>(StringComparer.OrdinalIgnoreCase);
        foreach (var s in staffList)
        {
            if (s.StaffID <= 0) continue;
            string fn = NormalizeName(s.StaffDetails?.StaffName);
            string sn = NormalizeName(s.StaffDetails?.StaffShortName);

            if (!string.IsNullOrEmpty(sn)) map[sn] = s.StaffID;
            if (!string.IsNullOrEmpty(fn) && !map.ContainsKey(fn)) map[fn] = s.StaffID;
        }
        return map;
    }

    /// <summary>Resolves against the cached staff directory. Returns 0 when no match is found.</summary>
    public static long ResolveStaffId(string? name)
    {
        var staffList = LoadCachedStaff();
        return ResolveStaffId(name, BuildExactNameMap(staffList), staffList);
    }

    /// <summary>
    /// Resolves a name to a StaffID using multi-tier matching:
    /// 1. Exact match on StaffShortName or StaffName
    /// 2. Whole-word / token match (e.g. "SYUKUR" matches "MOHD SYUKUR")
    /// 3. Substring fallback
    /// Returns 0 when no match is found.
    /// </summary>
    public static long ResolveStaffId(
        string? name,
        IReadOnlyDictionary<string, long> staffExactMap,
        IEnumerable<StaffRecord> staffList)
    {
        string clean = NormalizeName(name);
        if (string.IsNullOrWhiteSpace(clean)) return 0;

        // Tier 1: Exact match against ShortName or FullName
        if (staffExactMap.TryGetValue(clean, out long exactId))
            return exactId;

        var nameTokens = clean.Split(' ', StringSplitOptions.RemoveEmptyEntries);

        // Tier 2: Whole-word token match (prioritize StaffShortName first, then StaffName)
        long bestStaffId = 0;
        int bestMatchedWords = 0;

        foreach (var s in staffList)
        {
            if (s.StaffID <= 0) continue;

            string sn = NormalizeName(s.StaffDetails?.StaffShortName);
            string fn = NormalizeName(s.StaffDetails?.StaffName);

            foreach (var candidate in new[] { sn, fn })
            {
                if (string.IsNullOrEmpty(candidate)) continue;
                var candidateTokens = candidate.Split(' ', StringSplitOptions.RemoveEmptyEntries);

                // Check if all tokens in the name exist as whole words in candidate (or vice versa)
                bool nameInCandidate = nameTokens.All(nt =>
                    candidateTokens.Any(ct => ct.Equals(nt, StringComparison.OrdinalIgnoreCase)));
                bool candidateInName = candidateTokens.All(ct =>
                    nameTokens.Any(nt => nt.Equals(ct, StringComparison.OrdinalIgnoreCase)));

                if (nameInCandidate || candidateInName)
                {
                    int overlap = nameTokens.Count(nt =>
                        candidateTokens.Any(ct => ct.Equals(nt, StringComparison.OrdinalIgnoreCase)));
                    if (overlap > bestMatchedWords)
                    {
                        bestMatchedWords = overlap;
                        bestStaffId = s.StaffID;
                    }
                }
            }
        }

        if (bestStaffId > 0) return bestStaffId;

        // Tier 3: Substring containment fallback (only for names >= 3 chars to avoid false positives)
        if (clean.Length >= 3)
        {
            foreach (var (candidateName, id) in staffExactMap)
            {
                if (candidateName.Length >= 3 &&
                    (candidateName.Contains(clean, StringComparison.OrdinalIgnoreCase) ||
                     clean.Contains(candidateName, StringComparison.OrdinalIgnoreCase)))
                {
                    return id;
                }
            }
        }

        return 0;
    }
}
