namespace SKSB_App.Core.Localization;

using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using SKSB_App.Core.RuntimeUI.Dictionary;
using SKSB_App.Core.RuntimeUI.Guide;

/// <summary>One dictionary entry that a "Pick Text on Screen" capture (or Excel cell) resolved to.</summary>
public sealed class DictionaryPickerMatch
{
    public string DictType { get; init; } = CollectiveDictionaryRepository.DictTypeUi;
    public string EntryKey { get; init; } = string.Empty;
    public string FriendlyLocation { get; init; } = string.Empty;
    public string EnglishText { get; init; } = string.Empty;
    public string CurrentText { get; init; } = string.Empty;
    public string MatchReason { get; init; } = string.Empty;
    public int Score { get; init; }
}

/// <summary>
/// Resolves on-screen text (UI Automation name / WinForms tag) or an Excel cell context back to the
/// Collective Dictionary entries that produced it, plus plain-English location / category labels.
/// </summary>
public static class LocalizationTextMatcher
{
    // =========================================================================
    // 1. CONSTANTS & CACHES
    // =========================================================================
    private const string UiTagPrefix = "ui::";
    private const string TooltipTagPrefix = "tooltips::";

    private const int ScoreDirectTag = 1000;
    private const int ScoreExactActive = 820;
    private const int ScoreExactEnglish = 800;
    private const int ScoreTemplateBase = 600;
    private const int ScoreSibling = 500;
    private const int ScoreSubstring = 200;
    private const int MaxSubstringMatches = 8;
    private const int MaxProximityMatches = 20;
    private const int MaxProximitySubstringCandidates = 3;

    private static readonly string[] RibbonRoleSuffixes = { ".label", ".labelOnline", ".labelOffline", ".screentip", ".supertip" };

    private static readonly Regex PlaceholderRegex = new(@"\{[0-9]+\}", RegexOptions.Compiled);
    private static readonly Regex WhitespaceRegex = new(@"\s+", RegexOptions.Compiled);

    /// <summary>Template text -> compiled regex (null when the template is too generic to match safely).</summary>
    private static readonly ConcurrentDictionary<string, Regex?> _templateCache = new(StringComparer.Ordinal);

    /// <summary>Per-call snapshot of effective + English dictionaries for both dict types.</summary>
    private sealed class DictSet
    {
        public Dictionary<string, string> UiActive = new(StringComparer.Ordinal);
        public Dictionary<string, string> UiEnglish = new(StringComparer.Ordinal);
        public Dictionary<string, string> TipActive = new(StringComparer.Ordinal);
        public Dictionary<string, string> TipEnglish = new(StringComparer.Ordinal);

        public Dictionary<string, string> Active(string dictType) => IsTooltip(dictType) ? TipActive : UiActive;
        public Dictionary<string, string> English(string dictType) => IsTooltip(dictType) ? TipEnglish : UiEnglish;

        public bool Contains(string dictType, string key) => Active(dictType).ContainsKey(key) || English(dictType).ContainsKey(key);

        /// <summary>Normalized (key, raw, normalized, hasPlaceholder) rows per source dictionary, built once per DictSet.</summary>
        private readonly Dictionary<Dictionary<string, string>, List<NormalizedEntry>> _normalized = new(ReferenceEqualityComparer.Instance);

        public List<NormalizedEntry> GetNormalized(Dictionary<string, string> source)
        {
            if (_normalized.TryGetValue(source, out var rows)) return rows;
            rows = new List<NormalizedEntry>(source.Count);
            foreach (var kv in source)
            {
                if (string.IsNullOrWhiteSpace(kv.Value)) continue;
                string norm = Normalize(kv.Value);
                if (norm.Length == 0) continue;
                bool hasPlaceholder = kv.Value.IndexOf('{') >= 0 && PlaceholderRegex.IsMatch(kv.Value);
                rows.Add(new NormalizedEntry(kv.Key, norm, hasPlaceholder));
            }
            _normalized[source] = rows;
            return rows;
        }
    }

    private readonly record struct NormalizedEntry(string Key, string Norm, bool HasPlaceholder);

    // =========================================================================
    // 2. SCREEN PICK MATCHING
    // =========================================================================
    public static List<DictionaryPickerMatch> MatchFromScreenPick(string? controlTag, string? primaryText, string? helpText, string langCode)
    {
        var dicts = LoadDicts(langCode);
        return MatchFromScreenPickCore(dicts, controlTag, primaryText, helpText, allowSubstringFallback: true);
    }

    /// <summary>
    /// Resolves a click plus its proximity samples in strict priority order: Excel cell, exact clicked element,
    /// then nearby elements by ascending distance. Deduplicated (first occurrence wins) and capped.
    /// </summary>
    public static List<DictionaryPickerMatch> MatchFromProximityPick(ScreenPickResult pick, CellGuideSnapshot? excelCell, string langCode)
    {
        var ordered = new List<DictionaryPickerMatch>();
        if (pick == null) return ordered;

        var dicts = LoadDicts(langCode);
        var seen = new HashSet<(string, string)>();

        bool AddRange(IEnumerable<DictionaryPickerMatch> matches)
        {
            foreach (var m in matches)
            {
                if (ordered.Count >= MaxProximityMatches) return false;
                if (seen.Add((m.DictType.ToLowerInvariant(), m.EntryKey))) ordered.Add(m);
            }
            return ordered.Count < MaxProximityMatches;
        }

        // Tier 1: Excel cell.
        if (excelCell != null && !AddRange(MatchFromExcelCellCore(dicts, excelCell))) return ordered;

        // Tier 2: exact clicked element.
        if (!AddRange(MatchFromScreenPickCore(dicts, pick.ControlTag, pick.CapturedText, pick.HelpText, allowSubstringFallback: true)))
            return ordered;

        // Tier 3: nearby candidates (closest first), skipping duplicate captures.
        var nearby = (pick.NearbyCandidates ?? new List<ScreenPickCandidate>())
            .OrderBy(c => c.DistancePx)
            .ToList();
        var seenCandidates = new HashSet<(string, string, string)>();
        foreach (var cand in nearby)
        {
            var candId = ((cand.DirectTag ?? string.Empty).Trim().ToLowerInvariant(),
                          Normalize(cand.ElementName).ToLowerInvariant(),
                          Normalize(cand.ElementHelpText).ToLowerInvariant());
            if (!seenCandidates.Add(candId)) continue;

            if (!AddRange(MatchFromScreenPickCore(dicts, cand.DirectTag, cand.ElementName, cand.ElementHelpText, allowSubstringFallback: false)))
                return ordered;
        }

        // Last resort: substring fallback on the closest named neighbours.
        if (ordered.Count == 0)
        {
            foreach (var cand in nearby.Where(c => !string.IsNullOrWhiteSpace(c.ElementName)).Take(MaxProximitySubstringCandidates))
            {
                string norm = Normalize(cand.ElementName);
                if (norm.Length < 4) continue;
                var results = new Dictionary<(string, string), DictionaryPickerMatch>();
                RunSubstringFallback(results, dicts, norm);
                if (!AddRange(Rank(results.Values))) break;
            }
        }

        return ordered;
    }

    private static List<DictionaryPickerMatch> MatchFromScreenPickCore(
        DictSet dicts,
        string? controlTag,
        string? primaryText,
        string? helpText,
        bool allowSubstringFallback)
    {
        var results = new Dictionary<(string, string), DictionaryPickerMatch>();

        // Tier 1: direct WinForms tag ("ui::key" / "tooltips::key").
        if (TryParseTag(controlTag, out string tagType, out string tagKey) && dicts.Contains(tagType, tagKey))
        {
            AddOrUpgrade(results, BuildMatch(dicts, tagType, tagKey, ScoreDirectTag, "tag"));
        }

        string normPrimary = Normalize(primaryText);
        string normHelp = Normalize(helpText);
        var candidates = new List<string>(2);
        if (normPrimary.Length > 0) candidates.Add(normPrimary);
        if (normHelp.Length > 0 && !string.Equals(normHelp, normPrimary, StringComparison.OrdinalIgnoreCase)) candidates.Add(normHelp);

        if (candidates.Count > 0)
        {
            foreach (string dictType in new[] { CollectiveDictionaryRepository.DictTypeUi, CollectiveDictionaryRepository.DictTypeTooltips })
            {
                ScanExactAndTemplates(results, dicts, dictType, dicts.Active(dictType), candidates, isEnglish: false);
                if (!ReferenceEquals(dicts.Active(dictType), dicts.English(dictType)))
                    ScanExactAndTemplates(results, dicts, dictType, dicts.English(dictType), candidates, isEnglish: true);
            }
        }

        // Tier 4a: Ribbon sibling expansion (label <-> screentip <-> supertip).
        ExpandRibbonSiblings(results, dicts);

        // Tier 4b: safe substring fallback when nothing else matched.
        if (allowSubstringFallback && results.Count == 0 && normPrimary.Length >= 4)
        {
            RunSubstringFallback(results, dicts, normPrimary);
        }

        return Rank(results.Values);
    }

    private static void ScanExactAndTemplates(
        Dictionary<(string, string), DictionaryPickerMatch> results,
        DictSet dicts,
        string dictType,
        Dictionary<string, string> source,
        List<string> candidates,
        bool isEnglish)
    {
        foreach (var row in dicts.GetNormalized(source))
        {
            string normValue = row.Norm;
            bool hasPlaceholder = row.HasPlaceholder;

            foreach (string candidate in candidates)
            {
                // Tier 2: exact normalized match.
                if (string.Equals(normValue, candidate, StringComparison.OrdinalIgnoreCase))
                {
                    AddOrUpgrade(results, BuildMatch(dicts, dictType, row.Key,
                        isEnglish ? ScoreExactEnglish : ScoreExactActive, isEnglish ? "exact-en" : "exact"));
                    break;
                }

                // Tier 3: {0} placeholder template match.
                if (hasPlaceholder)
                {
                    var rx = GetTemplateRegex(normValue, out int literalChars);
                    if (rx != null && SafeIsMatch(rx, candidate))
                    {
                        int score = ScoreTemplateBase + Math.Min(150, literalChars) - (isEnglish ? 5 : 0);
                        AddOrUpgrade(results, BuildMatch(dicts, dictType, row.Key, score, "template"));
                        break;
                    }
                }
            }
        }
    }

    private static Regex? GetTemplateRegex(string normTemplate, out int literalChars)
    {
        literalChars = CountLiteralChars(normTemplate);
        int literal = literalChars;
        return _templateCache.GetOrAdd(normTemplate, t =>
        {
            if (literal < 3) return null;
            try
            {
                var sb = new StringBuilder("^");
                int last = 0;
                foreach (Match m in PlaceholderRegex.Matches(t))
                {
                    sb.Append(Regex.Escape(t.Substring(last, m.Index - last)));
                    sb.Append("(.+?)");
                    last = m.Index + m.Length;
                }
                sb.Append(Regex.Escape(t.Substring(last)));
                sb.Append('$');
                return new Regex(sb.ToString(), RegexOptions.IgnoreCase | RegexOptions.Singleline | RegexOptions.CultureInvariant,
                    TimeSpan.FromMilliseconds(50));
            }
            catch
            {
                return null;
            }
        });
    }

    private static bool SafeIsMatch(Regex rx, string input)
    {
        try { return rx.IsMatch(input); }
        catch (RegexMatchTimeoutException) { return false; }
    }

    private static int CountLiteralChars(string template)
    {
        string stripped = PlaceholderRegex.Replace(template, string.Empty);
        int count = 0;
        foreach (char ch in stripped)
            if (!char.IsWhiteSpace(ch)) count++;
        return count;
    }

    private static void ExpandRibbonSiblings(Dictionary<(string, string), DictionaryPickerMatch> results, DictSet dicts)
    {
        var seeds = results.Values
            .Where(m => string.Equals(m.DictType, CollectiveDictionaryRepository.DictTypeUi, StringComparison.OrdinalIgnoreCase))
            .ToList();

        foreach (var seed in seeds)
        {
            string? basePrefix = TryGetRibbonBasePrefix(seed.EntryKey);
            if (basePrefix == null) continue;

            foreach (string suffix in RibbonRoleSuffixes)
            {
                string siblingKey = basePrefix + suffix;
                if (string.Equals(siblingKey, seed.EntryKey, StringComparison.Ordinal)) continue;
                if (!dicts.Contains(CollectiveDictionaryRepository.DictTypeUi, siblingKey)) continue;

                var match = BuildMatch(dicts, CollectiveDictionaryRepository.DictTypeUi, siblingKey,
                    Math.Min(seed.Score - 1, ScoreSibling), "sibling");
                AddOrUpgrade(results, match);
            }
        }
    }

    private static string? TryGetRibbonBasePrefix(string key)
    {
        bool isRibbon = key.StartsWith("ribbon.", StringComparison.Ordinal);
        bool isLanguage = key.StartsWith("language.", StringComparison.Ordinal);
        if (!isRibbon && !isLanguage) return null;

        foreach (string suffix in RibbonRoleSuffixes)
        {
            if (key.EndsWith(suffix, StringComparison.Ordinal) && key.Length > suffix.Length)
                return key.Substring(0, key.Length - suffix.Length);
        }
        return null;
    }

    private static void RunSubstringFallback(Dictionary<(string, string), DictionaryPickerMatch> results, DictSet dicts, string normCaptured)
    {
        var hits = new List<DictionaryPickerMatch>();
        var seen = new HashSet<(string, string)>();

        void Scan(string dictType, Dictionary<string, string> source)
        {
            foreach (var kv in source)
            {
                if (seen.Contains((dictType, kv.Key))) continue;
                string normValue = Normalize(kv.Value);
                if (normValue.Length < 4) continue;

                bool hit = normValue.IndexOf(normCaptured, StringComparison.OrdinalIgnoreCase) >= 0
                        || normCaptured.IndexOf(normValue, StringComparison.OrdinalIgnoreCase) >= 0;
                if (!hit) continue;

                seen.Add((dictType, kv.Key));
                // Closer lengths rank higher.
                int diff = Math.Abs(normValue.Length - normCaptured.Length);
                int score = ScoreSubstring - Math.Min(150, diff);
                hits.Add(BuildMatch(dicts, dictType, kv.Key, score, "substring"));
            }
        }

        Scan(CollectiveDictionaryRepository.DictTypeUi, dicts.UiActive);
        Scan(CollectiveDictionaryRepository.DictTypeTooltips, dicts.TipActive);
        Scan(CollectiveDictionaryRepository.DictTypeUi, dicts.UiEnglish);
        Scan(CollectiveDictionaryRepository.DictTypeTooltips, dicts.TipEnglish);

        foreach (var m in hits.OrderByDescending(h => h.Score).Take(MaxSubstringMatches))
            AddOrUpgrade(results, m);
    }

    // =========================================================================
    // 3. EXCEL CELL MATCHING
    // =========================================================================
    public static List<DictionaryPickerMatch> MatchFromExcelCell(CellGuideSnapshot snap, string langCode)
    {
        if (snap == null) return new List<DictionaryPickerMatch>();
        return MatchFromExcelCellCore(LoadDicts(langCode), snap);
    }

    private static List<DictionaryPickerMatch> MatchFromExcelCellCore(DictSet dicts, CellGuideSnapshot snap)
    {
        var results = new Dictionary<(string, string), DictionaryPickerMatch>();
        if (snap == null) return new List<DictionaryPickerMatch>();

        string tip = CollectiveDictionaryRepository.DictTypeTooltips;
        string normTable = NormalizeKeyPart(snap.TableName);
        string normCol = NormalizeKeyPart(snap.ColumnHeader);

        if (normCol.Length > 0)
        {
            string header = (snap.ColumnHeader ?? string.Empty).Trim();
            string nameKey = ResolveColumnKey(dicts, normTable, normCol, "name");
            if (dicts.Contains(tip, nameKey))
            {
                string loc = Loc.T("dict.picker.role.colName");
                if (header.Length > 0) loc = $"{loc} · {header}";
                AddOrUpgrade(results, BuildMatch(dicts, tip, nameKey, 900, "excel-col-name", loc));
            }

            string descKey = ResolveColumnKey(dicts, normTable, normCol, "desc");
            if (dicts.Contains(tip, descKey))
            {
                string loc = Loc.T("dict.picker.role.colDesc");
                if (header.Length > 0) loc = $"{loc} · {header}";
                AddOrUpgrade(results, BuildMatch(dicts, tip, descKey, 880, "excel-col-desc", loc));
            }

            string sourceKey = ResolveColumnKey(dicts, normTable, normCol, "source");
            if (dicts.Contains(tip, sourceKey))
            {
                string loc = Loc.T("dict.loc.cellGuide") + (header.Length > 0 ? $" · {header}" : "");
                AddOrUpgrade(results, BuildMatch(dicts, tip, sourceKey, 860, "excel-col-source", loc));
            }
        }

        if (normTable.Length > 0)
        {
            string tableKey = $"table.{normTable}.name";
            if (dicts.Contains(tip, tableKey))
            {
                AddOrUpgrade(results, BuildMatch(dicts, tip, tableKey, 700, "excel-table", Loc.T("dict.picker.role.tableName")));
            }
        }

        return Rank(results.Values);
    }

    private static string ResolveColumnKey(DictSet dicts, string normTable, string normCol, string kind)
    {
        if (normTable.Length > 0)
        {
            string scoped = $"col.{normTable}.{normCol}.{kind}";
            if (dicts.Contains(CollectiveDictionaryRepository.DictTypeTooltips, scoped)) return scoped;
        }
        return $"col.{normCol}.{kind}";
    }

    // =========================================================================
    // 4. FRIENDLY LOCATION & CATEGORY LABELS
    // =========================================================================
    public static string GetFriendlyLocation(string dictType, string entryKey)
    {
        string key = entryKey ?? string.Empty;
        bool isRibbonLike = key.StartsWith("ribbon.", StringComparison.Ordinal) || key.StartsWith("language.", StringComparison.Ordinal);

        if (isRibbonLike && (key.EndsWith(".screentip", StringComparison.Ordinal) || key.EndsWith(".supertip", StringComparison.Ordinal)))
            return Loc.T("dict.loc.hoverTip");
        if (isRibbonLike)
            return Loc.T("dict.loc.ribbon");
        if (key.StartsWith("portal.", StringComparison.Ordinal) || key.StartsWith("module.", StringComparison.Ordinal))
            return Loc.T("dict.loc.portal");
        if (key.StartsWith("guide.spotlight.", StringComparison.Ordinal))
            return Loc.T("dict.loc.dailyTip");
        if (key.StartsWith("guide.", StringComparison.Ordinal))
            return Loc.T("dict.loc.cellGuide");
        if (key.StartsWith("col.", StringComparison.Ordinal) || key.StartsWith("table.", StringComparison.Ordinal))
            return Loc.T("dict.loc.columnTooltip");
        if (key.StartsWith("dict.", StringComparison.Ordinal))
            return Loc.T("dict.loc.dictionary");
        return Loc.T("dict.loc.popup");
    }

    /// <summary>Returns one of <c>"buttons"</c>, <c>"tips"</c>, <c>"popups"</c>.</summary>
    public static string GetFriendlyCategory(string dictType, string entryKey)
    {
        if (IsTooltip(dictType)) return "tips";
        string key = entryKey ?? string.Empty;
        if (key.StartsWith("ribbon.", StringComparison.Ordinal)
            || key.StartsWith("portal.", StringComparison.Ordinal)
            || key.StartsWith("module.", StringComparison.Ordinal)
            || key.StartsWith("language.", StringComparison.Ordinal))
        {
            return "buttons";
        }
        return "popups";
    }

    public static string GetFriendlyCategoryLabel(string category) => category switch
    {
        "buttons" => Loc.T("dict.cat.buttons"),
        "tips" => Loc.T("dict.cat.tips"),
        "popups" => Loc.T("dict.cat.popups"),
        _ => Loc.T("dict.cat.all")
    };

    // =========================================================================
    // 5. HELPERS
    // =========================================================================
    /// <summary>Trims and collapses all whitespace / newlines to single spaces.</summary>
    public static string Normalize(string? text)
    {
        if (string.IsNullOrWhiteSpace(text)) return string.Empty;
        return WhitespaceRegex.Replace(text.Trim(), " ");
    }

    private static bool IsTooltip(string dictType)
        => string.Equals(dictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);

    private static bool TryParseTag(string? tag, out string dictType, out string entryKey)
    {
        dictType = string.Empty;
        entryKey = string.Empty;
        if (string.IsNullOrWhiteSpace(tag)) return false;

        if (tag.StartsWith(UiTagPrefix, StringComparison.OrdinalIgnoreCase))
        {
            dictType = CollectiveDictionaryRepository.DictTypeUi;
            entryKey = tag.Substring(UiTagPrefix.Length).Trim();
        }
        else if (tag.StartsWith(TooltipTagPrefix, StringComparison.OrdinalIgnoreCase))
        {
            dictType = CollectiveDictionaryRepository.DictTypeTooltips;
            entryKey = tag.Substring(TooltipTagPrefix.Length).Trim();
        }
        return entryKey.Length > 0;
    }

    private static DictSet LoadDicts(string langCode)
    {
        string lang = string.IsNullOrWhiteSpace(langCode) ? "en" : langCode;
        string ui = CollectiveDictionaryRepository.DictTypeUi;
        string tip = CollectiveDictionaryRepository.DictTypeTooltips;

        var set = new DictSet
        {
            UiEnglish = Loc.GetEffectiveDictionary(ui, "en"),
            TipEnglish = Loc.GetEffectiveDictionary(tip, "en")
        };
        bool isEnglish = string.Equals(lang, "en", StringComparison.OrdinalIgnoreCase);
        set.UiActive = isEnglish ? set.UiEnglish : Loc.GetEffectiveDictionary(ui, lang);
        set.TipActive = isEnglish ? set.TipEnglish : Loc.GetEffectiveDictionary(tip, lang);
        return set;
    }

    private static DictionaryPickerMatch BuildMatch(DictSet dicts, string dictType, string key, int score, string reason, string? friendlyLocation = null)
    {
        dicts.English(dictType).TryGetValue(key, out string? en);
        if (string.IsNullOrEmpty(en))
        {
            string baseEn = Loc.GetBaseEntry(dictType, "en", key);
            if (!string.IsNullOrEmpty(baseEn)) en = baseEn;
        }
        dicts.Active(dictType).TryGetValue(key, out string? current);

        string location = friendlyLocation ?? GetFriendlyLocation(dictType, key);
        string? role = GetRibbonRoleLabel(dictType, key);
        if (friendlyLocation == null && role != null) location = $"{location} · {role}";

        return new DictionaryPickerMatch
        {
            DictType = dictType,
            EntryKey = key,
            EnglishText = en ?? string.Empty,
            CurrentText = current ?? en ?? string.Empty,
            FriendlyLocation = location,
            MatchReason = reason,
            Score = score
        };
    }

    private static string? GetRibbonRoleLabel(string dictType, string key)
    {
        if (IsTooltip(dictType) || TryGetRibbonBasePrefix(key) == null) return null;
        if (key.EndsWith(".screentip", StringComparison.Ordinal)) return Loc.T("dict.picker.role.screentip");
        if (key.EndsWith(".supertip", StringComparison.Ordinal)) return Loc.T("dict.picker.role.supertip");
        return Loc.T("dict.picker.role.label");
    }

    private static void AddOrUpgrade(Dictionary<(string, string), DictionaryPickerMatch> results, DictionaryPickerMatch match)
    {
        var id = (match.DictType.ToLowerInvariant(), match.EntryKey);
        if (!results.TryGetValue(id, out var existing) || existing.Score < match.Score)
            results[id] = match;
    }

    private static List<DictionaryPickerMatch> Rank(IEnumerable<DictionaryPickerMatch> matches)
        => matches
            .OrderByDescending(m => m.Score)
            .ThenBy(m => m.DictType, StringComparer.Ordinal)
            .ThenBy(m => m.EntryKey, StringComparer.Ordinal)
            .ToList();

    /// <summary>Mirror of <c>Loc.NormalizeKeyPart</c>: lower-case ASCII letters and digits only ("Staff ID" -> "staffid").</summary>
    private static string NormalizeKeyPart(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var sb = new StringBuilder(raw.Length);
        foreach (char ch in raw)
        {
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) sb.Append(ch);
            else if (ch >= 'A' && ch <= 'Z') sb.Append(char.ToLowerInvariant(ch));
        }
        return sb.ToString();
    }
}
