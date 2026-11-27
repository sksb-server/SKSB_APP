namespace SKSB_App.Core.Localization;

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using Dapper;
using Npgsql;
using SKSB_App.Core.Integration.Database;

// =========================================================================
// MODELS
// =========================================================================

public sealed class CollectiveOverrideRecord
{
    public string DictType { get; set; } = string.Empty;
    public string EntryKey { get; set; } = string.Empty;
    public string LangCode { get; set; } = string.Empty;
    public string Value { get; set; } = string.Empty;
    public long? SourceSuggestionId { get; set; }
    public string ApprovedBy { get; set; } = string.Empty;
    public DateTime ApprovedAt { get; set; }
}

public sealed class CollectiveSuggestionRecord
{
    public long SuggestionId { get; set; }
    public string DictType { get; set; } = string.Empty;
    public string EntryKey { get; set; } = string.Empty;
    public string LangCode { get; set; } = string.Empty;
    public string SuggestedValue { get; set; } = string.Empty;
    public string Note { get; set; } = string.Empty;
    public string SuggestedBy { get; set; } = string.Empty;
    public string MachineName { get; set; } = string.Empty;
    public string Status { get; set; } = "pending";
    public int VoteCount { get; set; }
    public DateTime CreatedAt { get; set; }
    public string ReviewedBy { get; set; } = string.Empty;
    public DateTime? ReviewedAt { get; set; }
    public bool HasCurrentUserVoted { get; set; }
}

public sealed class CollectiveVoteResult
{
    public int NewVoteCount { get; set; }
    public bool UserVoted { get; set; }
    public bool AutoApproved { get; set; }
    public string Status { get; set; } = "pending";
}

public sealed class CollectiveCachePayload
{
    public DateTime UpdatedAtUtc { get; set; }
    public Dictionary<string, Dictionary<string, string>> Ui { get; set; } = new(StringComparer.OrdinalIgnoreCase);
    public Dictionary<string, Dictionary<string, string>> Tooltips { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

public sealed class CollectiveSuggestionInput
{
    public string DictType { get; set; } = string.Empty;
    public string EntryKey { get; set; } = string.Empty;
    public string LangCode { get; set; } = string.Empty;
    public string SuggestedValue { get; set; } = string.Empty;
    public string Note { get; set; } = string.Empty;
}

public sealed class CollectiveBatchSubmitResult
{
    public int Submitted { get; set; }
    public int Skipped { get; set; }
    public List<string> Errors { get; set; } = new();
}

#if DEBUG
public sealed class CollectiveHistoryRecord
{
    public long HistoryId { get; set; }
    public string DictType { get; set; } = string.Empty;
    public string EntryKey { get; set; } = string.Empty;
    public string LangCode { get; set; } = string.Empty;
    public string? PreviousValue { get; set; }
    public string? NewValue { get; set; }
    public string Action { get; set; } = string.Empty;
    public string ChangedBy { get; set; } = string.Empty;
    public DateTime ChangedAt { get; set; }
}

public sealed class CollectiveDictionaryStateExport
{
    public DateTime ExportedAtUtc { get; set; }
    public List<CollectiveOverrideRecord> Overrides { get; set; } = new();
    public List<CollectiveSuggestionRecord> PendingSuggestions { get; set; } = new();
    public List<CollectiveSuggestionRecord> RejectedSuggestions { get; set; } = new();
    public List<CollectiveHistoryRecord> RecentHistory { get; set; } = new();
}
#endif

// =========================================================================
// REPOSITORY & LOCAL CACHE
// =========================================================================

/// <summary>
/// PostgreSQL-backed Collective Dictionary with community voting consensus and local JSON caching.
/// Always targets the shared <c>app</c> schema via <see cref="PostgreSqlService.ReleaseConnectionString"/>.
/// </summary>
public static class CollectiveDictionaryRepository
{
    public const int ConsensusVoteThreshold = 3;
    public const string DictTypeUi = "ui";
    public const string DictTypeTooltips = "tooltips";
    public const string CustomTipKeyPrefix = "guide.spotlight.custom.";
    public const string CustomContextTipKeyPrefix = "guide.spotlight.custom.ctx.";
    public static readonly string[] CustomTipContextTriggers =
    [
        "offline", "total", "header", "date_filter", "multiselect", "autosync",
        "dropdown", "auto_defaults", "history", "currency", "readonly", "required",
        "daily_attendance", "attendance_entries", "schedule_board",
        "staff_summary", "vehicle_summary", "project_summary"
    ];
    public const string CacheFileName = "collective_dict_cache.json";
    public const string AiHelperName = "AI Consistency Helper";
    public const string AiHelperVoterKey = "ai-consistency-helper@antigravity";

    public static bool IsCustomTipKey(string? key)
        => !string.IsNullOrEmpty(key) && key.StartsWith(CustomTipKeyPrefix, StringComparison.Ordinal);

    public static bool IsCustomContextTipKey(string? key)
        => !string.IsNullOrEmpty(key) && key.StartsWith(CustomContextTipKeyPrefix, StringComparison.Ordinal);

    public static string GetCustomContextTipPrefix(string trigger)
        => $"{CustomContextTipKeyPrefix}{trigger.Trim().ToLowerInvariant()}.";

    /// <summary>
    /// Builds a unique tooltip key under <see cref="CustomTipKeyPrefix"/> (or <see cref="CustomContextTipKeyPrefix"/>
    /// when a supported <paramref name="contextTrigger"/> is supplied) from an optional topic/title
    /// (normalized to lowercase ASCII letters/digits separated by '_', max 32 chars) plus a UTC timestamp suffix.
    /// </summary>
    public static string BuildCustomTipKey(string? topic, string? contextTrigger = null)
    {
        var sb = new StringBuilder(32);
        bool prevUnderscore = false;
        if (!string.IsNullOrWhiteSpace(topic))
        {
            foreach (char ch in topic.Trim())
            {
                if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9'))
                {
                    sb.Append(ch);
                    prevUnderscore = false;
                }
                else if (ch >= 'A' && ch <= 'Z')
                {
                    sb.Append(char.ToLowerInvariant(ch));
                    prevUnderscore = false;
                }
                else if (!prevUnderscore && sb.Length > 0)
                {
                    sb.Append('_');
                    prevUnderscore = true;
                }

                if (sb.Length >= 32) break;
            }
        }

        string slug = sb.ToString().Trim('_');
        if (slug.Length == 0) slug = "tip";
        string suffix = DateTime.UtcNow.ToString("yyyyMMddHHmmss");

        if (!string.IsNullOrWhiteSpace(contextTrigger))
        {
            string normalizedTrigger = contextTrigger.Trim().ToLowerInvariant();
            if (CustomTipContextTriggers.Contains(normalizedTrigger, StringComparer.Ordinal))
                return $"{CustomContextTipKeyPrefix}{normalizedTrigger}.{slug}_{suffix}";
        }

        return $"{CustomTipKeyPrefix}{slug}_{suffix}";
    }

    public static string CacheFilePath { get; } = Path.Combine(UserPreferences.DirectoryPath, CacheFileName);

    public static string CurrentVoterKey =>
        $"{Environment.UserName}@{Environment.MachineName}".Trim().ToLowerInvariant();

    public static string CurrentVoterName => Environment.UserName ?? string.Empty;

    private static readonly object _schemaLock = new();
    private static readonly object _cacheIoLock = new();
    private static bool _schemaEnsured;

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    static CollectiveDictionaryRepository()
    {
        DefaultTypeMap.MatchNamesWithUnderscores = true;
    }

    // =========================================================================
    // 1. CONNECTION & SCHEMA
    // =========================================================================
    private static NpgsqlConnection OpenConnection()
    {
#if DEBUG
        if (ServerEnvironment.SimulateOffline)
            throw new TimeoutException("[Simulated Offline Mode Active] Connection to PostgreSQL server blocked by user switch.");
#endif
        var csb = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString)
        {
            Timeout = 5,
            CommandTimeout = 30
        };
        var conn = new NpgsqlConnection(csb.ConnectionString);
        try
        {
            conn.Open();
        }
        catch
        {
            conn.Dispose();
            throw;
        }
        return conn;
    }

    /// <summary>Runs the embedded <c>collective_dictionary_schema.sql</c> once per process (idempotent DDL).</summary>
    public static void EnsureSchema()
    {
        if (_schemaEnsured) return;
        lock (_schemaLock)
        {
            if (_schemaEnsured) return;
            using Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("collective_dictionary_schema.sql");
            if (stream == null)
                throw new InvalidOperationException("Embedded resource 'collective_dictionary_schema.sql' not found.");
            using var reader = new StreamReader(stream, Encoding.UTF8);
            string sql = reader.ReadToEnd();

            using var conn = OpenConnection();
            conn.Execute(sql);
            _schemaEnsured = true;
        }
    }

    // =========================================================================
    // 2. PLACEHOLDER VALIDATOR
    // =========================================================================
    /// <summary>
    /// Ensures <paramref name="candidate"/> is non-empty, has balanced format braces, and preserves
    /// the exact set of integer format placeholder indices present in <paramref name="englishTemplate"/>.
    /// </summary>
    public static bool ValidatePlaceholders(string englishTemplate, string candidate, out string? errorMessage)
    {
        if (string.IsNullOrWhiteSpace(candidate))
        {
            errorMessage = Loc.T("dict.err.emptyValue");
            return false;
        }

        if (!TryExtractPlaceholderIndices(candidate, out var candidateIndices))
        {
            errorMessage = Loc.T("dict.err.invalidBraces");
            return false;
        }

        TryExtractPlaceholderIndices(englishTemplate ?? string.Empty, out var expectedIndices);

        if (!expectedIndices.SetEquals(candidateIndices))
        {
            string expectedText = FormatIndexList(expectedIndices);
            string actualText = FormatIndexList(candidateIndices);
            errorMessage = Loc.T("dict.err.placeholderMismatch", expectedText, actualText);
            return false;
        }

        errorMessage = null;
        return true;
    }

    /// <summary>Returns a display list of format placeholders in <paramref name="englishTemplate"/> (e.g. "{0}, {1}"), or empty string if none.</summary>
    public static string GetPlaceholderListDisplay(string englishTemplate)
    {
        if (string.IsNullOrEmpty(englishTemplate)) return string.Empty;
        if (!TryExtractPlaceholderIndices(englishTemplate, out var indices) || indices.Count == 0)
            return string.Empty;
        return string.Join(", ", indices.OrderBy(i => i).Select(i => $"{{{i}}}"));
    }

    private static string FormatIndexList(HashSet<int> indices)
    {
        if (indices.Count == 0) return "(none)";
        return string.Join(", ", indices.OrderBy(i => i).Select(i => $"{{{i}}}"));
    }

    private static bool TryExtractPlaceholderIndices(string text, out HashSet<int> indices)
    {
        indices = new HashSet<int>();
        if (string.IsNullOrEmpty(text)) return true;

        int i = 0;
        int len = text.Length;
        while (i < len)
        {
            char c = text[i];
            if (c == '{')
            {
                if (i + 1 < len && text[i + 1] == '{')
                {
                    i += 2; // escaped "{{"
                    continue;
                }

                int close = text.IndexOf('}', i + 1);
                if (close < 0) return false;

                string inside = text.Substring(i + 1, close - i - 1);
                if (inside.Length == 0 || inside.IndexOf('{') >= 0) return false;

                // Split index from optional alignment (",...") or format specifier (":...")
                int sep = inside.IndexOfAny(new[] { ',', ':' });
                string indexPart = sep >= 0 ? inside.Substring(0, sep) : inside;
                if (indexPart.Length == 0) return false;

                foreach (char d in indexPart)
                {
                    if (d < '0' || d > '9') return false;
                }

                if (!int.TryParse(indexPart, out int idx) || idx < 0) return false;
                indices.Add(idx);
                i = close + 1;
            }
            else if (c == '}')
            {
                if (i + 1 < len && text[i + 1] == '}')
                {
                    i += 2; // escaped "}}"
                    continue;
                }
                return false; // stray '}'
            }
            else
            {
                i++;
            }
        }
        return true;
    }

    private static bool ValidateEntryPlaceholders(string dictType, string entryKey, string candidate, out string? errorMessage)
    {
        string englishBase = Loc.GetBaseEntry(dictType, "en", entryKey);
        if (IsCustomTipKey(entryKey) && string.IsNullOrEmpty(englishBase))
            return ValidatePlaceholders(string.Empty, candidate, out errorMessage);
        return ValidatePlaceholders(englishBase, candidate, out errorMessage);
    }

    // =========================================================================
    // 3. LOCAL CACHE & SYNC
    // =========================================================================
    private static CollectiveCachePayload? TryReadCachePayload()
    {
        try
        {
            if (!File.Exists(CacheFilePath)) return null;

            string json;
            using (var fs = new FileStream(CacheFilePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (var reader = new StreamReader(fs, Encoding.UTF8))
            {
                json = reader.ReadToEnd();
            }

            if (string.IsNullOrWhiteSpace(json)) return null;
            return JsonSerializer.Deserialize<CollectiveCachePayload>(json, _jsonOptions);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[CollectiveDict] TryReadCachePayload failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>Loads cached collective overrides for the given dictionary type and language. Exception-safe.</summary>
    public static Dictionary<string, string> LoadCachedOverrides(string dictType, string langCode)
    {
        var result = new Dictionary<string, string>(StringComparer.Ordinal);
        if (string.IsNullOrWhiteSpace(dictType) || string.IsNullOrWhiteSpace(langCode)) return result;

        try
        {
            var payload = TryReadCachePayload();
            if (payload == null) return result;

            var bucket = string.Equals(dictType, DictTypeTooltips, StringComparison.OrdinalIgnoreCase)
                ? payload.Tooltips
                : payload.Ui;

            if (bucket != null)
            {
                foreach (var kv in bucket)
                {
                    if (!string.Equals(kv.Key, langCode, StringComparison.OrdinalIgnoreCase)) continue;
                    if (kv.Value != null)
                    {
                        foreach (var entry in kv.Value)
                        {
                            if (!string.IsNullOrEmpty(entry.Key) && entry.Value != null)
                                result[entry.Key] = entry.Value;
                        }
                    }
                    break;
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[CollectiveDict] LoadCachedOverrides failed: {ex.Message}");
        }

        return result;
    }

    /// <summary>
    /// Returns all distinct tooltip keys across all languages in the local collective cache
    /// that start with <paramref name="prefix"/>, sorted ordinally. Exception-safe.
    /// </summary>
    public static List<string> GetAllCachedTooltipKeysByPrefix(string prefix)
    {
        var keys = new HashSet<string>(StringComparer.Ordinal);
        if (string.IsNullOrEmpty(prefix)) return new List<string>();

        var payload = TryReadCachePayload();
        if (payload?.Tooltips != null)
        {
            foreach (var langBucket in payload.Tooltips.Values)
            {
                if (langBucket == null) continue;
                foreach (var entry in langBucket)
                {
                    if (!string.IsNullOrEmpty(entry.Key)
                        && !string.IsNullOrWhiteSpace(entry.Value)
                        && entry.Key.StartsWith(prefix, StringComparison.Ordinal))
                    {
                        keys.Add(entry.Key);
                    }
                }
            }
        }

        var sorted = keys.ToList();
        sorted.Sort(StringComparer.Ordinal);
        return sorted;
    }

    /// <summary>
    /// Looks up <paramref name="key"/> in the cached tooltip overrides, preferring English ("en")
    /// first and then any other language with a non-empty value. Exception-safe.
    /// </summary>
    public static bool TryGetAnyCachedTooltipValue(string key, out string value)
    {
        value = string.Empty;
        if (string.IsNullOrEmpty(key)) return false;

        var payload = TryReadCachePayload();
        if (payload?.Tooltips == null) return false;

        // Prefer English first
        foreach (var kv in payload.Tooltips)
        {
            if (string.Equals(kv.Key, "en", StringComparison.OrdinalIgnoreCase)
                && kv.Value != null
                && kv.Value.TryGetValue(key, out string? enVal)
                && !string.IsNullOrWhiteSpace(enVal))
            {
                value = enVal;
                return true;
            }
        }

        // Fall back to any other available language
        foreach (var kv in payload.Tooltips)
        {
            if (string.Equals(kv.Key, "en", StringComparison.OrdinalIgnoreCase) || kv.Value == null)
                continue;
            if (kv.Value.TryGetValue(key, out string? otherVal) && !string.IsNullOrWhiteSpace(otherVal))
            {
                value = otherVal;
                return true;
            }
        }

        return false;
    }

    /// <summary>
    /// Fetches all approved overrides from PostgreSQL, writes <see cref="CacheFilePath"/> atomically,
    /// and optionally reloads <see cref="Loc"/> dictionaries. Returns false cleanly when offline.
    /// </summary>
    public static bool SyncFromServer(bool reloadLoc = true)
    {
        try
        {
            EnsureSchema();
            List<CollectiveOverrideRecord> rows;
            using (var conn = OpenConnection())
            {
                const string sql = @"
SELECT dict_type, entry_key, lang_code, value, source_suggestion_id, approved_by, approved_at
FROM app.localization_overrides;";
                rows = conn.Query<CollectiveOverrideRecord>(sql).ToList();
            }

            var payload = new CollectiveCachePayload
            {
                UpdatedAtUtc = DateTime.UtcNow,
                Ui = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase),
                Tooltips = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase)
            };

            foreach (var row in rows)
            {
                if (string.IsNullOrWhiteSpace(row.EntryKey) || string.IsNullOrWhiteSpace(row.LangCode)) continue;
                var bucket = string.Equals(row.DictType, DictTypeTooltips, StringComparison.OrdinalIgnoreCase)
                    ? payload.Tooltips
                    : payload.Ui;

                if (!bucket.TryGetValue(row.LangCode, out var langDict))
                {
                    langDict = new Dictionary<string, string>(StringComparer.Ordinal);
                    bucket[row.LangCode] = langDict;
                }
                langDict[row.EntryKey] = row.Value ?? string.Empty;
            }

            lock (_cacheIoLock)
            {
                Directory.CreateDirectory(UserPreferences.DirectoryPath);
                string json = JsonSerializer.Serialize(payload, _jsonOptions);
                string tempPath = CacheFilePath + "." + Environment.ProcessId + ".tmp";
                File.WriteAllText(tempPath, json, Encoding.UTF8);
                File.Move(tempPath, CacheFilePath, overwrite: true);
            }

            if (reloadLoc) Loc.ReloadDictionaries(raiseEvent: true);
            return true;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[CollectiveDict] SyncFromServer skipped/failed: {ex.Message}");
            return false;
        }
    }

    /// <summary>Fires a background sync without blocking the caller or throwing exceptions.</summary>
    public static void SyncFromServerInBackground()
    {
        _ = Task.Run(() =>
        {
            try { SyncFromServer(reloadLoc: true); }
            catch { }
        });
    }

    // =========================================================================
    // 4. QUERIES & MUTATIONS
    // =========================================================================
    public static List<CollectiveSuggestionRecord> GetPendingSuggestions(string? langCode = null)
    {
        EnsureSchema();
        using var conn = OpenConnection();
        const string sql = @"
SELECT
    s.suggestion_id,
    s.dict_type,
    s.entry_key,
    s.lang_code,
    s.suggested_value,
    s.note,
    s.suggested_by,
    s.machine_name,
    s.status,
    s.vote_count,
    s.created_at,
    s.reviewed_by,
    s.reviewed_at,
    EXISTS (
        SELECT 1
        FROM app.localization_votes v
        WHERE v.suggestion_id = s.suggestion_id
          AND v.voter_key = @VoterKey
    ) AS has_current_user_voted
FROM app.localization_suggestions s
WHERE s.status = 'pending'
  AND (@LangCode IS NULL OR s.lang_code = @LangCode)
ORDER BY s.vote_count DESC, s.created_at DESC;";

        return conn.Query<CollectiveSuggestionRecord>(sql, new
        {
            VoterKey = CurrentVoterKey,
            LangCode = string.IsNullOrWhiteSpace(langCode) ? null : langCode.Trim()
        }).ToList();
    }

    public static List<CollectiveOverrideRecord> GetLiveOverrides(string? langCode = null)
    {
        EnsureSchema();
        using var conn = OpenConnection();
        const string sql = @"
SELECT dict_type, entry_key, lang_code, value, source_suggestion_id, approved_by, approved_at
FROM app.localization_overrides
WHERE (@LangCode IS NULL OR lang_code = @LangCode)
ORDER BY approved_at DESC;";

        return conn.Query<CollectiveOverrideRecord>(sql, new
        {
            LangCode = string.IsNullOrWhiteSpace(langCode) ? null : langCode.Trim()
        }).ToList();
    }

    public static long SubmitSuggestion(string dictType, string entryKey, string langCode, string suggestedValue, string note)
    {
        string normalizedType = NormalizeDictType(dictType);
        if (string.IsNullOrWhiteSpace(entryKey))
            throw new ArgumentException("Entry key is required.", nameof(entryKey));
        if (string.IsNullOrWhiteSpace(langCode))
            throw new ArgumentException("Language code is required.", nameof(langCode));

        if (!ValidateEntryPlaceholders(normalizedType, entryKey, suggestedValue, out string? errorMessage))
            throw new ArgumentException(errorMessage ?? "Invalid wording.", nameof(suggestedValue));

        EnsureSchema();
        using var conn = OpenConnection();
        using var trans = conn.BeginTransaction();

        const string insertSuggestion = @"
INSERT INTO app.localization_suggestions
    (dict_type, entry_key, lang_code, suggested_value, note, suggested_by, machine_name, status, vote_count)
VALUES
    (@DictType, @EntryKey, @LangCode, @SuggestedValue, @Note, @SuggestedBy, @MachineName, 'pending', 1)
RETURNING suggestion_id;";

        long suggestionId = conn.ExecuteScalar<long>(insertSuggestion, new
        {
            DictType = normalizedType,
            EntryKey = Truncate(entryKey.Trim(), 200),
            LangCode = Truncate(langCode.Trim(), 16),
            SuggestedValue = suggestedValue.Trim(),
            Note = (note ?? string.Empty).Trim(),
            SuggestedBy = Truncate(CurrentVoterName, 120),
            MachineName = Truncate(Environment.MachineName ?? string.Empty, 120)
        }, trans);

        const string insertVote = @"
INSERT INTO app.localization_votes (suggestion_id, voter_key, voter_name)
VALUES (@SuggestionId, @VoterKey, @VoterName)
ON CONFLICT (suggestion_id, voter_key) DO NOTHING;";

        conn.Execute(insertVote, new
        {
            SuggestionId = suggestionId,
            VoterKey = Truncate(CurrentVoterKey, 200),
            VoterName = Truncate(CurrentVoterName, 120)
        }, trans);

        trans.Commit();
        return suggestionId;
    }

    public static CollectiveVoteResult ToggleVote(long suggestionId)
    {
        EnsureSchema();
        bool autoApproved = false;
        CollectiveVoteResult result;

        using (var conn = OpenConnection())
        using (var trans = conn.BeginTransaction())
        {
            const string lockSql = @"
SELECT suggestion_id, dict_type, entry_key, lang_code, suggested_value, status, vote_count
FROM app.localization_suggestions
WHERE suggestion_id = @SuggestionId
FOR UPDATE;";

            var suggestion = conn.QueryFirstOrDefault<CollectiveSuggestionRecord>(
                lockSql, new { SuggestionId = suggestionId }, trans);

            if (suggestion == null)
                throw new InvalidOperationException($"Suggestion #{suggestionId} was not found.");

            if (!string.Equals(suggestion.Status, "pending", StringComparison.OrdinalIgnoreCase))
            {
                trans.Commit();
                return new CollectiveVoteResult
                {
                    NewVoteCount = suggestion.VoteCount,
                    UserVoted = false,
                    AutoApproved = false,
                    Status = suggestion.Status
                };
            }

            const string checkVoteSql = @"
SELECT COUNT(1)
FROM app.localization_votes
WHERE suggestion_id = @SuggestionId AND voter_key = @VoterKey;";

            bool alreadyVoted = conn.ExecuteScalar<int>(checkVoteSql, new
            {
                SuggestionId = suggestionId,
                VoterKey = CurrentVoterKey
            }, trans) > 0;

            int newCount;
            bool userVoted;
            string finalStatus = "pending";

            if (alreadyVoted)
            {
                conn.Execute(
                    "DELETE FROM app.localization_votes WHERE suggestion_id = @SuggestionId AND voter_key = @VoterKey;",
                    new { SuggestionId = suggestionId, VoterKey = CurrentVoterKey }, trans);

                newCount = Math.Max(0, suggestion.VoteCount - 1);
                userVoted = false;

                conn.Execute(
                    "UPDATE app.localization_suggestions SET vote_count = @VoteCount WHERE suggestion_id = @SuggestionId;",
                    new { VoteCount = newCount, SuggestionId = suggestionId }, trans);
            }
            else
            {
                conn.Execute(
                    @"INSERT INTO app.localization_votes (suggestion_id, voter_key, voter_name)
                      VALUES (@SuggestionId, @VoterKey, @VoterName)
                      ON CONFLICT (suggestion_id, voter_key) DO NOTHING;",
                    new
                    {
                        SuggestionId = suggestionId,
                        VoterKey = Truncate(CurrentVoterKey, 200),
                        VoterName = Truncate(CurrentVoterName, 120)
                    }, trans);

                newCount = suggestion.VoteCount + 1;
                userVoted = true;

                if (newCount >= ConsensusVoteThreshold)
                {
                    string approverLabel = $"consensus ({newCount} votes)";
                    PromoteOverrideInTransaction(
                        conn,
                        trans,
                        suggestion.DictType,
                        suggestion.EntryKey,
                        suggestion.LangCode,
                        suggestion.SuggestedValue,
                        suggestionId,
                        "consensus_approved",
                        approverLabel);

                    conn.Execute(
                        @"UPDATE app.localization_suggestions
                          SET vote_count = @VoteCount,
                              status = 'approved',
                              reviewed_by = @ReviewedBy,
                              reviewed_at = CURRENT_TIMESTAMP
                          WHERE suggestion_id = @SuggestionId;",
                        new
                        {
                            VoteCount = newCount,
                            ReviewedBy = Truncate(approverLabel, 120),
                            SuggestionId = suggestionId
                        }, trans);

                    autoApproved = true;
                    finalStatus = "approved";
                }
                else
                {
                    conn.Execute(
                        "UPDATE app.localization_suggestions SET vote_count = @VoteCount WHERE suggestion_id = @SuggestionId;",
                        new { VoteCount = newCount, SuggestionId = suggestionId }, trans);
                }
            }

            trans.Commit();
            result = new CollectiveVoteResult
            {
                NewVoteCount = newCount,
                UserVoted = userVoted,
                AutoApproved = autoApproved,
                Status = finalStatus
            };
        }

        if (autoApproved)
            SyncFromServer(reloadLoc: true);

        return result;
    }

#if DEBUG
    // =========================================================================
    // 5. DEBUG-ONLY ADMIN ACTIONS
    // =========================================================================
    public static void AdminApproveSuggestion(long suggestionId, string? editedValue = null)
    {
        EnsureSchema();
        using (var conn = OpenConnection())
        using (var trans = conn.BeginTransaction())
        {
            const string lockSql = @"
SELECT suggestion_id, dict_type, entry_key, lang_code, suggested_value, status, vote_count
FROM app.localization_suggestions
WHERE suggestion_id = @SuggestionId
FOR UPDATE;";

            var suggestion = conn.QueryFirstOrDefault<CollectiveSuggestionRecord>(
                lockSql, new { SuggestionId = suggestionId }, trans);

            if (suggestion == null)
                throw new InvalidOperationException($"Suggestion #{suggestionId} was not found.");

            string finalValue = editedValue != null ? editedValue.Trim() : suggestion.SuggestedValue;
            if (!ValidateEntryPlaceholders(suggestion.DictType, suggestion.EntryKey, finalValue, out string? errorMessage))
                throw new ArgumentException(errorMessage ?? "Invalid wording.", nameof(editedValue));

            string adminName = Truncate(CurrentVoterName, 120);

            PromoteOverrideInTransaction(
                conn,
                trans,
                suggestion.DictType,
                suggestion.EntryKey,
                suggestion.LangCode,
                finalValue,
                suggestionId,
                "admin_approved",
                adminName);

            conn.Execute(
                @"UPDATE app.localization_suggestions
                  SET suggested_value = @SuggestedValue,
                      status = 'approved',
                      reviewed_by = @ReviewedBy,
                      reviewed_at = CURRENT_TIMESTAMP
                  WHERE suggestion_id = @SuggestionId;",
                new
                {
                    SuggestedValue = finalValue,
                    ReviewedBy = adminName,
                    SuggestionId = suggestionId
                }, trans);

            trans.Commit();
        }

        SyncFromServer(reloadLoc: true);
    }

    public static void AdminRejectSuggestion(long suggestionId)
    {
        EnsureSchema();
        using var conn = OpenConnection();
        conn.Execute(
            @"UPDATE app.localization_suggestions
              SET status = 'rejected',
                  reviewed_by = @ReviewedBy,
                  reviewed_at = CURRENT_TIMESTAMP
              WHERE suggestion_id = @SuggestionId;",
            new
            {
                ReviewedBy = Truncate(CurrentVoterName, 120),
                SuggestionId = suggestionId
            });
    }

    public static void AdminRevertOverride(string dictType, string entryKey, string langCode)
    {
        string normalizedType = NormalizeDictType(dictType);
        EnsureSchema();
        using (var conn = OpenConnection())
        using (var trans = conn.BeginTransaction())
        {
            string? previousValue = conn.ExecuteScalar<string?>(
                @"SELECT value FROM app.localization_overrides
                  WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode;",
                new { DictType = normalizedType, EntryKey = entryKey, LangCode = langCode }, trans);

            conn.Execute(
                @"DELETE FROM app.localization_overrides
                  WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode;",
                new { DictType = normalizedType, EntryKey = entryKey, LangCode = langCode }, trans);

            conn.Execute(
                @"INSERT INTO app.localization_history
                    (dict_type, entry_key, lang_code, previous_value, new_value, action, changed_by)
                  VALUES
                    (@DictType, @EntryKey, @LangCode, @PreviousValue, NULL, 'reverted', @ChangedBy);",
                new
                {
                    DictType = normalizedType,
                    EntryKey = Truncate(entryKey, 200),
                    LangCode = Truncate(langCode, 16),
                    PreviousValue = previousValue,
                    ChangedBy = Truncate(CurrentVoterName, 120)
                }, trans);

            trans.Commit();
        }

        SyncFromServer(reloadLoc: true);
    }

    public static CollectiveBatchSubmitResult SubmitAiSuggestions(IEnumerable<CollectiveSuggestionInput> items)
    {
        var result = new CollectiveBatchSubmitResult();
        if (items == null) return result;

        EnsureSchema();
        using var conn = OpenConnection();

        const string skipCheckSql = @"
SELECT
    EXISTS (
        SELECT 1 FROM app.localization_overrides
        WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode
    )
    OR EXISTS (
        SELECT 1 FROM app.localization_suggestions
        WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode
          AND status = 'pending'
    )
    OR EXISTS (
        SELECT 1 FROM app.localization_suggestions
        WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode
          AND status = 'rejected' AND suggested_value = @SuggestedValue
    );";

        const string insertSuggestionSql = @"
INSERT INTO app.localization_suggestions
    (dict_type, entry_key, lang_code, suggested_value, note, suggested_by, machine_name, status, vote_count)
VALUES
    (@DictType, @EntryKey, @LangCode, @SuggestedValue, @Note, @SuggestedBy, @MachineName, 'pending', 1)
RETURNING suggestion_id;";

        const string insertVoteSql = @"
INSERT INTO app.localization_votes (suggestion_id, voter_key, voter_name)
VALUES (@SuggestionId, @VoterKey, @VoterName)
ON CONFLICT (suggestion_id, voter_key) DO NOTHING;";

        foreach (var item in items)
        {
            if (item == null) continue;

            string normalizedType = NormalizeDictType(item.DictType ?? string.Empty);
            string entryKey = Truncate((item.EntryKey ?? string.Empty).Trim(), 200);
            string langCode = Truncate((item.LangCode ?? string.Empty).Trim(), 16);
            string suggestedValue = (item.SuggestedValue ?? string.Empty).Trim();
            string note = (item.Note ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(entryKey))
            {
                result.Errors.Add($"{normalizedType}/<empty>/{langCode}: Entry key is required.");
                continue;
            }
            if (string.IsNullOrWhiteSpace(langCode))
            {
                result.Errors.Add($"{normalizedType}/{entryKey}/<empty>: Language code is required.");
                continue;
            }

            if (!ValidateEntryPlaceholders(normalizedType, entryKey, suggestedValue, out string? err))
            {
                result.Errors.Add($"{normalizedType}/{entryKey}/{langCode}: {err ?? "Invalid wording."}");
                continue;
            }

            bool shouldSkip = conn.ExecuteScalar<bool>(skipCheckSql, new
            {
                DictType = normalizedType,
                EntryKey = entryKey,
                LangCode = langCode,
                SuggestedValue = suggestedValue
            });

            if (shouldSkip)
            {
                result.Skipped++;
                continue;
            }

            using var trans = conn.BeginTransaction();
            long suggestionId = conn.ExecuteScalar<long>(insertSuggestionSql, new
            {
                DictType = normalizedType,
                EntryKey = entryKey,
                LangCode = langCode,
                SuggestedValue = suggestedValue,
                Note = note,
                SuggestedBy = Truncate(AiHelperName, 120),
                MachineName = "Antigravity"
            }, trans);

            conn.Execute(insertVoteSql, new
            {
                SuggestionId = suggestionId,
                VoterKey = Truncate(AiHelperVoterKey, 200),
                VoterName = Truncate(AiHelperName, 120)
            }, trans);

            trans.Commit();
            result.Submitted++;
        }

        return result;
    }

    public static CollectiveDictionaryStateExport ExportStateForReview(int historyLimit = 100)
    {
        int limit = historyLimit > 0 ? historyLimit : 100;
        EnsureSchema();
        using var conn = OpenConnection();

        const string overridesSql = @"
SELECT dict_type, entry_key, lang_code, value, source_suggestion_id, approved_by, approved_at
FROM app.localization_overrides
ORDER BY lang_code, dict_type, entry_key;";

        const string pendingSql = @"
SELECT
    suggestion_id, dict_type, entry_key, lang_code, suggested_value, note,
    suggested_by, machine_name, status, vote_count, created_at, reviewed_by, reviewed_at
FROM app.localization_suggestions
WHERE status = 'pending'
ORDER BY created_at DESC;";

        const string rejectedSql = @"
SELECT
    suggestion_id, dict_type, entry_key, lang_code, suggested_value, note,
    suggested_by, machine_name, status, vote_count, created_at, reviewed_by, reviewed_at
FROM app.localization_suggestions
WHERE status = 'rejected'
ORDER BY created_at DESC;";

        const string historySql = @"
SELECT history_id, dict_type, entry_key, lang_code, previous_value, new_value, action, changed_by, changed_at
FROM app.localization_history
ORDER BY changed_at DESC
LIMIT @Limit;";

        return new CollectiveDictionaryStateExport
        {
            ExportedAtUtc = DateTime.UtcNow,
            Overrides = conn.Query<CollectiveOverrideRecord>(overridesSql).ToList(),
            PendingSuggestions = conn.Query<CollectiveSuggestionRecord>(pendingSql).ToList(),
            RejectedSuggestions = conn.Query<CollectiveSuggestionRecord>(rejectedSql).ToList(),
            RecentHistory = conn.Query<CollectiveHistoryRecord>(historySql, new { Limit = limit }).ToList()
        };
    }
#endif

    // =========================================================================
    // HELPERS
    // =========================================================================
    private static void PromoteOverrideInTransaction(
        NpgsqlConnection conn,
        NpgsqlTransaction trans,
        string dictType,
        string entryKey,
        string langCode,
        string newValue,
        long sourceSuggestionId,
        string action,
        string changedBy)
    {
        string normalizedType = NormalizeDictType(dictType);

        string? previousValue = conn.ExecuteScalar<string?>(
            @"SELECT value FROM app.localization_overrides
              WHERE dict_type = @DictType AND entry_key = @EntryKey AND lang_code = @LangCode;",
            new { DictType = normalizedType, EntryKey = entryKey, LangCode = langCode }, trans);

        conn.Execute(
            @"INSERT INTO app.localization_overrides
                (dict_type, entry_key, lang_code, value, source_suggestion_id, approved_by, approved_at)
              VALUES
                (@DictType, @EntryKey, @LangCode, @Value, @SourceSuggestionId, @ApprovedBy, CURRENT_TIMESTAMP)
              ON CONFLICT (dict_type, entry_key, lang_code) DO UPDATE SET
                value = EXCLUDED.value,
                source_suggestion_id = EXCLUDED.source_suggestion_id,
                approved_by = EXCLUDED.approved_by,
                approved_at = CURRENT_TIMESTAMP;",
            new
            {
                DictType = normalizedType,
                EntryKey = Truncate(entryKey, 200),
                LangCode = Truncate(langCode, 16),
                Value = newValue,
                SourceSuggestionId = sourceSuggestionId,
                ApprovedBy = Truncate(changedBy, 120)
            }, trans);

        conn.Execute(
            @"INSERT INTO app.localization_history
                (dict_type, entry_key, lang_code, previous_value, new_value, action, changed_by)
              VALUES
                (@DictType, @EntryKey, @LangCode, @PreviousValue, @NewValue, @Action, @ChangedBy);",
            new
            {
                DictType = normalizedType,
                EntryKey = Truncate(entryKey, 200),
                LangCode = Truncate(langCode, 16),
                PreviousValue = previousValue,
                NewValue = newValue,
                Action = action,
                ChangedBy = Truncate(changedBy, 120)
            }, trans);
    }

    private static string NormalizeDictType(string dictType)
        => string.Equals(dictType, DictTypeTooltips, StringComparison.OrdinalIgnoreCase)
            ? DictTypeTooltips
            : DictTypeUi;

    private static string Truncate(string? value, int maxLength)
    {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value.Length <= maxLength ? value : value.Substring(0, maxLength);
    }
}
