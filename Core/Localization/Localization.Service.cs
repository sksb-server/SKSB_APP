namespace SKSB_App.Core.Localization;

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Text.Json;
using System.Threading;
using SKSB_App.Core.Bootstrap.Storage;

/// <summary>
/// Central UI string lookup for SKSB_App.
///
/// Dictionaries come from the embedded <c>Localization.ui.&lt;code&gt;.json</c> resources, then any external
/// <c>ui.&lt;code&gt;.json</c> override files are merged on top (admins can tweak text without recompiling).
/// Active dictionaries are immutable once published and swapped atomically, so <see cref="T(string)"/>
/// is safe from any thread (e.g. the LoadingOverlay background STA thread).
///
/// <see cref="LanguageChanged"/> may be raised on a thread-pool thread (when another process / add-in
/// changes <c>user_prefs.json</c>); subscribers must marshal to their own UI thread.
/// </summary>
public static class Loc
{
    // =========================================================================
    // 1. STATE
    // =========================================================================
    private const string ResourcePrefix = "Localization.ui.";
    private const string TooltipResourcePrefix = "Localization.tooltips.";
    private const string UiFilePrefix = "ui.";
    private const string TooltipFilePrefix = "tooltips.";

    private static readonly object _gate = new();
    private static volatile Dictionary<string, string> _fallback = new(StringComparer.Ordinal);
    private static volatile Dictionary<string, string> _active = new(StringComparer.Ordinal);
    private static volatile Dictionary<string, string> _tipFallback = new(StringComparer.Ordinal);
    private static volatile Dictionary<string, string> _tipActive = new(StringComparer.Ordinal);
    private static volatile int _currentLanguage;

    private static FileSystemWatcher? _prefsWatcher;
    private static Timer? _prefsDebounce;
    private static FileSystemEventHandler? _prefsChangedHandler;   // pinned static delegates (CLAUDE.md Rule 4)
    private static RenamedEventHandler? _prefsRenamedHandler;

    private static FileSystemWatcher? _dictWatcher;
    private static Timer? _dictDebounce;
    private static FileSystemEventHandler? _dictChangedHandler;    // pinned static delegates (CLAUDE.md Rule 4)
    private static RenamedEventHandler? _dictRenamedHandler;

    public static event Action? LanguageChanged;

    public static AppLanguage CurrentLanguage
    {
        get => (AppLanguage)_currentLanguage;
        private set => _currentLanguage = (int)value;
    }

    public static string CurrentCode => AppLanguageInfo.ToCode(CurrentLanguage);

    static Loc()
    {
        try
        {
            var language = AppLanguageInfo.FromCode(UserPreferences.Load().Language);
            var fallback = LoadDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
            _fallback = fallback;
            _active = language == AppLanguage.En ? fallback : LoadDictionary(AppLanguageInfo.ToCode(language));
            var tipFallback = LoadTooltipDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
            _tipFallback = tipFallback;
            _tipActive = language == AppLanguage.En ? tipFallback : LoadTooltipDictionary(AppLanguageInfo.ToCode(language));
            CurrentLanguage = language;
        }
        catch
        {
            CurrentLanguage = AppLanguage.En;
        }
    }

    // =========================================================================
    // 2. LANGUAGE SWITCHING
    // =========================================================================
    public static void SetLanguage(AppLanguage language, bool persist = true)
    {
        bool changed;
        lock (_gate)
        {
            changed = language != CurrentLanguage;
            if (changed)
            {
                // Reload English too so edited override files are picked up on every switch.
                var fallback = LoadDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
                var active = language == AppLanguage.En ? fallback : LoadDictionary(AppLanguageInfo.ToCode(language));
                var tipFallback = LoadTooltipDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
                var tipActive = language == AppLanguage.En ? tipFallback : LoadTooltipDictionary(AppLanguageInfo.ToCode(language));
                _fallback = fallback;
                _active = active;
                _tipFallback = tipFallback;
                _tipActive = tipActive;
                CurrentLanguage = language;
            }
        }

        if (persist)
        {
            var prefs = UserPreferences.Load();
            string code = AppLanguageInfo.ToCode(language);
            if (!string.Equals(prefs.Language, code, StringComparison.OrdinalIgnoreCase))
            {
                prefs.Language = code;
                UserPreferences.Save(prefs);
            }
        }

        if (changed) RaiseLanguageChanged();
    }

    /// <summary>Re-reads <c>user_prefs.json</c> and switches language if another process changed it.</summary>
    public static void ReloadFromPreferences()
    {
        try
        {
            var language = AppLanguageInfo.FromCode(UserPreferences.Load().Language);
            if (language != CurrentLanguage) SetLanguage(language, persist: false);
        }
        catch { }
    }

    /// <summary>Reloads embedded + collective + external dictionaries for the active language.</summary>
    public static void ReloadDictionaries(bool raiseEvent = true)
    {
        lock (_gate)
        {
            var language = CurrentLanguage;
            var fallback = LoadDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
            var active = language == AppLanguage.En ? fallback : LoadDictionary(AppLanguageInfo.ToCode(language));
            var tipFallback = LoadTooltipDictionary(AppLanguageInfo.ToCode(AppLanguage.En));
            var tipActive = language == AppLanguage.En ? tipFallback : LoadTooltipDictionary(AppLanguageInfo.ToCode(language));
            _fallback = fallback;
            _active = active;
            _tipFallback = tipFallback;
            _tipActive = tipActive;
        }

        if (raiseEvent) RaiseLanguageChanged();
    }

    /// <summary>
    /// Starts (once) watchers on <c>user_prefs.json</c> and <c>collective_dict_cache.json</c> so language switches
    /// and collective dictionary syncs in Portal or another add-in apply live here too.
    /// <see cref="LanguageChanged"/> then fires on a thread-pool thread.
    /// </summary>
    public static void WatchPreferences()
    {
        lock (_gate)
        {
            if (_prefsWatcher == null)
            {
                try
                {
                    Directory.CreateDirectory(UserPreferences.DirectoryPath);
                    _prefsDebounce = new Timer(_ => ReloadFromPreferences(), null, Timeout.Infinite, Timeout.Infinite);
                    _prefsChangedHandler = OnPrefsFileChanged;
                    _prefsRenamedHandler = OnPrefsFileRenamed;

                    var watcher = new FileSystemWatcher(UserPreferences.DirectoryPath, UserPreferences.FileName)
                    {
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size | NotifyFilters.CreationTime,
                        IncludeSubdirectories = false
                    };
                    watcher.Changed += _prefsChangedHandler;
                    watcher.Created += _prefsChangedHandler;
                    watcher.Renamed += _prefsRenamedHandler;
                    watcher.EnableRaisingEvents = true;
                    _prefsWatcher = watcher;
                }
                catch
                {
                    _prefsWatcher = null;
                }
            }

            if (_dictWatcher == null)
            {
                try
                {
                    Directory.CreateDirectory(UserPreferences.DirectoryPath);
                    _dictDebounce = new Timer(_ => ReloadDictionaries(raiseEvent: true), null, Timeout.Infinite, Timeout.Infinite);
                    _dictChangedHandler = OnDictFileChanged;
                    _dictRenamedHandler = OnDictFileRenamed;

                    var watcher = new FileSystemWatcher(UserPreferences.DirectoryPath, CollectiveDictionaryRepository.CacheFileName)
                    {
                        NotifyFilter = NotifyFilters.LastWrite | NotifyFilters.FileName | NotifyFilters.Size | NotifyFilters.CreationTime,
                        IncludeSubdirectories = false
                    };
                    watcher.Changed += _dictChangedHandler;
                    watcher.Created += _dictChangedHandler;
                    watcher.Renamed += _dictRenamedHandler;
                    watcher.EnableRaisingEvents = true;
                    _dictWatcher = watcher;
                }
                catch
                {
                    _dictWatcher = null;
                }
            }
        }
    }

    private static void OnPrefsFileChanged(object sender, FileSystemEventArgs e) => SchedulePrefsReload();

    private static void OnPrefsFileRenamed(object sender, RenamedEventArgs e)
    {
        if (string.Equals(e.Name, UserPreferences.FileName, StringComparison.OrdinalIgnoreCase)) SchedulePrefsReload();
    }

    private static void SchedulePrefsReload()
    {
        try { _prefsDebounce?.Change(300, Timeout.Infinite); } catch { }
    }

    private static void OnDictFileChanged(object sender, FileSystemEventArgs e) => ScheduleDictReload();

    private static void OnDictFileRenamed(object sender, RenamedEventArgs e)
    {
        if (string.Equals(e.Name, CollectiveDictionaryRepository.CacheFileName, StringComparison.OrdinalIgnoreCase)) ScheduleDictReload();
    }

    private static void ScheduleDictReload()
    {
        try { _dictDebounce?.Change(300, Timeout.Infinite); } catch { }
    }

    private static void RaiseLanguageChanged()
    {
        var handlers = LanguageChanged;
        if (handlers == null) return;
        foreach (Delegate d in handlers.GetInvocationList())
        {
            try { ((Action)d)(); }
            catch (Exception ex) { System.Diagnostics.Debug.WriteLine($"[Loc] LanguageChanged subscriber failed: {ex.Message}"); }
        }
    }

    // =========================================================================
    // 3. LOOKUP
    // =========================================================================
    public static string T(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        return TryT(key, out string value) ? value : key;
    }

    public static string T(string key, params object[] args)
    {
        string template = T(key);
        if (args == null || args.Length == 0) return template;
        try
        {
            return string.Format(CultureInfo.CurrentCulture, template, args);
        }
        catch (FormatException)
        {
            return template;
        }
    }

    /// <summary>Looks up <paramref name="key"/> in the active language, then English. False when neither has it.</summary>
    public static bool TryT(string key, out string value)
    {
        if (!string.IsNullOrEmpty(key))
        {
            var active = _active;
            if (active.TryGetValue(key, out var a) && a != null) { value = a; return true; }
            var fallback = _fallback;
            if (fallback.TryGetValue(key, out var f) && f != null) { value = f; return true; }
        }
        value = string.Empty;
        return false;
    }

    /// <summary>Translates known Manager module names ("StaffManager" or "Staff Manager"); unknown names pass through.</summary>
    public static string ModuleDisplayName(string moduleName)
    {
        if (string.IsNullOrWhiteSpace(moduleName)) return moduleName ?? string.Empty;
        string compact = moduleName.Replace(" ", string.Empty).Trim();
        string? id = compact.ToLowerInvariant() switch
        {
            "schedulemanager" => "schedule",
            "projectmanager" => "project",
            "staffmanager" => "staff",
            "vehiclemanager" => "vehicle",
            _ => null
        };
        if (id == null) return moduleName;
        return TryT($"module.{id}.name", out string name) ? name : moduleName;
    }

    // =========================================================================
    // 3a. TOOLTIP / CELL GUIDE LOOKUP (tooltips.<code>.json)
    // =========================================================================
    public static string Tip(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        return TryTip(key, out string value) ? value : key;
    }

    public static string Tip(string key, params object[] args)
    {
        string template = Tip(key);
        if (args == null || args.Length == 0) return template;
        try
        {
            return string.Format(CultureInfo.CurrentCulture, template, args);
        }
        catch (FormatException)
        {
            return template;
        }
    }

    /// <summary>Looks up <paramref name="key"/> in the active tooltip dictionary, then English, then any cached language for custom tips. False when none has it.</summary>
    public static bool TryTip(string key, out string value)
    {
        if (!string.IsNullOrEmpty(key))
        {
            var active = _tipActive;
            if (active.TryGetValue(key, out var a) && a != null) { value = a; return true; }
            var fallback = _tipFallback;
            if (fallback.TryGetValue(key, out var f) && f != null) { value = f; return true; }
            if (CollectiveDictionaryRepository.IsCustomTipKey(key)
                && CollectiveDictionaryRepository.TryGetAnyCachedTooltipValue(key, out var anyVal))
            {
                value = anyVal;
                return true;
            }
        }
        value = string.Empty;
        return false;
    }

    /// <summary>
    /// Returns all spotlight general tip keys: built-in <c>guide.spotlight.gen.&lt;id&gt;</c> keys
    /// followed by any approved custom community tips (<c>guide.spotlight.custom.*</c>), deduplicated.
    /// </summary>
    public static IReadOnlyList<string> GetAllGeneralTipKeys(IReadOnlyList<string> builtInIds)
    {
        var keys = new List<string>(builtInIds?.Count ?? 0);
        var seen = new HashSet<string>(StringComparer.Ordinal);
        if (builtInIds != null)
        {
            foreach (string id in builtInIds)
            {
                if (string.IsNullOrWhiteSpace(id)) continue;
                string fullKey = $"guide.spotlight.gen.{id}";
                if (seen.Add(fullKey)) keys.Add(fullKey);
            }
        }

        foreach (string customKey in CollectiveDictionaryRepository.GetAllCachedTooltipKeysByPrefix(
            CollectiveDictionaryRepository.CustomTipKeyPrefix))
        {
            if (CollectiveDictionaryRepository.IsCustomContextTipKey(customKey)) continue;
            if (seen.Add(customKey)) keys.Add(customKey);
        }

        return keys;
    }

    /// <summary>
    /// Returns all deduplicated approved custom contextual tip keys matching <paramref name="trigger"/>
    /// (<c>guide.spotlight.custom.ctx.&lt;trigger&gt;.*</c>), or an empty list if <paramref name="trigger"/> is blank.
    /// </summary>
    public static IReadOnlyList<string> GetCustomContextTipKeys(string trigger)
    {
        if (string.IsNullOrWhiteSpace(trigger)) return Array.Empty<string>();

        string prefix = CollectiveDictionaryRepository.GetCustomContextTipPrefix(trigger);
        var keys = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string customKey in CollectiveDictionaryRepository.GetAllCachedTooltipKeysByPrefix(prefix))
        {
            if (seen.Add(customKey)) keys.Add(customKey);
        }
        return keys;
    }

    /// <summary>
    /// Translated column name / 1-line description: <c>col.&lt;table&gt;.&lt;col&gt;.*</c> first, then <c>col.&lt;col&gt;.*</c>.
    /// Either value is null when no dictionary entry exists.
    /// </summary>
    public static (string? DisplayName, string? Description) ResolveColumnGuide(string? tableName, string? columnHeader)
    {
        string normCol = NormalizeKeyPart(columnHeader);
        if (normCol.Length == 0) return (null, null);
        string normTable = NormalizeKeyPart(tableName);

        string? name = null;
        string? desc = null;
        if (normTable.Length > 0)
        {
            if (TryTip($"col.{normTable}.{normCol}.name", out string tn)) name = tn;
            if (TryTip($"col.{normTable}.{normCol}.desc", out string td)) desc = td;
        }
        if (name == null && TryTip($"col.{normCol}.name", out string n)) name = n;
        if (desc == null && TryTip($"col.{normCol}.desc", out string d)) desc = d;
        return (name, desc);
    }

    /// <summary>Translated table display name (<c>table.&lt;table&gt;.name</c>); unknown tables pass through.</summary>
    public static string TableDisplayName(string? tableName)
    {
        if (string.IsNullOrWhiteSpace(tableName)) return tableName ?? string.Empty;
        string norm = NormalizeKeyPart(tableName);
        return norm.Length > 0 && TryTip($"table.{norm}.name", out string name) ? name : tableName;
    }

    /// <summary>Lower-case ASCII letters and digits only ("Staff ID" -> "staffid").</summary>
    private static string NormalizeKeyPart(string? raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return string.Empty;
        var sb = new System.Text.StringBuilder(raw.Length);
        foreach (char ch in raw)
        {
            if ((ch >= 'a' && ch <= 'z') || (ch >= '0' && ch <= '9')) sb.Append(ch);
            else if (ch >= 'A' && ch <= 'Z') sb.Append(char.ToLowerInvariant(ch));
        }
        return sb.ToString();
    }

    // =========================================================================
    // 4. DICTIONARY LOADING (embedded resource + collective + external overrides)
    // =========================================================================
    /// <summary>Loads the embedded base dictionary ONLY (no collective or external file overrides).</summary>
    public static Dictionary<string, string> GetBaseDictionary(string dictType, string langCode)
    {
        bool isTooltip = string.Equals(dictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);
        return LoadEmbeddedDictionary(isTooltip ? TooltipResourcePrefix : ResourcePrefix, isTooltip ? TooltipFilePrefix : UiFilePrefix, langCode);
    }

    /// <summary>Returns the embedded base entry for <paramref name="key"/> (or empty string if missing).</summary>
    public static string GetBaseEntry(string dictType, string langCode, string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        var baseDict = GetBaseDictionary(dictType, langCode);
        return baseDict.TryGetValue(key, out string? value) && value != null ? value : string.Empty;
    }

    /// <summary>Loads the full effective dictionary (embedded + collective cache + external overrides).</summary>
    public static Dictionary<string, string> GetEffectiveDictionary(string dictType, string langCode)
    {
        bool isTooltip = string.Equals(dictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);
        return isTooltip ? LoadTooltipDictionary(langCode) : LoadDictionary(langCode);
    }

    private static Dictionary<string, string> LoadDictionary(string code)
        => LoadDictionary(ResourcePrefix, UiFilePrefix, code);

    private static Dictionary<string, string> LoadTooltipDictionary(string code)
        => LoadDictionary(TooltipResourcePrefix, TooltipFilePrefix, code);

    private static Dictionary<string, string> LoadEmbeddedDictionary(string resourcePrefix, string filePrefix, string code)
    {
        var dict = new Dictionary<string, string>(StringComparer.Ordinal);
        try
        {
            using Stream? stream = typeof(Loc).Assembly.GetManifestResourceStream(resourcePrefix + code + ".json");
            if (stream != null)
            {
                using var reader = new StreamReader(stream);
                MergeJson(dict, reader.ReadToEnd());
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Loc] Embedded dictionary '{filePrefix}{code}' failed: {ex.Message}");
        }
        return dict;
    }

    private static Dictionary<string, string> LoadDictionary(string resourcePrefix, string filePrefix, string code)
    {
        var dict = LoadEmbeddedDictionary(resourcePrefix, filePrefix, code);

        string dictType = filePrefix == TooltipFilePrefix
            ? CollectiveDictionaryRepository.DictTypeTooltips
            : CollectiveDictionaryRepository.DictTypeUi;

        foreach (var kv in CollectiveDictionaryRepository.LoadCachedOverrides(dictType, code))
        {
            dict[kv.Key] = kv.Value;
        }

        foreach (string path in GetOverridePaths($"{filePrefix}{code}.json"))
        {
            try
            {
                if (!File.Exists(path)) continue;
                string json;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(fs))
                {
                    json = reader.ReadToEnd();
                }
                MergeJson(dict, json);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[Loc] Override '{path}' failed: {ex.Message}");
            }
        }

        return dict;
    }

    /// <summary>Override files in increasing priority (later files win).</summary>
    private static IEnumerable<string> GetOverridePaths(string fileName)
    {
        var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var paths = new List<string>();

        void Add(Func<string> build)
        {
            try
            {
                string full = Path.GetFullPath(build());
                if (seen.Add(full)) paths.Add(full);
            }
            catch { }
        }

        Add(() => Path.Combine(AppContext.BaseDirectory, "Localization", fileName));
        Add(() => Path.Combine(FolderRegistry.SKSB_App.GetPath(), "Localization", fileName));
#if DEBUG
        Add(() => Path.Combine(FolderRegistry.SKSB_App.GetPath(), "Core", "Localization", fileName));
#endif
        return paths;
    }

    private static void MergeJson(Dictionary<string, string> target, string json)
    {
        if (string.IsNullOrWhiteSpace(json)) return;
        var options = new JsonDocumentOptions { AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip };
        using var doc = JsonDocument.Parse(json, options);
        if (doc.RootElement.ValueKind != JsonValueKind.Object) return;
        foreach (var prop in doc.RootElement.EnumerateObject())
        {
            if (prop.Value.ValueKind == JsonValueKind.String)
                target[prop.Name] = prop.Value.GetString() ?? string.Empty;
        }
    }
}
