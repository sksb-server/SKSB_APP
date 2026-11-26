namespace SKSB_App.Core.Localization;

using System;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

/// <summary>Per-PC user preferences shared by Portal.exe and every Excel add-in.</summary>
public sealed class UserPreferencesData
{
    /// <summary>UI language code (<c>en</c>, <c>ms</c>, <c>zh</c>, <c>id</c>, <c>ms-en</c>).</summary>
    public string Language { get; set; } = "en";

    /// <summary>Whether the Cell Guide task pane opens automatically when a manager workbook starts. Defaults to true.</summary>
    public bool ShowGuideOnStartup { get; set; } = true;

    /// <summary>Round-trips keys written by newer versions (e.g. Step 2 guide preferences) so they are never lost.</summary>
    [JsonExtensionData]
    public Dictionary<string, JsonElement>? Extra { get; set; }
}

/// <summary>
/// Reads / writes <c>%LOCALAPPDATA%\SKSB_App\Data\user_prefs.json</c>. The path is fixed (independent of
/// Debug / Release folder resolution) so Portal and all Excel add-ins share one preference file.
/// Every operation is exception-safe: a locked or corrupt file never breaks startup.
/// </summary>
public static class UserPreferences
{
    private static readonly object _ioGate = new();

    private static readonly JsonSerializerOptions _jsonOptions = new()
    {
        WriteIndented = true,
        PropertyNameCaseInsensitive = true,
        PropertyNamingPolicy = JsonNamingPolicy.CamelCase,
        ReadCommentHandling = JsonCommentHandling.Skip,
        AllowTrailingCommas = true
    };

    /// <summary>
    /// <c>%LOCALAPPDATA%\SKSB_App\Data</c>: outside every build / publish output folder, so
    /// <c>dotnet clean</c> and client updates never wipe user preferences.
    /// </summary>
    public static string DirectoryPath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SKSB_App",
        "Data");

    public const string FileName = "user_prefs.json";

    public static string FilePath { get; } = Path.Combine(DirectoryPath, FileName);

    /// <summary>Pre-Data-folder location (<c>%LOCALAPPDATA%\SKSB_App\user_prefs.json</c>), migrated once on first load.</summary>
    public static string LegacyFilePath { get; } = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "SKSB_App",
        FileName);

    public static UserPreferencesData Load()
    {
        try
        {
            lock (_ioGate)
            {
                string sourcePath = FilePath;
                if (!File.Exists(FilePath))
                {
                    if (!File.Exists(LegacyFilePath)) return new UserPreferencesData();
                    sourcePath = LegacyFilePath;
                    TryMigrateLegacyFile();
                    if (File.Exists(FilePath)) sourcePath = FilePath;
                }

                string json;
                using (var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (var reader = new StreamReader(fs))
                {
                    json = reader.ReadToEnd();
                }

                if (string.IsNullOrWhiteSpace(json)) return new UserPreferencesData();
                var data = JsonSerializer.Deserialize<UserPreferencesData>(json, _jsonOptions) ?? new UserPreferencesData();
                if (string.IsNullOrWhiteSpace(data.Language)) data.Language = "en";
                return data;
            }
        }
        catch
        {
            return new UserPreferencesData();
        }
    }

    /// <summary>One-time copy of the legacy preferences file into the Data folder. I/O errors are ignored.</summary>
    private static void TryMigrateLegacyFile()
    {
        try
        {
            Directory.CreateDirectory(DirectoryPath);
            string tempPath = FilePath + "." + Environment.ProcessId + ".migrate.tmp";
            File.Copy(LegacyFilePath, tempPath, overwrite: true);
            if (File.Exists(FilePath))
            {
                File.Delete(tempPath); // another process migrated first
                return;
            }
            File.Move(tempPath, FilePath, overwrite: false);
        }
        catch
        {
        }
    }

    public static bool Save(UserPreferencesData data)
    {
        if (data == null) return false;
        try
        {
            lock (_ioGate)
            {
                Directory.CreateDirectory(DirectoryPath);
                string json = JsonSerializer.Serialize(data, _jsonOptions);

                // Write-then-move keeps the file intact if another process reads it mid-write.
                string tempPath = FilePath + "." + Environment.ProcessId + ".tmp";
                File.WriteAllText(tempPath, json);
                File.Move(tempPath, FilePath, overwrite: true);
                return true;
            }
        }
        catch
        {
            return false;
        }
    }
}
