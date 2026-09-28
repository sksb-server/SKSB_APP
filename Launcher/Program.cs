using Microsoft.Win32;
using System;
using System.Diagnostics;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;

using SKSB_App.Core.ExcelManager;
using SKSB_App.Core.FileManager.Constants;

namespace SKSB_App.Launcher;

static class Program
{
    private static readonly string RootPath = Folder.SKSB_App.GetPath();
    private static readonly string TemplatesFolder = Folder.Templates.GetPath();
    private static readonly string AddInsFolder = Folder.AddIns.GetPath();
    private static readonly string VersionFilePath = Path.Combine(RootPath, "version.json");

    [STAThread]
    static void Main()
    {
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        try
        {
            EnsureDirectoryStructure();

            // Run provision/update lifecycle
            SyncAppVersion().GetAwaiter().GetResult();

            // All checks complete; launch Portal
            LaunchOperationsPortal();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Failed to initialise SKSB App:\n\n{ex.Message}",
                "Deployment Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private static void EnsureDirectoryStructure()
    {
        Directory.CreateDirectory(RootPath);
        Directory.CreateDirectory(TemplatesFolder);
        Directory.CreateDirectory(AddInsFolder);
    }

    private static async Task SyncAppVersion()
    {
        AppVersion? currentVersion = ReturnCurrentVersion();

        // 1. FRESH INSTALLATION PATH
        if (currentVersion == null)
        {
            await DownloadPackageBundleAsync();

            currentVersion = new AppVersion
            {
                Version = "1.0.0",
                LastUpdatedUtc = DateTime.UtcNow,
                Environment = "Production"
            };
            await WriteCurrentVersionAsync(currentVersion);
            return;
        }

        // 2. EXISTING INSTALLATION -> CHECK FOR UPDATES
        AppVersion? remoteVersion = await CheckRemoteVersionAsync();
        if (remoteVersion == null) return; // Server unreachable or offline; continue to launch

        if (IsNewerVersion(remoteVersion.Version, currentVersion.Version))
        {
            DialogResult choice = MessageBox.Show(
                $"A new update is available (v{remoteVersion.Version}).\n\nWould you like to install it now?",
                "Update Available",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );

            if (choice == DialogResult.Yes)
            {
                await ApplyIncrementalUpdatesAsync(remoteVersion);
                await WriteCurrentVersionAsync(remoteVersion);
            }
        }
    }

    private static bool IsNewerVersion(string remoteVersion, string localVersion)
    {
        if (Version.TryParse(remoteVersion, out var remote) && Version.TryParse(localVersion, out var local))
        {
            return remote > local;
        }
        return !string.Equals(remoteVersion, localVersion, StringComparison.OrdinalIgnoreCase);
    }

    private static AppVersion? ReturnCurrentVersion()
    {
        if (!File.Exists(VersionFilePath)) return null;

        try
        {
            string json = File.ReadAllText(VersionFilePath);
            return JsonSerializer.Deserialize<AppVersion>(json, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }
        catch
        {
            return null;
        }
    }

    private static async Task WriteCurrentVersionAsync(AppVersion versionRecord)
    {
        versionRecord.LastUpdatedUtc = DateTime.UtcNow;
        string json = JsonSerializer.Serialize(versionRecord, new JsonSerializerOptions
        {
            WriteIndented = true
        });
        await File.WriteAllTextAsync(VersionFilePath, json);
    }

    private static async Task DownloadPackageBundleAsync()
    {
        // TODO: Call API (e.g. GET /api/client/bundle), unpack zip to RootPath
        await Task.CompletedTask;
    }

    private static async Task<AppVersion?> CheckRemoteVersionAsync()
    {
        // TODO: Query remote API (e.g. GET /api/version/check) and deserialise
        await Task.CompletedTask;
        return null;
    }

    private static async Task ApplyIncrementalUpdatesAsync(AppVersion targetManifest)
    {
        // TODO: Download changed files / diff package and overwrite local files
        await Task.CompletedTask;
    }

    /// <summary>
    /// Spawns the Portal desktop client once installations and checks pass.
    /// </summary>
    private static void LaunchOperationsPortal()
    {
        string portalExePath = Path.Combine(RootPath, "Portal.exe");

        if (!File.Exists(portalExePath))
        {
            // If testing in local development environment:
            string devPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Portal.exe");
            portalExePath = File.Exists(devPath) ? devPath : portalExePath;
        }

        if (!File.Exists(portalExePath))
        {
            MessageBox.Show($"Portal executable was not found at:\n{portalExePath}",
                            "Launch Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = portalExePath,
            WorkingDirectory = RootPath,
            UseShellExecute = true
        });
    }
}