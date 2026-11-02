namespace SKSB_App.Launcher;

using System;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Bootstrap.Storage;

public class LauncherPackageManifest
{
    [JsonPropertyName("version")]
    public string Version { get; set; } = "1.0.0";

    [JsonPropertyName("lastUpdatedUtc")]
    public DateTime LastUpdatedUtc { get; set; } = DateTime.UtcNow;

    [JsonPropertyName("environment")]
    public string Environment { get; set; } = "Production";
}

static class Program
{
    private const string ServerHost = "http://192.168.0.20:5000";
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromSeconds(25) };

    private static readonly string RootPath = FolderRegistry.SKSB_App.GetPath();
    private static readonly string TemplatesFolder = FolderRegistry.Templates.GetPath();
    private static readonly string AddInsFolder = FolderRegistry.AddIns.GetPath();
    private static readonly string VersionFilePath = Path.Combine(RootPath, "version.json");
    private static readonly string PortalExePath = Path.Combine(RootPath, "Portal.exe");

    private const string LauncherMutexName = "Global\\SKSB_App_SingleGateway_Mutex";

    [STAThread]
    static void Main()
    {
        // 1. Guard against duplicate running launcher processes
        using var mutex = new Mutex(true, LauncherMutexName, out bool isInitialInstance);
        if (!isInitialInstance) return;

        Application.SetHighDpiMode(HighDpiMode.PerMonitorV2);
        Application.EnableVisualStyles();
        Application.SetCompatibleTextRenderingDefault(false);

        try
        {
            EnsureDirectoryHierarchy();

            // 2. Detection Gate: Check if initial installation is needed
            bool isFirstInstall = !File.Exists(PortalExePath) || !File.Exists(VersionFilePath);

            if (isFirstInstall)
            {
                var progressForm = new InstallationProgressForm();
                progressForm.Shown += async (s, e) =>
                {
                    try
                    {
                        await RunFullInstallationAsync(progressForm);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Installation failed: {ex.Message}", "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        progressForm.Close();
                    }
                };
                Application.Run(progressForm);
            }
            else
            {
                // Smart incremental update check
                CheckAndApplyIncrementalUpdate();
            }

            // 3. Launch the Portal application
            LaunchOperationsPortal();
        }
        catch (Exception ex)
        {
            MessageBox.Show(
                $"Fatal startup failure in SKSB_App:\n\n{ex.Message}",
                "Gateway Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }
    }

    private static void EnsureDirectoryHierarchy()
    {
        FolderRegistry.SKSB_App.EnsureExists();
        FolderRegistry.Templates.EnsureExists();
        FolderRegistry.AddIns.EnsureExists();
        FolderRegistry.Data.EnsureExists();
    }

    private static async Task RunFullInstallationAsync(InstallationProgressForm progress)
    {
        progress.UpdateStatus("Connecting to SKSB server (192.168.0.20)...");

        KillExistingProcesses();

        var (downloaded, error) = await DownloadAndExtractFromServerAsync(progress);

        if (!downloaded)
        {
            progress.UpdateStatus("Server unreachable. Resolving local dependencies...");
            CopyLocalFolderArtifacts();
        }

        if (!File.Exists(VersionFilePath))
        {
            var defaultVer = new LauncherPackageManifest { Version = "1.0.0", LastUpdatedUtc = DateTime.UtcNow };
            WriteVersionManifest(defaultVer);
        }

        progress.UpdateStatus("Ready!");
        await Task.Delay(400);
    }

    private static void CheckAndApplyIncrementalUpdate()
    {
        LauncherPackageManifest? remoteVersion = null;
        var localVersion = ReadVersionManifest();

        try
        {
            string endpoint = $"{ServerHost}/api/update/version";
            string manifestJson = Http.GetStringAsync(endpoint).GetAwaiter().GetResult();
            remoteVersion = JsonSerializer.Deserialize<LauncherPackageManifest>(manifestJson, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }
        catch
        {
            // Server offline: silently continue to launch installed local files
            return;
        }

        if (remoteVersion != null && localVersion != null && IsNewerVersion(remoteVersion.Version, localVersion.Version))
        {
            // Check if Excel is currently running on the workstation
            bool isExcelRunning = Process.GetProcessesByName("excel").Length > 0;

            if (isExcelRunning)
            {
                DialogResult choice;
                using (var dummyForm = new Form { TopMost = true })
                {
                    string message =
                        $"A new system update (v{remoteVersion.Version}) is available.\n\n" +
                        "IMPORTANT: Excel is currently open. Installing this update requires closing Excel so file locks can be released.\n\n" +
                        "Please save any open work in your spreadsheets before continuing.\n\n" +
                        "Would you like to close Excel and install the update now?";

                    choice = MessageBox.Show(
                        dummyForm,
                        message,
                        "System Update - Excel Open",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Warning
                    );
                }

                // If user chooses 'No', skip updating for now so their unsaved work is preserved
                if (choice != DialogResult.Yes)
                {
                    return;
                }
            }

            // Execute the update
            var updateForm = new InstallationProgressForm();
            updateForm.Shown += async (s, e) =>
            {
                try
                {
                    KillExistingProcesses();
                    updateForm.UpdateStatus($"Updating system to v{remoteVersion.Version}...");

                    var (applied, err) = await DownloadAndExtractFromServerAsync(updateForm);
                    if (applied)
                    {
                        WriteVersionManifest(remoteVersion);
                        updateForm.UpdateStatus("Update applied successfully!");
                        await Task.Delay(300);
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"[AutoUpdate Failed]: {err}");
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"[AutoUpdate Exception]: {ex.Message}");
                }
                finally
                {
                    updateForm.Close();
                }
            };

            Application.Run(updateForm);
        }
    }

    private static async Task<(bool Success, string Error)> DownloadAndExtractFromServerAsync(InstallationProgressForm? progress)
    {
        string tempZip = Path.Combine(Path.GetTempPath(), $"SKSB_Setup_{Guid.NewGuid():N}.zip");
        string tempExtract = Path.Combine(Path.GetTempPath(), $"SKSB_Unpack_{Guid.NewGuid():N}");

        try
        {
            progress?.UpdateStatus("Downloading update package...");
            string packageUrl = $"{ServerHost}/api/update/download";

            using var response = await Http.GetAsync(packageUrl, HttpCompletionOption.ResponseHeadersRead);
            if (!response.IsSuccessStatusCode)
            {
                return (false, $"Server returned HTTP {response.StatusCode}");
            }

            await using (var fs = new FileStream(tempZip, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true))
            {
                await response.Content.CopyToAsync(fs);
            }

            progress?.UpdateStatus("Unpacking modules and Excel add-ins...");
            await Task.Run(() => ZipFile.ExtractToDirectory(tempZip, tempExtract, overwriteFiles: true));

            progress?.UpdateStatus("Configuring application environment...");
            await Task.Run(() => DispatchFilesToDestination(tempExtract));

            return (true, string.Empty);
        }
        catch (Exception ex)
        {
            return (false, ex.Message);
        }
        finally
        {
            try { if (File.Exists(tempZip)) File.Delete(tempZip); } catch { }
            try { if (Directory.Exists(tempExtract)) Directory.Delete(tempExtract, recursive: true); } catch { }
        }
    }

    private static void DispatchFilesToDestination(string sourceFolder)
    {
        foreach (string file in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
        {
            string fileName = Path.GetFileName(file);
            string ext = Path.GetExtension(file).ToLowerInvariant();

            string destinationPath = ext switch
            {
                ".xltx" or ".xltm" => Path.Combine(TemplatesFolder, fileName),
                ".xll" or ".dna" => Path.Combine(AddInsFolder, fileName),
                _ => Path.Combine(RootPath, fileName)
            };

            for (int i = 0; i < 3; i++)
            {
                try
                {
                    File.Copy(file, destinationPath, overwrite: true);
                    break;
                }
                catch when (i < 2)
                {
                    Thread.Sleep(200);
                }
            }
        }
    }

    private static void CopyLocalFolderArtifacts()
    {
        string baseDir = AppDomain.CurrentDomain.BaseDirectory;
        foreach (string file in Directory.GetFiles(baseDir, "*.*"))
        {
            string name = Path.GetFileName(file);
            string ext = Path.GetExtension(file).ToLowerInvariant();

            if (ext is ".exe" or ".dll" or ".json")
            {
                try { File.Copy(file, Path.Combine(RootPath, name), overwrite: true); } catch { }
            }
        }
    }

    private static void KillExistingProcesses()
    {
        try
        {
            foreach (var p in Process.GetProcessesByName("Portal"))
            {
                try { p.Kill(); p.WaitForExit(1000); } catch { }
            }
            foreach (var p in Process.GetProcessesByName("excel"))
            {
                try { p.Kill(); p.WaitForExit(1000); } catch { }
            }
        }
        catch { }
    }

    private static void LaunchOperationsPortal()
    {
#if DEBUG
        string localDebugPortal = Path.GetFullPath(Path.Combine(
            AppDomain.CurrentDomain.BaseDirectory,
            @"..\..\..\..\Portal\bin\Debug\net8.0-windows\Portal.exe"
        ));
        if (!File.Exists(localDebugPortal))
        {
            localDebugPortal = Path.Combine(FolderRegistry.SKSB_App.GetPath(), "Portal", "bin", "Debug", "net8.0-windows", "Portal.exe");
        }
        string portalExe = File.Exists(localDebugPortal) ? localDebugPortal : PortalExePath;
#else
        string portalExe = PortalExePath;
        if (!File.Exists(portalExe))
        {
            string localFallback = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Portal.exe");
            portalExe = File.Exists(localFallback) ? localFallback : portalExe;
        }
#endif

        if (!File.Exists(portalExe))
        {
            MessageBox.Show(
                $"Operations Portal was not found at:\n{portalExe}\n\n" +
                $"Please verify server deployment or run the installer again.",
                "Application Missing",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = portalExe,
            WorkingDirectory = Path.GetDirectoryName(portalExe) ?? RootPath,
            UseShellExecute = true
        });
    }

    private static LauncherPackageManifest? ReadVersionManifest()
    {
        if (!File.Exists(VersionFilePath)) return null;
        try
        {
            string json = File.ReadAllText(VersionFilePath);
            return JsonSerializer.Deserialize<LauncherPackageManifest>(json);
        }
        catch { return null; }
    }

    private static void WriteVersionManifest(LauncherPackageManifest ver)
    {
        try
        {
            string json = JsonSerializer.Serialize(ver, new JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(VersionFilePath, json);
        }
        catch { }
    }

    private static bool IsNewerVersion(string remote, string local)
    {
        if (Version.TryParse(remote, out var r) && Version.TryParse(local, out var l))
            return r > l;
        return !string.Equals(remote, local, StringComparison.OrdinalIgnoreCase);
    }
}