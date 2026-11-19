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
using Microsoft.Win32;
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
    // Long overall timeout for package downloads; the cheap version check uses its own short CTS.
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private static readonly TimeSpan VersionCheckTimeout = TimeSpan.FromSeconds(10);
    private static readonly string LogFilePath = Path.Combine(FolderRegistry.Data.GetPath(), "Launcher.log");

    private const string ServerOfflinePrefix = "SERVER OFFLINE";

    /// <summary>True when the exception means the server could not be reached (refused, DNS, timeout).</summary>
    private static bool IsServerUnreachable(Exception ex)
    {
        for (Exception? e = ex; e != null; e = e.InnerException)
        {
            if (e is System.Net.Sockets.SocketException or TimeoutException or TaskCanceledException) return true;
            if (e is HttpRequestException h && h.StatusCode == null) return true;
        }
        return false;
    }

    private static string DescribeServerError(Exception ex) =>
        IsServerUnreachable(ex)
            ? $"{ServerOfflinePrefix}: the SKSB update server ({ServerHost}) is not reachable.\nCheck that the server is running and that you are on the office network.\n\nDetails: {ex.Message}"
            : ex.Message;

    private static void LogLauncher(string message)
    {
        try
        {
            Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath)!);
            File.AppendAllText(LogFilePath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] {message}{System.Environment.NewLine}");
        }
        catch { }
    }

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
            EnsureExcelTrustedLocation();

            // 1b. Make sure the .NET 8 Desktop Runtime (needed by Portal and the Excel add-ins) is installed
            if (!EnsureDotNetRuntime()) return;

            // 2. Detection Gate: Check if initial installation is needed
            bool isFirstInstall = !File.Exists(PortalExePath) || !File.Exists(VersionFilePath);

            if (isFirstInstall)
            {
                bool installFailed = false;
                var progressForm = new InstallationProgressForm();
                progressForm.Shown += async (s, e) =>
                {
                    try
                    {
                        await RunFullInstallationAsync(progressForm);
                    }
                    catch (Exception ex)
                    {
                        installFailed = true;
                        LogLauncher($"[Install] Failed: {ex}");
                        MessageBox.Show($"Installation failed: {ex.Message}", "Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                    finally
                    {
                        progressForm.Close();
                    }
                };
                Application.Run(progressForm);
                if (installFailed) return;
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

    // =========================================================================
    // .NET 8 DESKTOP RUNTIME PREREQUISITE
    // =========================================================================
    private const string RuntimeUrlX64 = "https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x64.exe";
    private const string RuntimeUrlX86 = "https://aka.ms/dotnet/8.0/windowsdesktop-runtime-win-x86.exe";

    private static bool HasDesktopRuntime8(string programFilesDir)
    {
        try
        {
            string shared = Path.Combine(programFilesDir, "dotnet", "shared", "Microsoft.WindowsDesktop.App");
            return Directory.Exists(shared) &&
                   Directory.GetDirectories(shared, "8.*").Length > 0;
        }
        catch { return false; }
    }

    /// <summary>64-bit Excel needs only x64; 32-bit (or unknown) Excel also needs the x86 runtime.</summary>
    private static bool IsExcel64Bit()
    {
        try
        {
            using var key = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(@"SOFTWARE\Microsoft\Office\ClickToRun\Configuration");
            return string.Equals(key?.GetValue("Platform") as string, "x64", StringComparison.OrdinalIgnoreCase);
        }
        catch { return false; }
    }

    private static bool EnsureDotNetRuntime()
    {
        var missing = new List<(string Name, string Url)>();
        if (!HasDesktopRuntime8(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFiles)))
            missing.Add((".NET 8 Desktop Runtime (x64)", RuntimeUrlX64));
        if (!IsExcel64Bit() && !HasDesktopRuntime8(System.Environment.GetFolderPath(System.Environment.SpecialFolder.ProgramFilesX86)))
            missing.Add((".NET 8 Desktop Runtime (x86, for 32-bit Excel)", RuntimeUrlX86));

        if (missing.Count == 0) return true;

        var confirm = MessageBox.Show(
            "SKSB needs the following Microsoft component(s), which will be downloaded and installed now:\n\n  - " +
            string.Join("\n  - ", missing.Select(m => m.Name)) +
            "\n\nAn administrator permission prompt (UAC) will appear. Continue?",
            "SKSB Setup - Prerequisites", MessageBoxButtons.OKCancel, MessageBoxIcon.Information);
        if (confirm != DialogResult.OK) return false;

        string? failure = null;
        var form = new InstallationProgressForm();
        form.Shown += async (s, e) =>
        {
            try
            {
                foreach (var (name, url) in missing)
                {
                    form.UpdateStatus($"Downloading {name}...");
                    string exe = Path.Combine(Path.GetTempPath(), $"SKSB_Runtime_{Guid.NewGuid():N}.exe");
                    try
                    {
                        using (var resp = await Http.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
                        {
                            resp.EnsureSuccessStatusCode();
                            await using var fs = new FileStream(exe, FileMode.Create, FileAccess.Write, FileShare.None, 81920, useAsync: true);
                            await resp.Content.CopyToAsync(fs);
                        }

                        form.UpdateStatus($"Installing {name}...");
                        var psi = new ProcessStartInfo(exe, "/install /quiet /norestart") { UseShellExecute = true, Verb = "runas" };
                        using var proc = Process.Start(psi)!;
                        await proc.WaitForExitAsync();
                        // 0 = ok, 3010 = ok (reboot pending), 1638 = newer version already present
                        if (proc.ExitCode is not (0 or 3010 or 1638))
                            throw new InvalidOperationException($"Installer exited with code {proc.ExitCode}.");
                    }
                    finally { try { File.Delete(exe); } catch { } }
                }
            }
            catch (System.ComponentModel.Win32Exception wex) when (wex.NativeErrorCode == 1223)
            {
                failure = "The administrator permission prompt was cancelled.";
            }
            catch (Exception ex)
            {
                failure = IsServerUnreachable(ex)
                    ? $"Could not download the component. Check your internet connection.\n\n{ex.Message}"
                    : ex.Message;
            }
            finally { form.Close(); }
        };
        Application.Run(form);

        if (failure != null)
        {
            LogLauncher($"[Runtime] Install failed: {failure}");
            MessageBox.Show(
                $"The required .NET 8 Desktop Runtime could not be installed.\n\n{failure}\n\n" +
                "Please install it manually from https://dotnet.microsoft.com/download/dotnet/8.0 and start SKSB again.",
                "SKSB Setup - Prerequisites", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return false;
        }
        LogLauncher("[Runtime] .NET 8 Desktop Runtime installed.");
        return true;
    }

    private static void EnsureDirectoryHierarchy()
    {
        FolderRegistry.SKSB_App.EnsureExists();
        FolderRegistry.Templates.EnsureExists();
        FolderRegistry.AddIns.EnsureExists();
        FolderRegistry.Data.EnsureExists();
    }

    /// <summary>
    /// Registers RootPath (with subfolders) as an Excel Trusted Location for the current user so
    /// templates/add-ins under it open without security prompts. Never throws or blocks startup.
    /// </summary>
    private static void EnsureExcelTrustedLocation()
    {
        // 16.0 covers Office 2016/2019/2021/365; 15.0 (Office 2013) only if that Excel is present.
        foreach (string ver in new[] { "16.0", "15.0" })
        {
            try
            {
                if (ver != "16.0")
                {
                    using var excelKey = Registry.CurrentUser.OpenSubKey($@"Software\Microsoft\Office\{ver}\Excel");
                    if (excelKey == null) continue;
                }

                using var key = Registry.CurrentUser.CreateSubKey(
                    $@"Software\Microsoft\Office\{ver}\Excel\Security\Trusted Locations\SKSB_App");
                if (key == null)
                {
                    LogLauncher($"[TrustedLocation] Could not create registry key for Office {ver}.");
                    continue;
                }

                key.SetValue("Path", RootPath, RegistryValueKind.String);
                key.SetValue("AllowSubfolders", 1, RegistryValueKind.DWord);
                key.SetValue("Description", "SKSB_App", RegistryValueKind.String);
            }
            catch (Exception ex)
            {
                LogLauncher($"[TrustedLocation] Failed for Office {ver}: {ex.Message}");
            }
        }
    }

    private static async Task RunFullInstallationAsync(InstallationProgressForm progress)
    {
        progress.UpdateStatus("Connecting to SKSB server (192.168.0.20)...");

        KillExistingProcesses();

        var (downloaded, error) = await DownloadAndExtractFromServerAsync(progress);

        if (!downloaded)
        {
            LogLauncher($"[Install] Download failed: {error}");
            progress.UpdateStatus("Server unreachable. Resolving local dependencies...");
            CopyLocalFolderArtifacts();

            if (!File.Exists(PortalExePath))
            {
                // Do NOT write version.json: the next launch must retry a full install.
                throw new InvalidOperationException(
                    $"Could not install SKSB: no local Portal.exe is available and the package could not be downloaded.\n\n{error}\n\nLog: {LogFilePath}");
            }
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
            using var cts = new CancellationTokenSource(VersionCheckTimeout);
            string manifestJson = Http.GetStringAsync(endpoint, cts.Token).GetAwaiter().GetResult();
            remoteVersion = JsonSerializer.Deserialize<LauncherPackageManifest>(manifestJson, new JsonSerializerOptions
            {
                PropertyNameCaseInsensitive = true
            });
        }
        catch (Exception ex)
        {
            // Server offline: continue to launch installed local files (logged for diagnostics)
            LogLauncher($"[AutoUpdate] Version check skipped: {ex.Message}");
            if (IsServerUnreachable(ex))
            {
                MessageBox.Show(
                    $"{ServerOfflinePrefix}: cannot reach the SKSB update server ({ServerHost}).\n\n" +
                    $"Starting the installed version (v{localVersion?.Version ?? "unknown"}) without checking for updates.",
                    "Server Offline", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
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
                        LogLauncher($"[AutoUpdate Failed] v{remoteVersion.Version}: {err}");
                        MessageBox.Show(
                            $"The update to v{remoteVersion.Version} could not be installed.\n\n{err}\n\nThe currently installed version will be launched. Details: {LogFilePath}",
                            "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    }
                }
                catch (Exception ex)
                {
                    LogLauncher($"[AutoUpdate Exception]: {ex}");
                    MessageBox.Show(
                        $"The update failed unexpectedly:\n\n{ex.Message}\n\nDetails: {LogFilePath}",
                        "Update Failed", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
            return (false, DescribeServerError(ex));
        }
        finally
        {
            try { if (File.Exists(tempZip)) File.Delete(tempZip); } catch { }
            try { if (Directory.Exists(tempExtract)) Directory.Delete(tempExtract, recursive: true); } catch { }
        }
    }

    private static void DispatchFilesToDestination(string sourceFolder)
    {
        if (Directory.GetFiles(sourceFolder, "Portal.exe", SearchOption.AllDirectories).Length == 0)
            throw new InvalidDataException("The downloaded package does not contain Portal.exe.");

        foreach (string file in Directory.GetFiles(sourceFolder, "*.*", SearchOption.AllDirectories))
        {
            string fileName = Path.GetFileName(file);
            string ext = Path.GetExtension(file).ToLowerInvariant();

            string relative = Path.GetRelativePath(sourceFolder, file);
            string topDir = relative.Split(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar)[0];

            string destinationPath;
            if (relative.Contains(Path.DirectorySeparatorChar) && topDir.Equals("templates", StringComparison.OrdinalIgnoreCase))
                destinationPath = Path.Combine(TemplatesFolder, fileName);
            else if (relative.Contains(Path.DirectorySeparatorChar) && topDir.Equals("addins", StringComparison.OrdinalIgnoreCase))
                destinationPath = Path.Combine(AddInsFolder, fileName);
            else
                destinationPath = ext switch
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
                catch (Exception ex) when (i < 2)
                {
                    LogLauncher($"[Dispatch] Retry {i + 1} copying {fileName}: {ex.Message}");
                    Thread.Sleep(200);
                }
                // Final failure propagates; DownloadAndExtractFromServerAsync reports it.
            }
        }

        if (!File.Exists(PortalExePath))
            throw new FileNotFoundException("Portal.exe was not installed after extraction.", PortalExePath);
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