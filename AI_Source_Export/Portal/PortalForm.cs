namespace SKSB_App.Portal;

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Windows.Forms;
using SKSB_App.Core.Bootstrap.Storage;

public partial class PortalForm : Form
{
    private readonly string _currentVersion;

    public PortalForm()
    {
        _currentVersion = ResolveAppVersion();

#if DEBUG
        Text = $"SKSB Operations Portal (Debug) - v{_currentVersion}";
#else
        Text = $"SKSB Operations Portal - v{_currentVersion}";
#endif
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        BackColor = Color.FromArgb(248, 250, 252);

        // Enable automated DPI scaling
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24, 20, 24, 20);

        InitializeLayout();
    }

    private void InitializeLayout()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0),
            Padding = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // Header Panel: Title + Version Badge
        var headerPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 4)
        };

        var lblHeader = new Label
        {
#if DEBUG
            Text = "Operations Hub (Debug)",
#else
            Text = "Operations Hub",
#endif
            Font = new Font("Segoe UI", 16f, FontStyle.Bold),
            ForeColor = Color.FromArgb(15, 23, 42),
            AutoSize = true,
            Margin = new Padding(0, 0, 8, 0)
        };

        var lblVersion = new Label
        {
            Text = $"v{_currentVersion}",
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(79, 70, 229),
            BackColor = Color.FromArgb(238, 242, 255),
            Padding = new Padding(6, 2, 6, 2),
            AutoSize = true,
            Margin = new Padding(0, 6, 0, 0)
        };

        headerPanel.Controls.Add(lblHeader);
        headerPanel.Controls.Add(lblVersion);

        var lblSubheader = new Label
        {
            Text = "Select an operational module to launch in Excel:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        };

        // Module Buttons wired to ModuleRegistry matching current architecture
        var btnScheduler = CreateModuleButton(
            "Daily Scheduler",
            "Biometric logs, shifts & job site allocations",
            () => PortalRouter.Launch(ModuleRegistry.DailyScheduler));

        var btnProject = CreateModuleButton(
            "Project Manager",
            "iJob / PVT sites, CIDB compliance & client scopes",
            () => PortalRouter.Launch(ModuleRegistry.ProjectManager));

        var btnStaff = CreateModuleButton(
            "Staff Manager",
            "Profiles, vehicle & skill licenses, SQL Payroll import",
            () => PortalRouter.Launch(ModuleRegistry.StaffManager));

        var btnVehicle = CreateModuleButton(
            "Vehicle Manager",
            "Fleet tracking, Puspakom, road tax & workshop repairs",
            () => PortalRouter.Launch(ModuleRegistry.VehicleManager));

        // Footer: Runtime Status
        var lblFooter = new Label
        {
            Text = $"Running on .NET {Environment.Version.Major} | Deployment: {FolderRegistry.SKSB_App.GetPath()}",
            Font = new Font("Segoe UI", 7.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(148, 163, 184),
            AutoSize = true,
            Margin = new Padding(0, 10, 0, 0)
        };

        layout.Controls.Add(headerPanel);
        layout.Controls.Add(lblSubheader);
        layout.Controls.Add(btnScheduler);
        layout.Controls.Add(btnProject);
        layout.Controls.Add(btnStaff);
        layout.Controls.Add(btnVehicle);
        layout.Controls.Add(lblFooter);

        Controls.Add(layout);
    }

    private Button CreateModuleButton(string title, string subtitle, Action onClick)
    {
        var btn = new Button
        {
            Width = 360,
            Height = 64,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 0, 10)
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);
        btn.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            using (var titleFont = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(Color.FromArgb(30, 41, 59)))
            {
                e.Graphics.DrawString(title, titleFont, titleBrush, 14, 8);
            }

            using (var subFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (var subBrush = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                e.Graphics.DrawString(subtitle, subFont, subBrush, 14, 32);
            }
        };
        btn.Click += (s, e) => onClick();
        return btn;
    }

    private static string ResolveAppVersion()
    {
        try
        {
            // 1. Check version.json in FolderRegistry.SKSB_App
            string versionFile = Path.Combine(FolderRegistry.SKSB_App.GetPath(), "version.json");
            if (File.Exists(versionFile))
            {
                string json = File.ReadAllText(versionFile);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("version", out var prop))
                {
                    string? ver = prop.GetString();
                    if (!string.IsNullOrWhiteSpace(ver)) return ver.Trim().TrimStart('v');
                }
            }

            // 2. Check version.json adjacent to Portal.exe
            string localVersionFile = Path.Combine(AppContext.BaseDirectory, "version.json");
            if (File.Exists(localVersionFile))
            {
                string json = File.ReadAllText(localVersionFile);
                using var doc = JsonDocument.Parse(json);
                if (doc.RootElement.TryGetProperty("version", out var prop))
                {
                    string? ver = prop.GetString();
                    if (!string.IsNullOrWhiteSpace(ver)) return ver.Trim().TrimStart('v');
                }
            }
        }
        catch { }

        // 3. Fallback to Assembly Version
        try
        {
            var asm = Assembly.GetEntryAssembly() ?? Assembly.GetExecutingAssembly();
            var ver = asm.GetName().Version;
            if (ver != null) return $"{ver.Major}.{ver.Minor}.{ver.Build}";
        }
        catch { }

        return "1.0.0";
    }
}