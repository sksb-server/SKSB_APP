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

public class PortalForm : Form
{
    private readonly string _currentVersion;
    private TableLayoutPanel _layout = null!;

    public PortalForm()
    {
        SuspendLayout();
        _currentVersion = ResolveAppVersion();

#if DEBUG
        Text = $"SKSB Operations Portal (Debug) - v{_currentVersion}";
#else
        Text = $"SKSB Operations Portal - v{_currentVersion}";
#endif
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        BackColor = Color.FromArgb(248, 250, 252);

        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(24, 20, 24, 20);

        InitializeLayout();

        // Enable automated DPI scaling from a 96-DPI design baseline (Designer order: after controls are added).
        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        FitClientSizeToLayout();
        CenterToScreen();
    }

    protected override void OnDpiChanged(DpiChangedEventArgs e)
    {
        base.OnDpiChanged(e);
        FitClientSizeToLayout();
    }

    private void FitClientSizeToLayout()
    {
        MinimumSize = Size.Empty;
        _layout.PerformLayout();
        Size pref = _layout.GetPreferredSize(Size.Empty);
        ClientSize = new Size(pref.Width + Padding.Horizontal, pref.Height + Padding.Vertical);
        MinimumSize = Size;
    }

    private void InitializeLayout()
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0),
            Padding = new Padding(0)
        };
        _layout = layout;
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
            "Schedule Manager",
            "Biometric logs, shifts & job site allocations",
            () => PortalRouter.Launch(ModuleRegistry.ScheduleManager));

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
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
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

            float scale = btn.DeviceDpi / 96f;
            float x = 14f * scale;
            float titleY = 8f * scale;
            float subY;

            using (var titleFont = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(Color.FromArgb(30, 41, 59)))
            {
                e.Graphics.DrawString(title, titleFont, titleBrush, x, titleY);
                // Place subtitle below the measured title so they never overlap at any DPI.
                float titleH = e.Graphics.MeasureString(title, titleFont).Height;
                subY = Math.Max(32f * scale, titleY + titleH + (2f * scale));
            }

            using (var subFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (var subBrush = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                e.Graphics.DrawString(subtitle, subFont, subBrush, x, subY);
            }
        };
        btn.Click += (s, e) =>
        {
            if (!btn.Enabled) return;
            btn.Enabled = false;

            var reenableTimer = new System.Windows.Forms.Timer { Interval = 1500 };
            reenableTimer.Tick += (ts, te) =>
            {
                reenableTimer.Stop();
                reenableTimer.Dispose();
                if (!btn.IsDisposed) btn.Enabled = true;
            };
            reenableTimer.Start();

            onClick();
        };
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