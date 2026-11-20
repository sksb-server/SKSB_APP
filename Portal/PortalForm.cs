namespace SKSB_App.Portal;

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.BugReport;

public class PortalForm : Form
{
    private readonly string _currentVersion;
    private TableLayoutPanel _layout = null!;
    private Label _lblSubheader = null!;
    private Button _btnEditTemplate = null!;
    private bool _editTemplateMode;
#if !DEBUG
    private Button? _btnResetTemplates;
#endif
#if DEBUG
    private Button? _btnBugInbox;
#endif

    public PortalForm()
    {
        SuspendLayout();
        _currentVersion = ResolveAppVersion();

#if DEBUG
        string branchLabel = ResolveDebugBranchLabel();
        Text = string.IsNullOrWhiteSpace(branchLabel)
            ? $"SKSB Operations Portal (Debug) - v{_currentVersion}"
            : $"SKSB Operations Portal (Debug � {branchLabel}) - v{_currentVersion}";
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

#if DEBUG
    private static string ResolveDebugBranchLabel()
    {
        string root = FolderRegistry.SKSB_App.GetPath();
        try
        {
            string dotGit = Path.Combine(root, ".git");
            string? headPath = null;
            if (Directory.Exists(dotGit))
            {
                headPath = Path.Combine(dotGit, "HEAD");
            }
            else if (File.Exists(dotGit))
            {
                string firstLine = File.ReadLines(dotGit).FirstOrDefault() ?? string.Empty;
                string gitDir = firstLine.Replace("gitdir:", string.Empty, StringComparison.OrdinalIgnoreCase).Trim();
                if (gitDir.Length > 0)
                {
                    string resolved = Path.IsPathRooted(gitDir) ? gitDir : Path.GetFullPath(Path.Combine(root, gitDir));
                    headPath = Path.Combine(resolved, "HEAD");
                }
            }

            if (headPath != null && File.Exists(headPath))
            {
                string head = (File.ReadLines(headPath).FirstOrDefault() ?? string.Empty).Trim();
                const string prefix = "ref: refs/heads/";
                if (head.StartsWith(prefix, StringComparison.Ordinal)) return head[prefix.Length..];
                if (head.Length > 0) return head.Length > 7 ? head[..7] : head;
            }
        }
        catch { }

        string name = Path.GetFileName(root.TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar));
        return string.IsNullOrEmpty(name) ? "Debug" : name;
    }
#endif

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        FitClientSizeToLayout();
        CenterToScreen();
#if DEBUG
        _ = RefreshBugInboxCountAsync();
#endif
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

        _lblSubheader = new Label
        {
            Text = "Select an operational module to launch in Excel:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        };

        // Module Buttons wired to ModuleRegistry with Office Paper theme manager accents
        var btnScheduler = CreateModuleButton(
            "Schedule Manager",
            "Biometric logs, shifts & job site allocations",
            Color.FromArgb(0xF3, 0xA4, 0x47), // Paper Accent 2 � Warm Amber
            () => LaunchModule(ModuleRegistry.ScheduleManager));

        var btnProject = CreateModuleButton(
            "Project Manager",
            "iJob / PVT sites, CIDB compliance & client scopes",
            Color.FromArgb(0xA5, 0xB5, 0x92), // Paper Accent 1 � Sage Green
            () => LaunchModule(ModuleRegistry.ProjectManager));

        var btnStaff = CreateModuleButton(
            "Staff Manager",
            "Profiles, vehicle & skill licenses, SQL Payroll import",
            Color.FromArgb(0x80, 0x9E, 0xC2), // Paper Accent 6 � Steel Blue
            () => LaunchModule(ModuleRegistry.StaffManager));

        var btnVehicle = CreateModuleButton(
            "Vehicle Manager",
            "Fleet tracking, Puspakom, road tax & workshop repairs",
            Color.FromArgb(0xE7, 0xBC, 0x29), // Paper Accent 3 � Mustard Gold
            () => LaunchModule(ModuleRegistry.VehicleManager));

        // Utility action row: bug reporting (+ Debug-only inbox)
        var actionRow = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0, 4, 0, 0)
        };

        var btnReportBug = CreateUtilityButton("Report a Bug");
        btnReportBug.Click += (s, e) => OpenBugReportDialog();
        actionRow.Controls.Add(btnReportBug);

#if DEBUG
        _btnBugInbox = CreateUtilityButton("Bug Inbox");
        _btnBugInbox.Margin = new Padding(8, 0, 0, 0);
        _btnBugInbox.Click += async (s, e) =>
        {
            using (var inbox = new BugReportInboxDialog())
            {
                inbox.ShowDialog(this);
            }
            await RefreshBugInboxCountAsync();
        };
        actionRow.Controls.Add(_btnBugInbox);
#endif

        _btnEditTemplate = CreateUtilityButton("Edit Template Mode");
        _btnEditTemplate.Margin = new Padding(8, 0, 0, 0);
        _btnEditTemplate.Click += (s, e) =>
        {
            _editTemplateMode = !_editTemplateMode;
            UpdateEditTemplateVisuals();
        };
        actionRow.Controls.Add(_btnEditTemplate);
        UpdateEditTemplateVisuals();


#if !DEBUG
        _btnResetTemplates = CreateUtilityButton("Reset Templates");
        _btnResetTemplates.Margin = new Padding(8, 0, 0, 0);
        _btnResetTemplates.Enabled = PortalRouter.HasAnyCustomTemplates();
        _btnResetTemplates.Click += (s, e) => ResetTemplatesClicked();
        actionRow.Controls.Add(_btnResetTemplates);
#endif

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
        layout.Controls.Add(_lblSubheader);
        layout.Controls.Add(btnScheduler);
        layout.Controls.Add(btnProject);
        layout.Controls.Add(btnStaff);
        layout.Controls.Add(btnVehicle);
        layout.Controls.Add(actionRow);
        layout.Controls.Add(lblFooter);

        Controls.Add(layout);
    }

    private void LaunchModule(TemplateAddInPair pair)
    {
        bool edit = _editTemplateMode;
        PortalRouter.Launch(pair, editTemplate: edit);
#if !DEBUG
        if (edit && _btnResetTemplates != null)
            _btnResetTemplates.Enabled = PortalRouter.HasAnyCustomTemplates();
#endif
    }

    private void UpdateEditTemplateVisuals()
    {
        if (_editTemplateMode)
        {
            _btnEditTemplate.Text = "Edit Template Mode: ON";
            _btnEditTemplate.BackColor = Color.FromArgb(254, 243, 199);
            _btnEditTemplate.ForeColor = Color.FromArgb(180, 83, 9);
            _btnEditTemplate.FlatAppearance.BorderColor = Color.FromArgb(245, 158, 11);
            _btnEditTemplate.FlatAppearance.MouseOverBackColor = Color.FromArgb(253, 230, 138);
#if DEBUG
            _lblSubheader.Text = "EDIT TEMPLATE MODE: Opens repo .xltx directly (no DB load/save).";
#else
            _lblSubheader.Text = "EDIT TEMPLATE MODE: Opens your personal template copy (local only, never saved to server).";
#endif
            _lblSubheader.ForeColor = Color.FromArgb(180, 83, 9);
        }
        else
        {
            _btnEditTemplate.Text = "Edit Template Mode";
            _btnEditTemplate.BackColor = Color.White;
            _btnEditTemplate.ForeColor = Color.FromArgb(79, 70, 229);
            _btnEditTemplate.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
            _btnEditTemplate.FlatAppearance.MouseOverBackColor = Color.FromArgb(238, 242, 255);
            _lblSubheader.Text = "Select an operational module to launch in Excel:";
            _lblSubheader.ForeColor = Color.FromArgb(100, 116, 139);
        }
        if (_layout != null && IsHandleCreated) FitClientSizeToLayout();
    }

#if !DEBUG
    private void ResetTemplatesClicked()
    {
        var confirm = MessageBox.Show(this,
            "All your personal template customizations in Templates\\Custom will be removed and reset to the latest company defaults.\n\nContinue?",
            "Reset Templates", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirm == DialogResult.Yes)
        {
            int count = PortalRouter.ResetCustomTemplates(out string? err);
            if (err != null)
                MessageBox.Show(this, err, "Reset Templates", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show(this, $"Reset complete. {count} personal template(s) removed.", "Reset Templates", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        if (_btnResetTemplates != null) _btnResetTemplates.Enabled = PortalRouter.HasAnyCustomTemplates();
    }
#endif

    private static Button CreateUtilityButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(0, 30),
            Padding = new Padding(8, 0, 8, 0),
            BackColor = Color.White,
            ForeColor = Color.FromArgb(79, 70, 229),
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(0)
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(238, 242, 255);
        return btn;
    }

    private void OpenBugReportDialog()
    {
        try
        {
            // Capture the portal as the user currently sees it, before the dialog covers it.
            byte[]? screenshot = BugReportRepository.CaptureScreenPng();
            var context = new BugReportContext
            {
                SourceModule = "Portal",
                AppVersion = _currentVersion,
                IsOffline = ServerEnvironment.IsOffline,
                InitialScreenshotPng = screenshot
            };
            using var dialog = new BugReportSubmitDialog(context);
            dialog.ShowDialog(this);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not open the bug report form:\n{ex.Message}", "Report a Bug", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
#if DEBUG
        _ = RefreshBugInboxCountAsync();
#endif
    }

#if DEBUG
    private async Task RefreshBugInboxCountAsync()
    {
        if (_btnBugInbox == null) return;
        try
        {
            int count = await Task.Run(BugReportRepository.GetOpenBugCount);
            if (_btnBugInbox.IsDisposed) return;
            _btnBugInbox.Text = count > 0 ? $"Bug Inbox ({count} open)" : "Bug Inbox";
        }
        catch { /* never block or crash the portal over the inbox badge */ }
    }
#endif

    private Button CreateModuleButton(string title, string subtitle, Color accentColor, Action onClick)
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
            int barWidth = Math.Max(4, (int)Math.Round(6f * scale));
            using (var accentBrush = new SolidBrush(accentColor))
            {
                e.Graphics.FillRectangle(accentBrush, 1, 1, barWidth, Math.Max(0, btn.ClientSize.Height - 2));
            }

            float x = barWidth + (12f * scale);
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