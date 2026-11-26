namespace SKSB_App.Portal;

using System;
using System.Collections.Generic;
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
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.BugReport;
using SKSB_App.Core.RuntimeUI.Dictionary;
using SKSB_App.Core.RuntimeUI.Progress;

public class PortalForm : Form
{
    private const int ModuleButtonBaseWidth = 360;

    private readonly string _currentVersion;
    private TableLayoutPanel _layout = null!;
    private Label _lblHeader = null!;
    private Label _lblSubheader = null!;
    private Label _lblLanguage = null!;
    private ComboBox _cmbLanguage = null!;
    private Button _btnReportBug = null!;
    private Button _btnDictionary = null!;
    private Button _btnEditTemplate = null!;
    private Label _lblFooter = null!;
    private readonly List<ModuleCardBinding> _moduleCards = new();
    private bool _editTemplateMode;
    private bool _suppressLanguageEvent;
    private bool _languageRefreshPending;

    // Pinned static delegate (CLAUDE.md Rule 4); unhooked in OnFormClosed.
    private static Action? _languageChangedHandler;

#if !DEBUG
    private Button? _btnResetTemplates;
#endif
#if DEBUG
    private Button? _btnBugInbox;
    private Button? _btnOfficeManager;
    private int _openBugCount;
    private readonly string _debugBranchLabel;
#endif

    /// <summary>Live text source for a painted module card (repaints in the active language on Invalidate).</summary>
    private sealed class ModuleCardBinding
    {
        public ModuleCardBinding(Button button, Func<string> getTitle, Func<string> getSubtitle)
        {
            Button = button;
            GetTitle = getTitle;
            GetSubtitle = getSubtitle;
        }

        public Button Button { get; }
        public Func<string> GetTitle { get; }
        public Func<string> GetSubtitle { get; }
    }

    public PortalForm()
    {
        SuspendLayout();
        Loc.ReloadFromPreferences();
        Loc.WatchPreferences();
        _currentVersion = ResolveAppVersion();
#if DEBUG
        _debugBranchLabel = ResolveDebugBranchLabel();
#endif

        Text = BuildWindowTitle();
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        MaximizeBox = true;
        BackColor = Color.FromArgb(248, 250, 252);

        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(24, 20, 24, 20);

        InitializeLayout();
        ApplyLanguage();

        // Enable automated DPI scaling from a 96-DPI design baseline (Designer order: after controls are added).
        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();

        if (_languageChangedHandler != null) Loc.LanguageChanged -= _languageChangedHandler;
        _languageChangedHandler = OnLocLanguageChanged;
        Loc.LanguageChanged += _languageChangedHandler;
    }

    private string BuildWindowTitle()
    {
#if DEBUG
        return string.IsNullOrWhiteSpace(_debugBranchLabel)
            ? Loc.T("portal.titleDebug", _currentVersion)
            : Loc.T("portal.titleDebugBranch", _debugBranchLabel, _currentVersion);
#else
        return Loc.T("portal.title", _currentVersion);
#endif
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
        EnsureModuleButtonWidths();
        FitClientSizeToLayout();
        CenterToScreen();
        CollectiveDictionaryRepository.SyncFromServerInBackground();
#if DEBUG
        _ = RefreshBugInboxCountAsync();
#endif
    }

    protected override void OnActivated(EventArgs e)
    {
        base.OnActivated(e);
        if (_languageRefreshPending)
        {
            _languageRefreshPending = false;
            ApplyLanguage();
        }
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        if (_languageChangedHandler != null)
        {
            Loc.LanguageChanged -= _languageChangedHandler;
            _languageChangedHandler = null;
        }
        base.OnFormClosed(e);
    }

    protected override void OnDpiChanged(DpiChangedEventArgs e)
    {
        base.OnDpiChanged(e);
        EnsureModuleButtonWidths();
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

        // Header row: [Title + Version Badge] ............ [Language: ComboBox]
        var headerRow = new TableLayoutPanel
        {
            ColumnCount = 2,
            RowCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 4),
            Padding = new Padding(0)
        };
        headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        headerRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var headerPanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0)
        };

        _lblHeader = new Label
        {
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

        headerPanel.Controls.Add(_lblHeader);
        headerPanel.Controls.Add(lblVersion);

        var languagePanel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Anchor = AnchorStyles.Top | AnchorStyles.Right,
            Margin = new Padding(16, 4, 0, 0)
        };

        _lblLanguage = new Label
        {
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Margin = new Padding(0, 4, 6, 0)
        };

        _cmbLanguage = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            Width = 150,
            Margin = new Padding(0)
        };
        foreach (var language in AppLanguageInfo.All)
            _cmbLanguage.Items.Add(AppLanguageInfo.GetDisplayName(language));
        SyncLanguageComboSelection();
        _cmbLanguage.SelectedIndexChanged += (s, e) => OnLanguageComboChanged();

        languagePanel.Controls.Add(_lblLanguage);
        languagePanel.Controls.Add(_cmbLanguage);

        headerRow.Controls.Add(headerPanel, 0, 0);
        headerRow.Controls.Add(languagePanel, 1, 0);

        _lblSubheader = new Label
        {
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        };

        // Module Buttons wired to ModuleRegistry with Office Paper theme manager accents
        var btnScheduler = CreateModuleButton(
            () => Loc.T("module.schedule.name"),
            () => Loc.T("module.schedule.desc"),
            Color.FromArgb(0xF3, 0xA4, 0x47), // Paper Accent 2 - Warm Amber
            () => LaunchModule(ModuleRegistry.ScheduleManager));

        var btnProject = CreateModuleButton(
            () => Loc.T("module.project.name"),
            () => Loc.T("module.project.desc"),
            Color.FromArgb(0xA5, 0xB5, 0x92), // Paper Accent 1 - Sage Green
            () => LaunchModule(ModuleRegistry.ProjectManager));

        var btnStaff = CreateModuleButton(
            () => Loc.T("module.staff.name"),
            () => Loc.T("module.staff.desc"),
            Color.FromArgb(0x80, 0x9E, 0xC2), // Paper Accent 6 - Steel Blue
            () => LaunchModule(ModuleRegistry.StaffManager));

        var btnVehicle = CreateModuleButton(
            () => Loc.T("module.vehicle.name"),
            () => Loc.T("module.vehicle.desc"),
            Color.FromArgb(0xE7, 0xBC, 0x29), // Paper Accent 3 - Mustard Gold
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

        _btnReportBug = CreateUtilityButton(string.Empty);
        _btnReportBug.Click += (s, e) => OpenBugReportDialog();
        actionRow.Controls.Add(_btnReportBug);

        _btnDictionary = CreateUtilityButton(string.Empty);
        _btnDictionary.Margin = new Padding(8, 0, 0, 0);
        _btnDictionary.Click += (s, e) => OpenCollectiveDictionaryDialog();
        actionRow.Controls.Add(_btnDictionary);

#if DEBUG
        _btnBugInbox = CreateUtilityButton(string.Empty);
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

        _btnOfficeManager = CreateUtilityButton("Office Manager");
        _btnOfficeManager.Margin = new Padding(8, 0, 0, 0);
        _btnOfficeManager.Click += (s, e) =>
        {
            using var dlg = new OfficeManagerDialog();
            dlg.ShowDialog(this);
        };
        actionRow.Controls.Add(_btnOfficeManager);
#endif

        _btnEditTemplate = CreateUtilityButton(string.Empty);
        _btnEditTemplate.Margin = new Padding(8, 0, 0, 0);
        _btnEditTemplate.Click += (s, e) =>
        {
            _editTemplateMode = !_editTemplateMode;
            UpdateEditTemplateVisuals();
        };
        actionRow.Controls.Add(_btnEditTemplate);

#if !DEBUG
        _btnResetTemplates = CreateUtilityButton(string.Empty);
        _btnResetTemplates.Margin = new Padding(8, 0, 0, 0);
        _btnResetTemplates.Enabled = PortalRouter.HasAnyCustomTemplates();
        _btnResetTemplates.Click += (s, e) => ResetTemplatesClicked();
        actionRow.Controls.Add(_btnResetTemplates);
#endif

        // Footer: Runtime Status
        _lblFooter = new Label
        {
            Font = new Font("Segoe UI", 7.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(148, 163, 184),
            AutoSize = true,
            Margin = new Padding(0, 10, 0, 0)
        };

        layout.Controls.Add(headerRow);
        layout.Controls.Add(_lblSubheader);
        layout.Controls.Add(btnScheduler);
        layout.Controls.Add(btnProject);
        layout.Controls.Add(btnStaff);
        layout.Controls.Add(btnVehicle);
        layout.Controls.Add(actionRow);
        layout.Controls.Add(_lblFooter);

        Controls.Add(layout);
    }

    // =========================================================================
    // LANGUAGE
    // =========================================================================
    private void OnLanguageComboChanged()
    {
        if (_suppressLanguageEvent) return;
        int index = _cmbLanguage.SelectedIndex;
        if (index < 0 || index >= AppLanguageInfo.All.Count) return;

        Loc.SetLanguage(AppLanguageInfo.All[index]);
        ApplyLanguage();
    }

    /// <summary>Language changed by another process / add-in (raised on a thread-pool thread by the prefs watcher).</summary>
    private void OnLocLanguageChanged()
    {
        try
        {
            if (IsDisposed || !IsHandleCreated) return;
            if (InvokeRequired) BeginInvoke(new Action(OnExternalLanguageChanged));
            // Same-thread changes come from the ComboBox, which already calls ApplyLanguage().
        }
        catch { }
    }

    private void OnExternalLanguageChanged()
    {
        if (IsDisposed) return;
        if (Form.ActiveForm == this)
        {
            ApplyLanguage();
        }
        else
        {
            // Defer resizing/re-layout until Portal is focused again;
            // touching MinimumSize/ClientSize on a background WinForms window triggers Windows focus-steal.
            _languageRefreshPending = true;
        }
    }

    private void SyncLanguageComboSelection()
    {
        int index = AppLanguageInfo.IndexOf(Loc.CurrentLanguage);
        if (_cmbLanguage.SelectedIndex == index) return;
        _suppressLanguageEvent = true;
        try { _cmbLanguage.SelectedIndex = index; }
        finally { _suppressLanguageEvent = false; }
    }

    private void ApplyLanguage()
    {
        if (IsDisposed) return;
        _languageRefreshPending = false;
        SuspendLayout();
        try
        {
            Text = BuildWindowTitle();
#if DEBUG
            _lblHeader.Text = Loc.T("portal.headerDebug");
#else
            _lblHeader.Text = Loc.T("portal.header");
#endif
            _lblLanguage.Text = Loc.T("language.label");
            SyncLanguageComboSelection();

            _btnReportBug.Text = Loc.T("portal.reportBug");
            _btnDictionary.Text = Loc.T("portal.dictionary");
#if DEBUG
            UpdateBugInboxText();
#else
            if (_btnResetTemplates != null) _btnResetTemplates.Text = Loc.T("portal.resetTemplates");
#endif
            _lblFooter.Text = Loc.T("portal.footer", Environment.Version.Major, FolderRegistry.SKSB_App.GetPath());

            foreach (var card in _moduleCards) card.Button.Invalidate();
        }
        finally
        {
            ResumeLayout(true);
        }

        // Updates _lblSubheader + _btnEditTemplate and refits the window when the handle exists.
        UpdateEditTemplateVisuals();
        if (IsHandleCreated) EnsureModuleButtonWidths();
        if (_layout != null && IsHandleCreated) FitClientSizeToLayout();
    }

    /// <summary>Widens the painted module cards so longer translated titles / subtitles never clip.</summary>
    private void EnsureModuleButtonWidths()
    {
        if (_moduleCards.Count == 0) return;
        try
        {
            float scale = DeviceDpi / 96f;
            int required = (int)Math.Round(ModuleButtonBaseWidth * scale);

            using var g = _moduleCards[0].Button.CreateGraphics();
            using var titleFont = new Font("Segoe UI", 10.5f, FontStyle.Bold);
            using var subFont = new Font("Segoe UI", 8.5f, FontStyle.Regular);
            foreach (var card in _moduleCards)
            {
                int barWidth = Math.Max(4, (int)Math.Round(6f * scale));
                float textX = barWidth + (12f * scale);
                float titleW = g.MeasureString(card.GetTitle(), titleFont).Width;
                float subW = g.MeasureString(card.GetSubtitle(), subFont).Width;
                int needed = (int)Math.Ceiling(textX + Math.Max(titleW, subW) + (16f * scale));
                required = Math.Max(required, needed);
            }

            foreach (var card in _moduleCards)
            {
                if (card.Button.Width != required) card.Button.Width = required;
            }
        }
        catch { }
    }

    // =========================================================================
    // ACTIONS
    // =========================================================================
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
            _btnEditTemplate.Text = Loc.T("portal.editTemplateOn");
            _btnEditTemplate.BackColor = Color.FromArgb(254, 243, 199);
            _btnEditTemplate.ForeColor = Color.FromArgb(180, 83, 9);
            _btnEditTemplate.FlatAppearance.BorderColor = Color.FromArgb(245, 158, 11);
            _btnEditTemplate.FlatAppearance.MouseOverBackColor = Color.FromArgb(253, 230, 138);
#if DEBUG
            _lblSubheader.Text = Loc.T("portal.subheader.editDebug");
#else
            _lblSubheader.Text = Loc.T("portal.subheader.edit");
#endif
            _lblSubheader.ForeColor = Color.FromArgb(180, 83, 9);
        }
        else
        {
            _btnEditTemplate.Text = Loc.T("portal.editTemplate");
            _btnEditTemplate.BackColor = Color.White;
            _btnEditTemplate.ForeColor = Color.FromArgb(79, 70, 229);
            _btnEditTemplate.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
            _btnEditTemplate.FlatAppearance.MouseOverBackColor = Color.FromArgb(238, 242, 255);
            _lblSubheader.Text = Loc.T("portal.subheader");
            _lblSubheader.ForeColor = Color.FromArgb(100, 116, 139);
        }
        if (_layout != null && IsHandleCreated) FitClientSizeToLayout();
    }

#if !DEBUG
    private void ResetTemplatesClicked()
    {
        string title = Loc.T("portal.resetTemplates");
        var confirm = MessageBox.Show(this,
            Loc.T("portal.resetTemplates.confirm"),
            title, MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
        if (confirm == DialogResult.Yes)
        {
            int count = PortalRouter.ResetCustomTemplates(out string? err);
            if (err != null)
                MessageBox.Show(this, err, title, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            else
                MessageBox.Show(this, Loc.T("portal.resetTemplates.done", count), title, MessageBoxButtons.OK, MessageBoxIcon.Information);
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
            MessageBox.Show(this, Loc.T("common.bugReport.openFailed", ex.Message), Loc.T("portal.reportBug"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
#if DEBUG
        _ = RefreshBugInboxCountAsync();
#endif
    }

    private void OpenCollectiveDictionaryDialog()
    {
        try
        {
            using var dlg = new CollectiveDictionaryDialog();
            dlg.ShowDialog(this);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("dict.openFailed", ex.Message), Loc.T("portal.dictionary"),
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

#if DEBUG
    private async Task RefreshBugInboxCountAsync()
    {
        if (_btnBugInbox == null) return;
        try
        {
            int count = await Task.Run(BugReportRepository.GetOpenBugCount);
            if (_btnBugInbox.IsDisposed) return;
            _openBugCount = count;
            UpdateBugInboxText();
        }
        catch { /* never block or crash the portal over the inbox badge */ }
    }

    private void UpdateBugInboxText()
    {
        if (_btnBugInbox == null || _btnBugInbox.IsDisposed) return;
        _btnBugInbox.Text = _openBugCount > 0 ? Loc.T("portal.bugInboxCount", _openBugCount) : Loc.T("portal.bugInbox");
    }
#endif

    private Button CreateModuleButton(Func<string> getTitle, Func<string> getSubtitle, Color accentColor, Action onClick)
    {
        var btn = new Button
        {
            Width = ModuleButtonBaseWidth,
            Height = 64,
            Anchor = AnchorStyles.Left | AnchorStyles.Right,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 0, 10)
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);

        var binding = new ModuleCardBinding(btn, getTitle, getSubtitle);
        _moduleCards.Add(binding);
        btn.AccessibleName = getTitle();

        btn.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // Read live so a language switch + Invalidate() repaints translated text immediately.
            string title = binding.GetTitle();
            string subtitle = binding.GetSubtitle();
            btn.AccessibleName = title;

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
