namespace SKSB_App.Core.RuntimeUI.Dictionary;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Localization;

/// <summary>
/// Shared WinForms dialog for browsing translations, submitting wording suggestions,
/// voting on pending community suggestions (3 votes auto-approves), and (in Debug)
/// moderating or reverting live overrides.
/// </summary>
public sealed class CollectiveDictionaryDialog : Form
{
    private static readonly Color BgPage = Color.FromArgb(248, 250, 252);
    private static readonly Color BgCard = Color.White;
    private static readonly Color BorderSubtle = Color.FromArgb(226, 232, 240);
    private static readonly Color AccentPrimary = Color.FromArgb(79, 70, 229);
    private static readonly Color AccentPrimaryHover = Color.FromArgb(67, 56, 202);
    private static readonly Color AccentSoft = Color.FromArgb(238, 242, 255);
    private static readonly Color TextPrimary = Color.FromArgb(15, 23, 42);
    private static readonly Color TextMuted = Color.FromArgb(100, 116, 139);
    private static readonly Color WarnBg = Color.FromArgb(254, 243, 199);
    private static readonly Color WarnFg = Color.FromArgb(146, 64, 14);
    private static readonly Color DangerFg = Color.FromArgb(185, 28, 28);
    private static readonly Color DangerSoft = Color.FromArgb(254, 226, 226);

    private sealed class BrowseEntry
    {
        public string DictType { get; init; } = CollectiveDictionaryRepository.DictTypeUi;
        public string EntryKey { get; init; } = string.Empty;
        public string English { get; init; } = string.Empty;
        public string EnglishDisplay { get; init; } = string.Empty;
        public string BuiltInDefault { get; init; } = string.Empty;
        public string CurrentActive { get; init; } = string.Empty;
        public bool IsCollectiveOverride { get; init; }
    }

    // Top bar
    private ComboBox _cmbLanguage = null!;
    private ComboBox _cmbType = null!;
    private TextBox _txtSearch = null!;
    private Button _btnSync = null!;
    private Button _btnNewTip = null!;
    private Label _lblStatus = null!;
    private Panel _offlineBanner = null!;
    private Label _lblOfflineBanner = null!;

    // Tabs
    private TabControl _tabs = null!;
    private TabPage _tabBrowse = null!;
    private TabPage _tabVotes = null!;
#if DEBUG
    private TabPage _tabOverrides = null!;
#endif

    // Tab 1: Browse & Suggest
    private SplitContainer _browseSplit = null!;
    private DataGridView _dgvBrowse = null!;
    private Label _lblEditorKey = null!;
    private TextBox _txtEditorEnglish = null!;
    private TextBox _txtEditorDefault = null!;
    private TextBox _txtEditorCurrent = null!;
    private Label _lblPlaceholderHint = null!;
    private TextBox _txtSuggested = null!;
    private TextBox _txtNote = null!;
    private Button _btnSubmit = null!;

    // Tab 2: Community Votes
    private DataGridView _dgvVotes = null!;
    private Button _btnToggleVote = null!;
#if DEBUG
    private Button _btnAdminApprove = null!;
    private Button _btnAdminEditApprove = null!;
    private Button _btnAdminReject = null!;
#endif

    // Tab 3 (Debug): Live Overrides
#if DEBUG
    private DataGridView _dgvOverrides = null!;
    private Button _btnRevertOverride = null!;
#endif

    // State
    private List<BrowseEntry> _allBrowseEntries = new();
    private List<CollectiveSuggestionRecord> _pendingSuggestions = new();
#if DEBUG
    private List<CollectiveOverrideRecord> _liveOverrides = new();
#endif
    private bool _suppressFilters;
    private bool _isBusy;

    public CollectiveDictionaryDialog()
    {
        SuspendLayout();

        Text = Loc.T("dict.windowTitle");
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimumSize = new Size(900, 600);
        ClientSize = new Size(1040, 680);
        BackColor = BgPage;
        ForeColor = TextPrimary;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);

        InitializeLayout();

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        ApplyInitialSplitterLayout();
        RebuildBrowseSource();
        await RefreshServerTabsAsync();
    }

    private void ApplyInitialSplitterLayout()
    {
        if (_browseSplit == null || _browseSplit.IsDisposed)
            return;

        int total = _browseSplit.Height;
        int splitterWidth = _browseSplit.SplitterWidth;
        if (total <= splitterWidth + 2)
            return;

        int min1 = Math.Min(140, total / 4);
        int min2 = Math.Min(180, total / 3);
        if (min1 + min2 + splitterWidth >= total)
            return;

        try
        {
            _browseSplit.Panel1MinSize = min1;
            _browseSplit.Panel2MinSize = min2;
            _browseSplit.SplitterDistance = Math.Clamp(
                (int)(total * 0.52),
                min1,
                Math.Max(min1, total - min2 - splitterWidth));
        }
        catch
        {
            // Ignore layout race if the form is minimized or resized during show.
        }
    }

    private string SelectedLangCode
    {
        get
        {
            int idx = _cmbLanguage?.SelectedIndex ?? 0;
            if (idx < 0 || idx >= AppLanguageInfo.All.Count) idx = 0;
            return AppLanguageInfo.ToCode(AppLanguageInfo.All[idx]);
        }
    }

    private string? SelectedDictTypeFilter => _cmbType?.SelectedIndex switch
    {
        1 => CollectiveDictionaryRepository.DictTypeUi,
        2 => CollectiveDictionaryRepository.DictTypeTooltips,
        _ => null
    };

    // =========================================================================
    // LAYOUT
    // =========================================================================
    private void InitializeLayout()
    {
        var root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 3,
            Padding = new Padding(12),
            BackColor = BgPage
        };
        root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        // 1. Top filter bar
        var topBar = BuildTopFilterBar();
        root.Controls.Add(topBar, 0, 0);

        // 2. Offline banner (hidden by default)
        _offlineBanner = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = WarnBg,
            Padding = new Padding(10, 6, 10, 6),
            Margin = new Padding(0, 0, 0, 8),
            AutoSize = true,
            Visible = false
        };
        _lblOfflineBanner = new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ForeColor = WarnFg,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold)
        };
        _offlineBanner.Controls.Add(_lblOfflineBanner);
        root.Controls.Add(_offlineBanner, 0, 1);

        // 3. TabControl
        _tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };

        _tabBrowse = new TabPage(Loc.T("dict.tab.browse")) { BackColor = BgPage, Padding = new Padding(8) };
        BuildBrowseTab(_tabBrowse);
        _tabs.TabPages.Add(_tabBrowse);

        _tabVotes = new TabPage(Loc.T("dict.tab.votes", 0)) { BackColor = BgPage, Padding = new Padding(8) };
        BuildVotesTab(_tabVotes);
        _tabs.TabPages.Add(_tabVotes);

#if DEBUG
        _tabOverrides = new TabPage(Loc.T("dict.tab.overrides", 0)) { BackColor = BgPage, Padding = new Padding(8) };
        BuildOverridesTab(_tabOverrides);
        _tabs.TabPages.Add(_tabOverrides);
#endif

        root.Controls.Add(_tabs, 0, 2);
        Controls.Add(root);
    }

    private Control BuildTopFilterBar()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 9,
            RowCount = 1,
            AutoSize = true,
            BackColor = BgCard,
            Padding = new Padding(10, 8, 10, 8),
            Margin = new Padding(0, 0, 0, 8)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Lang label
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Lang combo
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Type label
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Type combo
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Search label
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // Search box
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Sync btn
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // New Tip btn
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize)); // Status label

        var lblLang = new Label
        {
            Text = Loc.T("dict.filter.language"),
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            ForeColor = TextMuted,
            Margin = new Padding(0, 0, 6, 0)
        };

        _cmbLanguage = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 155,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 0, 14, 0)
        };
        foreach (var lang in AppLanguageInfo.All)
            _cmbLanguage.Items.Add(AppLanguageInfo.GetDisplayName(lang));

        _suppressFilters = true;
        _cmbLanguage.SelectedIndex = AppLanguageInfo.IndexOf(Loc.CurrentLanguage);

        var lblType = new Label
        {
            Text = Loc.T("dict.filter.type"),
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            ForeColor = TextMuted,
            Margin = new Padding(0, 0, 6, 0)
        };

        _cmbType = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 140,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 0, 14, 0)
        };
        _cmbType.Items.Add(Loc.T("dict.filter.typeAll"));
        _cmbType.Items.Add(Loc.T("dict.filter.typeUi"));
        _cmbType.Items.Add(Loc.T("dict.filter.typeTooltips"));
        _cmbType.SelectedIndex = 0;
        _suppressFilters = false;

        var lblSearch = new Label
        {
            Text = Loc.T("dict.filter.search"),
            AutoSize = true,
            Anchor = AnchorStyles.Left,
            ForeColor = TextMuted,
            Margin = new Padding(0, 0, 6, 0)
        };

        _txtSearch = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.filter.searchPlaceholder"),
            Margin = new Padding(0, 2, 12, 2)
        };

        _btnSync = CreateSecondaryButton(Loc.T("dict.btn.sync"));
        _btnSync.Margin = new Padding(0, 0, 8, 0);
        _btnSync.Click += async (s, e) => await OnSyncClickedAsync();

        _btnNewTip = CreatePrimaryButton(Loc.T("dict.btn.newTip"));
        _btnNewTip.Margin = new Padding(0, 0, 10, 0);
        _btnNewTip.Click += async (s, e) => await OnSuggestNewTipClickedAsync();

        _lblStatus = new Label
        {
            Text = string.Empty,
            AutoSize = true,
            Anchor = AnchorStyles.Right,
            ForeColor = TextMuted,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular)
        };

        _cmbLanguage.SelectedIndexChanged += async (s, e) =>
        {
            if (_suppressFilters) return;
            RebuildBrowseSource();
            await RefreshServerTabsAsync();
        };
        _cmbType.SelectedIndexChanged += (s, e) =>
        {
            if (_suppressFilters) return;
            ApplyBrowseFilter();
            PopulateVotesGrid();
#if DEBUG
            PopulateOverridesGrid();
#endif
        };
        _txtSearch.TextChanged += (s, e) =>
        {
            if (_suppressFilters) return;
            ApplyBrowseFilter();
            PopulateVotesGrid();
#if DEBUG
            PopulateOverridesGrid();
#endif
        };

        panel.Controls.Add(lblLang, 0, 0);
        panel.Controls.Add(_cmbLanguage, 1, 0);
        panel.Controls.Add(lblType, 2, 0);
        panel.Controls.Add(_cmbType, 3, 0);
        panel.Controls.Add(lblSearch, 4, 0);
        panel.Controls.Add(_txtSearch, 5, 0);
        panel.Controls.Add(_btnSync, 6, 0);
        panel.Controls.Add(_btnNewTip, 7, 0);
        panel.Controls.Add(_lblStatus, 8, 0);

        return panel;
    }

    private void BuildBrowseTab(TabPage page)
    {
        _browseSplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            BackColor = BorderSubtle
        };

        // Top: Grid
        _dgvBrowse = CreateStyledGrid();
        _dgvBrowse.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Type",
            HeaderText = Loc.T("dict.col.type"),
            FillWeight = 12,
            MinimumWidth = 80
        });
        _dgvBrowse.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "English",
            HeaderText = Loc.T("dict.col.english"),
            FillWeight = 38,
            MinimumWidth = 180
        });
        _dgvBrowse.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Current",
            HeaderText = Loc.T("dict.col.current"),
            FillWeight = 38,
            MinimumWidth = 180
        });
        _dgvBrowse.Columns.Add(new DataGridViewTextBoxColumn
        {
            Name = "Source",
            HeaderText = Loc.T("dict.col.source"),
            FillWeight = 12,
            MinimumWidth = 85
        });
        _dgvBrowse.SelectionChanged += (s, e) => PopulateEditorFromSelection();
        _browseSplit.Panel1.Controls.Add(_dgvBrowse);

        // Bottom: Editor panel
        var editorScroll = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = BgCard,
            Padding = new Padding(12, 10, 12, 10)
        };

        var editorLayout = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            BackColor = BgCard
        };
        editorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 130));
        editorLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        int row = 0;

        // Key
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.key")), 0, row);
        _lblEditorKey = new Label
        {
            AutoSize = true,
            ForeColor = TextMuted,
            Font = new Font("Consolas", 8.5f, FontStyle.Regular),
            Margin = new Padding(0, 3, 0, 4)
        };
        editorLayout.Controls.Add(_lblEditorKey, 1, row++);

        // English
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.english")), 0, row);
        _txtEditorEnglish = CreateReadOnlyField();
        editorLayout.Controls.Add(_txtEditorEnglish, 1, row++);

        // Built-in Default
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.default")), 0, row);
        _txtEditorDefault = CreateReadOnlyField();
        editorLayout.Controls.Add(_txtEditorDefault, 1, row++);

        // Current Active
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.current")), 0, row);
        _txtEditorCurrent = CreateReadOnlyField();
        editorLayout.Controls.Add(_txtEditorCurrent, 1, row++);

        // Placeholder hint
        editorLayout.Controls.Add(new Label { AutoSize = true }, 0, row);
        _lblPlaceholderHint = new Label
        {
            AutoSize = true,
            ForeColor = AccentPrimary,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Visible = false,
            Margin = new Padding(0, 2, 0, 4)
        };
        editorLayout.Controls.Add(_lblPlaceholderHint, 1, row++);

        // Suggested Wording
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.suggested")), 0, row);
        _txtSuggested = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Height = 58,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Margin = new Padding(0, 2, 0, 6)
        };
        editorLayout.Controls.Add(_txtSuggested, 1, row++);

        // Note + Submit row
        editorLayout.Controls.Add(CreateFieldLabel(Loc.T("dict.editor.note")), 0, row);
        var noteRow = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0)
        };
        noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _txtNote = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.editor.notePlaceholder"),
            Margin = new Padding(0, 3, 10, 0)
        };
        _btnSubmit = CreatePrimaryButton(Loc.T("dict.btn.submit"));
        _btnSubmit.Click += async (s, e) => await OnSubmitSuggestionClickedAsync();

        noteRow.Controls.Add(_txtNote, 0, 0);
        noteRow.Controls.Add(_btnSubmit, 1, 0);
        editorLayout.Controls.Add(noteRow, 1, row);

        editorScroll.Controls.Add(editorLayout);
        _browseSplit.Panel2.Controls.Add(editorScroll);
        page.Controls.Add(_browseSplit);
    }

    private void BuildVotesTab(TabPage page)
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = BgPage
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _dgvVotes = CreateStyledGrid();
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = Loc.T("dict.col.type"), FillWeight = 9, MinimumWidth = 70 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "English", HeaderText = Loc.T("dict.col.english"), FillWeight = 20, MinimumWidth = 120 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "Current", HeaderText = Loc.T("dict.col.current"), FillWeight = 20, MinimumWidth = 120 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "Suggested", HeaderText = Loc.T("dict.col.suggested"), FillWeight = 24, MinimumWidth = 140 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "Votes", HeaderText = Loc.T("dict.col.votes"), FillWeight = 9, MinimumWidth = 75 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "By", HeaderText = Loc.T("dict.col.by"), FillWeight = 8, MinimumWidth = 75 });
        _dgvVotes.Columns.Add(new DataGridViewTextBoxColumn { Name = "Note", HeaderText = Loc.T("dict.col.note"), FillWeight = 10, MinimumWidth = 90 });
        _dgvVotes.SelectionChanged += (s, e) => UpdateVotesActionButtons();

        var actionBar = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            Padding = new Padding(0, 8, 0, 0)
        };
        actionBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));
        actionBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 50f));

        var leftFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.LeftToRight, AutoSize = true, Dock = DockStyle.Left };
        _btnToggleVote = CreatePrimaryButton(Loc.T("dict.btn.upvote"));
        _btnToggleVote.Enabled = false;
        _btnToggleVote.Click += async (s, e) => await OnToggleVoteClickedAsync();
        leftFlow.Controls.Add(_btnToggleVote);
        actionBar.Controls.Add(leftFlow, 0, 0);

#if DEBUG
        var rightFlow = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Dock = DockStyle.Right };
        _btnAdminReject = CreateDangerButton(Loc.T("dict.btn.reject"));
        _btnAdminReject.Enabled = false;
        _btnAdminReject.Margin = new Padding(8, 0, 0, 0);
        _btnAdminReject.Click += async (s, e) => await OnAdminRejectClickedAsync();

        _btnAdminEditApprove = CreateSecondaryButton(Loc.T("dict.btn.editApprove"));
        _btnAdminEditApprove.Enabled = false;
        _btnAdminEditApprove.Margin = new Padding(8, 0, 0, 0);
        _btnAdminEditApprove.Click += async (s, e) => await OnAdminEditApproveClickedAsync();

        _btnAdminApprove = CreateSecondaryButton(Loc.T("dict.btn.approveNow"));
        _btnAdminApprove.Enabled = false;
        _btnAdminApprove.Click += async (s, e) => await OnAdminApproveClickedAsync();

        rightFlow.Controls.Add(_btnAdminReject);
        rightFlow.Controls.Add(_btnAdminEditApprove);
        rightFlow.Controls.Add(_btnAdminApprove);
        actionBar.Controls.Add(rightFlow, 1, 0);
#endif

        layout.Controls.Add(_dgvVotes, 0, 0);
        layout.Controls.Add(actionBar, 0, 1);
        page.Controls.Add(layout);
    }

#if DEBUG
    private void BuildOverridesTab(TabPage page)
    {
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = BgPage
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _dgvOverrides = CreateStyledGrid();
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "Type", HeaderText = Loc.T("dict.col.type"), FillWeight = 8, MinimumWidth = 65 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "Key", HeaderText = Loc.T("dict.col.key"), FillWeight = 16, MinimumWidth = 120 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "English", HeaderText = Loc.T("dict.col.english"), FillWeight = 18, MinimumWidth = 120 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "Default", HeaderText = Loc.T("dict.col.default"), FillWeight = 18, MinimumWidth = 120 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "Override", HeaderText = Loc.T("dict.col.override"), FillWeight = 20, MinimumWidth = 130 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "ApprovedBy", HeaderText = Loc.T("dict.col.approvedBy"), FillWeight = 10, MinimumWidth = 85 });
        _dgvOverrides.Columns.Add(new DataGridViewTextBoxColumn { Name = "ApprovedAt", HeaderText = Loc.T("dict.col.approvedAt"), FillWeight = 10, MinimumWidth = 95 });
        _dgvOverrides.SelectionChanged += (s, e) =>
        {
            _btnRevertOverride.Enabled = !_isBusy && GetSelectedOverride() != null;
        };

        var actionBar = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 8, 0, 0)
        };
        _btnRevertOverride = CreateDangerButton(Loc.T("dict.btn.revert"));
        _btnRevertOverride.Enabled = false;
        _btnRevertOverride.Click += async (s, e) => await OnRevertOverrideClickedAsync();
        actionBar.Controls.Add(_btnRevertOverride);

        layout.Controls.Add(_dgvOverrides, 0, 0);
        layout.Controls.Add(actionBar, 0, 1);
        page.Controls.Add(layout);
    }
#endif

    // =========================================================================
    // DATA LOADING & FILTERING
    // =========================================================================
    private void RebuildBrowseSource()
    {
        string lang = SelectedLangCode;
        var entries = new List<BrowseEntry>();

        AddEntriesForType(entries, CollectiveDictionaryRepository.DictTypeUi, lang);
        AddEntriesForType(entries, CollectiveDictionaryRepository.DictTypeTooltips, lang);

        _allBrowseEntries = entries;
        ApplyBrowseFilter();
    }

    private static void AddEntriesForType(List<BrowseEntry> target, string dictType, string langCode)
    {
        var enBase = Loc.GetBaseDictionary(dictType, "en");
        var enCollective = CollectiveDictionaryRepository.LoadCachedOverrides(dictType, "en");
        var langBase = Loc.GetBaseDictionary(dictType, langCode);
        var effective = Loc.GetEffectiveDictionary(dictType, langCode);
        var collective = CollectiveDictionaryRepository.LoadCachedOverrides(dictType, langCode);

        var allKeys = new HashSet<string>(enBase.Keys, StringComparer.Ordinal);
        foreach (var k in langBase.Keys) allKeys.Add(k);
        foreach (var k in effective.Keys) allKeys.Add(k);
        if (string.Equals(dictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase))
        {
            foreach (var k in CollectiveDictionaryRepository.GetAllCachedTooltipKeysByPrefix(
                CollectiveDictionaryRepository.CustomTipKeyPrefix))
            {
                allKeys.Add(k);
            }
        }

        foreach (string key in allKeys.OrderBy(k => k, StringComparer.Ordinal))
        {
            enBase.TryGetValue(key, out string? enVal);
            if (string.IsNullOrEmpty(enVal) && enCollective.TryGetValue(key, out string? enOverride) && !string.IsNullOrWhiteSpace(enOverride))
                enVal = enOverride;

            langBase.TryGetValue(key, out string? baseVal);
            effective.TryGetValue(key, out string? effVal);

            string enDisplay = FormatEnglishDisplay(key, enVal);
            string currentVal = effVal ?? baseVal ?? enVal ?? string.Empty;
            if (string.IsNullOrEmpty(currentVal)
                && CollectiveDictionaryRepository.IsCustomTipKey(key)
                && CollectiveDictionaryRepository.TryGetAnyCachedTooltipValue(key, out string anyVal))
            {
                currentVal = anyVal;
            }

            target.Add(new BrowseEntry
            {
                DictType = dictType,
                EntryKey = key,
                English = enVal ?? string.Empty,
                EnglishDisplay = enDisplay,
                BuiltInDefault = baseVal ?? enVal ?? string.Empty,
                CurrentActive = currentVal,
                IsCollectiveOverride = collective.ContainsKey(key)
            });
        }
    }

    private static string FormatEnglishDisplay(string entryKey, string? enVal)
    {
        if (!string.IsNullOrWhiteSpace(enVal)) return enVal;
        if (CollectiveDictionaryRepository.IsCustomTipKey(entryKey))
        {
            string tag = Loc.T("dict.customTip.noEnglish");
            if (CollectiveDictionaryRepository.TryGetAnyCachedTooltipValue(entryKey, out string anyVal))
                return $"{tag} {anyVal}";
            return tag;
        }
        return string.Empty;
    }

    private void ApplyBrowseFilter()
    {
        string? typeFilter = SelectedDictTypeFilter;
        string q = (_txtSearch.Text ?? string.Empty).Trim();

        string? selectedKey = GetSelectedBrowseEntry()?.EntryKey;
        string? selectedType = GetSelectedBrowseEntry()?.DictType;

        _dgvBrowse.SuspendLayout();
        _dgvBrowse.Rows.Clear();

        int restoreRow = -1;
        string uiLabel = Loc.T("dict.type.ui");
        string tipLabel = Loc.T("dict.type.tooltips");
        string defaultLabel = Loc.T("dict.source.default");
        string collectiveLabel = Loc.T("dict.source.collective");

        foreach (var item in _allBrowseEntries)
        {
            if (typeFilter != null && !string.Equals(item.DictType, typeFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            if (q.Length > 0 &&
                item.EntryKey.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                item.EnglishDisplay.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                item.CurrentActive.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                item.BuiltInDefault.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            int idx = _dgvBrowse.Rows.Add(
                item.DictType == CollectiveDictionaryRepository.DictTypeTooltips ? tipLabel : uiLabel,
                SingleLinePreview(item.EnglishDisplay),
                SingleLinePreview(item.CurrentActive),
                item.IsCollectiveOverride ? collectiveLabel : defaultLabel);

            _dgvBrowse.Rows[idx].Tag = item;
            if (item.IsCollectiveOverride)
                _dgvBrowse.Rows[idx].DefaultCellStyle.BackColor = AccentSoft;

            if (selectedKey != null && item.EntryKey == selectedKey && item.DictType == selectedType)
                restoreRow = idx;
        }

        _dgvBrowse.ResumeLayout();

        if (restoreRow >= 0 && restoreRow < _dgvBrowse.Rows.Count)
            _dgvBrowse.Rows[restoreRow].Selected = true;
        else if (_dgvBrowse.Rows.Count > 0)
            _dgvBrowse.Rows[0].Selected = true;

        _lblStatus.Text = Loc.T("dict.status.ready", _dgvBrowse.Rows.Count);
        PopulateEditorFromSelection();
    }

    private async Task RefreshServerTabsAsync()
    {
        string lang = SelectedLangCode;
        try
        {
            var pending = await Task.Run(() => CollectiveDictionaryRepository.GetPendingSuggestions(lang));
#if DEBUG
            var overrides = await Task.Run(() => CollectiveDictionaryRepository.GetLiveOverrides(lang));
#endif
            if (IsDisposed) return;

            _pendingSuggestions = pending;
#if DEBUG
            _liveOverrides = overrides;
#endif
            SetOfflineBanner(null);
            PopulateVotesGrid();
#if DEBUG
            PopulateOverridesGrid();
#endif
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            _pendingSuggestions = new List<CollectiveSuggestionRecord>();
#if DEBUG
            _liveOverrides = new List<CollectiveOverrideRecord>();
#endif
            SetOfflineBanner(ex.Message);
            PopulateVotesGrid();
#if DEBUG
            PopulateOverridesGrid();
#endif
        }
    }

    private void PopulateVotesGrid()
    {
        string? typeFilter = SelectedDictTypeFilter;
        string q = (_txtSearch.Text ?? string.Empty).Trim();
        string lang = SelectedLangCode;

        var uiEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeUi, "en");
        var tipEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var tipEnCollective = CollectiveDictionaryRepository.LoadCachedOverrides(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var uiEff = Loc.GetEffectiveDictionary(CollectiveDictionaryRepository.DictTypeUi, lang);
        var tipEff = Loc.GetEffectiveDictionary(CollectiveDictionaryRepository.DictTypeTooltips, lang);

        string uiLabel = Loc.T("dict.type.ui");
        string tipLabel = Loc.T("dict.type.tooltips");

        _dgvVotes.SuspendLayout();
        _dgvVotes.Rows.Clear();

        foreach (var s in _pendingSuggestions)
        {
            if (typeFilter != null && !string.Equals(s.DictType, typeFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            bool isTip = string.Equals(s.DictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);
            (isTip ? tipEn : uiEn).TryGetValue(s.EntryKey, out string? enVal);
            if (isTip && string.IsNullOrEmpty(enVal) && tipEnCollective.TryGetValue(s.EntryKey, out string? enOv))
                enVal = enOv;
            (isTip ? tipEff : uiEff).TryGetValue(s.EntryKey, out string? curVal);
            string enDisplay = FormatEnglishDisplay(s.EntryKey, enVal);
            curVal ??= enVal ?? string.Empty;

            if (q.Length > 0 &&
                s.EntryKey.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                enDisplay.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                curVal.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                s.SuggestedValue.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                s.Note.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            string voteDisplay = s.HasCurrentUserVoted
                ? $"{s.VoteCount} / {CollectiveDictionaryRepository.ConsensusVoteThreshold} ✓"
                : $"{s.VoteCount} / {CollectiveDictionaryRepository.ConsensusVoteThreshold}";

            int idx = _dgvVotes.Rows.Add(
                isTip ? tipLabel : uiLabel,
                SingleLinePreview(enDisplay),
                SingleLinePreview(curVal),
                SingleLinePreview(s.SuggestedValue),
                voteDisplay,
                s.SuggestedBy,
                SingleLinePreview(s.Note));
            _dgvVotes.Rows[idx].Tag = s;
        }

        _dgvVotes.ResumeLayout();
        _tabVotes.Text = Loc.T("dict.tab.votes", _pendingSuggestions.Count);
        UpdateVotesActionButtons();
    }

#if DEBUG
    private void PopulateOverridesGrid()
    {
        string? typeFilter = SelectedDictTypeFilter;
        string q = (_txtSearch.Text ?? string.Empty).Trim();
        string lang = SelectedLangCode;

        var uiEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeUi, "en");
        var tipEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var tipEnCollective = CollectiveDictionaryRepository.LoadCachedOverrides(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var uiBase = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeUi, lang);
        var tipBase = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeTooltips, lang);

        string uiLabel = Loc.T("dict.type.ui");
        string tipLabel = Loc.T("dict.type.tooltips");

        _dgvOverrides.SuspendLayout();
        _dgvOverrides.Rows.Clear();

        foreach (var o in _liveOverrides)
        {
            if (typeFilter != null && !string.Equals(o.DictType, typeFilter, StringComparison.OrdinalIgnoreCase))
                continue;

            bool isTip = string.Equals(o.DictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);
            (isTip ? tipEn : uiEn).TryGetValue(o.EntryKey, out string? enVal);
            if (isTip && string.IsNullOrEmpty(enVal) && tipEnCollective.TryGetValue(o.EntryKey, out string? enOv))
                enVal = enOv;
            (isTip ? tipBase : uiBase).TryGetValue(o.EntryKey, out string? defVal);
            string enDisplay = FormatEnglishDisplay(o.EntryKey, enVal);
            defVal ??= enVal ?? string.Empty;

            if (q.Length > 0 &&
                o.EntryKey.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                enDisplay.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                defVal.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                o.Value.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            int idx = _dgvOverrides.Rows.Add(
                isTip ? tipLabel : uiLabel,
                o.EntryKey,
                SingleLinePreview(enDisplay),
                SingleLinePreview(defVal),
                SingleLinePreview(o.Value),
                o.ApprovedBy,
                o.ApprovedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));
            _dgvOverrides.Rows[idx].Tag = o;
        }

        _dgvOverrides.ResumeLayout();
        _tabOverrides.Text = Loc.T("dict.tab.overrides", _liveOverrides.Count);
        _btnRevertOverride.Enabled = !_isBusy && GetSelectedOverride() != null;
    }
#endif

    // =========================================================================
    // EDITOR & SELECTION HELPERS
    // =========================================================================
    private BrowseEntry? GetSelectedBrowseEntry()
        => _dgvBrowse.SelectedRows.Count > 0 ? _dgvBrowse.SelectedRows[0].Tag as BrowseEntry : null;

    private CollectiveSuggestionRecord? GetSelectedSuggestion()
        => _dgvVotes.SelectedRows.Count > 0 ? _dgvVotes.SelectedRows[0].Tag as CollectiveSuggestionRecord : null;

#if DEBUG
    private CollectiveOverrideRecord? GetSelectedOverride()
        => _dgvOverrides.SelectedRows.Count > 0 ? _dgvOverrides.SelectedRows[0].Tag as CollectiveOverrideRecord : null;
#endif

    private void PopulateEditorFromSelection()
    {
        var entry = GetSelectedBrowseEntry();
        if (entry == null)
        {
            _lblEditorKey.Text = string.Empty;
            _txtEditorEnglish.Text = string.Empty;
            _txtEditorDefault.Text = string.Empty;
            _txtEditorCurrent.Text = string.Empty;
            _lblPlaceholderHint.Visible = false;
            _txtSuggested.Text = string.Empty;
            _txtNote.Text = string.Empty;
            _btnSubmit.Enabled = false;
            return;
        }

        _lblEditorKey.Text = $"{entry.DictType} :: {entry.EntryKey}";
        _txtEditorEnglish.Text = entry.EnglishDisplay;
        _txtEditorDefault.Text = entry.BuiltInDefault;
        _txtEditorCurrent.Text = entry.CurrentActive;

        string placeholders = CollectiveDictionaryRepository.GetPlaceholderListDisplay(entry.English);
        if (placeholders.Length > 0)
        {
            _lblPlaceholderHint.Text = Loc.T("dict.editor.placeholderHint", placeholders);
            _lblPlaceholderHint.Visible = true;
        }
        else
        {
            _lblPlaceholderHint.Visible = false;
        }

        _txtSuggested.Text = entry.CurrentActive;
        _txtNote.Text = string.Empty;
        _btnSubmit.Enabled = !_isBusy;
    }

    private void UpdateVotesActionButtons()
    {
        var sel = GetSelectedSuggestion();
        bool hasSelection = sel != null && !_isBusy;
        _btnToggleVote.Enabled = hasSelection;
        _btnToggleVote.Text = sel != null && sel.HasCurrentUserVoted
            ? Loc.T("dict.btn.removeVote")
            : Loc.T("dict.btn.upvote");

#if DEBUG
        _btnAdminApprove.Enabled = hasSelection;
        _btnAdminEditApprove.Enabled = hasSelection;
        _btnAdminReject.Enabled = hasSelection;
#endif
    }

    private void SetOfflineBanner(string? errorDetail)
    {
        if (string.IsNullOrWhiteSpace(errorDetail))
        {
            _offlineBanner.Visible = false;
        }
        else
        {
            _lblOfflineBanner.Text = Loc.T("dict.offlineBanner", errorDetail);
            _offlineBanner.Visible = true;
            _lblStatus.Text = Loc.T("dict.status.offline");
        }
    }

    private void SetBusy(bool busy)
    {
        _isBusy = busy;
        _btnSync.Enabled = !busy;
        _btnNewTip.Enabled = !busy;
        _btnSubmit.Enabled = !busy && GetSelectedBrowseEntry() != null;
        UpdateVotesActionButtons();
#if DEBUG
        _btnRevertOverride.Enabled = !busy && GetSelectedOverride() != null;
#endif
        UseWaitCursor = busy;
    }

    // =========================================================================
    // USER ACTIONS
    // =========================================================================
    private async Task OnSyncClickedAsync()
    {
        SetBusy(true);
        _btnSync.Text = Loc.T("dict.btn.syncing");
        try
        {
            bool ok = await Task.Run(() => CollectiveDictionaryRepository.SyncFromServer(reloadLoc: true));
            if (IsDisposed) return;

            RebuildBrowseSource();
            await RefreshServerTabsAsync();
            _lblStatus.Text = ok ? Loc.T("dict.status.synced") : Loc.T("dict.status.offline");
        }
        finally
        {
            if (!IsDisposed)
            {
                _btnSync.Text = Loc.T("dict.btn.sync");
                SetBusy(false);
            }
        }
    }

    private async Task OnSuggestNewTipClickedAsync()
    {
        if (!ShowSuggestNewTipDialog(out string langCode, out string contextTrigger, out string topic, out string tipText, out string note))
            return;

        string key = CollectiveDictionaryRepository.BuildCustomTipKey(topic, contextTrigger);
        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.SubmitSuggestion(
                CollectiveDictionaryRepository.DictTypeTooltips, key, langCode, tipText, note));
            if (IsDisposed) return;

            if (!string.Equals(SelectedLangCode, langCode, StringComparison.OrdinalIgnoreCase))
            {
                int targetIdx = AppLanguageInfo.IndexOf(AppLanguageInfo.FromCode(langCode));
                if (targetIdx >= 0 && targetIdx < _cmbLanguage.Items.Count)
                {
                    _suppressFilters = true;
                    _cmbLanguage.SelectedIndex = targetIdx;
                    _suppressFilters = false;
                    RebuildBrowseSource();
                }
            }

            await RefreshServerTabsAsync();
            MessageBox.Show(
                this,
                Loc.T("dict.newTip.submitted", CollectiveDictionaryRepository.ConsensusVoteThreshold),
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            _tabs.SelectedTab = _tabVotes;
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.submitFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

    private bool ShowSuggestNewTipDialog(out string langCode, out string contextTrigger, out string topic, out string tipText, out string note)
    {
        langCode = SelectedLangCode;
        contextTrigger = string.Empty;
        topic = string.Empty;
        tipText = string.Empty;
        note = string.Empty;

        using var dlg = new Form
        {
            Text = Loc.T("dict.newTip.title"),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ClientSize = new Size(560, 420),
            BackColor = BgPage,
            ForeColor = TextPrimary,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            AutoScaleDimensions = new SizeF(96f, 96f),
            AutoScaleMode = AutoScaleMode.Dpi
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 7,
            Padding = new Padding(14),
            BackColor = BgPage
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 115));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Subtitle
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Language
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // When to Show
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Topic
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));// Tip text
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Note
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));     // Buttons

        var lblSubtitle = new Label
        {
            Text = Loc.T("dict.newTip.subtitle"),
            AutoSize = true,
            MaximumSize = new Size(520, 0),
            ForeColor = TextMuted,
            Margin = new Padding(0, 0, 0, 10)
        };
        layout.Controls.Add(lblSubtitle, 0, 0);
        layout.SetColumnSpan(lblSubtitle, 2);

        // Language
        layout.Controls.Add(CreateFieldLabel(Loc.T("dict.newTip.language")), 0, 1);
        var cmbLang = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Left,
            Width = 180,
            Margin = new Padding(0, 2, 0, 6)
        };
        foreach (var lang in AppLanguageInfo.All)
            cmbLang.Items.Add(AppLanguageInfo.GetDisplayName(lang));
        cmbLang.SelectedIndex = Math.Clamp(_cmbLanguage.SelectedIndex, 0, Math.Max(0, cmbLang.Items.Count - 1));
        layout.Controls.Add(cmbLang, 1, 1);

        // When to Show (General vs Contextual trigger)
        layout.Controls.Add(CreateFieldLabel(Loc.T("dict.newTip.trigger")), 0, 2);
        var triggerCodes = new List<string> { string.Empty };
        var cmbTrigger = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 2, 0, 6)
        };
        cmbTrigger.Items.Add(Loc.T("dict.newTip.trigger.general"));
        foreach (string trig in CollectiveDictionaryRepository.CustomTipContextTriggers)
        {
            triggerCodes.Add(trig);
            cmbTrigger.Items.Add(Loc.T($"dict.newTip.trigger.{trig}"));
        }
        cmbTrigger.SelectedIndex = 0;
        layout.Controls.Add(cmbTrigger, 1, 2);

        // Optional Topic/Title
        layout.Controls.Add(CreateFieldLabel(Loc.T("dict.newTip.topic")), 0, 3);
        var txtTopic = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.newTip.topicPlaceholder"),
            Margin = new Padding(0, 2, 0, 6)
        };
        layout.Controls.Add(txtTopic, 1, 3);

        // Tip Text (required)
        layout.Controls.Add(CreateFieldLabel(Loc.T("dict.newTip.text")), 0, 4);
        var txtTip = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            PlaceholderText = Loc.T("dict.newTip.textPlaceholder"),
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Margin = new Padding(0, 2, 0, 6)
        };
        layout.Controls.Add(txtTip, 1, 4);

        // Optional Note
        layout.Controls.Add(CreateFieldLabel(Loc.T("dict.newTip.note")), 0, 5);
        var txtNote = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.editor.notePlaceholder"),
            Margin = new Padding(0, 2, 0, 8)
        };
        layout.Controls.Add(txtNote, 1, 5);

        // Buttons
        var btnFlow = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 4, 0, 0)
        };
        var btnSubmit = CreatePrimaryButton(Loc.T("dict.newTip.submit"));
        var btnCancel = CreateSecondaryButton(Loc.T("common.cancel"));
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Margin = new Padding(0, 0, 8, 0);

        btnSubmit.Click += (s, e) =>
        {
            string candidate = (txtTip.Text ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(candidate))
            {
                MessageBox.Show(dlg, Loc.T("dict.newTip.emptyError"), dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTip.Focus();
                return;
            }
            if (!CollectiveDictionaryRepository.ValidatePlaceholders(string.Empty, candidate, out string? err))
            {
                MessageBox.Show(dlg, err, dlg.Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTip.Focus();
                return;
            }
            dlg.DialogResult = DialogResult.OK;
            dlg.Close();
        };

        btnFlow.Controls.Add(btnSubmit);
        btnFlow.Controls.Add(btnCancel);
        layout.Controls.Add(btnFlow, 0, 6);
        layout.SetColumnSpan(btnFlow, 2);

        dlg.Controls.Add(layout);
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            int idx = cmbLang.SelectedIndex;
            if (idx < 0 || idx >= AppLanguageInfo.All.Count) idx = 0;
            langCode = AppLanguageInfo.ToCode(AppLanguageInfo.All[idx]);
            int trigIdx = cmbTrigger.SelectedIndex;
            contextTrigger = trigIdx >= 0 && trigIdx < triggerCodes.Count ? triggerCodes[trigIdx] : string.Empty;
            topic = (txtTopic.Text ?? string.Empty).Trim();
            tipText = (txtTip.Text ?? string.Empty).Trim();
            note = (txtNote.Text ?? string.Empty).Trim();
            return true;
        }
        return false;
    }

    private async Task OnSubmitSuggestionClickedAsync()
    {
        var entry = GetSelectedBrowseEntry();
        if (entry == null)
        {
            MessageBox.Show(this, Loc.T("dict.err.selectEntry"), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string suggested = (_txtSuggested.Text ?? string.Empty).Trim();
        string note = (_txtNote.Text ?? string.Empty).Trim();

        if (!CollectiveDictionaryRepository.ValidatePlaceholders(entry.English, suggested, out string? err))
        {
            MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.Equals(suggested, entry.CurrentActive.Trim(), StringComparison.Ordinal))
        {
            MessageBox.Show(this, Loc.T("dict.err.sameAsCurrent"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return;
        }

        string lang = SelectedLangCode;
        SetBusy(true);
        _btnSubmit.Text = Loc.T("dict.btn.submitting");
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.SubmitSuggestion(
                entry.DictType, entry.EntryKey, lang, suggested, note));
            if (IsDisposed) return;

            _txtNote.Text = string.Empty;
            await RefreshServerTabsAsync();
            MessageBox.Show(this, Loc.T("dict.msg.submitted", CollectiveDictionaryRepository.ConsensusVoteThreshold),
                Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            _tabs.SelectedTab = _tabVotes;
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.submitFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed)
            {
                _btnSubmit.Text = Loc.T("dict.btn.submit");
                SetBusy(false);
            }
        }
    }

    private async Task OnToggleVoteClickedAsync()
    {
        var sel = GetSelectedSuggestion();
        if (sel == null) return;

        SetBusy(true);
        try
        {
            var result = await Task.Run(() => CollectiveDictionaryRepository.ToggleVote(sel.SuggestionId));
            if (IsDisposed) return;

            RebuildBrowseSource();
            await RefreshServerTabsAsync();

            if (result.AutoApproved)
            {
                MessageBox.Show(
                    this,
                    Loc.T("dict.msg.autoApproved", result.NewVoteCount, CollectiveDictionaryRepository.ConsensusVoteThreshold),
                    Text,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.voteFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

#if DEBUG
    private async Task OnAdminApproveClickedAsync()
    {
        var sel = GetSelectedSuggestion();
        if (sel == null) return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminApproveSuggestion(sel.SuggestionId));
            if (IsDisposed) return;

            RebuildBrowseSource();
            await RefreshServerTabsAsync();
            MessageBox.Show(this, Loc.T("dict.msg.approved"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.adminFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

    private async Task OnAdminEditApproveClickedAsync()
    {
        var sel = GetSelectedSuggestion();
        if (sel == null) return;

        string englishBase = Loc.GetBaseEntry(sel.DictType, "en", sel.EntryKey);
        if (!ShowEditApprovePrompt(englishBase, sel.SuggestedValue, out string? edited) || edited == null)
            return;

        if (!CollectiveDictionaryRepository.ValidatePlaceholders(englishBase, edited, out string? err))
        {
            MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminApproveSuggestion(sel.SuggestionId, edited));
            if (IsDisposed) return;

            RebuildBrowseSource();
            await RefreshServerTabsAsync();
            MessageBox.Show(this, Loc.T("dict.msg.approved"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.adminFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

    private async Task OnAdminRejectClickedAsync()
    {
        var sel = GetSelectedSuggestion();
        if (sel == null) return;

        if (MessageBox.Show(this, Loc.T("dict.confirm.reject"), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminRejectSuggestion(sel.SuggestionId));
            if (IsDisposed) return;

            await RefreshServerTabsAsync();
            MessageBox.Show(this, Loc.T("dict.msg.rejected"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.adminFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

    private async Task OnRevertOverrideClickedAsync()
    {
        var sel = GetSelectedOverride();
        if (sel == null) return;

        if (MessageBox.Show(this, Loc.T("dict.confirm.revert", sel.EntryKey), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminRevertOverride(sel.DictType, sel.EntryKey, sel.LangCode));
            if (IsDisposed) return;

            RebuildBrowseSource();
            await RefreshServerTabsAsync();
            MessageBox.Show(this, Loc.T("dict.msg.reverted"), Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            if (!IsDisposed)
                MessageBox.Show(this, Loc.T("dict.err.adminFailed", ex.Message), Text, MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            if (!IsDisposed) SetBusy(false);
        }
    }

    private bool ShowEditApprovePrompt(string englishBase, string initialValue, out string? editedValue)
    {
        editedValue = null;
        using var dlg = new Form
        {
            Text = Loc.T("dict.editApprove.title"),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false,
            ClientSize = new Size(540, 260),
            BackColor = BgPage,
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            AutoScaleDimensions = new SizeF(96f, 96f),
            AutoScaleMode = AutoScaleMode.Dpi
        };

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 4,
            Padding = new Padding(14)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var lblPrompt = new Label
        {
            Text = Loc.T("dict.editApprove.prompt"),
            AutoSize = true,
            ForeColor = TextPrimary,
            Margin = new Padding(0, 0, 0, 4)
        };

        string placeholders = CollectiveDictionaryRepository.GetPlaceholderListDisplay(englishBase);
        var lblHint = new Label
        {
            Text = placeholders.Length > 0 ? Loc.T("dict.editor.placeholderHint", placeholders) : string.Empty,
            AutoSize = true,
            ForeColor = AccentPrimary,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
            Visible = placeholders.Length > 0,
            Margin = new Padding(0, 0, 0, 6)
        };

        var txt = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ScrollBars = ScrollBars.Vertical,
            Text = initialValue,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular)
        };

        var btnFlow = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Padding = new Padding(0, 8, 0, 0)
        };
        var btnOk = CreatePrimaryButton(Loc.T("common.ok"));
        btnOk.DialogResult = DialogResult.OK;
        var btnCancel = CreateSecondaryButton(Loc.T("common.cancel"));
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Margin = new Padding(0, 0, 8, 0);

        btnFlow.Controls.Add(btnOk);
        btnFlow.Controls.Add(btnCancel);

        layout.Controls.Add(lblPrompt, 0, 0);
        layout.Controls.Add(lblHint, 0, 1);
        layout.Controls.Add(txt, 0, 2);
        layout.Controls.Add(btnFlow, 0, 3);

        dlg.Controls.Add(layout);
        dlg.AcceptButton = btnOk;
        dlg.CancelButton = btnCancel;

        if (dlg.ShowDialog(this) == DialogResult.OK)
        {
            editedValue = txt.Text;
            return true;
        }
        return false;
    }
#endif

    // =========================================================================
    // UI FACTORY HELPERS
    // =========================================================================
    private static DataGridView CreateStyledGrid()
    {
        var dgv = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            RowHeadersVisible = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
            BackgroundColor = BgCard,
            BorderStyle = BorderStyle.FixedSingle,
            GridColor = BorderSubtle,
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize
        };
        dgv.ColumnHeadersDefaultCellStyle.BackColor = BgPage;
        dgv.ColumnHeadersDefaultCellStyle.ForeColor = TextPrimary;
        dgv.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
        dgv.DefaultCellStyle.SelectionBackColor = AccentPrimary;
        dgv.DefaultCellStyle.SelectionForeColor = Color.White;
        return dgv;
    }

    private static Label CreateFieldLabel(string text) => new()
    {
        Text = text,
        AutoSize = true,
        ForeColor = TextMuted,
        Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
        Margin = new Padding(0, 4, 8, 4)
    };

    private static TextBox CreateReadOnlyField() => new()
    {
        Dock = DockStyle.Fill,
        ReadOnly = true,
        BorderStyle = BorderStyle.FixedSingle,
        BackColor = BgPage,
        ForeColor = TextPrimary,
        Margin = new Padding(0, 2, 0, 4)
    };

    private static Button CreatePrimaryButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(96, 30),
            Padding = new Padding(10, 0, 10, 0),
            BackColor = AccentPrimary,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderSize = 0;
        btn.FlatAppearance.MouseOverBackColor = AccentPrimaryHover;
        return btn;
    }

    private static Button CreateSecondaryButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(88, 30),
            Padding = new Padding(10, 0, 10, 0),
            BackColor = BgCard,
            ForeColor = AccentPrimary,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderColor = BorderSubtle;
        btn.FlatAppearance.MouseOverBackColor = AccentSoft;
        return btn;
    }

    private static Button CreateDangerButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            MinimumSize = new Size(88, 30),
            Padding = new Padding(10, 0, 10, 0),
            BackColor = BgCard,
            ForeColor = DangerFg,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand
        };
        btn.FlatAppearance.BorderColor = Color.FromArgb(252, 165, 165);
        btn.FlatAppearance.MouseOverBackColor = DangerSoft;
        return btn;
    }

    private static string SingleLinePreview(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text.Replace("\r\n", " ↵ ").Replace("\n", " ↵ ");
    }
}
