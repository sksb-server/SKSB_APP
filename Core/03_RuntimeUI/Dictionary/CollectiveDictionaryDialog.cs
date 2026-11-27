namespace SKSB_App.Core.RuntimeUI.Dictionary;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Localization;

/// <summary>
/// Friendly multi-view Collective Dictionary: a Home view with three large action cards, a card-based
/// Upvote view (3 votes auto-approves), an Update view with "Pick Text on Screen" plus search fallback,
/// a Disambiguation view for ambiguous picks, a focused Edit Card view, and (Debug) Live Overrides.
///
/// "Pick Text on Screen" is orchestrated by <see cref="DictionaryLauncher"/>: this dialog sets
/// <see cref="RequestedScreenPick"/> and closes; the launcher re-opens it with the resolved matches.
/// </summary>
public sealed class CollectiveDictionaryDialog : Form
{
    // =========================================================================
    // 1. PALETTE, TYPES & STATE
    // =========================================================================
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
    private static readonly Color SuccessFg = Color.FromArgb(4, 120, 87);
    private static readonly Color SuccessSoft = Color.FromArgb(209, 250, 229);
    private static readonly Color CardHover = Color.FromArgb(245, 247, 255);

    private const string CatAll = "all";

    private sealed class BrowseEntry
    {
        public string DictType { get; init; } = CollectiveDictionaryRepository.DictTypeUi;
        public string EntryKey { get; init; } = string.Empty;
        public string English { get; init; } = string.Empty;
        public string EnglishDisplay { get; init; } = string.Empty;
        public string BuiltInDefault { get; init; } = string.Empty;
        public string CurrentActive { get; init; } = string.Empty;
        public bool IsCollectiveOverride { get; init; }
        public string Location { get; init; } = string.Empty;
        public string Category { get; init; } = CatAll;
    }

    private enum DictView { Home, Upvote, Update, Disambiguation, Edit, Overrides }

    // Top bar
    private ComboBox _cmbLanguage = null!;
    private Button _btnSync = null!;
    private Label _lblStatus = null!;
    private Panel _offlineBanner = null!;
    private Label _lblOfflineBanner = null!;

    // Views
    private Panel _viewHost = null!;
    private DictView _currentView = DictView.Home;
    private DictView _editReturnView = DictView.Update;
    private Label? _homeUpvoteBadge;
#if DEBUG
    private LinkLabel? _homeOverridesLink;
#endif

    // Update view state
    private string _searchText = string.Empty;
    private string _categoryFilter = CatAll;
    private DataGridView? _dgvUpdate;
    private readonly List<Button> _categoryPills = new();

    // Edit view state
    private BrowseEntry? _editEntry;
    private string? _editFriendlyLocation;
    private TextBox? _txtSuggested;
    private TextBox? _txtNote;
    private Button? _btnSubmit;
    private ComboBox? _cmbPickedMatches;
    private int _selectedPickedMatchIndex = -1;
    private bool _suppressPickerDropdown;

#if DEBUG
    // Overrides view state
    private DataGridView? _dgvOverrides;
    private Button? _btnRevertOverride;
#endif

    // Data
    private List<BrowseEntry> _allBrowseEntries = new();
    private List<CollectiveSuggestionRecord> _pendingSuggestions = new();
#if DEBUG
    private List<CollectiveOverrideRecord> _liveOverrides = new();
#endif
    private readonly IReadOnlyList<DictionaryPickerMatch> _pickedMatches;
    private readonly bool _showNoMatchBanner;
    private readonly bool _openUpdateView;
    private bool _suppressFilters;
    private bool _isBusy;
    private string _selectedLangCode;

    /// <summary>True when the user pressed "Pick Text on Screen" (dialog closed with <see cref="DialogResult.Retry"/>).</summary>
    public bool RequestedScreenPick { get; private set; }

    /// <summary>Language code currently selected in the dialog's language selector.</summary>
    public string SelectedLangCode => _selectedLangCode;

    public CollectiveDictionaryDialog() : this(null, null, false, false) { }

    internal CollectiveDictionaryDialog(
        string? initialLangCode,
        IReadOnlyList<DictionaryPickerMatch>? pickedMatches = null,
        bool showNoMatchBanner = false,
        bool openUpdateView = false)
    {
        _pickedMatches = pickedMatches ?? Array.Empty<DictionaryPickerMatch>();
        _showNoMatchBanner = showNoMatchBanner;
        _openUpdateView = openUpdateView;
        _selectedLangCode = string.IsNullOrWhiteSpace(initialLangCode)
            ? AppLanguageInfo.ToCode(Loc.CurrentLanguage)
            : AppLanguageInfo.ToCode(AppLanguageInfo.FromCode(initialLangCode));

        SuspendLayout();

        Text = Loc.T("dict.windowTitle");
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimumSize = new Size(820, 580);
        ClientSize = new Size(980, 680);
        BackColor = BgPage;
        ForeColor = TextPrimary;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);

        InitializeLayout();

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        // Views are built after the form's DPI auto-scaling so dynamic sizes use S() exactly once.
        RebuildBrowseSource();
        ShowInitialView();
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        if (_currentView == DictView.Edit) FocusSuggestedBox();
        await RefreshServerDataAsync();
    }

    private void ShowInitialView()
    {
        if (_pickedMatches.Count >= 1)
        {
            // Open straight on the Edit Card; the in-card dropdown lists every other nearby pick.
            _selectedPickedMatchIndex = 0;
            OpenEditForMatch(_pickedMatches[0], DictView.Update);
        }
        else if (_openUpdateView)
        {
            ShowView(DictView.Update);
        }
        else
        {
            ShowView(DictView.Home);
        }
    }

    /// <summary>Scales a 96-DPI pixel value to the form's current DPI (for controls built after auto-scaling).</summary>
    private int S(int px) => (int)Math.Round(px * DeviceDpi / 96f);

    // =========================================================================
    // 2. SHELL LAYOUT (top bar + offline banner + view host)
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

        root.Controls.Add(BuildTopBar(), 0, 0);

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

        _viewHost = new Panel
        {
            Dock = DockStyle.Fill,
            BackColor = BgPage,
            Margin = new Padding(0)
        };
        root.Controls.Add(_viewHost, 0, 2);

        Controls.Add(root);
    }

    private Control BuildTopBar()
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 5,
            RowCount = 1,
            AutoSize = true,
            BackColor = BgCard,
            Padding = new Padding(10, 6, 10, 6),
            Margin = new Padding(0, 0, 0, 8)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Lang label
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Lang combo
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f)); // Spacer
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Status
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));      // Sync

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
            Width = 165,
            Anchor = AnchorStyles.Left,
            Margin = new Padding(0, 0, 14, 0)
        };
        foreach (var lang in AppLanguageInfo.All)
            _cmbLanguage.Items.Add(AppLanguageInfo.GetDisplayName(lang));

        _suppressFilters = true;
        int idx = AppLanguageInfo.IndexOf(AppLanguageInfo.FromCode(_selectedLangCode));
        _cmbLanguage.SelectedIndex = Math.Clamp(idx, 0, Math.Max(0, _cmbLanguage.Items.Count - 1));
        _suppressFilters = false;

        _lblStatus = new Label
        {
            Text = string.Empty,
            AutoSize = true,
            Anchor = AnchorStyles.Right,
            ForeColor = TextMuted,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            Margin = new Padding(0, 0, 10, 0)
        };

        _btnSync = CreateSecondaryButton(Loc.T("dict.btn.sync"));
        _btnSync.Anchor = AnchorStyles.Right;
        _btnSync.Margin = new Padding(0);
        _btnSync.Click += async (s, e) => await OnSyncClickedAsync();

        _cmbLanguage.SelectedIndexChanged += async (s, e) =>
        {
            if (_suppressFilters) return;
            _selectedLangCode = ReadLangCodeFromCombo();
            RebuildBrowseSource();
            RefreshCurrentViewForLanguage();
            await RefreshServerDataAsync();
        };

        panel.Controls.Add(lblLang, 0, 0);
        panel.Controls.Add(_cmbLanguage, 1, 0);
        panel.Controls.Add(new Label { AutoSize = true, Margin = new Padding(0) }, 2, 0);
        panel.Controls.Add(_lblStatus, 3, 0);
        panel.Controls.Add(_btnSync, 4, 0);
        return panel;
    }

    private string ReadLangCodeFromCombo()
    {
        int idx = _cmbLanguage?.SelectedIndex ?? 0;
        if (idx < 0 || idx >= AppLanguageInfo.All.Count) idx = 0;
        return AppLanguageInfo.ToCode(AppLanguageInfo.All[idx]);
    }

    private bool IsEnglishSelected => string.Equals(SelectedLangCode, "en", StringComparison.OrdinalIgnoreCase);

    // =========================================================================
    // 3. VIEW SWITCHING
    // =========================================================================
    private void ShowView(DictView view)
    {
        _currentView = view;
        _homeUpvoteBadge = null;
#if DEBUG
        _homeOverridesLink = null;
        _dgvOverrides = null;
        _btnRevertOverride = null;
#endif
        _dgvUpdate = null;
        _categoryPills.Clear();
        _txtSuggested = null;
        _txtNote = null;
        _btnSubmit = null;
        _cmbPickedMatches = null;

        Control content = view switch
        {
            DictView.Upvote => BuildUpvoteView(),
            DictView.Update => BuildUpdateView(),
            DictView.Disambiguation => BuildDisambiguationView(),
            DictView.Edit => BuildEditView(),
#if DEBUG
            DictView.Overrides => BuildOverridesView(),
#endif
            _ => BuildHomeView()
        };

        _viewHost.SuspendLayout();
        var old = _viewHost.Controls.Cast<Control>().ToList();
        _viewHost.Controls.Clear();
        foreach (var c in old) c.Dispose();
        content.Dock = DockStyle.Fill;
        _viewHost.Controls.Add(content);
        _viewHost.ResumeLayout(true);

        _viewHost.Enabled = !_isBusy;
        UpdateStatusLabel();
    }

    private void RefreshCurrentViewForLanguage()
    {
        switch (_currentView)
        {
            case DictView.Edit:
                if (_editEntry != null)
                {
                    var refreshed = FindBrowseEntry(_editEntry.DictType, _editEntry.EntryKey);
                    if (refreshed != null) _editEntry = refreshed;
                }
                ShowView(DictView.Edit);
                break;
            default:
                ShowView(_currentView);
                break;
        }
    }

    /// <summary>Re-renders server-backed views after pending suggestions / overrides change.</summary>
    private void RefreshServerBackedView()
    {
        switch (_currentView)
        {
            case DictView.Home:
                UpdateHomeBadges();
                break;
            case DictView.Upvote:
#if DEBUG
            case DictView.Overrides:
#endif
                ShowView(_currentView);
                break;
        }
    }

    private void RequestScreenPick()
    {
        if (_isBusy) return;
        RequestedScreenPick = true;
        DialogResult = DialogResult.Retry;
        Close();
    }

    // =========================================================================
    // 4. HOME VIEW
    // =========================================================================
    private Control BuildHomeView()
    {
        var scroll = new Panel { AutoScroll = true, BackColor = BgPage, Padding = new Padding(S(8), S(8), S(8), S(8)) };

        var stack = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            BackColor = BgPage,
            Padding = new Padding(0)
        };
        stack.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        var lblTitle = new Label
        {
            Text = Loc.T("dict.home.title"),
            AutoSize = true,
            ForeColor = TextPrimary,
            Font = new Font("Segoe UI", 15f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, S(2))
        };
        var lblSubtitle = new Label
        {
            Text = Loc.T("dict.home.subtitle"),
            AutoSize = true,
            ForeColor = TextMuted,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Margin = new Padding(0, 0, 0, S(14))
        };
        stack.Controls.Add(lblTitle);
        stack.Controls.Add(lblSubtitle);

        _homeUpvoteBadge = CreateBadge(string.Empty, SuccessSoft, SuccessFg);
        var cardUpvote = CreateHomeCard("▲", AccentSoft, AccentPrimary,
            Loc.T("dict.home.upvoteTitle"), Loc.T("dict.home.upvoteDesc"), _homeUpvoteBadge,
            () => ShowView(DictView.Upvote));
        var cardUpdate = CreateHomeCard("Aa", Color.FromArgb(224, 242, 254), Color.FromArgb(3, 105, 161),
            Loc.T("dict.home.updateTitle"), Loc.T("dict.home.updateDesc"), null,
            () => ShowView(DictView.Update));
        var cardNewTip = CreateHomeCard("+", WarnBg, WarnFg,
            Loc.T("dict.home.newTipTitle"), Loc.T("dict.home.newTipDesc"), null,
            async () => await OnSuggestNewTipClickedAsync());

        stack.Controls.Add(cardUpvote);
        stack.Controls.Add(cardUpdate);
        stack.Controls.Add(cardNewTip);

#if DEBUG
        _homeOverridesLink = new LinkLabel
        {
            AutoSize = true,
            Text = Loc.T("dict.home.debugOverrides", _liveOverrides.Count),
            LinkColor = TextMuted,
            ActiveLinkColor = AccentPrimary,
            Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
            Margin = new Padding(S(2), S(10), 0, 0)
        };
        _homeOverridesLink.LinkClicked += (s, e) => ShowView(DictView.Overrides);
        stack.Controls.Add(_homeOverridesLink);
#endif

        scroll.Controls.Add(stack);
        UpdateHomeBadges();
        return scroll;
    }

    private void UpdateHomeBadges()
    {
        if (_homeUpvoteBadge != null && !_homeUpvoteBadge.IsDisposed)
        {
            int waiting = _pendingSuggestions.Count;
            if (waiting > 0)
            {
                _homeUpvoteBadge.Text = Loc.T("dict.home.upvoteBadge", waiting);
                _homeUpvoteBadge.BackColor = AccentSoft;
                _homeUpvoteBadge.ForeColor = AccentPrimary;
            }
            else
            {
                _homeUpvoteBadge.Text = Loc.T("dict.home.upvoteCaughtUp");
                _homeUpvoteBadge.BackColor = SuccessSoft;
                _homeUpvoteBadge.ForeColor = SuccessFg;
            }
        }
#if DEBUG
        if (_homeOverridesLink != null && !_homeOverridesLink.IsDisposed)
            _homeOverridesLink.Text = Loc.T("dict.home.debugOverrides", _liveOverrides.Count);
#endif
    }

    private Control CreateHomeCard(string icon, Color iconBack, Color iconFore, string title, string description, Label? badge, Action onClick)
    {
        var card = new HoverCard(BgCard, CardHover, BorderSubtle, AccentPrimary)
        {
            Dock = DockStyle.Top,
            Height = S(96),
            Margin = new Padding(0, 0, 0, S(12)),
            Padding = new Padding(S(18), S(14), S(18), S(14)),
            Cursor = Cursors.Hand
        };

        var grid = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 3,
            RowCount = 2,
            BackColor = Color.Transparent,
            Margin = new Padding(0)
        };
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, S(58)));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        grid.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));
        grid.RowStyles.Add(new RowStyle(SizeType.Percent, 50f));

        var lblIcon = CreateIconBadge(icon, iconBack, iconFore);
        lblIcon.Anchor = AnchorStyles.Left;
        grid.Controls.Add(lblIcon, 0, 0);
        grid.SetRowSpan(lblIcon, 2);

        var lblTitle = new Label
        {
            Text = title,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.BottomLeft,
            ForeColor = TextPrimary,
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            BackColor = Color.Transparent,
            AutoEllipsis = true,
            Margin = new Padding(S(6), 0, 0, S(1))
        };
        grid.Controls.Add(lblTitle, 1, 0);

        var lblDesc = new Label
        {
            Text = description,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.TopLeft,
            ForeColor = TextMuted,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            BackColor = Color.Transparent,
            AutoEllipsis = true,
            Margin = new Padding(S(6), S(2), 0, 0)
        };
        grid.Controls.Add(lblDesc, 1, 1);

        if (badge != null)
        {
            badge.Anchor = AnchorStyles.Right;
            grid.Controls.Add(badge, 2, 0);
            grid.SetRowSpan(badge, 2);
        }
        else
        {
            var chevron = new Label
            {
                Text = "›",
                AutoSize = true,
                Anchor = AnchorStyles.Right,
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 18f, FontStyle.Regular),
                BackColor = Color.Transparent
            };
            grid.Controls.Add(chevron, 2, 0);
            grid.SetRowSpan(chevron, 2);
        }

        card.Controls.Add(grid);
        card.WireClick(onClick);
        return card;
    }

    /// <summary>44x44 soft-tinted rounded square with a crisp Segoe UI glyph (no emoji fonts).</summary>
    private Label CreateIconBadge(string glyph, Color back, Color fore)
    {
        var lbl = new Label
        {
            Text = string.Empty,
            AutoSize = false,
            Size = new Size(S(44), S(44)),
            BackColor = Color.Transparent,
            ForeColor = fore,
            Font = new Font("Segoe UI", 15f, FontStyle.Bold),
            Margin = new Padding(0)
        };
        lbl.Paint += (s, e) =>
        {
            var g = e.Graphics;
            g.SmoothingMode = SmoothingMode.AntiAlias;
            var rect = new Rectangle(0, 0, lbl.Width - 1, lbl.Height - 1);
            int d = Math.Min(S(20), Math.Min(rect.Width, rect.Height));
            using (var path = new GraphicsPath())
            {
                path.AddArc(rect.Left, rect.Top, d, d, 180, 90);
                path.AddArc(rect.Right - d, rect.Top, d, d, 270, 90);
                path.AddArc(rect.Right - d, rect.Bottom - d, d, d, 0, 90);
                path.AddArc(rect.Left, rect.Bottom - d, d, d, 90, 90);
                path.CloseFigure();
                using var brush = new SolidBrush(back);
                g.FillPath(brush, path);
            }
            TextRenderer.DrawText(g, glyph, lbl.Font, new Rectangle(0, 0, lbl.Width, lbl.Height), fore,
                TextFormatFlags.HorizontalCenter | TextFormatFlags.VerticalCenter | TextFormatFlags.SingleLine | TextFormatFlags.NoPadding);
        };
        return lbl;
    }

    private Label CreateBadge(string text, Color back, Color fore) => new()
    {
        Text = text,
        AutoSize = true,
        BackColor = back,
        ForeColor = fore,
        Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
        Padding = new Padding(S(8), S(3), S(8), S(3)),
        Margin = new Padding(0, 0, S(8), 0)
    };

    // =========================================================================
    // 5. UPVOTE VIEW
    // =========================================================================
    private Control BuildUpvoteView()
    {
        var layout = CreateViewLayout(out var headerFlow);
        headerFlow.Controls.Add(CreateBackButton(() => ShowView(DictView.Home)));
        headerFlow.Controls.Add(CreateViewTitle(Loc.T("dict.votes.title")));
        headerFlow.Controls.Add(CreateBadge(Loc.T("dict.home.upvoteBadge", _pendingSuggestions.Count), AccentSoft, AccentPrimary));

        var list = CreateCardList();
        if (_pendingSuggestions.Count == 0)
        {
            list.Controls.Add(CreateEmptyStateCard(Loc.T("dict.votes.emptyTitle"), Loc.T("dict.votes.emptyDesc")));
        }
        else
        {
            foreach (var s in _pendingSuggestions)
                list.Controls.Add(CreateVoteCard(s));
        }

        layout.Controls.Add(list, 0, 1);
        return layout;
    }

    private Control CreateVoteCard(CollectiveSuggestionRecord s)
    {
        var entry = FindBrowseEntry(s.DictType, s.EntryKey);
        string english = entry?.EnglishDisplay ?? FormatEnglishDisplay(s.EntryKey, Loc.GetBaseEntry(s.DictType, "en", s.EntryKey));
        string current = entry?.CurrentActive ?? english;
        int threshold = CollectiveDictionaryRepository.ConsensusVoteThreshold;

        var card = CreateStackCard();

        // Top row: location badge, author, vote count
        var top = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            BackColor = Color.Transparent,
            Margin = new Padding(0, 0, 0, S(8))
        };
        top.Controls.Add(CreateBadge(LocalizationTextMatcher.GetFriendlyLocation(s.DictType, s.EntryKey), AccentSoft, AccentPrimary));
        top.Controls.Add(new Label
        {
            Text = Loc.T("dict.voteCard.by", string.IsNullOrWhiteSpace(s.SuggestedBy) ? "?" : s.SuggestedBy),
            AutoSize = true,
            ForeColor = TextMuted,
            Margin = new Padding(0, S(3), S(10), 0)
        });
        top.Controls.Add(new Label
        {
            Text = $"{s.VoteCount}/{threshold}",
            AutoSize = true,
            ForeColor = s.HasCurrentUserVoted ? SuccessFg : TextMuted,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, S(3), 0, 0)
        });
        card.Controls.Add(top);

        // Before -> After
        card.Controls.Add(CreateCaption(Loc.T("dict.voteCard.before")));
        card.Controls.Add(CreateWrapLabel(current, TextPrimary, 9.5f, FontStyle.Regular, S(6)));
        card.Controls.Add(CreateCaption("↓ " + Loc.T("dict.voteCard.after"), AccentPrimary));
        card.Controls.Add(CreateWrapLabel(s.SuggestedValue, AccentPrimary, 10f, FontStyle.Bold, S(6)));

        if (!IsEnglishSelected && !string.IsNullOrWhiteSpace(english))
            card.Controls.Add(CreateWrapLabel($"{Loc.T("dict.editor.english")} {english}", TextMuted, 8.5f, FontStyle.Regular, S(4)));

        if (!string.IsNullOrWhiteSpace(s.Note))
            card.Controls.Add(CreateWrapLabel($"“{s.Note.Trim()}”", TextMuted, 8.5f, FontStyle.Italic, S(4)));

        // Bottom action bar
        var actions = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            BackColor = Color.Transparent,
            Margin = new Padding(0, S(8), 0, 0)
        };

        Button btnVote;
        if (s.HasCurrentUserVoted)
        {
            btnVote = CreateSecondaryButton(Loc.T("dict.voteCard.upvoted", s.VoteCount, threshold));
            btnVote.ForeColor = SuccessFg;
            btnVote.BackColor = SuccessSoft;
            btnVote.FlatAppearance.BorderColor = Color.FromArgb(110, 231, 183);
            btnVote.FlatAppearance.MouseOverBackColor = Color.FromArgb(167, 243, 208);
        }
        else
        {
            btnVote = CreatePrimaryButton(Loc.T("dict.voteCard.upvote", s.VoteCount, threshold));
        }
        btnVote.Margin = new Padding(0, 0, S(8), 0);
        btnVote.Click += async (o, e) => await OnToggleVoteClickedAsync(s);
        actions.Controls.Add(btnVote);

#if DEBUG
        var btnApprove = CreateSecondaryButton(Loc.T("dict.btn.approveNow"));
        btnApprove.Margin = new Padding(S(12), 0, S(6), 0);
        btnApprove.Click += async (o, e) => await OnAdminApproveClickedAsync(s);
        var btnEditApprove = CreateSecondaryButton(Loc.T("dict.btn.editApprove"));
        btnEditApprove.Margin = new Padding(0, 0, S(6), 0);
        btnEditApprove.Click += async (o, e) => await OnAdminEditApproveClickedAsync(s);
        var btnReject = CreateDangerButton(Loc.T("dict.btn.reject"));
        btnReject.Margin = new Padding(0);
        btnReject.Click += async (o, e) => await OnAdminRejectClickedAsync(s);
        actions.Controls.Add(btnApprove);
        actions.Controls.Add(btnEditApprove);
        actions.Controls.Add(btnReject);
#endif

        card.Controls.Add(actions);
        return card;
    }

    private Control CreateEmptyStateCard(string title, string description)
    {
        var card = CreateStackCard();
        var icon = CreateIconBadge("✓", SuccessSoft, SuccessFg);
        icon.Margin = new Padding(0, 0, 0, S(6));
        card.Controls.Add(icon);
        card.Controls.Add(CreateWrapLabel(title, TextPrimary, 11.5f, FontStyle.Bold, S(2)));
        card.Controls.Add(CreateWrapLabel(description, TextMuted, 9.5f, FontStyle.Regular, 0));
        return card;
    }

    // =========================================================================
    // 6. UPDATE VIEW (Pick on screen + search fallback)
    // =========================================================================
    private Control BuildUpdateView()
    {
        var layout = new TableLayoutPanel
        {
            ColumnCount = 1,
            RowCount = 6,
            BackColor = BgPage,
            Padding = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // header
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // no-match banner
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // hero
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // search + pills
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f)); // grid
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));      // edit selected

        var header = CreateHeaderFlow();
        header.Controls.Add(CreateBackButton(() => ShowView(DictView.Home)));
        header.Controls.Add(CreateViewTitle(Loc.T("dict.update.title")));
        layout.Controls.Add(header, 0, 0);

        if (_showNoMatchBanner)
        {
            var banner = new Panel
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                BackColor = WarnBg,
                Padding = new Padding(S(10), S(8), S(10), S(8)),
                Margin = new Padding(0, 0, 0, S(10))
            };
            banner.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Text = Loc.T("dict.picker.noMatchBanner"),
                ForeColor = WarnFg,
                Font = new Font("Segoe UI", 9f, FontStyle.Bold)
            });
            layout.Controls.Add(banner, 0, 1);
        }

        var btnPick = CreatePrimaryButton(Loc.T("dict.btn.pickOnScreen"));
        btnPick.Dock = DockStyle.Fill;
        btnPick.AutoSize = false;
        btnPick.Height = S(52);
        btnPick.Font = new Font("Segoe UI", 12f, FontStyle.Bold);
        btnPick.Margin = new Padding(0, 0, 0, S(12));
        btnPick.Click += (s, e) => RequestScreenPick();
        layout.Controls.Add(btnPick, 0, 2);

        // Search + pills
        var searchRow = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 2,
            BackColor = BgPage,
            Margin = new Padding(0, 0, 0, S(8))
        };
        searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        searchRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        var txtSearch = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.update.searchHint"),
            Text = _searchText,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular),
            Margin = new Padding(0, S(3), S(12), 0)
        };
        txtSearch.TextChanged += (s, e) =>
        {
            _searchText = txtSearch.Text ?? string.Empty;
            PopulateUpdateGrid();
        };
        txtSearch.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Down && _dgvUpdate != null && _dgvUpdate.Rows.Count > 0)
            {
                _dgvUpdate.Focus();
                e.Handled = true;
            }
        };

        var pills = new FlowLayoutPanel
        {
            AutoSize = true,
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            BackColor = BgPage,
            Margin = new Padding(0)
        };
        foreach (string cat in new[] { CatAll, "buttons", "tips", "popups" })
        {
            var pill = CreateSecondaryButton(LocalizationTextMatcher.GetFriendlyCategoryLabel(cat));
            pill.Tag = cat;
            pill.MinimumSize = new Size(0, S(30));
            pill.Margin = new Padding(0, 0, S(6), 0);
            pill.Click += (s, e) =>
            {
                _categoryFilter = cat;
                UpdatePillVisuals();
                PopulateUpdateGrid();
            };
            _categoryPills.Add(pill);
            pills.Controls.Add(pill);
        }
        UpdatePillVisuals();

        searchRow.Controls.Add(txtSearch, 0, 0);
        searchRow.Controls.Add(pills, 1, 0);
        layout.Controls.Add(searchRow, 0, 3);

        // Grid
        _dgvUpdate = CreateStyledGrid();
        _dgvUpdate.Columns.Add(new DataGridViewTextBoxColumn { Name = "Current", HeaderText = Loc.T("dict.col.current"), FillWeight = 44, MinimumWidth = 160 });
        _dgvUpdate.Columns.Add(new DataGridViewTextBoxColumn { Name = "English", HeaderText = Loc.T("dict.col.english"), FillWeight = 38, MinimumWidth = 140, Visible = !IsEnglishSelected });
        _dgvUpdate.Columns.Add(new DataGridViewTextBoxColumn { Name = "Location", HeaderText = Loc.T("dict.col.location"), FillWeight = 18, MinimumWidth = 110 });
        _dgvUpdate.CellClick += (s, e) =>
        {
            if (e.RowIndex >= 0) OpenEditForSelectedGridRow();
        };
        _dgvUpdate.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter)
            {
                OpenEditForSelectedGridRow();
                e.Handled = true;
            }
        };
        layout.Controls.Add(_dgvUpdate, 0, 4);

        var footer = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            FlowDirection = FlowDirection.RightToLeft,
            BackColor = BgPage,
            Margin = new Padding(0, S(8), 0, 0)
        };
        var btnEdit = CreateSecondaryButton(Loc.T("dict.btn.editSelected"));
        btnEdit.Click += (s, e) => OpenEditForSelectedGridRow();
        footer.Controls.Add(btnEdit);
        layout.Controls.Add(footer, 0, 5);

        PopulateUpdateGrid();
        return layout;
    }

    private void UpdatePillVisuals()
    {
        foreach (var pill in _categoryPills)
        {
            bool active = string.Equals(pill.Tag as string, _categoryFilter, StringComparison.Ordinal);
            pill.BackColor = active ? AccentPrimary : BgCard;
            pill.ForeColor = active ? Color.White : AccentPrimary;
            pill.FlatAppearance.BorderColor = active ? AccentPrimary : BorderSubtle;
            pill.FlatAppearance.MouseOverBackColor = active ? AccentPrimaryHover : AccentSoft;
        }
    }

    private void PopulateUpdateGrid()
    {
        var dgv = _dgvUpdate;
        if (dgv == null || dgv.IsDisposed) return;

        string q = LocalizationTextMatcher.Normalize(_searchText);
        dgv.SuspendLayout();
        dgv.Rows.Clear();

        foreach (var item in _allBrowseEntries)
        {
            if (_categoryFilter != CatAll && !string.Equals(item.Category, _categoryFilter, StringComparison.Ordinal))
                continue;

            if (q.Length > 0 &&
                item.CurrentActive.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                item.EnglishDisplay.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                item.EntryKey.IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0 &&
                LocalizationTextMatcher.Normalize(item.CurrentActive).IndexOf(q, StringComparison.OrdinalIgnoreCase) < 0)
            {
                continue;
            }

            int idx = dgv.Rows.Add(
                SingleLinePreview(item.CurrentActive),
                SingleLinePreview(item.EnglishDisplay),
                item.Location);
            dgv.Rows[idx].Tag = item;
            if (item.IsCollectiveOverride)
                dgv.Rows[idx].DefaultCellStyle.BackColor = AccentSoft;
        }

        dgv.ResumeLayout();
        if (dgv.Rows.Count > 0)
        {
            dgv.ClearSelection();
            dgv.Rows[0].Selected = true;
        }
        _lblStatus.Text = Loc.T("dict.status.ready", dgv.Rows.Count);
    }

    private void OpenEditForSelectedGridRow()
    {
        if (_dgvUpdate == null || _dgvUpdate.SelectedRows.Count == 0) return;
        if (_dgvUpdate.SelectedRows[0].Tag is not BrowseEntry entry) return;
        _selectedPickedMatchIndex = -1;
        OpenEdit(entry, null, DictView.Update);
    }

    // =========================================================================
    // 7. DISAMBIGUATION VIEW
    // =========================================================================
    private Control BuildDisambiguationView()
    {
        var layout = CreateViewLayout(out var header, extraHeaderRow: true);
        header.Controls.Add(CreateBackButton(() => ShowView(DictView.Update)));
        header.Controls.Add(CreateViewTitle(Loc.T("dict.picker.disambigTitle")));
        var btnPick = CreateSecondaryButton(Loc.T("dict.btn.pickAnother"));
        btnPick.Margin = new Padding(S(12), S(2), 0, 0);
        btnPick.Click += (s, e) => RequestScreenPick();
        header.Controls.Add(btnPick);

        var lblSub = new Label
        {
            Text = Loc.T("dict.picker.disambigSubtitle"),
            AutoSize = true,
            ForeColor = TextMuted,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Margin = new Padding(0, 0, 0, S(10))
        };
        layout.Controls.Add(lblSub, 0, 1);

        var list = CreateCardList();
        for (int i = 0; i < _pickedMatches.Count; i++)
        {
            var m = _pickedMatches[i];
            int matchIndex = i;
            var entry = FindBrowseEntry(m.DictType, m.EntryKey);
            string current = entry?.CurrentActive ?? m.CurrentText;
            string english = entry?.EnglishDisplay ?? m.EnglishText;

            var card = CreateStackCard(clickable: true);
            card.Controls.Add(CreateBadge(m.FriendlyLocation, AccentSoft, AccentPrimary));
            card.Controls.Add(CreateWrapLabel(current, TextPrimary, 10.5f, FontStyle.Bold, S(4), topMargin: S(8)));
            if (!IsEnglishSelected && !string.IsNullOrWhiteSpace(english) && !string.Equals(english, current, StringComparison.Ordinal))
                card.Controls.Add(CreateWrapLabel($"{Loc.T("dict.editor.english")} {english}", TextMuted, 8.5f, FontStyle.Regular, 0));
            ((HoverStackCard)card).WireClick(() =>
            {
                _selectedPickedMatchIndex = matchIndex;
                OpenEditForMatch(m, DictView.Disambiguation);
            });
            list.Controls.Add(card);
        }

        layout.Controls.Add(list, 0, 2);
        return layout;
    }

    // =========================================================================
    // 8. EDIT CARD VIEW
    // =========================================================================
    private void OpenEditForMatch(DictionaryPickerMatch match, DictView returnView)
    {
        var entry = FindBrowseEntry(match.DictType, match.EntryKey) ?? new BrowseEntry
        {
            DictType = match.DictType,
            EntryKey = match.EntryKey,
            English = match.EnglishText,
            EnglishDisplay = FormatEnglishDisplay(match.EntryKey, match.EnglishText),
            BuiltInDefault = match.EnglishText,
            CurrentActive = match.CurrentText,
            Location = LocalizationTextMatcher.GetFriendlyLocation(match.DictType, match.EntryKey),
            Category = LocalizationTextMatcher.GetFriendlyCategory(match.DictType, match.EntryKey)
        };
        OpenEdit(entry, match.FriendlyLocation, returnView);
    }

    private void OpenEdit(BrowseEntry entry, string? friendlyLocation, DictView returnView)
    {
        _editEntry = entry;
        _editFriendlyLocation = string.IsNullOrWhiteSpace(friendlyLocation) ? null : friendlyLocation;
        _editReturnView = returnView;
        ShowView(DictView.Edit);
        if (IsHandleCreated) FocusSuggestedBox();
    }

    private void FocusSuggestedBox()
    {
        if (_txtSuggested == null || _txtSuggested.IsDisposed) return;
        try
        {
            _txtSuggested.Focus();
            _txtSuggested.SelectAll();
        }
        catch { }
    }

    private Control BuildEditView()
    {
        var entry = _editEntry;
        var layout = CreateViewLayout(out var header);
        header.Controls.Add(CreateBackButton(() => ShowView(_editReturnView == DictView.Disambiguation && _pickedMatches.Count >= 2
            ? DictView.Disambiguation
            : DictView.Update)));
        header.Controls.Add(CreateViewTitle(Loc.T("dict.edit.title")));
        var btnPick = CreateSecondaryButton(Loc.T("dict.btn.pickAnother"));
        btnPick.Margin = new Padding(S(12), S(2), 0, 0);
        btnPick.Click += (s, e) => RequestScreenPick();
        header.Controls.Add(btnPick);

        var scroll = new Panel { Dock = DockStyle.Fill, AutoScroll = true, BackColor = BgPage };
        var card = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 1,
            BackColor = BgCard,
            Padding = new Padding(S(18), S(14), S(18), S(14)),
            Margin = new Padding(0)
        };
        card.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        card.Paint += (s, e) => DrawCardBorder(e.Graphics, card.ClientRectangle, BorderSubtle);

        if (entry == null)
        {
            card.Controls.Add(CreateWrapLabel(Loc.T("dict.err.selectEntry"), TextMuted, 9.5f, FontStyle.Regular, 0));
            scroll.Controls.Add(card);
            layout.Controls.Add(scroll, 0, 1);
            return layout;
        }

        if (_pickedMatches.Count >= 1 && _selectedPickedMatchIndex >= 0 && _selectedPickedMatchIndex < _pickedMatches.Count)
            card.Controls.Add(BuildPickedMatchesBar());

        string location = _editFriendlyLocation ?? entry.Location;
        var badge = CreateBadge(location, AccentSoft, AccentPrimary);
        badge.Margin = new Padding(0, 0, 0, S(12));
        card.Controls.Add(badge);

        // Currently says
        card.Controls.Add(CreateCaption(Loc.T("dict.edit.currentlySays")));
        var txtCurrent = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BorderStyle = BorderStyle.FixedSingle,
            BackColor = BgPage,
            ForeColor = TextPrimary,
            Font = new Font("Segoe UI", 10f, FontStyle.Regular),
            Height = S(58),
            Text = NormalizeNewLines(entry.CurrentActive),
            Margin = new Padding(0, S(2), 0, S(4))
        };
        card.Controls.Add(txtCurrent);

        if (!IsEnglishSelected && !string.IsNullOrWhiteSpace(entry.EnglishDisplay)
            && !string.Equals(entry.EnglishDisplay, entry.CurrentActive, StringComparison.Ordinal))
        {
            card.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Text = $"{Loc.T("dict.editor.english")} {SingleLinePreview(entry.EnglishDisplay)}",
                ForeColor = TextMuted,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Regular),
                Margin = new Padding(0, 0, 0, S(6))
            });
        }

        // Your improved wording
        card.Controls.Add(CreateCaption(Loc.T("dict.edit.yourWording"), AccentPrimary, topMargin: S(8)));
        _txtSuggested = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            AcceptsReturn = true,
            ScrollBars = ScrollBars.Vertical,
            Font = new Font("Segoe UI", 10.5f, FontStyle.Regular),
            Height = S(84),
            Text = NormalizeNewLines(entry.CurrentActive),
            Margin = new Padding(0, S(2), 0, S(4))
        };
        card.Controls.Add(_txtSuggested);

        string placeholders = CollectiveDictionaryRepository.GetPlaceholderListDisplay(entry.English);
        if (placeholders.Length > 0)
        {
            card.Controls.Add(new Label
            {
                Dock = DockStyle.Fill,
                AutoSize = true,
                Text = Loc.T("dict.editor.placeholderHint", placeholders),
                ForeColor = AccentPrimary,
                Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
                Margin = new Padding(0, 0, 0, S(6))
            });
        }

        // Reason / note + submit
        card.Controls.Add(CreateCaption(Loc.T("dict.edit.reasonLabel"), topMargin: S(8)));
        var noteRow = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 2,
            BackColor = Color.Transparent,
            Margin = new Padding(0, S(2), 0, 0)
        };
        noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        noteRow.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        _txtNote = new TextBox
        {
            Dock = DockStyle.Fill,
            PlaceholderText = Loc.T("dict.editor.notePlaceholder"),
            Margin = new Padding(0, S(4), S(10), 0)
        };
        _btnSubmit = CreatePrimaryButton(Loc.T("dict.btn.submit"));
        _btnSubmit.Margin = new Padding(0);
        _btnSubmit.Click += async (s, e) => await OnSubmitSuggestionClickedAsync();
        noteRow.Controls.Add(_txtNote, 0, 0);
        noteRow.Controls.Add(_btnSubmit, 1, 0);
        card.Controls.Add(noteRow);

#if DEBUG
        card.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Text = $"{entry.DictType} :: {entry.EntryKey}",
            ForeColor = Color.FromArgb(148, 163, 184),
            Font = new Font("Consolas", 8f, FontStyle.Regular),
            Margin = new Padding(0, S(14), 0, 0)
        });
#endif

        scroll.Controls.Add(card);
        layout.Controls.Add(scroll, 0, 1);
        return layout;
    }

    /// <summary>Highlighted bar at the top of the Edit Card listing every possible pick near the click.</summary>
    private Control BuildPickedMatchesBar()
    {
        var bar = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            ColumnCount = 1,
            RowCount = 2,
            BackColor = AccentSoft,
            Padding = new Padding(S(10), S(8), S(10), S(10)),
            Margin = new Padding(0, 0, 0, S(14))
        };
        bar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        bar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bar.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bar.Paint += (s, e) => DrawCardBorder(e.Graphics, bar.ClientRectangle, AccentPrimary);

        bar.Controls.Add(new Label
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            Text = Loc.T("dict.edit.pickedDropdown", _pickedMatches.Count),
            ForeColor = AccentPrimary,
            BackColor = Color.Transparent,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, S(4))
        }, 0, 0);

        var cmb = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Dock = DockStyle.Fill,
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            Margin = new Padding(0)
        };

        _suppressPickerDropdown = true;
        try
        {
            foreach (var m in _pickedMatches)
            {
                string currentText = FindBrowseEntry(m.DictType, m.EntryKey)?.CurrentActive ?? m.CurrentText;
                cmb.Items.Add($"{m.FriendlyLocation} — {SingleLinePreview(currentText, 65)}");
            }
            cmb.SelectedIndex = Math.Clamp(_selectedPickedMatchIndex, 0, cmb.Items.Count - 1);
        }
        finally
        {
            _suppressPickerDropdown = false;
        }

        cmb.SelectedIndexChanged += (s, e) =>
        {
            if (_suppressPickerDropdown) return;
            int idx = cmb.SelectedIndex;
            if (idx < 0 || idx >= _pickedMatches.Count || idx == _selectedPickedMatchIndex) return;

            // Rebuild outside the combo's own event (the view rebuild disposes this combo).
            BeginInvoke(new Action(() =>
            {
                if (IsDisposed || _currentView != DictView.Edit) return;
                _selectedPickedMatchIndex = idx;
                OpenEditForMatch(_pickedMatches[idx], _editReturnView);
                try { _cmbPickedMatches?.Focus(); } catch { }
            }));
        };

        _cmbPickedMatches = cmb;
        bar.Controls.Add(cmb, 0, 1);
        return bar;
    }

    // =========================================================================
    // 9. DEBUG: LIVE OVERRIDES VIEW
    // =========================================================================
#if DEBUG
    private Control BuildOverridesView()
    {
        var layout = new TableLayoutPanel
        {
            ColumnCount = 1,
            RowCount = 3,
            BackColor = BgPage
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var header = CreateHeaderFlow();
        header.Controls.Add(CreateBackButton(() => ShowView(DictView.Home)));
        header.Controls.Add(CreateViewTitle(Loc.T("dict.tab.overrides", _liveOverrides.Count)));
        layout.Controls.Add(header, 0, 0);

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
            if (_btnRevertOverride != null) _btnRevertOverride.Enabled = !_isBusy && GetSelectedOverride() != null;
        };
        layout.Controls.Add(_dgvOverrides, 0, 1);

        var actionBar = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Padding = new Padding(0, S(8), 0, 0)
        };
        _btnRevertOverride = CreateDangerButton(Loc.T("dict.btn.revert"));
        _btnRevertOverride.Enabled = false;
        _btnRevertOverride.Click += async (s, e) => await OnRevertOverrideClickedAsync();
        actionBar.Controls.Add(_btnRevertOverride);
        layout.Controls.Add(actionBar, 0, 2);

        PopulateOverridesGrid();
        return layout;
    }

    private void PopulateOverridesGrid()
    {
        var dgv = _dgvOverrides;
        if (dgv == null || dgv.IsDisposed) return;
        string lang = SelectedLangCode;

        var uiEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeUi, "en");
        var tipEn = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var tipEnCollective = CollectiveDictionaryRepository.LoadCachedOverrides(CollectiveDictionaryRepository.DictTypeTooltips, "en");
        var uiBase = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeUi, lang);
        var tipBase = Loc.GetBaseDictionary(CollectiveDictionaryRepository.DictTypeTooltips, lang);

        string uiLabel = Loc.T("dict.type.ui");
        string tipLabel = Loc.T("dict.type.tooltips");

        dgv.SuspendLayout();
        dgv.Rows.Clear();
        foreach (var o in _liveOverrides)
        {
            bool isTip = string.Equals(o.DictType, CollectiveDictionaryRepository.DictTypeTooltips, StringComparison.OrdinalIgnoreCase);
            (isTip ? tipEn : uiEn).TryGetValue(o.EntryKey, out string? enVal);
            if (isTip && string.IsNullOrEmpty(enVal) && tipEnCollective.TryGetValue(o.EntryKey, out string? enOv))
                enVal = enOv;
            (isTip ? tipBase : uiBase).TryGetValue(o.EntryKey, out string? defVal);
            string enDisplay = FormatEnglishDisplay(o.EntryKey, enVal);
            defVal ??= enVal ?? string.Empty;

            int idx = dgv.Rows.Add(
                isTip ? tipLabel : uiLabel,
                o.EntryKey,
                SingleLinePreview(enDisplay),
                SingleLinePreview(defVal),
                SingleLinePreview(o.Value),
                o.ApprovedBy,
                o.ApprovedAt.ToLocalTime().ToString("yyyy-MM-dd HH:mm"));
            dgv.Rows[idx].Tag = o;
        }
        dgv.ResumeLayout();
        if (_btnRevertOverride != null) _btnRevertOverride.Enabled = !_isBusy && GetSelectedOverride() != null;
    }

    private CollectiveOverrideRecord? GetSelectedOverride()
        => _dgvOverrides != null && _dgvOverrides.SelectedRows.Count > 0
            ? _dgvOverrides.SelectedRows[0].Tag as CollectiveOverrideRecord
            : null;
#endif

    // =========================================================================
    // 10. DATA LOADING
    // =========================================================================
    private void RebuildBrowseSource()
    {
        string lang = SelectedLangCode;
        var entries = new List<BrowseEntry>();
        AddEntriesForType(entries, CollectiveDictionaryRepository.DictTypeUi, lang);
        AddEntriesForType(entries, CollectiveDictionaryRepository.DictTypeTooltips, lang);
        _allBrowseEntries = entries;
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
                IsCollectiveOverride = collective.ContainsKey(key),
                Location = LocalizationTextMatcher.GetFriendlyLocation(dictType, key),
                Category = LocalizationTextMatcher.GetFriendlyCategory(dictType, key)
            });
        }
    }

    private BrowseEntry? FindBrowseEntry(string dictType, string entryKey)
        => _allBrowseEntries.FirstOrDefault(e =>
            string.Equals(e.DictType, dictType, StringComparison.OrdinalIgnoreCase)
            && string.Equals(e.EntryKey, entryKey, StringComparison.Ordinal));

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

    private async Task RefreshServerDataAsync()
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
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            _pendingSuggestions = new List<CollectiveSuggestionRecord>();
#if DEBUG
            _liveOverrides = new List<CollectiveOverrideRecord>();
#endif
            SetOfflineBanner(ex.Message);
        }

        RefreshServerBackedView();
    }

    private void UpdateStatusLabel()
    {
        if (_offlineBanner.Visible)
        {
            _lblStatus.Text = Loc.T("dict.status.offline");
            return;
        }
        if (_currentView != DictView.Update)
            _lblStatus.Text = Loc.T("dict.status.ready", _allBrowseEntries.Count);
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
        _cmbLanguage.Enabled = !busy;
        if (_viewHost != null && !_viewHost.IsDisposed) _viewHost.Enabled = !busy;
        UseWaitCursor = busy;
    }

    // =========================================================================
    // 11. USER ACTIONS
    // =========================================================================
    private async Task OnSyncClickedAsync()
    {
        SetBusy(true);
        _btnSync.Text = Loc.T("dict.btn.syncing");
        bool ok = false;
        try
        {
            ok = await Task.Run(() => CollectiveDictionaryRepository.SyncFromServer(reloadLoc: true));
            if (IsDisposed) return;
            RebuildBrowseSource();
        }
        finally
        {
            if (!IsDisposed)
            {
                _btnSync.Text = Loc.T("dict.btn.sync");
                SetBusy(false);
            }
        }

        if (IsDisposed) return;
        if (_currentView == DictView.Update) PopulateUpdateGrid();
        await RefreshServerDataAsync();
        if (!IsDisposed) _lblStatus.Text = ok ? Loc.T("dict.status.synced") : Loc.T("dict.status.offline");
    }

    private async Task OnSuggestNewTipClickedAsync()
    {
        if (_isBusy) return;
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
                    _selectedLangCode = ReadLangCodeFromCombo();
                    RebuildBrowseSource();
                }
            }

            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
            MessageBox.Show(
                this,
                Loc.T("dict.newTip.submitted", CollectiveDictionaryRepository.ConsensusVoteThreshold),
                Text,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
            ShowView(DictView.Upvote);
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
        var entry = _editEntry;
        if (entry == null || _txtSuggested == null)
        {
            MessageBox.Show(this, Loc.T("dict.err.selectEntry"), Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        string suggested = DenormalizeNewLines((_txtSuggested.Text ?? string.Empty).Trim());
        string note = (_txtNote?.Text ?? string.Empty).Trim();

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
        if (_btnSubmit != null) _btnSubmit.Text = Loc.T("dict.btn.submitting");
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.SubmitSuggestion(
                entry.DictType, entry.EntryKey, lang, suggested, note));
            if (IsDisposed) return;

            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
            MessageBox.Show(this, Loc.T("dict.msg.submitted", CollectiveDictionaryRepository.ConsensusVoteThreshold),
                Text, MessageBoxButtons.OK, MessageBoxIcon.Information);
            ShowView(DictView.Upvote);
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
                if (_btnSubmit != null && !_btnSubmit.IsDisposed) _btnSubmit.Text = Loc.T("dict.btn.submit");
                SetBusy(false);
            }
        }
    }

    private async Task OnToggleVoteClickedAsync(CollectiveSuggestionRecord suggestion)
    {
        if (suggestion == null || _isBusy) return;

        SetBusy(true);
        try
        {
            var result = await Task.Run(() => CollectiveDictionaryRepository.ToggleVote(suggestion.SuggestionId));
            if (IsDisposed) return;

            RebuildBrowseSource();
            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;

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
    private async Task OnAdminApproveClickedAsync(CollectiveSuggestionRecord sel)
    {
        if (sel == null || _isBusy) return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminApproveSuggestion(sel.SuggestionId));
            if (IsDisposed) return;

            RebuildBrowseSource();
            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
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

    private async Task OnAdminEditApproveClickedAsync(CollectiveSuggestionRecord sel)
    {
        if (sel == null || _isBusy) return;

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
            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
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

    private async Task OnAdminRejectClickedAsync(CollectiveSuggestionRecord sel)
    {
        if (sel == null || _isBusy) return;

        if (MessageBox.Show(this, Loc.T("dict.confirm.reject"), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
            return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminRejectSuggestion(sel.SuggestionId));
            if (IsDisposed) return;

            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
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
        if (sel == null || _isBusy) return;

        if (MessageBox.Show(this, Loc.T("dict.confirm.revert", sel.EntryKey), Text, MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
            return;

        SetBusy(true);
        try
        {
            await Task.Run(() => CollectiveDictionaryRepository.AdminRevertOverride(sel.DictType, sel.EntryKey, sel.LangCode));
            if (IsDisposed) return;

            RebuildBrowseSource();
            SetBusy(false);
            await RefreshServerDataAsync();
            if (IsDisposed) return;
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
    // 12. VIEW BUILDING BLOCKS
    // =========================================================================
    /// <summary>Header row + content area (row 1, or row 2 when <paramref name="extraHeaderRow"/>).</summary>
    private TableLayoutPanel CreateViewLayout(out FlowLayoutPanel header, bool extraHeaderRow = false)
    {
        var layout = new TableLayoutPanel
        {
            ColumnCount = 1,
            RowCount = extraHeaderRow ? 3 : 2,
            BackColor = BgPage,
            Padding = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        if (extraHeaderRow) layout.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        header = CreateHeaderFlow();
        layout.Controls.Add(header, 0, 0);
        return layout;
    }

    private FlowLayoutPanel CreateHeaderFlow() => new()
    {
        Dock = DockStyle.Fill,
        AutoSize = true,
        FlowDirection = FlowDirection.LeftToRight,
        WrapContents = false,
        BackColor = BgPage,
        Margin = new Padding(0, 0, 0, S(10))
    };

    private Button CreateBackButton(Action onClick)
    {
        var btn = CreateSecondaryButton(Loc.T("dict.nav.back"));
        btn.Margin = new Padding(0, S(2), S(12), 0);
        btn.Click += (s, e) => { if (!_isBusy) onClick(); };
        return btn;
    }

    private Label CreateViewTitle(string text) => new()
    {
        Text = text,
        AutoSize = true,
        ForeColor = TextPrimary,
        Font = new Font("Segoe UI", 13f, FontStyle.Bold),
        Margin = new Padding(0, S(3), S(10), 0)
    };

    private Label CreateCaption(string text, Color? color = null, int topMargin = 0) => new()
    {
        Text = text,
        AutoSize = true,
        ForeColor = color ?? TextMuted,
        Font = new Font("Segoe UI", 8.5f, FontStyle.Bold),
        Margin = new Padding(0, topMargin, 0, S(2))
    };

    private static Label CreateWrapLabel(string text, Color color, float size, FontStyle style, int bottomMargin, int topMargin = 0) => new()
    {
        Text = NormalizeNewLines(text),
        AutoSize = true,
        ForeColor = color,
        Font = new Font("Segoe UI", size, style),
        Margin = new Padding(0, topMargin, 0, bottomMargin),
        UseMnemonic = false
    };

    /// <summary>Vertical auto-scrolling list whose cards always track the container width.</summary>
    private FlowLayoutPanel CreateCardList()
    {
        var list = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            BackColor = BgPage,
            Padding = new Padding(0, 0, S(4), S(8)),
            Margin = new Padding(0)
        };
        list.Layout += (s, e) => FitCardsToWidth(list);
        list.ControlAdded += (s, e) => FitCardsToWidth(list);
        return list;
    }

    private bool _fittingCards;

    private void FitCardsToWidth(FlowLayoutPanel list)
    {
        if (_fittingCards || list.IsDisposed) return;
        _fittingCards = true;
        try
        {
            int width = list.ClientSize.Width - list.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth - S(2);
            if (width < S(200)) width = S(200);
            foreach (Control card in list.Controls)
            {
                int cardWidth = width - card.Margin.Horizontal;
                if (card.MinimumSize.Width == cardWidth && card.MaximumSize.Width == cardWidth) continue;
                card.MinimumSize = new Size(cardWidth, 0);
                card.MaximumSize = new Size(cardWidth, 0);
                int labelWidth = cardWidth - card.Padding.Horizontal - S(4);
                foreach (Control child in card.Controls)
                {
                    if (child is Label lbl && lbl.AutoSize) lbl.MaximumSize = new Size(Math.Max(S(80), labelWidth), 0);
                    else if (child is FlowLayoutPanel flow) flow.MaximumSize = new Size(Math.Max(S(80), labelWidth), 0);
                }
            }
        }
        finally
        {
            _fittingCards = false;
        }
    }

    /// <summary>White bordered card that auto-sizes its height to stacked children.</summary>
    private FlowLayoutPanel CreateStackCard(bool clickable = false)
    {
        FlowLayoutPanel card = clickable
            ? new HoverStackCard(BgCard, CardHover, BorderSubtle, AccentPrimary) { Cursor = Cursors.Hand }
            : new HoverStackCard(BgCard, BgCard, BorderSubtle, BorderSubtle);
        card.FlowDirection = FlowDirection.TopDown;
        card.WrapContents = false;
        card.AutoSize = true;
        card.AutoSizeMode = AutoSizeMode.GrowAndShrink;
        card.Padding = new Padding(S(16), S(12), S(16), S(12));
        card.Margin = new Padding(0, 0, 0, S(10));
        return card;
    }

    private static void DrawCardBorder(Graphics g, Rectangle bounds, Color border)
    {
        g.SmoothingMode = SmoothingMode.AntiAlias;
        using var pen = new Pen(border, 1f);
        var r = new Rectangle(bounds.X, bounds.Y, bounds.Width - 1, bounds.Height - 1);
        using var path = RoundedRect(r, 8);
        g.DrawPath(pen, path);
    }

    private static GraphicsPath RoundedRect(Rectangle r, int radius)
    {
        var path = new GraphicsPath();
        int d = Math.Max(2, Math.Min(radius * 2, Math.Min(r.Width, r.Height)));
        path.AddArc(r.X, r.Y, d, d, 180, 90);
        path.AddArc(r.Right - d, r.Y, d, d, 270, 90);
        path.AddArc(r.Right - d, r.Bottom - d, d, d, 0, 90);
        path.AddArc(r.X, r.Bottom - d, d, d, 90, 90);
        path.CloseFigure();
        return path;
    }

    /// <summary>Clickable rounded card panel with hover highlight (used for Home cards).</summary>
    private sealed class HoverCard : Panel
    {
        private readonly Color _normal;
        private readonly Color _hover;
        private readonly Color _border;
        private readonly Color _borderHover;
        private bool _isHover;

        public HoverCard(Color normal, Color hover, Color border, Color borderHover)
        {
            _normal = normal;
            _hover = hover;
            _border = border;
            _borderHover = borderHover;
            BackColor = normal;
            DoubleBuffered = true;
            ResizeRedraw = true;
        }

        public void WireClick(Action onClick) => CardWiring.Wire(this, onClick, SetHover);

        private void SetHover(bool hover)
        {
            if (_isHover == hover) return;
            _isHover = hover;
            BackColor = hover ? _hover : _normal;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            DrawCardBorder(e.Graphics, ClientRectangle, _isHover ? _borderHover : _border);
        }
    }

    /// <summary>Auto-sizing stacked card (FlowLayoutPanel) with optional click + hover highlight.</summary>
    private sealed class HoverStackCard : FlowLayoutPanel
    {
        private readonly Color _normal;
        private readonly Color _hover;
        private readonly Color _border;
        private readonly Color _borderHover;
        private bool _isHover;

        public HoverStackCard(Color normal, Color hover, Color border, Color borderHover)
        {
            _normal = normal;
            _hover = hover;
            _border = border;
            _borderHover = borderHover;
            BackColor = normal;
            DoubleBuffered = true;
            ResizeRedraw = true;
        }

        public void WireClick(Action onClick) => CardWiring.Wire(this, onClick, SetHover);

        private void SetHover(bool hover)
        {
            if (_isHover == hover) return;
            _isHover = hover;
            BackColor = hover ? _hover : _normal;
            Invalidate();
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);
            DrawCardBorder(e.Graphics, ClientRectangle, _isHover ? _borderHover : _border);
        }
    }

    private static class CardWiring
    {
        public static void Wire(Control card, Action onClick, Action<bool> setHover)
        {
            void Attach(Control c)
            {
                if (c is Button || c is TextBoxBase) return;
                c.Click += (s, e) => onClick();
                c.MouseEnter += (s, e) => setHover(true);
                c.MouseLeave += (s, e) =>
                {
                    if (card.IsDisposed) return;
                    var p = card.PointToClient(Cursor.Position);
                    setHover(card.ClientRectangle.Contains(p));
                };
                if (c != card) c.Cursor = Cursors.Hand;
                foreach (Control child in c.Controls) Attach(child);
                c.ControlAdded += (s, e) => { if (e.Control != null) Attach(e.Control); };
            }
            Attach(card);
        }
    }

    // =========================================================================
    // 13. UI FACTORY HELPERS
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
        dgv.Cursor = Cursors.Hand;
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
            Cursor = Cursors.Hand,
            UseMnemonic = false
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
            Cursor = Cursors.Hand,
            UseMnemonic = false
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
            Cursor = Cursors.Hand,
            UseMnemonic = false
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

    private static string SingleLinePreview(string? text, int maxLen)
    {
        string single = SingleLinePreview(text);
        if (maxLen > 3 && single.Length > maxLen) single = single.Substring(0, maxLen - 3) + "...";
        return single;
    }

    /// <summary>Dictionary text uses "\n"; WinForms multi-line text boxes need "\r\n".</summary>
    private static string NormalizeNewLines(string? text)
    {
        if (string.IsNullOrEmpty(text)) return string.Empty;
        return text.Replace("\r\n", "\n").Replace("\n", "\r\n");
    }

    private static string DenormalizeNewLines(string text) => text.Replace("\r\n", "\n");
}
