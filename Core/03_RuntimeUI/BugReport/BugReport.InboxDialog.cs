namespace SKSB_App.Core.RuntimeUI.BugReport;

#if DEBUG
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Diagnostics;

/// <summary>
/// DEBUG-only triage window for <c>public.bug_reports</c>: browse, preview attachments,
/// read the runtime log tail, update status / dev notes and delete reports.
/// </summary>
public sealed class BugReportInboxDialog : Form
{
    // =========================================================================
    // 1. PALETTE, STATE & CONTROLS
    // =========================================================================
    private static readonly Color Slate900 = Color.FromArgb(15, 23, 42);
    private static readonly Color Slate800 = Color.FromArgb(30, 41, 59);
    private static readonly Color Slate500 = Color.FromArgb(100, 116, 139);
    private static readonly Color Slate400 = Color.FromArgb(148, 163, 184);
    private static readonly Color Slate200 = Color.FromArgb(226, 232, 240);
    private static readonly Color Slate50 = Color.FromArgb(248, 250, 252);
    private static readonly Color Indigo600 = Color.FromArgb(79, 70, 229);

    private static readonly string[] FilterOptions =
    {
        "Active (Open & In Progress)", "All",
        BugReportStatus.Open, BugReportStatus.InProgress, BugReportStatus.Resolved, BugReportStatus.Closed
    };

    private readonly Dictionary<long, byte[]> _dataCache = new();
    private readonly Dictionary<long, Image> _imageCache = new();
    private List<BugReportSummary> _reports = new();
    private BugReportSummary? _current;
    private int _selectionVersion;
    private bool _busy;
    private Font? _statusFont;

    private ComboBox _cmbFilter = null!;
    private Button _btnRefresh = null!;
    private Label _lblSummary = null!;
    private SplitContainer _split = null!;
    private DataGridView _grid = null!;
    private Label _lblBugTitle = null!;
    private Label _lblBugMeta = null!;
    private TextBox _txtDescription = null!;
    private ListBox _lstAttachments = null!;
    private Button _btnOpen = null!;
    private Button _btnSaveAs = null!;
    private PictureBox _picPreview = null!;
    private Label _lblPreviewInfo = null!;
    private TextBox _txtLog = null!;
    private ComboBox _cmbStatus = null!;
    private TextBox _txtDevNotes = null!;
    private Button _btnSaveTriage = null!;
    private Button _btnDelete = null!;
    private TableLayoutPanel _detailRoot = null!;

    public BugReportInboxDialog()
    {
        SuspendLayout();
        Text = "SKSB Bug Reports Inbox (Debug)";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = true;
        MaximizeBox = true;
        ShowInTaskbar = false;
        BackColor = Slate50;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(12);
        ClientSize = new Size(1100, 700);
        MinimumSize = new Size(820, 520);

        InitializeLayout();

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    // =========================================================================
    // 2. LAYOUT
    // =========================================================================
    private void InitializeLayout()
    {
        // --- Top bar ---
        var topBar = new FlowLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(0, 0, 0, 8)
        };
        var lblInbox = new Label
        {
            Text = "Bug Reports",
            Font = new Font("Segoe UI", 13f, FontStyle.Bold),
            ForeColor = Slate900,
            AutoSize = true,
            Margin = new Padding(0, 0, 16, 0)
        };
        var lblFilter = new Label { Text = "Status:", AutoSize = true, ForeColor = Slate800, Margin = new Padding(0, 6, 4, 0) };
        _cmbFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200, Margin = new Padding(0, 2, 8, 0) };
        _cmbFilter.Items.AddRange(FilterOptions);
        _cmbFilter.SelectedIndex = 0;
        _cmbFilter.SelectedIndexChanged += async (s, e) => await ReloadAsync(_current?.BugId);
        _btnRefresh = CreateButton("Refresh", primary: false);
        _btnRefresh.Margin = new Padding(0, 1, 12, 0);
        _btnRefresh.Click += async (s, e) => await ReloadAsync(_current?.BugId);
        _lblSummary = new Label { AutoSize = true, ForeColor = Slate500, Margin = new Padding(0, 6, 0, 0) };
        topBar.Controls.AddRange(new Control[] { lblInbox, lblFilter, _cmbFilter, _btnRefresh, _lblSummary });

        // --- Split ---
        _split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            BackColor = Slate200,
            SplitterWidth = 5
        };
        _split.Panel1.BackColor = Color.White;
        _split.Panel2.BackColor = Slate50;

        _grid = BuildGrid();
        _split.Panel1.Controls.Add(_grid);
        _split.Panel2.Padding = new Padding(10, 0, 0, 0);
        _split.Panel2.Controls.Add(BuildDetailPane());

        Controls.Add(_split);
        Controls.Add(topBar);
    }

    private DataGridView BuildGrid()
    {
        var grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            MultiSelect = false,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            RowHeadersVisible = false,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            GridColor = Slate200,
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells
        };
        grid.ColumnHeadersDefaultCellStyle.BackColor = Slate50;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Slate800;
        grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9f, FontStyle.Bold);
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 231, 255);
        grid.DefaultCellStyle.SelectionForeColor = Slate900;

        DataGridViewTextBoxColumn Col(string name, string header, int weight, DataGridViewAutoSizeColumnMode mode = DataGridViewAutoSizeColumnMode.Fill)
            => new() { Name = name, HeaderText = header, FillWeight = weight, AutoSizeMode = mode, SortMode = DataGridViewColumnSortMode.NotSortable };

        grid.Columns.Add(Col("colId", "#", 10, DataGridViewAutoSizeColumnMode.AllCells));
        grid.Columns.Add(Col("colStatus", "Status", 18, DataGridViewAutoSizeColumnMode.AllCells));
        grid.Columns.Add(Col("colTitle", "Title", 60));
        grid.Columns.Add(Col("colModule", "Module", 22));
        grid.Columns.Add(Col("colBy", "By", 18));
        grid.Columns.Add(Col("colFiles", "Files", 8, DataGridViewAutoSizeColumnMode.AllCells));
        grid.Columns.Add(Col("colCreated", "Created", 24, DataGridViewAutoSizeColumnMode.AllCells));

        grid.CellFormatting += OnGridCellFormatting;
        grid.SelectionChanged += async (s, e) => await OnGridSelectionChangedAsync();
        return grid;
    }

    private Control BuildDetailPane()
    {
        _detailRoot = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 3, Margin = new Padding(0) };
        _detailRoot.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _detailRoot.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _detailRoot.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        _detailRoot.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        // --- Header ---
        var header = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, WrapContents = false, AutoSize = true, Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 6) };
        _lblBugTitle = new Label
        {
            Text = "Select a bug report",
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            ForeColor = Slate900,
            AutoSize = true,
            MaximumSize = new Size(560, 0),
            Margin = new Padding(0, 0, 0, 2)
        };
        _lblBugMeta = new Label { Text = string.Empty, ForeColor = Slate500, AutoSize = true, MaximumSize = new Size(560, 0), Margin = new Padding(0) };
        header.Controls.Add(_lblBugTitle);
        header.Controls.Add(_lblBugMeta);
        header.SizeChanged += (s, e) =>
        {
            int w = Math.Max(100, header.ClientSize.Width - 4);
            _lblBugTitle.MaximumSize = new Size(w, 0);
            _lblBugMeta.MaximumSize = new Size(w, 0);
        };

        // --- Tabs ---
        var tabs = new TabControl { Dock = DockStyle.Fill, Margin = new Padding(0, 0, 0, 6) };
        var tabDetails = new TabPage("Details & Attachments") { BackColor = Color.White, Padding = new Padding(8) };
        var tabLog = new TabPage("Runtime Trace Log") { BackColor = Color.White, Padding = new Padding(8) };

        var details = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 1, RowCount = 4 };
        details.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        details.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        details.RowStyles.Add(new RowStyle(SizeType.Percent, 35f));
        details.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        details.RowStyles.Add(new RowStyle(SizeType.Percent, 65f));

        _txtDescription = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            BackColor = Slate50,
            Margin = new Padding(0, 0, 0, 8)
        };

        var attachArea = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 2, Margin = new Padding(0) };
        attachArea.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40f));
        attachArea.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60f));
        attachArea.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        attachArea.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        _lstAttachments = new ListBox { Dock = DockStyle.Fill, IntegralHeight = false, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0, 0, 8, 0) };
        _lstAttachments.SelectedIndexChanged += async (s, e) => await UpdatePreviewAsync();
        _lstAttachments.DoubleClick += async (s, e) => await OpenSelectedAttachmentAsync();

        var attachButtons = new FlowLayoutPanel { AutoSize = true, Dock = DockStyle.Fill, WrapContents = true, Margin = new Padding(0, 6, 0, 0) };
        _btnOpen = CreateButton("Open in Default App", primary: false);
        _btnSaveAs = CreateButton("Save As...", primary: false);
        _btnOpen.Click += async (s, e) => await OpenSelectedAttachmentAsync();
        _btnSaveAs.Click += async (s, e) => await SaveSelectedAttachmentAsync();
        attachButtons.Controls.Add(_btnOpen);
        attachButtons.Controls.Add(_btnSaveAs);

        var previewHost = new Panel { Dock = DockStyle.Fill, BackColor = Slate50, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0) };
        _picPreview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Slate50 };
        _lblPreviewInfo = new Label { Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleCenter, ForeColor = Slate400, Text = "No attachment selected" };
        previewHost.Controls.Add(_picPreview);
        previewHost.Controls.Add(_lblPreviewInfo);
        _lblPreviewInfo.BringToFront();

        attachArea.Controls.Add(_lstAttachments, 0, 0);
        attachArea.Controls.Add(previewHost, 1, 0);
        attachArea.SetRowSpan(previewHost, 2);
        attachArea.Controls.Add(attachButtons, 0, 1);

        details.Controls.Add(CreateSectionLabel("Description"), 0, 0);
        details.Controls.Add(_txtDescription, 0, 1);
        details.Controls.Add(CreateSectionLabel("Attachments"), 0, 2);
        details.Controls.Add(attachArea, 0, 3);
        tabDetails.Controls.Add(details);

        _txtLog = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            ReadOnly = true,
            WordWrap = false,
            ScrollBars = ScrollBars.Both,
            Font = new Font("Consolas", 9f, FontStyle.Regular),
            BackColor = Slate50
        };
        tabLog.Controls.Add(_txtLog);

        tabs.TabPages.Add(tabDetails);
        tabs.TabPages.Add(tabLog);

        // --- Triage bar ---
        var triage = new TableLayoutPanel { Dock = DockStyle.Fill, AutoSize = true, ColumnCount = 6, RowCount = 1, Margin = new Padding(0) };
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        triage.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));

        _cmbStatus = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 120, Margin = new Padding(0, 3, 12, 0) };
        _cmbStatus.Items.AddRange(BugReportStatus.AllStatuses);
        _txtDevNotes = new TextBox { Dock = DockStyle.Fill, PlaceholderText = "Developer notes (root cause, fix commit, etc.)", Margin = new Padding(0, 3, 8, 0) };
        _btnSaveTriage = CreateButton("Save Triage", primary: true);
        _btnSaveTriage.Click += async (s, e) => await SaveTriageAsync();
        _btnDelete = CreateButton("Delete", primary: false);
        _btnDelete.ForeColor = Color.FromArgb(185, 28, 28);
        _btnDelete.Margin = new Padding(6, 0, 0, 0);
        _btnDelete.Click += async (s, e) => await DeleteCurrentAsync();

        triage.Controls.Add(new Label { Text = "Status:", AutoSize = true, ForeColor = Slate800, Margin = new Padding(0, 7, 4, 0) }, 0, 0);
        triage.Controls.Add(_cmbStatus, 1, 0);
        triage.Controls.Add(new Label { Text = "Dev Notes:", AutoSize = true, ForeColor = Slate800, Margin = new Padding(0, 7, 4, 0) }, 2, 0);
        triage.Controls.Add(_txtDevNotes, 3, 0);
        triage.Controls.Add(_btnSaveTriage, 4, 0);
        triage.Controls.Add(_btnDelete, 5, 0);

        _detailRoot.Controls.Add(header, 0, 0);
        _detailRoot.Controls.Add(tabs, 0, 1);
        _detailRoot.Controls.Add(triage, 0, 2);
        _detailRoot.Enabled = false;
        return _detailRoot;
    }

    private static Label CreateSectionLabel(string text) => new()
    {
        Text = text,
        Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        ForeColor = Slate800,
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 3)
    };

    private static Button CreateButton(string text, bool primary)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            MinimumSize = new Size(0, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = primary ? Indigo600 : Color.White,
            ForeColor = primary ? Color.White : Slate800,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 6, 0)
        };
        if (primary) btn.FlatAppearance.BorderSize = 0;
        else btn.FlatAppearance.BorderColor = Slate200;
        return btn;
    }

    protected override async void OnShown(EventArgs e)
    {
        base.OnShown(e);
        try { _split.SplitterDistance = (int)(_split.ClientSize.Width * 0.46); } catch { }
        await ReloadAsync(null);
    }

    // =========================================================================
    // 3. LIST LOADING & GRID
    // =========================================================================
    private string CurrentFilterKey()
    {
        return _cmbFilter.SelectedIndex switch
        {
            0 => "Active",
            1 => "All",
            _ => _cmbFilter.SelectedItem?.ToString() ?? "All"
        };
    }

    private async Task ReloadAsync(long? reselectBugId)
    {
        if (_busy) return;
        string filter = CurrentFilterKey();
        SetBusy(true, "Loading...");
        try
        {
            _reports = await Task.Run(() => BugReportRepository.ListBugReports(filter));
            PopulateGrid(reselectBugId);
            int open = _reports.Count(r => r.Status == BugReportStatus.Open);
            int inProgress = _reports.Count(r => r.Status == BugReportStatus.InProgress);
            _lblSummary.Text = $"{_reports.Count} report(s) shown  •  {open} open  •  {inProgress} in progress";
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[BugReport] Inbox load failed: {ex.Message}");
            _lblSummary.Text = "Failed to load — server unreachable?";
            MessageBox.Show(this, $"Failed to load bug reports:\n\n{ex.Message}", "Bug Inbox", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetBusy(false, null);
        }
    }

    private void PopulateGrid(long? reselectBugId)
    {
        _grid.SuspendLayout();
        _grid.Rows.Clear();
        int reselectIndex = -1;
        foreach (var r in _reports)
        {
            int idx = _grid.Rows.Add(
                r.BugId,
                r.Status,
                r.Title,
                r.SourceModule,
                r.ReportedBy,
                r.AttachmentCount,
                ToLocal(r.CreatedAt).ToString("yyyy-MM-dd HH:mm"));
            _grid.Rows[idx].Tag = r;
            if (reselectBugId.HasValue && r.BugId == reselectBugId.Value) reselectIndex = idx;
        }
        _grid.ResumeLayout();

        _grid.ClearSelection();
        if (_grid.Rows.Count == 0)
        {
            ShowDetails(null);
            return;
        }
        int target = reselectIndex >= 0 ? reselectIndex : 0;
        _grid.Rows[target].Selected = true;
        _grid.CurrentCell = _grid.Rows[target].Cells[0];
    }

    private void OnGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || _grid.Columns[e.ColumnIndex].Name != "colStatus" || e.CellStyle == null) return;
        string status = e.Value?.ToString() ?? string.Empty;
        (Color back, Color fore) = status switch
        {
            BugReportStatus.Open => (Color.FromArgb(254, 226, 226), Color.FromArgb(153, 27, 27)),
            BugReportStatus.InProgress => (Color.FromArgb(254, 243, 199), Color.FromArgb(146, 64, 14)),
            BugReportStatus.Resolved => (Color.FromArgb(220, 252, 231), Color.FromArgb(22, 101, 52)),
            BugReportStatus.Closed => (Color.FromArgb(241, 245, 249), Color.FromArgb(71, 85, 105)),
            _ => (Color.White, Slate800)
        };
        e.CellStyle.BackColor = back;
        e.CellStyle.ForeColor = fore;
        e.CellStyle.SelectionBackColor = back;
        e.CellStyle.SelectionForeColor = fore;
        _statusFont ??= new Font(_grid.Font, FontStyle.Bold);
        e.CellStyle.Font = _statusFont;
    }

    private async Task OnGridSelectionChangedAsync()
    {
        var summary = _grid.SelectedRows.Count > 0 ? _grid.SelectedRows[0].Tag as BugReportSummary : null;
        if (ReferenceEquals(summary, _current)) return;
        ShowDetails(summary);
        if (summary == null) return;

        int version = ++_selectionVersion;
        _lstAttachments.Items.Clear();
        if (summary.AttachmentCount == 0)
        {
            ShowPreviewMessage("No attachments");
            return;
        }

        ShowPreviewMessage("Loading attachments...");
        try
        {
            var metas = await Task.Run(() => BugReportRepository.GetAttachmentsMeta(summary.BugId));
            if (version != _selectionVersion) return;
            foreach (var m in metas) _lstAttachments.Items.Add(m);
            if (_lstAttachments.Items.Count > 0) _lstAttachments.SelectedIndex = 0;
            else ShowPreviewMessage("No attachments");
        }
        catch (Exception ex)
        {
            if (version == _selectionVersion) ShowPreviewMessage($"Failed to load attachments:\n{ex.Message}");
        }
    }

    // =========================================================================
    // 4. DETAIL PANE
    // =========================================================================
    private void ShowDetails(BugReportSummary? r)
    {
        _current = r;
        _lstAttachments.Items.Clear();
        _picPreview.Image = null; // cached images are disposed on close, not here

        if (r == null)
        {
            _detailRoot.Enabled = false;
            _lblBugTitle.Text = "Select a bug report";
            _lblBugMeta.Text = string.Empty;
            _txtDescription.Text = string.Empty;
            _txtLog.Text = string.Empty;
            _txtDevNotes.Text = string.Empty;
            _cmbStatus.SelectedIndex = -1;
            ShowPreviewMessage("No attachment selected");
            return;
        }

        _detailRoot.Enabled = true;
        _lblBugTitle.Text = $"#{r.BugId} — {r.Title}";
        var meta = new List<string> { string.IsNullOrWhiteSpace(r.SourceModule) ? "Unknown module" : r.SourceModule };
        if (!string.IsNullOrWhiteSpace(r.ActiveSheet)) meta.Add(r.ActiveSheet);
        meta.Add($"v{r.AppVersion}");
        meta.Add(r.BuildMode + (r.IsOffline ? " (Offline)" : string.Empty));
        meta.Add($"{r.ReportedBy}@{r.MachineName}");
        meta.Add(ToLocal(r.CreatedAt).ToString("yyyy-MM-dd HH:mm"));
        _lblBugMeta.Text = string.Join("  •  ", meta);

        _txtDescription.Text = string.IsNullOrWhiteSpace(r.Description) ? "(No description)" : NormalizeNewLines(r.Description);
        _txtLog.Text = string.IsNullOrWhiteSpace(r.RuntimeLogTail) ? "(No runtime log attached)" : NormalizeNewLines(r.RuntimeLogTail);
        _txtDevNotes.Text = r.DevNotes;
        int statusIdx = Array.IndexOf(BugReportStatus.AllStatuses, r.Status);
        _cmbStatus.SelectedIndex = statusIdx >= 0 ? statusIdx : 0;
    }

    private async Task UpdatePreviewAsync()
    {
        if (_lstAttachments.SelectedItem is not BugReportAttachmentMeta meta)
        {
            _picPreview.Image = null;
            ShowPreviewMessage(_lstAttachments.Items.Count == 0 ? "No attachments" : "Select an attachment to preview");
            return;
        }

        if (!meta.IsImage)
        {
            _picPreview.Image = null;
            string kind = meta.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase) ? "Video" : "File";
            ShowPreviewMessage($"{kind}: {meta.FileName}\n{BugReportFormat.FormatSize(meta.FileSize)}\n\nDouble-click or use 'Open in Default App' to view.");
            return;
        }

        if (_imageCache.TryGetValue(meta.AttachmentId, out var cached))
        {
            ShowPreviewImage(cached);
            return;
        }

        int version = _selectionVersion;
        _picPreview.Image = null;
        ShowPreviewMessage("Loading preview...");
        try
        {
            byte[]? data = await GetDataAsync(meta.AttachmentId);
            if (version != _selectionVersion || !ReferenceEquals(_lstAttachments.SelectedItem, meta)) return;
            if (data == null)
            {
                ShowPreviewMessage("Attachment data not found.");
                return;
            }
            Image img;
            using (var ms = new MemoryStream(data))
            using (var decoded = Image.FromStream(ms))
            {
                img = new Bitmap(decoded);
            }
            _imageCache[meta.AttachmentId] = img;
            ShowPreviewImage(img);
        }
        catch (Exception ex)
        {
            ShowPreviewMessage($"Preview unavailable:\n{ex.Message}");
        }
    }

    private void ShowPreviewImage(Image img)
    {
        _picPreview.Image = img;
        _lblPreviewInfo.Visible = false;
    }

    private void ShowPreviewMessage(string text)
    {
        _lblPreviewInfo.Text = text;
        _lblPreviewInfo.Visible = true;
        _lblPreviewInfo.BringToFront();
    }

    // =========================================================================
    // 5. ATTACHMENT ACTIONS
    // =========================================================================
    private async Task<byte[]?> GetDataAsync(long attachmentId)
    {
        if (_dataCache.TryGetValue(attachmentId, out var data)) return data;
        data = await Task.Run(() => BugReportRepository.GetAttachmentData(attachmentId));
        if (data != null) _dataCache[attachmentId] = data;
        return data;
    }

    private async Task OpenSelectedAttachmentAsync()
    {
        if (_lstAttachments.SelectedItem is not BugReportAttachmentMeta meta) return;
        try
        {
            UseWaitCursor = true;
            byte[]? data = await GetDataAsync(meta.AttachmentId);
            if (data == null)
            {
                MessageBox.Show(this, "Attachment data not found.", "Open Attachment", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            string dir = Path.Combine(Path.GetTempPath(), "SKSB_BugAttachments");
            Directory.CreateDirectory(dir);
            string path = Path.Combine(dir, $"Bug_{meta.BugId}_{meta.AttachmentId}_{SanitizeFileName(meta.FileName)}");
            await File.WriteAllBytesAsync(path, data);
            Process.Start(new ProcessStartInfo(path) { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not open attachment:\n{ex.Message}", "Open Attachment", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    private async Task SaveSelectedAttachmentAsync()
    {
        if (_lstAttachments.SelectedItem is not BugReportAttachmentMeta meta) return;
        using var sfd = new SaveFileDialog
        {
            FileName = SanitizeFileName(meta.FileName),
            Filter = "All Files|*.*",
            OverwritePrompt = true
        };
        if (sfd.ShowDialog(this) != DialogResult.OK) return;
        try
        {
            UseWaitCursor = true;
            byte[]? data = await GetDataAsync(meta.AttachmentId);
            if (data == null)
            {
                MessageBox.Show(this, "Attachment data not found.", "Save Attachment", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            await File.WriteAllBytesAsync(sfd.FileName, data);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Could not save attachment:\n{ex.Message}", "Save Attachment", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            UseWaitCursor = false;
        }
    }

    // =========================================================================
    // 6. TRIAGE
    // =========================================================================
    private async Task SaveTriageAsync()
    {
        if (_current == null || _busy) return;
        long bugId = _current.BugId;
        string status = _cmbStatus.SelectedItem?.ToString() ?? BugReportStatus.Open;
        string notes = _txtDevNotes.Text;

        SetBusy(true, "Saving triage...");
        try
        {
            await Task.Run(() => BugReportRepository.UpdateBugTriage(bugId, status, notes));
        }
        catch (Exception ex)
        {
            SetBusy(false, null);
            MessageBox.Show(this, $"Failed to save triage:\n\n{ex.Message}", "Save Triage", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        SetBusy(false, null);
        _current = null; // force detail refresh on reselect
        await ReloadAsync(bugId);
    }

    private async Task DeleteCurrentAsync()
    {
        if (_current == null || _busy) return;
        long bugId = _current.BugId;
        var answer = MessageBox.Show(this,
            $"Permanently delete Bug Report #{bugId} — \"{_current.Title}\" and all of its attachments?\n\nThis cannot be undone.",
            "Delete Bug Report", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button2);
        if (answer != DialogResult.Yes) return;

        SetBusy(true, "Deleting...");
        try
        {
            await Task.Run(() => BugReportRepository.DeleteBugReport(bugId));
        }
        catch (Exception ex)
        {
            SetBusy(false, null);
            MessageBox.Show(this, $"Failed to delete bug report:\n\n{ex.Message}", "Delete Bug Report", MessageBoxButtons.OK, MessageBoxIcon.Error);
            return;
        }
        SetBusy(false, null);
        _current = null;
        await ReloadAsync(null);
    }

    // =========================================================================
    // 7. HELPERS & CLEANUP
    // =========================================================================
    private void SetBusy(bool busy, string? status)
    {
        _busy = busy;
        UseWaitCursor = busy;
        _cmbFilter.Enabled = !busy;
        _btnRefresh.Enabled = !busy;
        _btnSaveTriage.Enabled = !busy;
        _btnDelete.Enabled = !busy;
        if (busy && status != null) _lblSummary.Text = status;
    }

    private static DateTime ToLocal(DateTime dt) => dt.Kind == DateTimeKind.Utc ? dt.ToLocalTime() : dt;

    private static string NormalizeNewLines(string text) => text.Replace("\r\n", "\n").Replace("\n", Environment.NewLine);

    private static string SanitizeFileName(string name)
    {
        if (string.IsNullOrWhiteSpace(name)) return "attachment.bin";
        var invalid = Path.GetInvalidFileNameChars();
        var chars = name.Select(c => invalid.Contains(c) ? '_' : c).ToArray();
        string clean = new string(chars).Trim();
        return clean.Length == 0 ? "attachment.bin" : clean;
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _picPreview.Image = null;
        foreach (var img in _imageCache.Values)
        {
            try { img.Dispose(); } catch { }
        }
        _imageCache.Clear();
        _dataCache.Clear();
        _statusFont?.Dispose();
        _statusFont = null;
        base.OnFormClosed(e);
    }
}
#endif
