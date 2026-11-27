namespace SKSB_App.Core.RuntimeUI.Progress;

#if DEBUG
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows.Forms;
using Npgsql;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.BugReport;

/// <summary>
/// DEBUG-only diagnostics window displaying per-PC relative loading speeds,
/// operation breakdowns, and per-step bottleneck analytics (software vs. hardware).
/// </summary>
public sealed class OfficeManagerDialog : Form
{
    // =========================================================================
    // 1. PALETTE, MODELS & STATE
    // =========================================================================
    private static readonly Color Slate900 = Color.FromArgb(15, 23, 42);
    private static readonly Color Slate800 = Color.FromArgb(30, 41, 59);
    private static readonly Color Slate500 = Color.FromArgb(100, 116, 139);
    private static readonly Color Slate400 = Color.FromArgb(148, 163, 184);
    private static readonly Color Slate200 = Color.FromArgb(226, 232, 240);
    private static readonly Color Slate50 = Color.FromArgb(248, 250, 252);
    private static readonly Color Indigo600 = Color.FromArgb(79, 70, 229);
    private static readonly Color Emerald600 = Color.FromArgb(5, 150, 105);
    private static readonly Color Amber600 = Color.FromArgb(217, 119, 6);
    private static readonly Color Red600 = Color.FromArgb(220, 38, 38);

    private sealed class TimeframeOption
    {
        public string Label { get; init; } = string.Empty;
        public int? Days { get; init; }
        public override string ToString() => Label;
    }

    private sealed class VersionOption
    {
        public string Label { get; init; } = string.Empty;
        public string? Version { get; init; }
        public override string ToString() => Label;
    }

    private sealed class PcRankingRow
    {
        public string MachineName { get; init; } = string.Empty;
        public int OperationsMeasured { get; init; }
        public int TotalSamples { get; init; }
        public double RelativeDurationRatio { get; init; }
        public double AvgMs { get; init; }
        public DateTime? LastSeenAt { get; init; }
        public bool IsGlobal { get; init; }
    }

    private sealed class OperationTimingRow
    {
        public string OperationKey { get; init; } = string.Empty;
        public string MachineName { get; init; } = string.Empty;
        public double AvgMs { get; init; }
        public double MinMs { get; init; }
        public double MaxMs { get; init; }
        public double LastMs { get; init; }
        public int SampleCount { get; init; }
        public DateTime UpdatedAt { get; init; }
    }

    private sealed class StepPcRow
    {
        public string MachineName { get; init; } = string.Empty;
        public double AvgMs { get; init; }
        public double MinMs { get; init; }
        public double MaxMs { get; init; }
        public int SampleCount { get; init; }
        public DateTime LastSeenAt { get; init; }
    }

    private sealed class BottleneckRow
    {
        public string OperationKey { get; init; } = string.Empty;
        public string StepKey { get; init; } = string.Empty;
        public double ServerAvgMs { get; init; }
        public double MaxMs { get; init; }
        public double MinMs { get; init; }
        public double FastestPcAvgMs { get; init; }
        public double SlowestPcAvgMs { get; init; }
        public int PcCount { get; init; }
        public int TotalSamples { get; init; }
        public string Diagnosis { get; init; } = "Normal";
        public List<StepPcRow> PerPc { get; init; } = new();
    }

    private sealed class TelemetrySnapshot
    {
        public List<PcRankingRow> Rankings { get; init; } = new();
        public List<OperationTimingRow> Timings { get; init; } = new();
        public List<BottleneckRow> Bottlenecks { get; init; } = new();
        public List<string> AvailableVersions { get; init; } = new();
        public bool IsLegacyFallback { get; init; }
        public bool StepTableMissing { get; init; }
    }

    private readonly string _currentAppVersion;
    private List<PcRankingRow> _pcRows = new();
    private List<OperationTimingRow> _allTimings = new();
    private Dictionary<string, OperationTimingRow> _globalTimings = new(StringComparer.OrdinalIgnoreCase);
    private List<BottleneckRow> _bottlenecks = new();
    private bool _busy;
    private bool _suppressSelection;
    private bool _suppressFilterEvents;
    private bool _bottleneckSplitterInitialized;

    private Font? _titleFont;
    private Font? _sectionFont;
    private Font? _boldCellFont;
    private Font? _headerCellFont;

    private Button _btnRefresh = null!;
    private Button _btnDeletePc = null!;
    private ComboBox _cboTimeframe = null!;
    private ComboBox _cboVersion = null!;
    private Label _lblSummary = null!;
    private TabControl _tabs = null!;

    // Tab 1: PC Speed Ranking
    private SplitContainer _split = null!;
    private DataGridView _pcGrid = null!;
    private Label _lblOpsHeader = null!;
    private DataGridView _opGrid = null!;

    // Tab 2: Hanging Functions / Bottlenecks
    private SplitContainer _bottleneckSplit = null!;
    private Label _lblBottleneckHeader = null!;
    private DataGridView _bottleneckGrid = null!;
    private Label _lblStepPcHeader = null!;
    private DataGridView _stepPcGrid = null!;

    public OfficeManagerDialog()
    {
        _currentAppVersion = ResolveCurrentVersion();

        SuspendLayout();
        Text = "Office Manager — PC Speeds & Bottlenecks (Debug)";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = true;
        MaximizeBox = true;
        ShowInTaskbar = false;
        BackColor = Slate50;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(12);
        ClientSize = new Size(1120, 660);
        MinimumSize = new Size(860, 500);

        _titleFont = new Font("Segoe UI", 13f, FontStyle.Bold);
        _sectionFont = new Font("Segoe UI", 9.5f, FontStyle.Bold);
        _boldCellFont = new Font("Segoe UI", 9f, FontStyle.Bold);
        _headerCellFont = new Font("Segoe UI", 9f, FontStyle.Bold);

        InitializeLayout();

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    private static string ResolveCurrentVersion()
    {
        try
        {
            string? v = BugReportRepository.ResolveAppVersion();
            return string.IsNullOrWhiteSpace(v) ? "unknown" : v.Trim();
        }
        catch
        {
            return "unknown";
        }
    }

    // =========================================================================
    // 2. LAYOUT
    // =========================================================================
    private void InitializeLayout()
    {
        // --- Top bar ---
        var topBar = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            ColumnCount = 2,
            RowCount = 1,
            Padding = new Padding(0, 0, 0, 8),
            Margin = new Padding(0)
        };
        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        topBar.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        topBar.RowStyles.Add(new RowStyle(SizeType.AutoSize));

        var leftBar = new FlowLayoutPanel
        {
            AutoSize = true,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            Margin = new Padding(0)
        };

        var lblTitle = new Label
        {
            Text = "Office Manager",
            Font = _titleFont,
            ForeColor = Slate900,
            AutoSize = true,
            Margin = new Padding(0, 2, 14, 0)
        };

        var lblWindow = new Label
        {
            Text = "Window:",
            ForeColor = Slate800,
            AutoSize = true,
            Margin = new Padding(0, 6, 4, 0)
        };

        _cboTimeframe = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 152,
            Margin = new Padding(0, 2, 10, 0)
        };
        _cboTimeframe.Items.Add(new TimeframeOption { Label = "Last 7 days", Days = 7 });
        _cboTimeframe.Items.Add(new TimeframeOption { Label = "Last 30 days", Days = 30 });
        _cboTimeframe.Items.Add(new TimeframeOption { Label = "Last 3 months (90d)", Days = 90 });
        _cboTimeframe.Items.Add(new TimeframeOption { Label = "All time", Days = null });
        _cboTimeframe.SelectedIndex = 2; // Last 3 months (90d) default
        _cboTimeframe.SelectedIndexChanged += async (s, e) =>
        {
            if (!_suppressFilterEvents) await ReloadAsync(GetSelectedMachineName());
        };

        var lblVersion = new Label
        {
            Text = "Version:",
            ForeColor = Slate800,
            AutoSize = true,
            Margin = new Padding(0, 6, 4, 0)
        };

        _cboVersion = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Width = 150,
            Margin = new Padding(0, 2, 12, 0)
        };
        _cboVersion.Items.Add(new VersionOption { Label = "All versions", Version = null });
        _cboVersion.SelectedIndex = 0;
        _cboVersion.SelectedIndexChanged += async (s, e) =>
        {
            if (!_suppressFilterEvents) await ReloadAsync(GetSelectedMachineName());
        };

        _btnRefresh = CreateButton("Refresh", primary: false);
        _btnRefresh.Margin = new Padding(0, 1, 8, 0);
        _btnRefresh.Click += async (s, e) => await ReloadAsync(GetSelectedMachineName());

        _btnDeletePc = CreateButton("Delete Selected PC", primary: false);
        _btnDeletePc.ForeColor = Red600;
        _btnDeletePc.Enabled = false;
        _btnDeletePc.Margin = new Padding(0, 1, 12, 0);
        _btnDeletePc.Click += async (s, e) => await DeleteSelectedPcAsync();

        leftBar.Controls.AddRange(new Control[]
        {
            lblTitle, lblWindow, _cboTimeframe, lblVersion, _cboVersion, _btnRefresh, _btnDeletePc
        });

        _lblSummary = new Label
        {
            Text = "Loading...",
            ForeColor = Slate500,
            AutoSize = true,
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleRight,
            Margin = new Padding(8, 4, 0, 0)
        };

        topBar.Controls.Add(leftBar, 0, 0);
        topBar.Controls.Add(_lblSummary, 1, 0);

        // --- TabControl ---
        _tabs = new TabControl
        {
            Dock = DockStyle.Fill,
            Margin = new Padding(0)
        };
        _tabs.SelectedIndexChanged += (s, e) =>
        {
            if (_tabs.SelectedIndex == 1 && !_bottleneckSplitterInitialized && _bottleneckSplit.ClientSize.Height > 100)
            {
                try
                {
                    _bottleneckSplit.SplitterDistance = (int)(_bottleneckSplit.ClientSize.Height * 0.55);
                    _bottleneckSplitterInitialized = true;
                }
                catch { }
            }
        };

        // --- Tab 1: PC Speed Ranking ---
        var tabRanking = new TabPage("PC Speed Ranking")
        {
            BackColor = Slate50,
            Padding = new Padding(6)
        };

        _split = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Vertical,
            BackColor = Slate200,
            SplitterWidth = 5
        };
        _split.Panel1.BackColor = Slate50;
        _split.Panel1.Padding = new Padding(0, 0, 4, 0);
        _split.Panel2.BackColor = Slate50;
        _split.Panel2.Padding = new Padding(4, 0, 0, 0);

        var leftPanel = BuildSectionContainer("PC Speed Ranking (vs. Server Average)", out _);
        _pcGrid = BuildPcGrid();
        leftPanel.Controls.Add(_pcGrid, 0, 1);
        _split.Panel1.Controls.Add(leftPanel);

        var rightPanel = BuildSectionContainer("Operations", out _lblOpsHeader);
        _opGrid = BuildOpGrid();
        rightPanel.Controls.Add(_opGrid, 0, 1);
        _split.Panel2.Controls.Add(rightPanel);

        tabRanking.Controls.Add(_split);

        // --- Tab 2: Hanging Functions / Bottlenecks ---
        var tabBottlenecks = new TabPage("Hanging Functions / Bottlenecks")
        {
            BackColor = Slate50,
            Padding = new Padding(6)
        };

        _bottleneckSplit = new SplitContainer
        {
            Dock = DockStyle.Fill,
            Orientation = Orientation.Horizontal,
            BackColor = Slate200,
            SplitterWidth = 5
        };
        _bottleneckSplit.Panel1.BackColor = Slate50;
        _bottleneckSplit.Panel1.Padding = new Padding(0, 0, 0, 4);
        _bottleneckSplit.Panel2.BackColor = Slate50;
        _bottleneckSplit.Panel2.Padding = new Padding(0, 4, 0, 0);

        var topBottleneckPanel = BuildSectionContainer("Slowest Functions / Steps (Software vs. Hardware)", out _lblBottleneckHeader);
        _bottleneckGrid = BuildBottleneckGrid();
        topBottleneckPanel.Controls.Add(_bottleneckGrid, 0, 1);
        _bottleneckSplit.Panel1.Controls.Add(topBottleneckPanel);

        var bottomStepPcPanel = BuildSectionContainer("Per-PC Breakdown", out _lblStepPcHeader);
        _stepPcGrid = BuildStepPcGrid();
        bottomStepPcPanel.Controls.Add(_stepPcGrid, 0, 1);
        _bottleneckSplit.Panel2.Controls.Add(bottomStepPcPanel);

        tabBottlenecks.Controls.Add(_bottleneckSplit);

        _tabs.TabPages.Add(tabRanking);
        _tabs.TabPages.Add(tabBottlenecks);

        Controls.Add(_tabs);
        Controls.Add(topBar);
    }

    private TableLayoutPanel BuildSectionContainer(string headerText, out Label headerLabel)
    {
        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 2,
            Margin = new Padding(0)
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        panel.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        panel.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        headerLabel = new Label
        {
            Text = headerText,
            Font = _sectionFont,
            ForeColor = Slate800,
            AutoSize = true,
            Margin = new Padding(2, 0, 0, 6)
        };
        panel.Controls.Add(headerLabel, 0, 0);
        return panel;
    }

    private DataGridView CreateBaseGrid()
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
            BorderStyle = BorderStyle.FixedSingle,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            GridColor = Slate200,
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize,
            AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells,
            Margin = new Padding(0)
        };
        grid.ColumnHeadersDefaultCellStyle.BackColor = Slate50;
        grid.ColumnHeadersDefaultCellStyle.ForeColor = Slate800;
        grid.ColumnHeadersDefaultCellStyle.Font = _headerCellFont;
        grid.DefaultCellStyle.ForeColor = Slate900;
        grid.DefaultCellStyle.SelectionBackColor = Color.FromArgb(224, 231, 255);
        grid.DefaultCellStyle.SelectionForeColor = Slate900;
        EnableDoubleBuffering(grid);
        return grid;
    }

    private DataGridView BuildPcGrid()
    {
        var grid = CreateBaseGrid();

        grid.Columns.Add(CreateCol("colRank", "Rank", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleCenter));
        grid.Columns.Add(CreateCol("colPcName", "PC Name", 40, DataGridViewAutoSizeColumnMode.Fill));
        grid.Columns.Add(CreateCol("colSpeed", "Speed", 32, DataGridViewAutoSizeColumnMode.AllCells));
        grid.Columns.Add(CreateCol("colAvgLoad", "Avg Load", 14, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colOps", "Ops", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colSamples", "Samples", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colLastSeen", "Last Seen", 20, DataGridViewAutoSizeColumnMode.AllCells));

        grid.CellFormatting += OnPcGridCellFormatting;
        grid.SelectionChanged += (s, e) => OnPcSelectionChanged();
        return grid;
    }

    private DataGridView BuildOpGrid()
    {
        var grid = CreateBaseGrid();

        grid.Columns.Add(CreateCol("colOperation", "Operation", 44, DataGridViewAutoSizeColumnMode.Fill));
        grid.Columns.Add(CreateCol("colPcAvg", "PC Avg", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colServerAvg", "Server Avg", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colVsAvg", "vs. Avg", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colMin", "Min", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colMax", "Max", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colLast", "Last", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colOpSamples", "Samples", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colUpdated", "Updated", 18, DataGridViewAutoSizeColumnMode.AllCells));

        grid.CellFormatting += OnOpGridCellFormatting;
        return grid;
    }

    private DataGridView BuildBottleneckGrid()
    {
        var grid = CreateBaseGrid();

        grid.Columns.Add(CreateCol("colStep", "Function / Step", 36, DataGridViewAutoSizeColumnMode.Fill));
        grid.Columns.Add(CreateCol("colStepOp", "Operation", 26, DataGridViewAutoSizeColumnMode.Fill));
        grid.Columns.Add(CreateCol("colStepServerAvg", "Server Avg", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepMax", "Max (Hang)", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepMin", "Min", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepSpread", "PC Spread", 20, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcs", "PCs", 8, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepSamples", "Samples", 10, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepDiagnosis", "Diagnosis", 22, DataGridViewAutoSizeColumnMode.AllCells));

        grid.CellFormatting += OnBottleneckGridCellFormatting;
        grid.SelectionChanged += (s, e) => OnBottleneckSelectionChanged();
        return grid;
    }

    private DataGridView BuildStepPcGrid()
    {
        var grid = CreateBaseGrid();

        grid.Columns.Add(CreateCol("colStepPcName", "PC Name", 36, DataGridViewAutoSizeColumnMode.Fill));
        grid.Columns.Add(CreateCol("colStepPcAvg", "PC Avg", 14, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcVsAvg", "vs. Server Avg", 14, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcMin", "Min", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcMax", "Max (Hang)", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcSamples", "Samples", 12, DataGridViewAutoSizeColumnMode.AllCells, DataGridViewContentAlignment.MiddleRight));
        grid.Columns.Add(CreateCol("colStepPcLastSeen", "Last Seen", 18, DataGridViewAutoSizeColumnMode.AllCells));

        grid.CellFormatting += OnStepPcGridCellFormatting;
        return grid;
    }

    private static DataGridViewTextBoxColumn CreateCol(
        string name,
        string header,
        int fillWeight,
        DataGridViewAutoSizeColumnMode mode,
        DataGridViewContentAlignment alignment = DataGridViewContentAlignment.MiddleLeft)
    {
        var col = new DataGridViewTextBoxColumn
        {
            Name = name,
            HeaderText = header,
            FillWeight = fillWeight,
            AutoSizeMode = mode,
            SortMode = DataGridViewColumnSortMode.NotSortable
        };
        col.DefaultCellStyle.Alignment = alignment;
        return col;
    }

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
            Cursor = Cursors.Hand
        };
        if (primary) btn.FlatAppearance.BorderSize = 0;
        else btn.FlatAppearance.BorderColor = Slate200;
        return btn;
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        _ = ReloadAsync(null);
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        try { _split.SplitterDistance = (int)(_split.ClientSize.Width * 0.44); } catch { }
        if (_bottleneckSplit.ClientSize.Height > 100)
        {
            try
            {
                _bottleneckSplit.SplitterDistance = (int)(_bottleneckSplit.ClientSize.Height * 0.55);
                _bottleneckSplitterInitialized = true;
            }
            catch { }
        }
    }

    // =========================================================================
    // 3. DATA LOADING & SERVER QUERIES
    // =========================================================================
    private int? GetSelectedWindowDays() => (_cboTimeframe.SelectedItem as TimeframeOption)?.Days;

    private string? GetSelectedVersionFilter() => (_cboVersion.SelectedItem as VersionOption)?.Version;

    private async Task ReloadAsync(string? reselectMachine)
    {
        if (_busy || IsDisposed) return;

        if (ServerEnvironment.IsOffline)
        {
            ClearGrids();
            _lblSummary.ForeColor = Slate500;
            _lblSummary.Text = "Offline mode — server telemetry unavailable";
            UpdateDeleteButtonState();
            return;
        }

        int? windowDays = GetSelectedWindowDays();
        string? versionFilter = GetSelectedVersionFilter();

        SetBusy(true, "Loading telemetry...");
        try
        {
            TelemetrySnapshot snapshot = await Task.Run(() => FetchTelemetryFromServer(windowDays, versionFilter));
            if (IsDisposed) return;

            SyncVersionDropdown(snapshot.AvailableVersions, versionFilter);

            _allTimings = snapshot.Timings;
            _globalTimings = _allTimings
                .Where(t => string.Equals(t.MachineName, LoadingTimeEstimator.GlobalMachineKey, StringComparison.OrdinalIgnoreCase))
                .ToDictionary(t => t.OperationKey, t => t, StringComparer.OrdinalIgnoreCase);

            _pcRows = BuildPcDisplayRows(snapshot.Rankings, _globalTimings);
            PopulatePcGrid(reselectMachine);

            _bottlenecks = snapshot.Bottlenecks;
            PopulateBottleneckGrid(snapshot.StepTableMissing);

            int pcCount = _pcRows.Count(r => !r.IsGlobal);
            int opCount = _globalTimings.Count > 0
                ? _globalTimings.Count
                : _allTimings.Select(t => t.OperationKey).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            int sampleCount = _pcRows.Where(r => !r.IsGlobal).Sum(r => r.TotalSamples);

            _lblSummary.ForeColor = Slate500;
            string legacyNote = snapshot.IsLegacyFallback ? " · (legacy all-time data)" : string.Empty;
            _lblSummary.Text = $"{pcCount} PC(s) · {opCount} operation(s) · {sampleCount} sample(s){legacyNote}";
        }
        catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
        {
            if (IsDisposed) return;
            ClearGrids();
            _lblSummary.ForeColor = Amber600;
            _lblSummary.Text = "Telemetry table not created yet — run a loading operation first";
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            AppLogger.Warn($"[OfficeManagerDialog] Telemetry load failed: {ex.Message}");
            ClearGrids();
            _lblSummary.ForeColor = Red600;
            _lblSummary.Text = $"Unable to load telemetry: {ex.Message}";
        }
        finally
        {
            if (!IsDisposed) SetBusy(false, null);
        }
    }

    private void SyncVersionDropdown(List<string> versions, string? selectedVersion)
    {
        _suppressFilterEvents = true;
        try
        {
            _cboVersion.Items.Clear();
            _cboVersion.Items.Add(new VersionOption { Label = "All versions", Version = null });

            int targetIdx = 0;
            foreach (string v in versions)
            {
                bool isCurrent = string.Equals(v, _currentAppVersion, StringComparison.OrdinalIgnoreCase);
                string label = isCurrent ? $"{v} (Current)" : v;
                int idx = _cboVersion.Items.Add(new VersionOption { Label = label, Version = v });
                if (selectedVersion != null && string.Equals(v, selectedVersion, StringComparison.OrdinalIgnoreCase))
                    targetIdx = idx;
            }
            _cboVersion.SelectedIndex = targetIdx;
        }
        finally
        {
            _suppressFilterEvents = false;
        }
    }

    private static NpgsqlConnection OpenTelemetryConnection()
    {
        string cs = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString)
        {
            Timeout = 3,
            CommandTimeout = 5
        }.ConnectionString;
        var conn = new NpgsqlConnection(cs);
        try { conn.Open(); }
        catch { conn.Dispose(); throw; }
        return conn;
    }

    private static TelemetrySnapshot FetchTelemetryFromServer(int? windowDays, string? versionFilter)
    {
        using var conn = OpenTelemetryConnection();

        // 1. Distinct app versions from samples
        var availableVersions = new List<string>();
        bool samplesTableMissing = false;
        try
        {
            const string versionSql = @"
SELECT app_version
FROM app.operation_timing_samples
GROUP BY app_version
ORDER BY MAX(recorded_at) DESC, app_version DESC;";
            using var cmd = new NpgsqlCommand(versionSql, conn);
            using var r = cmd.ExecuteReader();
            while (r.Read())
            {
                if (!r.IsDBNull(0))
                {
                    string v = r.GetString(0);
                    if (!string.IsNullOrWhiteSpace(v)) availableVersions.Add(v);
                }
            }
        }
        catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
        {
            samplesTableMissing = true;
        }

        // 2. Operation timings & PC rankings from app.operation_timing_samples
        var rankings = new List<PcRankingRow>();
        var timings = new List<OperationTimingRow>();
        bool isLegacyFallback = false;

        if (!samplesTableMissing)
        {
            string where = BuildFilterWhereClause(windowDays, versionFilter);

            string pcOpSql = $@"
SELECT operation_key,
       machine_name,
       AVG(duration_ms) AS avg_ms,
       MIN(duration_ms) AS min_ms,
       MAX(duration_ms) AS max_ms,
       (ARRAY_AGG(duration_ms ORDER BY recorded_at DESC))[1] AS last_ms,
       COUNT(*)::int AS sample_count,
       MAX(recorded_at) AS updated_at
FROM app.operation_timing_samples
{where}
GROUP BY operation_key, machine_name
ORDER BY operation_key ASC, machine_name ASC;";

            using (var cmd = new NpgsqlCommand(pcOpSql, conn))
            {
                AddFilterParams(cmd, windowDays, versionFilter);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    timings.Add(new OperationTimingRow
                    {
                        OperationKey = r.GetString(0),
                        MachineName = r.GetString(1),
                        AvgMs = r.GetDouble(2),
                        MinMs = r.GetDouble(3),
                        MaxMs = r.GetDouble(4),
                        LastMs = r.GetDouble(5),
                        SampleCount = r.GetInt32(6),
                        UpdatedAt = r.GetDateTime(7)
                    });
                }
            }

            if (timings.Count > 0)
            {
                string globalOpSql = $@"
SELECT operation_key,
       AVG(duration_ms) AS avg_ms,
       MIN(duration_ms) AS min_ms,
       MAX(duration_ms) AS max_ms,
       (ARRAY_AGG(duration_ms ORDER BY recorded_at DESC))[1] AS last_ms,
       COUNT(*)::int AS sample_count,
       MAX(recorded_at) AS updated_at
FROM app.operation_timing_samples
{where}
GROUP BY operation_key
ORDER BY operation_key ASC;";

                var globalByOp = new Dictionary<string, OperationTimingRow>(StringComparer.OrdinalIgnoreCase);
                using (var cmd = new NpgsqlCommand(globalOpSql, conn))
                {
                    AddFilterParams(cmd, windowDays, versionFilter);
                    using var r = cmd.ExecuteReader();
                    while (r.Read())
                    {
                        var row = new OperationTimingRow
                        {
                            OperationKey = r.GetString(0),
                            MachineName = LoadingTimeEstimator.GlobalMachineKey,
                            AvgMs = r.GetDouble(1),
                            MinMs = r.GetDouble(2),
                            MaxMs = r.GetDouble(3),
                            LastMs = r.GetDouble(4),
                            SampleCount = r.GetInt32(5),
                            UpdatedAt = r.GetDateTime(6)
                        };
                        globalByOp[row.OperationKey] = row;
                        timings.Add(row);
                    }
                }

                // Compute PC speed rankings from per-PC vs server-wide averages
                foreach (var pcGroup in timings
                    .Where(t => !string.Equals(t.MachineName, LoadingTimeEstimator.GlobalMachineKey, StringComparison.OrdinalIgnoreCase))
                    .GroupBy(t => t.MachineName, StringComparer.OrdinalIgnoreCase))
                {
                    var ratios = new List<double>();
                    foreach (var op in pcGroup)
                    {
                        if (globalByOp.TryGetValue(op.OperationKey, out var gOp) && gOp.AvgMs > 0.0)
                            ratios.Add(op.AvgMs / gOp.AvgMs);
                    }

                    rankings.Add(new PcRankingRow
                    {
                        MachineName = pcGroup.Key,
                        OperationsMeasured = pcGroup.Count(),
                        TotalSamples = pcGroup.Sum(x => x.SampleCount),
                        RelativeDurationRatio = ratios.Count > 0 ? Math.Round(ratios.Average(), 2) : 1.0,
                        AvgMs = Math.Round(pcGroup.Average(x => x.AvgMs), 0),
                        LastSeenAt = pcGroup.Max(x => x.UpdatedAt),
                        IsGlobal = false
                    });
                }

                rankings.Sort((a, b) =>
                {
                    int cmp = a.RelativeDurationRatio.CompareTo(b.RelativeDurationRatio);
                    return cmp != 0 ? cmp : string.Compare(a.MachineName, b.MachineName, StringComparison.OrdinalIgnoreCase);
                });
            }
        }

        // 3. Fallback to legacy app.pc_speed_ranking / app.operation_timings if samples empty/missing and no version filter
        if (timings.Count == 0 && versionFilter == null)
        {
            isLegacyFallback = true;
            const string rankingSql = @"
SELECT machine_name, operations_measured, total_samples, relative_duration_ratio, avg_ms, last_seen_at
FROM app.pc_speed_ranking
ORDER BY relative_duration_ratio ASC NULLS LAST, machine_name ASC;";

            using (var cmd = new NpgsqlCommand(rankingSql, conn))
            using (var r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    rankings.Add(new PcRankingRow
                    {
                        MachineName = r.GetString(0),
                        OperationsMeasured = Convert.ToInt32(r.GetValue(1)),
                        TotalSamples = Convert.ToInt32(r.GetValue(2)),
                        RelativeDurationRatio = r.IsDBNull(3) ? 1.0 : Convert.ToDouble(r.GetValue(3)),
                        AvgMs = r.IsDBNull(4) ? 0.0 : Convert.ToDouble(r.GetValue(4)),
                        LastSeenAt = r.IsDBNull(5) ? null : r.GetDateTime(5),
                        IsGlobal = false
                    });
                }
            }

            const string timingsSql = @"
SELECT operation_key, machine_name, avg_ms, min_ms, max_ms, last_ms, sample_count, updated_at
FROM app.operation_timings
ORDER BY operation_key ASC, machine_name ASC;";

            using (var cmd = new NpgsqlCommand(timingsSql, conn))
            using (var r = cmd.ExecuteReader())
            {
                while (r.Read())
                {
                    timings.Add(new OperationTimingRow
                    {
                        OperationKey = r.GetString(0),
                        MachineName = r.GetString(1),
                        AvgMs = r.GetDouble(2),
                        MinMs = r.GetDouble(3),
                        MaxMs = r.GetDouble(4),
                        LastMs = r.GetDouble(5),
                        SampleCount = r.GetInt32(6),
                        UpdatedAt = r.GetDateTime(7)
                    });
                }
            }
        }

        // 4. Per-step bottleneck samples from app.operation_step_samples
        var bottlenecks = new List<BottleneckRow>();
        bool stepTableMissing = false;
        try
        {
            string where = BuildFilterWhereClause(windowDays, versionFilter);
            string stepSql = $@"
SELECT operation_key,
       step_key,
       machine_name,
       AVG(duration_ms) AS avg_ms,
       MIN(duration_ms) AS min_ms,
       MAX(duration_ms) AS max_ms,
       COUNT(*)::int AS cnt,
       MAX(recorded_at) AS last_seen
FROM app.operation_step_samples
{where}
GROUP BY operation_key, step_key, machine_name;";

            var rawSteps = new Dictionary<(string Op, string Step), List<StepPcRow>>();
            using (var cmd = new NpgsqlCommand(stepSql, conn))
            {
                AddFilterParams(cmd, windowDays, versionFilter);
                using var r = cmd.ExecuteReader();
                while (r.Read())
                {
                    string opKey = r.GetString(0);
                    string stepKey = r.GetString(1);
                    var pcRow = new StepPcRow
                    {
                        MachineName = r.GetString(2),
                        AvgMs = r.GetDouble(3),
                        MinMs = r.GetDouble(4),
                        MaxMs = r.GetDouble(5),
                        SampleCount = r.GetInt32(6),
                        LastSeenAt = r.GetDateTime(7)
                    };

                    var key = (opKey, stepKey);
                    if (!rawSteps.TryGetValue(key, out var list))
                    {
                        list = new List<StepPcRow>();
                        rawSteps[key] = list;
                    }
                    list.Add(pcRow);
                }
            }

            foreach (var kv in rawSteps)
            {
                var pcs = kv.Value.OrderByDescending(p => p.AvgMs).ThenBy(p => p.MachineName, StringComparer.OrdinalIgnoreCase).ToList();
                int totalSamples = pcs.Sum(p => p.SampleCount);
                if (totalSamples <= 0) continue;

                double weightedSum = pcs.Sum(p => p.AvgMs * p.SampleCount);
                double serverAvgMs = weightedSum / totalSamples;
                double maxMs = pcs.Max(p => p.MaxMs);
                double minMs = pcs.Min(p => p.MinMs);
                double fastestPcAvgMs = pcs.Min(p => p.AvgMs);
                double slowestPcAvgMs = pcs.Max(p => p.AvgMs);
                int pcCount = pcs.Count;

                string diagnosis = ClassifyBottleneck(serverAvgMs, maxMs, fastestPcAvgMs, slowestPcAvgMs, pcCount);

                bottlenecks.Add(new BottleneckRow
                {
                    OperationKey = kv.Key.Op,
                    StepKey = kv.Key.Step,
                    ServerAvgMs = serverAvgMs,
                    MaxMs = maxMs,
                    MinMs = minMs,
                    FastestPcAvgMs = fastestPcAvgMs,
                    SlowestPcAvgMs = slowestPcAvgMs,
                    PcCount = pcCount,
                    TotalSamples = totalSamples,
                    Diagnosis = diagnosis,
                    PerPc = pcs
                });
            }

            bottlenecks.Sort((a, b) =>
            {
                int cmp = b.ServerAvgMs.CompareTo(a.ServerAvgMs);
                return cmp != 0 ? cmp : b.MaxMs.CompareTo(a.MaxMs);
            });
        }
        catch (PostgresException pgEx) when (pgEx.SqlState == "42P01")
        {
            stepTableMissing = true;
        }

        return new TelemetrySnapshot
        {
            Rankings = rankings,
            Timings = timings,
            Bottlenecks = bottlenecks,
            AvailableVersions = availableVersions,
            IsLegacyFallback = isLegacyFallback,
            StepTableMissing = stepTableMissing
        };
    }

    private static string BuildFilterWhereClause(int? windowDays, string? versionFilter)
    {
        var clauses = new List<string>();
        if (windowDays.HasValue)
            clauses.Add("recorded_at >= CURRENT_TIMESTAMP - (@WindowDays * INTERVAL '1 day')");
        if (!string.IsNullOrWhiteSpace(versionFilter))
            clauses.Add("app_version = @Version");
        return clauses.Count == 0 ? string.Empty : "WHERE " + string.Join(" AND ", clauses);
    }

    private static void AddFilterParams(NpgsqlCommand cmd, int? windowDays, string? versionFilter)
    {
        if (windowDays.HasValue)
            cmd.Parameters.AddWithValue("WindowDays", windowDays.Value);
        if (!string.IsNullOrWhiteSpace(versionFilter))
            cmd.Parameters.AddWithValue("Version", versionFilter);
    }

    private static string ClassifyBottleneck(
        double serverAvgMs,
        double maxMs,
        double fastestPcAvgMs,
        double slowestPcAvgMs,
        int pcCount)
    {
        double spreadRatio = fastestPcAvgMs > 0.0 ? slowestPcAvgMs / fastestPcAvgMs : 1.0;

        if (serverAvgMs >= 1500.0 && pcCount >= 2 && (fastestPcAvgMs >= 1200.0 || spreadRatio <= 1.8))
            return "Software Bottleneck";

        if (pcCount == 1 && serverAvgMs >= 1500.0)
            return "Slow (1 PC)";

        if (pcCount >= 2 && spreadRatio > 1.8 && fastestPcAvgMs < 1200.0)
            return "Hardware / PC-Specific";

        if (maxMs >= serverAvgMs * 3.0 && maxMs >= 2000.0)
            return "Intermittent Hang";

        return "Normal";
    }

    private static List<PcRankingRow> BuildPcDisplayRows(
        List<PcRankingRow> rankings,
        Dictionary<string, OperationTimingRow> globalTimings)
    {
        var list = new List<PcRankingRow>(rankings);

        if (globalTimings.Count > 0)
        {
            var values = globalTimings.Values;
            list.Add(new PcRankingRow
            {
                MachineName = LoadingTimeEstimator.GlobalMachineKey,
                OperationsMeasured = globalTimings.Count,
                TotalSamples = values.Sum(v => v.SampleCount),
                RelativeDurationRatio = 1.0,
                AvgMs = values.Average(v => v.AvgMs),
                LastSeenAt = values.Max(v => v.UpdatedAt),
                IsGlobal = true
            });
        }

        return list;
    }

    // =========================================================================
    // 4. GRID BINDING & FORMATTING
    // =========================================================================
    private void ClearGrids()
    {
        _pcRows.Clear();
        _allTimings.Clear();
        _globalTimings.Clear();
        _bottlenecks.Clear();
        _pcGrid.Rows.Clear();
        _opGrid.Rows.Clear();
        _bottleneckGrid.Rows.Clear();
        _stepPcGrid.Rows.Clear();
        _lblOpsHeader.Text = "Operations";
        _lblBottleneckHeader.Text = "Slowest Functions / Steps (Software vs. Hardware)";
        _lblStepPcHeader.Text = "Per-PC Breakdown";
    }

    private void PopulatePcGrid(string? reselectMachine)
    {
        _suppressSelection = true;
        try
        {
            _pcGrid.SuspendLayout();
            _pcGrid.Rows.Clear();

            int rank = 0;
            int reselectIndex = -1;

            for (int i = 0; i < _pcRows.Count; i++)
            {
                PcRankingRow row = _pcRows[i];
                string rankText;
                string pcDisplay;
                string speedText;

                if (row.IsGlobal)
                {
                    rankText = "—";
                    pcDisplay = $"Server Average ({LoadingTimeEstimator.GlobalMachineKey})";
                    speedText = "1.00× (Avg)";
                }
                else
                {
                    rank++;
                    rankText = $"#{rank}";
                    bool isThisPc = string.Equals(row.MachineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase);
                    pcDisplay = isThisPc ? $"{row.MachineName} (This PC)" : row.MachineName;
                    speedText = FormatSpeedRatio(row.RelativeDurationRatio);
                }

                string lastSeen = row.LastSeenAt.HasValue
                    ? row.LastSeenAt.Value.ToLocalTime().ToString("dd MMM HH:mm")
                    : "—";

                int idx = _pcGrid.Rows.Add(
                    rankText,
                    pcDisplay,
                    speedText,
                    FormatSeconds(row.AvgMs),
                    row.OperationsMeasured,
                    row.TotalSamples,
                    lastSeen);

                _pcGrid.Rows[idx].Tag = row;

                if (reselectMachine != null &&
                    string.Equals(row.MachineName, reselectMachine, StringComparison.OrdinalIgnoreCase))
                {
                    reselectIndex = idx;
                }
            }

            _pcGrid.ResumeLayout();

            if (_pcGrid.Rows.Count > 0)
            {
                int target = reselectIndex >= 0 ? reselectIndex : 0;
                _pcGrid.ClearSelection();
                _pcGrid.Rows[target].Selected = true;
                _pcGrid.CurrentCell = _pcGrid.Rows[target].Cells[0];
            }
        }
        finally
        {
            _suppressSelection = false;
        }

        OnPcSelectionChanged();
    }

    private void OnPcSelectionChanged()
    {
        if (_suppressSelection) return;

        PcRankingRow? selected = GetSelectedPcRow();
        UpdateDeleteButtonState();

        _opGrid.SuspendLayout();
        _opGrid.Rows.Clear();

        if (selected == null)
        {
            _lblOpsHeader.Text = "Operations";
            _opGrid.ResumeLayout();
            return;
        }

        string displayTitle = selected.IsGlobal
            ? $"Server Average ({LoadingTimeEstimator.GlobalMachineKey})"
            : string.Equals(selected.MachineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase)
                ? $"{selected.MachineName} (This PC)"
                : selected.MachineName;

        _lblOpsHeader.Text = $"Operations — {displayTitle}";

        var ops = _allTimings
            .Where(t => string.Equals(t.MachineName, selected.MachineName, StringComparison.OrdinalIgnoreCase))
            .OrderByDescending(t => t.AvgMs)
            .ThenBy(t => t.OperationKey, StringComparer.OrdinalIgnoreCase);

        foreach (OperationTimingRow op in ops)
        {
            _globalTimings.TryGetValue(op.OperationKey, out OperationTimingRow? globalOp);

            string serverAvgText = globalOp != null ? FormatSeconds(globalOp.AvgMs) : "—";
            string vsAvgText;
            double? pctDiff = null;

            if (selected.IsGlobal || globalOp == null || globalOp.AvgMs <= 0.0)
            {
                vsAvgText = "—";
            }
            else
            {
                pctDiff = (op.AvgMs - globalOp.AvgMs) / globalOp.AvgMs * 100.0;
                int rounded = (int)Math.Round(pctDiff.Value);
                vsAvgText = rounded > 0 ? $"+{rounded}%" : $"{rounded}%";
            }

            int idx = _opGrid.Rows.Add(
                op.OperationKey,
                FormatSeconds(op.AvgMs),
                serverAvgText,
                vsAvgText,
                FormatSeconds(op.MinMs),
                FormatSeconds(op.MaxMs),
                FormatSeconds(op.LastMs),
                op.SampleCount,
                op.UpdatedAt.ToLocalTime().ToString("dd MMM HH:mm"));

            _opGrid.Rows[idx].Tag = pctDiff;
        }

        _opGrid.ResumeLayout();
    }

    private void PopulateBottleneckGrid(bool stepTableMissing)
    {
        _suppressSelection = true;
        try
        {
            _bottleneckGrid.SuspendLayout();
            _bottleneckGrid.Rows.Clear();

            if (stepTableMissing)
            {
                _lblBottleneckHeader.Text = "Slowest Functions / Steps — step telemetry table not created yet (run a Release build first)";
            }
            else if (_bottlenecks.Count == 0)
            {
                _lblBottleneckHeader.Text = "Slowest Functions / Steps — no step samples in selected filter";
            }
            else
            {
                _lblBottleneckHeader.Text = $"Slowest Functions / Steps ({_bottlenecks.Count} measured)";
            }

            foreach (BottleneckRow b in _bottlenecks)
            {
                string spreadText;
                if (b.PcCount >= 2)
                {
                    double ratio = b.FastestPcAvgMs > 0.0 ? b.SlowestPcAvgMs / b.FastestPcAvgMs : 1.0;
                    spreadText = $"{FormatSeconds(b.FastestPcAvgMs)} – {FormatSeconds(b.SlowestPcAvgMs)} ({ratio:0.0}×)";
                }
                else
                {
                    spreadText = "—";
                }

                int idx = _bottleneckGrid.Rows.Add(
                    b.StepKey,
                    b.OperationKey,
                    FormatSeconds(b.ServerAvgMs),
                    FormatSeconds(b.MaxMs),
                    FormatSeconds(b.MinMs),
                    spreadText,
                    b.PcCount,
                    b.TotalSamples,
                    b.Diagnosis);

                _bottleneckGrid.Rows[idx].Tag = b;
            }

            _bottleneckGrid.ResumeLayout();

            if (_bottleneckGrid.Rows.Count > 0)
            {
                _bottleneckGrid.ClearSelection();
                _bottleneckGrid.Rows[0].Selected = true;
                _bottleneckGrid.CurrentCell = _bottleneckGrid.Rows[0].Cells[0];
            }
        }
        finally
        {
            _suppressSelection = false;
        }

        OnBottleneckSelectionChanged();
    }

    private void OnBottleneckSelectionChanged()
    {
        if (_suppressSelection) return;

        BottleneckRow? selected = _bottleneckGrid.SelectedRows.Count > 0
            ? _bottleneckGrid.SelectedRows[0].Tag as BottleneckRow
            : null;

        _stepPcGrid.SuspendLayout();
        _stepPcGrid.Rows.Clear();

        if (selected == null)
        {
            _lblStepPcHeader.Text = "Per-PC Breakdown";
            _stepPcGrid.ResumeLayout();
            return;
        }

        _lblStepPcHeader.Text = $"Per-PC Breakdown — {selected.StepKey} ({selected.OperationKey})";

        foreach (StepPcRow pc in selected.PerPc)
        {
            bool isThisPc = string.Equals(pc.MachineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase);
            string pcDisplay = isThisPc ? $"{pc.MachineName} (This PC)" : pc.MachineName;

            string vsAvgText;
            double? ratio = null;
            if (selected.ServerAvgMs > 0.0)
            {
                ratio = pc.AvgMs / selected.ServerAvgMs;
                int pct = (int)Math.Round((ratio.Value - 1.0) * 100.0);
                vsAvgText = pct > 0 ? $"+{pct}%" : $"{pct}%";
            }
            else
            {
                vsAvgText = "—";
            }

            int idx = _stepPcGrid.Rows.Add(
                pcDisplay,
                FormatSeconds(pc.AvgMs),
                vsAvgText,
                FormatSeconds(pc.MinMs),
                FormatSeconds(pc.MaxMs),
                pc.SampleCount,
                pc.LastSeenAt.ToLocalTime().ToString("dd MMM HH:mm"));

            _stepPcGrid.Rows[idx].Tag = (pc.MachineName, ratio);
        }

        _stepPcGrid.ResumeLayout();
    }

    private void OnPcGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.CellStyle == null || e.RowIndex < 0 || e.RowIndex >= _pcGrid.Rows.Count) return;
        if (_pcGrid.Rows[e.RowIndex].Tag is not PcRankingRow row) return;

        string colName = _pcGrid.Columns[e.ColumnIndex].Name;
        if (row.IsGlobal)
        {
            e.CellStyle.BackColor = Slate50;
            e.CellStyle.ForeColor = Slate500;
            if (colName == "colPcName") e.CellStyle.Font = _boldCellFont;
            return;
        }

        if (colName == "colPcName" &&
            string.Equals(row.MachineName, Environment.MachineName, StringComparison.OrdinalIgnoreCase))
        {
            e.CellStyle.Font = _boldCellFont;
            e.CellStyle.ForeColor = Indigo600;
        }
        else if (colName == "colSpeed")
        {
            e.CellStyle.ForeColor = GetSpeedColor(row.RelativeDurationRatio);
            if (row.RelativeDurationRatio < 0.95 || row.RelativeDurationRatio > 1.05)
                e.CellStyle.Font = _boldCellFont;
        }
    }

    private void OnOpGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.CellStyle == null || e.RowIndex < 0 || e.RowIndex >= _opGrid.Rows.Count) return;
        if (_opGrid.Columns[e.ColumnIndex].Name != "colVsAvg") return;

        if (_opGrid.Rows[e.RowIndex].Tag is double pctDiff)
        {
            double ratio = 1.0 + (pctDiff / 100.0);
            e.CellStyle.ForeColor = GetSpeedColor(ratio);
            if (ratio < 0.95 || ratio > 1.05)
                e.CellStyle.Font = _boldCellFont;
        }
        else
        {
            e.CellStyle.ForeColor = Slate400;
        }
    }

    private void OnBottleneckGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.CellStyle == null || e.RowIndex < 0 || e.RowIndex >= _bottleneckGrid.Rows.Count) return;
        if (_bottleneckGrid.Rows[e.RowIndex].Tag is not BottleneckRow row) return;

        string colName = _bottleneckGrid.Columns[e.ColumnIndex].Name;
        if (colName == "colStepDiagnosis")
        {
            switch (row.Diagnosis)
            {
                case "Software Bottleneck":
                    e.CellStyle.ForeColor = Red600;
                    e.CellStyle.Font = _boldCellFont;
                    break;
                case "Slow (1 PC)":
                case "Intermittent Hang":
                    e.CellStyle.ForeColor = Amber600;
                    e.CellStyle.Font = _boldCellFont;
                    break;
                case "Hardware / PC-Specific":
                    e.CellStyle.ForeColor = Indigo600;
                    e.CellStyle.Font = _boldCellFont;
                    break;
                default:
                    e.CellStyle.ForeColor = Emerald600;
                    break;
            }
        }
        else if (colName == "colStepMax" && row.MaxMs >= row.ServerAvgMs * 3.0 && row.MaxMs >= 2000.0)
        {
            e.CellStyle.ForeColor = Amber600;
            e.CellStyle.Font = _boldCellFont;
        }
    }

    private void OnStepPcGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.CellStyle == null || e.RowIndex < 0 || e.RowIndex >= _stepPcGrid.Rows.Count) return;
        if (_stepPcGrid.Rows[e.RowIndex].Tag is not ValueTuple<string, double?> tag) return;

        string colName = _stepPcGrid.Columns[e.ColumnIndex].Name;
        if (colName == "colStepPcName" &&
            string.Equals(tag.Item1, Environment.MachineName, StringComparison.OrdinalIgnoreCase))
        {
            e.CellStyle.Font = _boldCellFont;
            e.CellStyle.ForeColor = Indigo600;
        }
        else if (colName == "colStepPcVsAvg")
        {
            if (tag.Item2 is double ratio)
            {
                e.CellStyle.ForeColor = GetSpeedColor(ratio);
                if (ratio < 0.95 || ratio > 1.05)
                    e.CellStyle.Font = _boldCellFont;
            }
            else
            {
                e.CellStyle.ForeColor = Slate400;
            }
        }
    }

    // =========================================================================
    // 5. ACTIONS
    // =========================================================================
    private async Task DeleteSelectedPcAsync()
    {
        PcRankingRow? selected = GetSelectedPcRow();
        if (_busy || selected == null || selected.IsGlobal) return;

        DialogResult confirm = MessageBox.Show(
            this,
            $"Delete all recorded loading timings for PC '{selected.MachineName}'?\n\n" +
            $"({selected.OperationsMeasured} operation(s), {selected.TotalSamples} sample(s))\n" +
            "Server averages (__global__) will not be affected.",
            "Delete Selected PC",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Warning);

        if (confirm != DialogResult.Yes) return;

        SetBusy(true, $"Deleting timings for {selected.MachineName}...");
        try
        {
            string machine = selected.MachineName;
            await Task.Run(() => DeleteMachineTimings(machine));
        }
        catch (Exception ex)
        {
            if (IsDisposed) return;
            AppLogger.Warn($"[OfficeManagerDialog] Delete failed for {selected.MachineName}: {ex.Message}");
            SetBusy(false, null);
            MessageBox.Show(
                this,
                $"Failed to delete timings for '{selected.MachineName}':\n\n{ex.Message}",
                "Delete Selected PC",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
            return;
        }

        if (IsDisposed) return;
        SetBusy(false, null);
        await ReloadAsync(null);
    }

    private static void DeleteMachineTimings(string machineName)
    {
        using var conn = OpenTelemetryConnection();

        ExecDeleteSafe(conn, "DELETE FROM app.operation_timings WHERE machine_name = @Machine;", machineName, ignoreMissingTable: false);
        ExecDeleteSafe(conn, "DELETE FROM app.operation_timing_samples WHERE machine_name = @Machine;", machineName, ignoreMissingTable: true);
        ExecDeleteSafe(conn, "DELETE FROM app.operation_step_samples WHERE machine_name = @Machine;", machineName, ignoreMissingTable: true);
    }

    private static void ExecDeleteSafe(NpgsqlConnection conn, string sql, string machineName, bool ignoreMissingTable)
    {
        try
        {
            using var cmd = new NpgsqlCommand(sql, conn);
            cmd.Parameters.AddWithValue("Machine", machineName);
            cmd.ExecuteNonQuery();
        }
        catch (PostgresException pgEx) when (ignoreMissingTable && pgEx.SqlState == "42P01")
        {
            // Table not yet created; nothing to delete.
        }
    }

    // =========================================================================
    // 6. HELPERS & CLEANUP
    // =========================================================================
    private PcRankingRow? GetSelectedPcRow()
    {
        if (_pcGrid.SelectedRows.Count == 0) return null;
        return _pcGrid.SelectedRows[0].Tag as PcRankingRow;
    }

    private string? GetSelectedMachineName() => GetSelectedPcRow()?.MachineName;

    private void SetBusy(bool busy, string? status)
    {
        _busy = busy;
        UseWaitCursor = busy;
        _btnRefresh.Enabled = !busy;
        _cboTimeframe.Enabled = !busy;
        _cboVersion.Enabled = !busy;
        UpdateDeleteButtonState();
        if (busy && status != null)
        {
            _lblSummary.ForeColor = Slate500;
            _lblSummary.Text = status;
        }
    }

    private void UpdateDeleteButtonState()
    {
        PcRankingRow? selected = GetSelectedPcRow();
        _btnDeletePc.Enabled = !_busy && selected != null && !selected.IsGlobal;
    }

    private static string FormatSeconds(double ms) => $"{ms / 1000.0:0.00}s";

    private static string FormatSpeedRatio(double ratio)
    {
        if (ratio < 0.995)
            return $"{ratio:0.00}× ({Math.Round((1.0 - ratio) * 100)}% faster)";
        if (ratio <= 1.005)
            return $"{ratio:0.00}× (Avg)";
        return $"{ratio:0.00}× ({Math.Round((ratio - 1.0) * 100)}% slower)";
    }

    private static Color GetSpeedColor(double ratio)
    {
        if (ratio < 0.95) return Emerald600;
        if (ratio <= 1.05) return Slate500;
        if (ratio <= 1.25) return Amber600;
        return Red600;
    }

    private static void EnableDoubleBuffering(DataGridView grid)
    {
        typeof(DataGridView)
            .GetProperty("DoubleBuffered", BindingFlags.Instance | BindingFlags.NonPublic)
            ?.SetValue(grid, true);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        _titleFont?.Dispose();
        _sectionFont?.Dispose();
        _boldCellFont?.Dispose();
        _headerCellFont?.Dispose();
        _titleFont = null;
        _sectionFont = null;
        _boldCellFont = null;
        _headerCellFont = null;
        base.OnFormClosed(e);
    }
}
#endif
