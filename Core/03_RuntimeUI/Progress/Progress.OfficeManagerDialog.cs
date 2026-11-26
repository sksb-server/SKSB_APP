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

/// <summary>
/// DEBUG-only diagnostics window displaying per-PC relative loading speeds
/// from <c>app.pc_speed_ranking</c> and per-operation breakdowns from <c>app.operation_timings</c>.
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

    private sealed class TelemetrySnapshot
    {
        public List<PcRankingRow> Rankings { get; init; } = new();
        public List<OperationTimingRow> Timings { get; init; } = new();
    }

    private List<PcRankingRow> _pcRows = new();
    private List<OperationTimingRow> _allTimings = new();
    private Dictionary<string, OperationTimingRow> _globalTimings = new(StringComparer.OrdinalIgnoreCase);
    private bool _busy;
    private bool _suppressSelection;

    private Font? _titleFont;
    private Font? _sectionFont;
    private Font? _boldCellFont;
    private Font? _headerCellFont;

    private Button _btnRefresh = null!;
    private Button _btnDeletePc = null!;
    private Label _lblSummary = null!;
    private SplitContainer _split = null!;
    private DataGridView _pcGrid = null!;
    private Label _lblOpsHeader = null!;
    private DataGridView _opGrid = null!;

    public OfficeManagerDialog()
    {
        SuspendLayout();
        Text = "Office Manager — PC Speeds (Debug)";
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = true;
        MaximizeBox = true;
        ShowInTaskbar = false;
        BackColor = Slate50;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(12);
        ClientSize = new Size(1060, 620);
        MinimumSize = new Size(780, 460);

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
            Text = "Office Manager — PC Speeds",
            Font = _titleFont,
            ForeColor = Slate900,
            AutoSize = true,
            Margin = new Padding(0, 0, 16, 0)
        };

        _btnRefresh = CreateButton("Refresh", primary: false);
        _btnRefresh.Margin = new Padding(0, 1, 8, 0);
        _btnRefresh.Click += async (s, e) => await ReloadAsync(GetSelectedMachineName());

        _btnDeletePc = CreateButton("Delete Selected PC", primary: false);
        _btnDeletePc.ForeColor = Red600;
        _btnDeletePc.Enabled = false;
        _btnDeletePc.Margin = new Padding(0, 1, 12, 0);
        _btnDeletePc.Click += async (s, e) => await DeleteSelectedPcAsync();

        leftBar.Controls.AddRange(new Control[] { lblTitle, _btnRefresh, _btnDeletePc });

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

        // --- Split container (Left ~44% PC Ranking, Right ~56% Operation Breakdown) ---
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

        Controls.Add(_split);
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
    }

    // =========================================================================
    // 3. DATA LOADING & SERVER QUERIES
    // =========================================================================
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

        SetBusy(true, "Loading PC speed telemetry...");
        try
        {
            TelemetrySnapshot snapshot = await Task.Run(FetchTelemetryFromServer);
            if (IsDisposed) return;

            _allTimings = snapshot.Timings;
            _globalTimings = _allTimings
                .Where(t => string.Equals(t.MachineName, LoadingTimeEstimator.GlobalMachineKey, StringComparison.OrdinalIgnoreCase))
                .ToDictionary(t => t.OperationKey, t => t, StringComparer.OrdinalIgnoreCase);

            _pcRows = BuildPcDisplayRows(snapshot.Rankings, _globalTimings);
            PopulatePcGrid(reselectMachine);

            int pcCount = _pcRows.Count(r => !r.IsGlobal);
            int opCount = _globalTimings.Count > 0
                ? _globalTimings.Count
                : _allTimings.Select(t => t.OperationKey).Distinct(StringComparer.OrdinalIgnoreCase).Count();
            int sampleCount = _pcRows.Where(r => !r.IsGlobal).Sum(r => r.TotalSamples);

            _lblSummary.ForeColor = Slate500;
            _lblSummary.Text = $"{pcCount} PC(s) · {opCount} operation(s) · {sampleCount} sample(s)";
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

    private static TelemetrySnapshot FetchTelemetryFromServer()
    {
        using var conn = OpenTelemetryConnection();

        var rankings = new List<PcRankingRow>();
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

        var timings = new List<OperationTimingRow>();
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

        return new TelemetrySnapshot { Rankings = rankings, Timings = timings };
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
        _pcGrid.Rows.Clear();
        _opGrid.Rows.Clear();
        _lblOpsHeader.Text = "Operations";
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
        using var cmd = new NpgsqlCommand(
            "DELETE FROM app.operation_timings WHERE machine_name = @Machine;", conn);
        cmd.Parameters.AddWithValue("Machine", machineName);
        cmd.ExecuteNonQuery();
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
