namespace SKSB_App.Core.WorkbookSave.Validation;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;

/// <summary>
/// Explicit COM default interface for <see cref="ValidationTaskPaneControl"/>. On .NET 8 + Excel-DNA,
/// <c>[ComVisible(true)]</c> alone does not produce a usable class interface, so the Custom Task Pane
/// ActiveX host fails ("Unable to create specified ActiveX control") without this.
/// </summary>
[ComVisible(true)]
[Guid("A3F6E2C1-8B4D-4E9A-9C1F-5D7B2E8A4C61")]
public interface IValidationTaskPaneControl { }

/// <summary>
/// Content of the docked "Validation Issues" Excel Custom Task Pane. Pure WinForms: it never touches Excel COM.
/// Every callback handed to <see cref="SetIssues"/> / <see cref="IssueSelected"/> is expected to be pre-wrapped
/// in <c>ExcelAsyncUtil.QueueAsMacro</c> by <see cref="ValidationTaskPaneManager"/>.
/// </summary>
[ComVisible(true)]
[Guid("B7D9C4A2-3E5F-4A8B-8D2E-6F1C9A3B5E72")]
[ProgId("SKSB_App.ValidationTaskPaneControl")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IValidationTaskPaneControl))]
public class ValidationTaskPaneControl : UserControl, IValidationTaskPaneControl
{
    // =========================================================================
    // 1. PALETTE & STATE
    // =========================================================================
    private static readonly Color ToolbarBack = ColorTranslator.FromHtml("#F8FAFC");
    private static readonly Color Divider = ColorTranslator.FromHtml("#E2E8F0");
    private static readonly Color TextPrimary = ColorTranslator.FromHtml("#1E293B");
    private static readonly Color TextMuted = ColorTranslator.FromHtml("#64748B");
    private static readonly Color ErrorFore = ColorTranslator.FromHtml("#B91C1C");
    private static readonly Color ErrorBack = ColorTranslator.FromHtml("#FEE2E2");
    private static readonly Color WarningFore = ColorTranslator.FromHtml("#B45309");
    private static readonly Color WarningBack = ColorTranslator.FromHtml("#FEF3C7");
    private static readonly Color SuccessFore = ColorTranslator.FromHtml("#15803D");
    private static readonly Color AccentBack = ColorTranslator.FromHtml("#DBEAFE");
    private static readonly Color AccentBorder = ColorTranslator.FromHtml("#3B82F6");
    private static readonly Color AltRowBack = ColorTranslator.FromHtml("#F8FAFC");
    private static readonly Color SelectionBack = ColorTranslator.FromHtml("#E0ECFF");

    private static readonly Font RegularFont = new("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
    private static readonly Font BoldFont = new("Segoe UI", 9f, FontStyle.Bold, GraphicsUnit.Point);
    private static readonly Font PillFont = new("Segoe UI Semibold", 8.5f, FontStyle.Bold, GraphicsUnit.Point);

    private enum IssueFilter { All, Errors, Warnings }

    private readonly Panel _toolbar;
    private readonly Label _summaryLabel;
    private readonly Button _btnAll;
    private readonly Button _btnErrors;
    private readonly Button _btnWarnings;
    private readonly Button _btnRevalidate;
    private readonly Button _btnProceed;
    private readonly Button _btnDebugQuit;
    private readonly Button _btnRevertHistory;
    private readonly Button _btnCloneHistory;
    private readonly Button _btnKeepServer;
    private readonly DataGridView _grid;

    private static readonly Color ProceedBack = ColorTranslator.FromHtml("#16A34A");
    private static readonly Color ProceedBorder = ColorTranslator.FromHtml("#15803D");
    private static readonly Color DangerBack = ColorTranslator.FromHtml("#DC2626");
    private static readonly Color DangerBorder = ColorTranslator.FromHtml("#B91C1C");

    private List<ValidationIssue> _errors = new();
    private List<ValidationIssue> _warnings = new();
    private IssueFilter _filter = IssueFilter.All;
    private Action? _onRevalidate;
    private Action? _onProceedWithSave;
    private Action? _onDebugQuitAnyway;
    private Action? _onRevertHistoryEdits;
    private Action? _onCloneHistoryAsNewRows;
    private Action? _onKeepServerValues;
    private bool _isClosing;
    private bool _isConflictMode;

    /// <summary>Raised when the user clicks / double-clicks an issue row.</summary>
    public Action<ValidationIssue>? IssueSelected { get; set; }

    // =========================================================================
    // 2. CONSTRUCTION & LAYOUT
    // =========================================================================
    public ValidationTaskPaneControl()
    {
        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        BackColor = Color.White;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular, GraphicsUnit.Point);
        DoubleBuffered = true;

        SuspendLayout();

        _toolbar = new Panel
        {
            Dock = DockStyle.Top,
            Height = 40,
            BackColor = ToolbarBack,
            Padding = new Padding(8, 0, 8, 0)
        };
        _toolbar.Paint += (_, e) =>
        {
            using var pen = new Pen(Divider);
            e.Graphics.DrawLine(pen, 0, _toolbar.Height - 1, _toolbar.Width, _toolbar.Height - 1);
        };

        var leftFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Left,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            FlowDirection = FlowDirection.LeftToRight,
            BackColor = Color.Transparent,
            Padding = new Padding(0, 6, 0, 0)
        };

        _summaryLabel = new Label
        {
            AutoSize = true,
            Font = new Font("Segoe UI Semibold", 9.5f, FontStyle.Bold, GraphicsUnit.Point),
            ForeColor = TextPrimary,
            Margin = new Padding(0, 6, 14, 0),
            Text = "No validation run yet"
        };

        _btnAll = CreateToolbarButton("All (0)");
        _btnErrors = CreateToolbarButton("Errors (0)");
        _btnWarnings = CreateToolbarButton("Warnings (0)");
        _btnAll.Click += (_, _) => ApplyFilter(IssueFilter.All);
        _btnErrors.Click += (_, _) => ApplyFilter(IssueFilter.Errors);
        _btnWarnings.Click += (_, _) => ApplyFilter(IssueFilter.Warnings);

        leftFlow.Controls.Add(_summaryLabel);
        leftFlow.Controls.Add(_btnAll);
        leftFlow.Controls.Add(_btnErrors);
        leftFlow.Controls.Add(_btnWarnings);

        var rightFlow = new FlowLayoutPanel
        {
            Dock = DockStyle.Right,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            WrapContents = false,
            FlowDirection = FlowDirection.RightToLeft,
            BackColor = Color.Transparent,
            Padding = new Padding(0, 6, 0, 0)
        };

        _btnProceed = CreateToolbarButton("Proceed with Save");
        _btnProceed.BackColor = ProceedBack;
        _btnProceed.ForeColor = Color.White;
        _btnProceed.FlatAppearance.BorderColor = ProceedBorder;
        _btnProceed.Visible = false;
        _btnProceed.Click += (_, _) =>
        {
            if (_isConflictMode) InvokeConflictAction(_onProceedWithSave);
            else _onProceedWithSave?.Invoke();
        };

        _btnKeepServer = CreateToolbarButton("Keep Server Values");
        _btnKeepServer.BackColor = AccentBack;
        _btnKeepServer.ForeColor = ColorTranslator.FromHtml("#1D4ED8");
        _btnKeepServer.FlatAppearance.BorderColor = AccentBorder;
        _btnKeepServer.Visible = false;
        _btnKeepServer.Click += (_, _) => InvokeConflictAction(_onKeepServerValues);

        _btnRevalidate = CreateToolbarButton("Re-validate");
        _btnRevalidate.Click += (_, _) => _onRevalidate?.Invoke();

        _btnDebugQuit = CreateToolbarButton("Debug: Quit Anyway");
        _btnDebugQuit.ForeColor = ErrorFore;
        _btnDebugQuit.Visible = false;
        _btnDebugQuit.Click += (_, _) => _onDebugQuitAnyway?.Invoke();

        _btnRevertHistory = CreateToolbarButton("Revert History Edits");
        _btnRevertHistory.BackColor = WarningBack;
        _btnRevertHistory.ForeColor = WarningFore;
        _btnRevertHistory.FlatAppearance.BorderColor = WarningFore;
        _btnRevertHistory.Visible = false;
        _btnRevertHistory.Click += (_, _) => InvokeHistoryAction(_onRevertHistoryEdits);

        _btnCloneHistory = CreateToolbarButton("Clone Edits as New Rows");
        _btnCloneHistory.BackColor = AccentBack;
        _btnCloneHistory.ForeColor = ColorTranslator.FromHtml("#1D4ED8");
        _btnCloneHistory.FlatAppearance.BorderColor = AccentBorder;
        _btnCloneHistory.Visible = false;
        _btnCloneHistory.Click += (_, _) => InvokeHistoryAction(_onCloneHistoryAsNewRows);

        // RightToLeft flow: first added sits at the far right.
        rightFlow.Controls.Add(_btnProceed);
        rightFlow.Controls.Add(_btnKeepServer);
        rightFlow.Controls.Add(_btnRevalidate);
        rightFlow.Controls.Add(_btnCloneHistory);
        rightFlow.Controls.Add(_btnRevertHistory);
        rightFlow.Controls.Add(_btnDebugQuit);

        _toolbar.Controls.Add(rightFlow);
        _toolbar.Controls.Add(leftFlow);

        _grid = new DataGridView
        {
            Dock = DockStyle.Fill,
            ReadOnly = true,
            SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            MultiSelect = false,
            RowHeadersVisible = false,
            AllowUserToAddRows = false,
            AllowUserToDeleteRows = false,
            AllowUserToResizeRows = false,
            AllowUserToOrderColumns = false,
            AutoGenerateColumns = false,
            BackgroundColor = Color.White,
            BorderStyle = BorderStyle.None,
            CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal,
            GridColor = Divider,
            EnableHeadersVisualStyles = false,
            ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.DisableResizing,
            ColumnHeadersHeight = 28,
            ColumnHeadersBorderStyle = DataGridViewHeaderBorderStyle.Single,
            StandardTab = true
        };
        _grid.RowTemplate.Height = 24;
        _grid.ColumnHeadersDefaultCellStyle.BackColor = ToolbarBack;
        _grid.ColumnHeadersDefaultCellStyle.ForeColor = TextMuted;
        _grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = ToolbarBack;
        _grid.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI Semibold", 9f, FontStyle.Bold, GraphicsUnit.Point);
        _grid.DefaultCellStyle.ForeColor = TextPrimary;
        _grid.DefaultCellStyle.SelectionBackColor = SelectionBack;
        _grid.DefaultCellStyle.SelectionForeColor = TextPrimary;
        _grid.DefaultCellStyle.Padding = new Padding(4, 0, 4, 0);
        _grid.AlternatingRowsDefaultCellStyle.BackColor = AltRowBack;

        _grid.Columns.Add(CreateColumn("colType", "Type", 78));
        _grid.Columns.Add(CreateColumn("colSheet", "Sheet", 140));
        _grid.Columns.Add(CreateColumn("colCell", "Cell", 70));
        _grid.Columns.Add(CreateColumn("colColumn", "Column", 150));
        var details = CreateColumn("colDetails", "Details", 300);
        details.AutoSizeMode = DataGridViewAutoSizeColumnMode.Fill;
        details.DefaultCellStyle.WrapMode = DataGridViewTriState.False;
        _grid.Columns.Add(details);

        _grid.CellFormatting += OnGridCellFormatting;
        _grid.CellClick += OnGridCellActivated;
        _grid.CellDoubleClick += OnGridCellActivated;
        _grid.KeyDown += OnGridKeyDown;

        Controls.Add(_grid);
        Controls.Add(_toolbar);

        ResumeLayout(performLayout: true);
        RefreshToolbar();
    }

    private static Button CreateToolbarButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.White,
            ForeColor = TextPrimary,
            Height = 26,
            Margin = new Padding(0, 0, 6, 0),
            Padding = new Padding(6, 0, 6, 0),
            Cursor = Cursors.Hand,
            UseVisualStyleBackColor = false
        };
        btn.FlatAppearance.BorderColor = Divider;
        btn.FlatAppearance.BorderSize = 1;
        return btn;
    }

    private static DataGridViewTextBoxColumn CreateColumn(string name, string header, int width) => new()
    {
        Name = name,
        HeaderText = header,
        Width = width,
        SortMode = DataGridViewColumnSortMode.NotSortable,
        ReadOnly = true
    };

    // =========================================================================
    // 3. PUBLIC API
    // =========================================================================
    public void SetIssues(
        List<ValidationIssue> errors,
        List<ValidationIssue> warnings,
        Action? onRevalidate,
        Action? onProceedWithSave,
        bool isClosing,
        Action? onDebugQuitAnyway,
        Action? onRevertHistoryEdits = null,
        Action? onCloneHistoryAsNewRows = null,
        Action? onKeepServerValues = null)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try
            {
                BeginInvoke(new Action(() => SetIssues(errors, warnings, onRevalidate, onProceedWithSave, isClosing, onDebugQuitAnyway,
                    onRevertHistoryEdits, onCloneHistoryAsNewRows, onKeepServerValues)));
            }
            catch { }
            return;
        }

        _errors = errors ?? new List<ValidationIssue>();
        _warnings = warnings ?? new List<ValidationIssue>();
        _onRevalidate = onRevalidate;
        _onProceedWithSave = onProceedWithSave;
        _onRevertHistoryEdits = onRevertHistoryEdits;
        _onCloneHistoryAsNewRows = onCloneHistoryAsNewRows;
        _onKeepServerValues = onKeepServerValues;
        _isClosing = isClosing;
#if DEBUG
        _onDebugQuitAnyway = isClosing ? onDebugQuitAnyway : null;
#else
        _onDebugQuitAnyway = null;
#endif

        // Keep the current filter unless it would now show nothing while other issues exist.
        if (_filter == IssueFilter.Errors && _errors.Count == 0) _filter = IssueFilter.All;
        if (_filter == IssueFilter.Warnings && _warnings.Count == 0) _filter = IssueFilter.All;

        RefreshToolbar();
        RebuildRows();
    }

    // =========================================================================
    // 4. RENDERING
    // =========================================================================
    private void ApplyFilter(IssueFilter filter)
    {
        _filter = filter;
        RefreshToolbar();
        RebuildRows();
    }

    private void RefreshToolbar()
    {
        int e = _errors.Count;
        int w = _warnings.Count;

        // Conflict Mode: the save was rolled back by optimistic concurrency; only the two resolutions are offered.
        bool hasConflicts = e == 0 && _warnings.Any(x => x.IsConcurrencyConflict);
        _isConflictMode = hasConflicts;

        if (hasConflicts)
        {
            _summaryLabel.Text = $"{w} Conflict{(w == 1 ? "" : "s")}";
            _summaryLabel.ForeColor = WarningFore;
        }
        else if (e == 0 && w == 0)
        {
            _summaryLabel.Text = "✔ All checks passed";
            _summaryLabel.ForeColor = SuccessFore;
        }
        else
        {
            var parts = new List<string>();
            if (e > 0) parts.Add($"{e} Error{(e == 1 ? "" : "s")}");
            if (w > 0) parts.Add($"{w} Warning{(w == 1 ? "" : "s")}");
            _summaryLabel.Text = string.Join(", ", parts);
            _summaryLabel.ForeColor = e > 0 ? ErrorFore : WarningFore;
        }

        _btnAll.Text = $"All ({e + w})";
        _btnErrors.Text = $"Errors ({e})";
        _btnWarnings.Text = $"Warnings ({w})";
        StyleFilterButton(_btnAll, _filter == IssueFilter.All);
        StyleFilterButton(_btnErrors, _filter == IssueFilter.Errors);
        StyleFilterButton(_btnWarnings, _filter == IssueFilter.Warnings);

        if (hasConflicts)
        {
            // Red so the destructive overwrite is never clicked by habit; blue "Keep Server Values" is the safe path.
            _btnProceed.Text = _isClosing ? "Overwrite && Close" : "Overwrite Server && Save";
            _btnProceed.BackColor = DangerBack;
            _btnProceed.ForeColor = Color.White;
            _btnProceed.FlatAppearance.BorderColor = DangerBorder;
            _btnProceed.Visible = _onProceedWithSave != null;
            _btnProceed.Enabled = _btnProceed.Visible;

            _btnKeepServer.Visible = _onKeepServerValues != null;
            _btnKeepServer.Enabled = _btnKeepServer.Visible;

            _btnRevalidate.Visible = false;
            _btnRevertHistory.Visible = false;
            _btnCloneHistory.Visible = false;
        }
        else
        {
            _btnProceed.Text = _isClosing ? "Save && Close" : "Proceed with Save";
            _btnProceed.BackColor = ProceedBack;
            _btnProceed.ForeColor = Color.White;
            _btnProceed.FlatAppearance.BorderColor = ProceedBorder;
            _btnKeepServer.Visible = false;

            _btnRevalidate.Visible = _onRevalidate != null;
            bool canProceed = e == 0 && _onProceedWithSave != null;
            _btnProceed.Visible = canProceed;
            _btnProceed.Enabled = canProceed;

            bool hasStacked = _warnings.Any(i => i.IsStackedHistoryEdit);
            _btnRevertHistory.Visible = hasStacked && _onRevertHistoryEdits != null;
            _btnRevertHistory.Enabled = _btnRevertHistory.Visible;
            _btnCloneHistory.Visible = hasStacked && _onCloneHistoryAsNewRows != null;
            _btnCloneHistory.Enabled = _btnCloneHistory.Visible;
        }
#if DEBUG
        _btnDebugQuit.Visible = _onDebugQuitAnyway != null;
#else
        _btnDebugQuit.Visible = false;
#endif
    }

    /// <summary>
    /// Disables both history buttons until the next <see cref="SetIssues"/> (the action re-validates), so a double
    /// click can never queue the same revert / clone twice.
    /// </summary>
    private void InvokeHistoryAction(Action? action)
    {
        if (action == null) return;
        _btnRevertHistory.Enabled = false;
        _btnCloneHistory.Enabled = false;
        action();
    }

    /// <summary>Same double-click guard for the Conflict Mode buttons (re-enabled by the next <see cref="SetIssues"/>).</summary>
    private void InvokeConflictAction(Action? action)
    {
        if (action == null) return;
        _btnKeepServer.Enabled = false;
        _btnProceed.Enabled = false;
        action();
    }

    private static void StyleFilterButton(Button btn, bool active)
    {
        btn.BackColor = active ? AccentBack : Color.White;
        btn.FlatAppearance.BorderColor = active ? AccentBorder : Divider;
        btn.Font = active ? BoldFont : RegularFont;
    }

    private void RebuildRows()
    {
        IEnumerable<ValidationIssue> source = _filter switch
        {
            IssueFilter.Errors => _errors,
            IssueFilter.Warnings => _warnings,
            _ => _errors.Concat(_warnings)
        };

        _grid.SuspendLayout();
        try
        {
            _grid.Rows.Clear();
            foreach (var issue in source)
            {
                int idx = _grid.Rows.Add(
                    issue.Severity == ValidationSeverity.Error ? "Error"
                        : issue.IsConcurrencyConflict ? "Conflict"
                        : "Warning",
                    issue.SheetName,
                    issue.CellAddress,
                    issue.ColumnHeader,
                    issue.Message);
                var row = _grid.Rows[idx];
                row.Tag = issue;
                row.Cells[4].ToolTipText = issue.Message;
            }
            _grid.ClearSelection();
        }
        finally
        {
            _grid.ResumeLayout();
        }
    }

    private void OnGridCellFormatting(object? sender, DataGridViewCellFormattingEventArgs e)
    {
        if (e.RowIndex < 0 || e.ColumnIndex != 0 || e.CellStyle == null) return;
        if (_grid.Rows[e.RowIndex].Tag is not ValidationIssue issue) return;

        bool isError = issue.Severity == ValidationSeverity.Error;
        e.CellStyle.BackColor = isError ? ErrorBack : WarningBack;
        e.CellStyle.ForeColor = isError ? ErrorFore : WarningFore;
        e.CellStyle.SelectionBackColor = isError ? ErrorBack : WarningBack;
        e.CellStyle.SelectionForeColor = isError ? ErrorFore : WarningFore;
        e.CellStyle.Font = PillFont;
        e.CellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
    }

    // =========================================================================
    // 5. NAVIGATION
    // =========================================================================
    private void OnGridCellActivated(object? sender, DataGridViewCellEventArgs e)
    {
        if (e.RowIndex < 0 || e.RowIndex >= _grid.Rows.Count) return;
        RaiseIssueSelected(_grid.Rows[e.RowIndex]);
    }

    private void OnGridKeyDown(object? sender, KeyEventArgs e)
    {
        if (e.KeyCode != Keys.Enter || _grid.CurrentRow == null) return;
        e.Handled = true;
        RaiseIssueSelected(_grid.CurrentRow);
    }

    private void RaiseIssueSelected(DataGridViewRow row)
    {
        if (row.Tag is ValidationIssue issue) IssueSelected?.Invoke(issue);
    }
}
