namespace SKSB_App.Core.RuntimeUI.Guide;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.Types.Attributes;

[ComVisible(true)]
[Guid("D4A8F1B3-6C2E-4F7A-9B3D-8E1C5A2F7B94")]
public interface IGuideTaskPaneControl { }

/// <summary>
/// Content of the floating / dockable "Cell Guide" Custom Task Pane. Pure WinForms: it only renders
/// <see cref="CellGuideSnapshot"/> DTOs and never touches Excel COM. Must stay public, [ComVisible(true)]
/// and keep a parameterless constructor (required by <c>CustomTaskPaneFactory.CreateCustomTaskPane</c>).
/// </summary>
[ComVisible(true)]
[Guid("E9C2B5D7-4A1F-48E3-B6C8-2D9A7F4E1C35")]
[ProgId("SKSB_App.GuideTaskPaneControl")]
[ClassInterface(ClassInterfaceType.None)]
[ComDefaultInterface(typeof(IGuideTaskPaneControl))]
public class GuideTaskPaneControl : UserControl, IGuideTaskPaneControl
{
    // =========================================================================
    // 1. THEME
    // =========================================================================
    private static readonly Color PaneBack = Color.FromArgb(248, 249, 251);
    private static readonly Color CardBack = Color.White;
    private static readonly Color CardBorder = Color.FromArgb(222, 226, 232);
    private static readonly Color TextPrimary = Color.FromArgb(32, 33, 36);
    private static readonly Color TextMuted = Color.FromArgb(95, 99, 104);
    private static readonly Color Accent = Color.FromArgb(26, 115, 232);
    private static readonly Color TipCardBack = Color.FromArgb(254, 249, 231);
    private static readonly Color TipCardBorder = Color.FromArgb(245, 224, 163);
    private static readonly Color TipCardAccent = Color.FromArgb(146, 94, 0);

    private static readonly (Color Back, Color Fore) BadgeRequired = (Color.FromArgb(252, 232, 230), Color.FromArgb(197, 34, 31));
    private static readonly (Color Back, Color Fore) BadgeWarning = (Color.FromArgb(254, 247, 224), Color.FromArgb(176, 96, 0));
    private static readonly (Color Back, Color Fore) BadgeOptional = (Color.FromArgb(241, 243, 244), Color.FromArgb(95, 99, 104));
    private static readonly (Color Back, Color Fore) BadgeLocked = (Color.FromArgb(232, 234, 237), Color.FromArgb(60, 64, 67));
    private static readonly (Color Back, Color Fore) BadgeCalculated = (Color.FromArgb(243, 232, 253), Color.FromArgb(132, 48, 206));
    private static readonly (Color Back, Color Fore) BadgeDropdown = (Color.FromArgb(232, 240, 254), Color.FromArgb(25, 103, 210));
    private static readonly (Color Back, Color Fore) BadgeSync = (Color.FromArgb(230, 244, 234), Color.FromArgb(19, 115, 51));
    private static readonly (Color Back, Color Fore) BadgeType = (Color.FromArgb(228, 247, 251), Color.FromArgb(0, 113, 140));

    private static readonly string[] GeneralTipIds =
    [
        "ctrl_s", "insert_top", "trim_empty", "dedupe", "total_row", "dropdown", "language", "dock_pane",
        "dictionary", "template_portal", "validate", "report_bug", "header_sort", "reload_data",
        "total_row_filter", "insert_row", "delete_row", "offline_safe", "cell_guide_toggle", "date_filter", "auto_defaults", "multi_select", "dynamic_cost"
    ];

    // =========================================================================
    // 2. LAYOUT & IDLE FLOATING TRANSPARENCY (Win32 P/Invoke — never touches XLMAIN)
    // =========================================================================
    private const uint GA_ROOT = 2;
    private const int GWL_EXSTYLE = -20;
    private const int WS_EX_LAYERED = 0x00080000;
    private const uint LWA_ALPHA = 0x2;

    private const double IdleFadeDelaySeconds = 8.0;
    private const byte InactiveAlpha = 125;
    private const byte ActiveAlpha = 255;
    private const byte FadeStep = 26;

    [StructLayout(LayoutKind.Sequential)]
    private struct RECT
    {
        public int Left;
        public int Top;
        public int Right;
        public int Bottom;
    }

    [DllImport("user32.dll", ExactSpelling = true)]
    private static extern IntPtr GetAncestor(IntPtr hwnd, uint gaFlags);

    [DllImport("user32.dll", CharSet = CharSet.Unicode)]
    private static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int GetWindowLong(IntPtr hWnd, int nIndex);

    [DllImport("user32.dll", SetLastError = true)]
    private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

    [DllImport("user32.dll", SetLastError = true)]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetLayeredWindowAttributes(IntPtr hwnd, uint crKey, byte bAlpha, uint dwFlags);

    private readonly Panel _scrollHost;
    private readonly TableLayoutPanel _content;
    private readonly System.Windows.Forms.Timer _fadeTimer;
    private CellGuideSnapshot _snapshot = CellGuideSnapshot.Empty();
    private int _lastRenderWidth = -1;
    private int _tipIndex = Random.Shared.Next(GeneralTipIds.Length);
    private bool _tipDismissed;
    private bool _preferGeneralTipForCurrentCell;

    private IntPtr _excelMainHwnd;
    private bool _isFloating = true;
    private IntPtr _lastFloatingRootHwnd;
    private DateTime _lastActivityUtc = DateTime.UtcNow;
    private byte _currentAlpha = ActiveAlpha;

    public GuideTaskPaneControl()
    {
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9F, FontStyle.Regular, GraphicsUnit.Point);
        BackColor = PaneBack;
        DoubleBuffered = true;

        _scrollHost = new Panel
        {
            Dock = DockStyle.Fill,
            AutoScroll = true,
            BackColor = PaneBack,
            Padding = new Padding(10, 8, 10, 10)
        };

        _content = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            ColumnCount = 1,
            BackColor = PaneBack,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        _content.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100F));

        _scrollHost.Controls.Add(_content);
        Controls.Add(_scrollHost);

        _scrollHost.Resize += OnHostResize;
        _scrollHost.MouseEnter += OnScrollHostMouseEnter;

        _fadeTimer = new System.Windows.Forms.Timer { Interval = 250 };
        _fadeTimer.Tick += OnFadeTimerTick;
        _fadeTimer.Start();

        Render();
    }

    /// <summary>Configures the Excel main window guard handle and initial floating state. Safe to call from any thread.</summary>
    public void ConfigureHostWindow(IntPtr excelMainHwnd, bool isFloating)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(() => ConfigureHostWindow(excelMainHwnd, isFloating))); } catch { }
            return;
        }
        _excelMainHwnd = excelMainHwnd;
        ApplyFloatingState(isFloating);
    }

    /// <summary>Updates whether the hosting Custom Task Pane is floating or docked. Safe to call from any thread.</summary>
    public void SetFloatingState(bool isFloating)
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(() => SetFloatingState(isFloating))); } catch { }
            return;
        }
        ApplyFloatingState(isFloating);
    }

    private void ApplyFloatingState(bool isFloating)
    {
        _isFloating = isFloating;
        if (!isFloating)
        {
            RestoreOpacity();
            _lastFloatingRootHwnd = IntPtr.Zero;
        }
        else
        {
            _lastActivityUtc = DateTime.UtcNow;
            RestoreOpacity();
        }
    }

    private void MarkActive()
    {
        _lastActivityUtc = DateTime.UtcNow;
        if (_currentAlpha != ActiveAlpha)
        {
            RestoreOpacity();
        }
    }

    private void OnScrollHostMouseEnter(object? sender, EventArgs e)
    {
        MarkActive();
    }

    private void OnFadeTimerTick(object? sender, EventArgs e)
    {
        if (IsDisposed || !IsHandleCreated || !Visible || !_isFloating)
        {
            if (_currentAlpha != ActiveAlpha) RestoreOpacity();
            return;
        }

        IntPtr rootHwnd = TryGetSafeFloatingRootHwnd();
        if (rootHwnd == IntPtr.Zero) return;

        if (GetWindowRect(rootHwnd, out RECT rect))
        {
            Point cursor = Control.MousePosition;
            if (cursor.X >= rect.Left && cursor.X <= rect.Right && cursor.Y >= rect.Top && cursor.Y <= rect.Bottom)
            {
                MarkActive();
                return;
            }
        }

        if ((DateTime.UtcNow - _lastActivityUtc).TotalSeconds >= IdleFadeDelaySeconds)
        {
            if (_currentAlpha > InactiveAlpha)
            {
                byte next = (byte)Math.Max(InactiveAlpha, _currentAlpha - FadeStep);
                ApplyFloatingAlpha(rootHwnd, next);
            }
        }
    }

    private IntPtr TryGetSafeFloatingRootHwnd()
    {
        if (!_isFloating || IsDisposed || !IsHandleCreated) return IntPtr.Zero;
        IntPtr rootHwnd = GetAncestor(Handle, GA_ROOT);
        return IsSafeFloatingHwnd(rootHwnd) ? rootHwnd : IntPtr.Zero;
    }

    private bool IsSafeFloatingHwnd(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return false;
        if (_excelMainHwnd != IntPtr.Zero && hwnd == _excelMainHwnd) return false;
        var sb = new StringBuilder(64);
        int len = GetClassName(hwnd, sb, sb.Capacity);
        if (len <= 0) return false;
        if (string.Equals(sb.ToString(), "XLMAIN", StringComparison.OrdinalIgnoreCase)) return false;
        return true;
    }

    private void ApplyFloatingAlpha(IntPtr rootHwnd, byte alpha)
    {
        if (!IsSafeFloatingHwnd(rootHwnd)) return;
        try
        {
            int exStyle = GetWindowLong(rootHwnd, GWL_EXSTYLE);
            if ((exStyle & WS_EX_LAYERED) == 0)
            {
                SetWindowLong(rootHwnd, GWL_EXSTYLE, exStyle | WS_EX_LAYERED);
            }
            if (SetLayeredWindowAttributes(rootHwnd, 0, alpha, LWA_ALPHA))
            {
                _currentAlpha = alpha;
                _lastFloatingRootHwnd = rootHwnd;
            }
        }
        catch { }
    }

    private void RestoreOpacity()
    {
        IntPtr target = _lastFloatingRootHwnd;
        if (target == IntPtr.Zero && _isFloating && !IsDisposed && IsHandleCreated)
        {
            target = TryGetSafeFloatingRootHwnd();
        }
        if (target != IntPtr.Zero && IsSafeFloatingHwnd(target))
        {
            try
            {
                int exStyle = GetWindowLong(target, GWL_EXSTYLE);
                if ((exStyle & WS_EX_LAYERED) != 0)
                {
                    SetLayeredWindowAttributes(target, 0, ActiveAlpha, LWA_ALPHA);
                }
            }
            catch { }
        }
        _currentAlpha = ActiveAlpha;
    }

    /// <summary>Re-renders the panel for <paramref name="snapshot"/>. Safe to call from any thread.</summary>
    public void UpdateContext(CellGuideSnapshot snapshot)
    {
        if (IsDisposed) return;
        var next = snapshot ?? CellGuideSnapshot.Empty();
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(() => UpdateContext(next))); } catch { }
            return;
        }
        _snapshot = next;
        _preferGeneralTipForCurrentCell = false;
        MarkActive();
        Render();
    }

    /// <summary>Re-renders the current snapshot in the newly active language.</summary>
    public void RefreshLanguage()
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(RefreshLanguage)); } catch { }
            return;
        }
        MarkActive();
        Render();
    }

    /// <summary>Advances to the next general workflow tip. Safe to call from any thread.</summary>
    public void AdvanceTip()
    {
        if (IsDisposed) return;
        if (InvokeRequired)
        {
            try { BeginInvoke(new Action(AdvanceTip)); } catch { }
            return;
        }
        var allTipKeys = Loc.GetAllGeneralTipKeys(GeneralTipIds);
        if (allTipKeys.Count > 0)
            _tipIndex = (_tipIndex + 1) % allTipKeys.Count;
        _preferGeneralTipForCurrentCell = false;
        MarkActive();
        if (!_tipDismissed) Render();
    }

    private void OnHostResize(object? sender, EventArgs e)
    {
        int width = ContentWidth;
        if (Math.Abs(width - _lastRenderWidth) > 4) Render();
    }

    private int ContentWidth
    {
        get
        {
            int w = _scrollHost.ClientSize.Width - _scrollHost.Padding.Horizontal - SystemInformation.VerticalScrollBarWidth;
            return Math.Max(160, w);
        }
    }

    // =========================================================================
    // 3. RENDERING
    // =========================================================================
    private void Render()
    {
        if (IsDisposed) return;
        int width = ContentWidth;
        _lastRenderWidth = width;

        _content.SuspendLayout();
        try
        {
            var old = _content.Controls.Cast<Control>().ToList();
            _content.Controls.Clear();
            _content.RowStyles.Clear();
            foreach (var c in old) c.Dispose();

            var s = _snapshot;
            if (!_tipDismissed)
                AddRow(BuildSpotlightTipCard(width));
            else
                AddRow(BuildShowTipLink(width));
            AddRow(BuildContextStrip(s, width));

            switch (s.Zone)
            {
                case CellZone.None:
                    AddRow(BuildBannerCard(Loc.Tip("guide.banner.noSelection"), width));
                    break;
                case CellZone.OutsideTable:
                    AddRow(BuildBannerCard(Loc.Tip("guide.banner.outside"), width));
                    break;
                default:
                    BuildTableSections(s, width);
                    break;
            }
        }
        finally
        {
            _content.ResumeLayout(true);
        }
    }

    private void BuildTableSections(CellGuideSnapshot s, int width)
    {
        if (s.SelectedCellCount > 1)
            AddRow(MakeLabel(Loc.Tip("guide.banner.multi", s.SelectedCellCount, s.CellAddress), width, TextMuted, italic: true));

        // ---- Column title card ----
        var (displayName, description) = Loc.ResolveColumnGuide(s.TableName, s.ColumnHeader);
        string header = string.IsNullOrWhiteSpace(s.ColumnHeader) ? Loc.Tip("guide.label.unnamedColumn") : s.ColumnHeader!;
        string title = displayName ?? header;

        var titleCard = NewCard(width);
        titleCard.Controls.Add(MakeLabel(title, CardInner(width), TextPrimary, size: 13F, bold: true));
        if (!string.Equals(title, header, StringComparison.OrdinalIgnoreCase))
            titleCard.Controls.Add(MakeLabel(Loc.Tip("guide.label.originalHeader", header), CardInner(width), TextMuted));
        if (s.Zone == CellZone.Data || s.Zone == CellZone.Total)
            titleCard.Controls.Add(BuildBadges(s, CardInner(width)));
        AddRow(titleCard);

        // ---- Zone banner (header / total) ----
        if (s.Zone == CellZone.Header)
            AddRow(BuildBannerCard(Loc.Tip("guide.banner.header"), width));
        else if (s.Zone == CellZone.Total)
            AddRow(BuildBannerCard(Loc.Tip("guide.banner.total"), width));

        // ---- About this field ----
        var about = NewCard(width);
        about.Controls.Add(MakeSectionHeader(Loc.Tip("guide.section.about"), CardInner(width)));
        if (!string.IsNullOrWhiteSpace(description))
            about.Controls.Add(MakeLabel(description!, CardInner(width), TextPrimary));
        foreach (string tip in BuildFieldTips(s))
            about.Controls.Add(MakeLabel("• " + tip, CardInner(width), TextPrimary));
        if (s.Zone == CellZone.Data)
        {
            string value = string.IsNullOrEmpty(s.CurrentValuePreview) ? Loc.Tip("guide.label.empty") : s.CurrentValuePreview!;
            about.Controls.Add(MakeLabel(Loc.Tip("guide.label.currentValue", value), CardInner(width), TextMuted));
        }
        AddRow(about);

        // ---- Rules & format ----
        if (s.Zone == CellZone.Data)
        {
            var rules = BuildRuleLines(s);
            if (rules.Count > 0)
            {
                var card = NewCard(width);
                card.Controls.Add(MakeSectionHeader(Loc.Tip("guide.section.rules"), CardInner(width)));
                foreach (string line in rules)
                    card.Controls.Add(MakeLabel("• " + line, CardInner(width), TextPrimary));
                AddRow(card);
            }
        }

        // ---- Table nature ----
        if (s.TableNature.HasValue)
        {
            var card = NewCard(width);
            card.Controls.Add(MakeSectionHeader(Loc.Tip("guide.section.table"), CardInner(width)));
            string natureKey = s.TableNature.Value == TableNature.Additive ? "additive" : "absolute";
            card.Controls.Add(MakeLabel(Loc.Tip($"guide.nature.{natureKey}.desc"), CardInner(width), TextPrimary));
            if (s.IsHistoryPreservingTable)
                card.Controls.Add(MakeLabel(Loc.Tip("guide.nature.history.desc"), CardInner(width), TextMuted));
            AddRow(card);
        }
    }

    // =========================================================================
    // 4. CONTENT SYNTHESIS (plain-language guidance from rule metadata)
    // =========================================================================
    private static List<string> BuildFieldTips(CellGuideSnapshot s)
    {
        var tips = new List<string>();
        if (s.Zone == CellZone.Header)
        {
            tips.Add(Loc.Tip("guide.tip.header.sort"));
            return tips;
        }
        if (s.Zone != CellZone.Data) return tips;

        if (s.IsStrictlyCalculated) tips.Add(Loc.Tip("guide.tip.calculated"));
        else if (s.IsReadOnly) tips.Add(Loc.Tip("guide.tip.readonly"));
        else if (s.HasRule)
        {
            if (s.IsRequired) tips.Add(Loc.Tip("guide.tip.required"));
            else if (s.IsWarningOnly) tips.Add(Loc.Tip("guide.tip.warningOnly"));
            else tips.Add(Loc.Tip("guide.tip.optional"));
        }

        if (!s.IsLocked && s.HasDropdown) tips.Add(Loc.Tip("guide.tip.dropdown"));
        if (!s.IsLocked && s.IsMultiSelect) tips.Add(Loc.Tip("guide.tip.multiselect"));
        if (s.SyncedFromSources.Count > 0) tips.Add(Loc.Tip("guide.tip.syncTarget", JoinHeaders(s.TableName, s.SyncedFromSources)));
        if (s.SyncTargets.Count > 0) tips.Add(Loc.Tip("guide.tip.syncSource", JoinHeaders(s.TableName, s.SyncTargets)));
        if (s.HasExcelFormula && !s.IsStrictlyCalculated) tips.Add(Loc.Tip("guide.tip.excelFormula"));
        return tips;
    }

    private static List<string> BuildRuleLines(CellGuideSnapshot s)
    {
        var lines = new List<string>();

        string? format = FormatDescription(s);
        if (format != null) lines.Add(Loc.Tip("guide.label.format", format));
        string? example = FormatExample(s);
        if (example != null && !s.IsLocked) lines.Add(Loc.Tip("guide.label.example", example));

        if (!string.IsNullOrWhiteSpace(s.DateMustBeGreaterThanHeader))
        {
            string refName = ColumnName(s.TableName, s.DateMustBeGreaterThanHeader!);
            lines.Add(Loc.Tip(s.DateMustBeGreaterThanAllowEqual ? "guide.tip.dateRuleEqual" : "guide.tip.dateRule", refName));
        }

        if (s.DefaultKindDescription is { } def && !s.IsLocked) lines.Add(def);
        if (!string.IsNullOrWhiteSpace(s.FormulaHint)) lines.Add(Loc.Tip("guide.label.formula", s.FormulaHint!));

        if (!string.IsNullOrWhiteSpace(s.ExcelValidationMessage))
        {
            string text = string.IsNullOrWhiteSpace(s.ExcelValidationTitle)
                ? s.ExcelValidationMessage!
                : $"{s.ExcelValidationTitle}: {s.ExcelValidationMessage}";
            lines.Add(Loc.Tip("guide.label.excelNote", text));
        }
        return lines;
    }

    private static string? FormatDescription(CellGuideSnapshot s)
    {
        string? key = s.FormatKind switch
        {
            ColumnFormatKind.Text => "guide.format.text",
            ColumnFormatKind.Integer => "guide.format.integer",
            ColumnFormatKind.Id => "guide.format.id",
            ColumnFormatKind.Decimal => "guide.format.decimal",
            ColumnFormatKind.Currency => "guide.format.currency",
            ColumnFormatKind.PricePerLiter => "guide.format.pricePerLiter",
            ColumnFormatKind.Liters => "guide.format.liters",
            ColumnFormatKind.Kilometers => "guide.format.kilometers",
            ColumnFormatKind.Hours => "guide.format.hours",
            ColumnFormatKind.Shifts => "guide.format.shifts",
            ColumnFormatKind.Years => "guide.format.years",
            ColumnFormatKind.Days => "guide.format.days",
            ColumnFormatKind.OdometerDynamic => "guide.format.odometer",
            ColumnFormatKind.CurrencyPerHour => "guide.format.currencyPerHour",
            ColumnFormatKind.WageDynamic => "guide.format.wage",
            ColumnFormatKind.Date => "guide.format.date",
            ColumnFormatKind.Time => "guide.format.time",
            _ => null
        };
        if (key == null && s.IsDateColumn) key = "guide.format.date";
        if (key == null && s.IsNumericColumn) key = "guide.format.number";
        return key == null ? null : Loc.Tip(key);
    }

    private static string? FormatExample(CellGuideSnapshot s)
    {
        string? key = s.FormatKind switch
        {
            ColumnFormatKind.Integer or ColumnFormatKind.Years or ColumnFormatKind.Days => "guide.example.integer",
            ColumnFormatKind.Decimal or ColumnFormatKind.Liters or ColumnFormatKind.Hours
                or ColumnFormatKind.Kilometers or ColumnFormatKind.Shifts => "guide.example.decimal",
            ColumnFormatKind.Currency or ColumnFormatKind.PricePerLiter
                or ColumnFormatKind.CurrencyPerHour or ColumnFormatKind.WageDynamic => "guide.example.currency",
            ColumnFormatKind.Date => "guide.example.date",
            ColumnFormatKind.Time => "guide.example.time",
            _ => null
        };
        if (key == null && s.IsDateColumn) key = "guide.example.date";
        return key == null ? null : Loc.Tip(key);
    }

    private static string ColumnName(string? tableName, string header)
        => Loc.ResolveColumnGuide(tableName, header).DisplayName ?? header;

    private static string JoinHeaders(string? tableName, IEnumerable<string> headers)
        => string.Join(", ", headers.Select(h => ColumnName(tableName, h)));

    // =========================================================================
    // 5. BUILDING BLOCKS
    // =========================================================================
    private void AddRow(Control control)
    {
        _content.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        _content.Controls.Add(control, 0, _content.RowStyles.Count - 1);
    }

    private static int CardInner(int width) => Math.Max(120, width - 26);

    private static FlowLayoutPanel NewCard(int width)
    {
        var card = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = CardBack,
            Padding = new Padding(10, 8, 10, 8),
            Margin = new Padding(0, 0, 0, 8),
            MinimumSize = new Size(width, 0),
            MaximumSize = new Size(width, 0)
        };
        card.Paint += PaintCardBorder;
        return card;
    }

    private static void PaintCardBorder(object? sender, PaintEventArgs e)
    {
        if (sender is not Control c) return;
        using var pen = new Pen(CardBorder);
        e.Graphics.DrawRectangle(pen, 0, 0, c.Width - 1, c.Height - 1);
    }

    private static void PaintTipCardBorder(object? sender, PaintEventArgs e)
    {
        if (sender is not Control c) return;
        using var pen = new Pen(TipCardBorder);
        e.Graphics.DrawRectangle(pen, 0, 0, c.Width - 1, c.Height - 1);
    }

    private static bool TryParseCellRow(string? address, out int row)
    {
        row = 0;
        if (string.IsNullOrWhiteSpace(address)) return false;
        int end = address.Length;
        int colon = address.IndexOf(':');
        if (colon >= 0) end = colon;

        int i = end - 1;
        while (i >= 0 && char.IsDigit(address[i])) i--;
        int start = i + 1;
        if (start >= end) return false;
        return int.TryParse(address.AsSpan(start, end - start), out row);
    }

    private (string TipKey, bool IsContextual) ResolveSpotlightTip()
    {
        if (!_preferGeneralTipForCurrentCell)
        {
            string? trigger = null;
            string? builtInKey = null;

            if (ServerEnvironment.IsOffline)
            {
                trigger = "offline";
                builtInKey = "guide.spotlight.ctx.offline";
            }
            else if (_snapshot.Zone == CellZone.Total)
            {
                trigger = "total";
                builtInKey = "guide.spotlight.gen.total_row_filter";
            }
            else if (_snapshot.Zone == CellZone.Header)
            {
                trigger = "header";
                builtInKey = "guide.spotlight.gen.header_sort";
            }
            else if (_snapshot.Zone == CellZone.OutsideTable
                && TryParseCellRow(_snapshot.CellAddress, out int outsideRow)
                && outsideRow is >= 1 and <= 6)
            {
                trigger = "date_filter";
                builtInKey = "guide.spotlight.gen.date_filter";
            }
            else if (_snapshot.Zone == CellZone.Data)
            {
                if (string.Equals(_snapshot.ColumnHeader, "Actual Project Cost", StringComparison.OrdinalIgnoreCase)
                    && (string.Equals(_snapshot.TableName, "GeneralProjectsTable", StringComparison.OrdinalIgnoreCase)
                        || string.Equals(_snapshot.TableName, "InternalProjectsTable", StringComparison.OrdinalIgnoreCase)))
                {
                    return ("guide.spotlight.ctx.dynamic_cost", true);
                }
                else if (_snapshot.IsMultiSelect && !_snapshot.IsLocked)
                {
                    trigger = "multiselect";
                    builtInKey = "guide.spotlight.ctx.multiselect";
                }
                else if (_snapshot.SyncedFromSources.Count > 0)
                {
                    trigger = "autosync";
                    builtInKey = "guide.spotlight.ctx.autosync";
                }
                else if (_snapshot.HasDropdown && !_snapshot.IsLocked)
                {
                    trigger = "dropdown";
                    builtInKey = "guide.spotlight.gen.dropdown";
                }
                else if (_snapshot.HasDefaultValue && !_snapshot.IsLocked && string.IsNullOrEmpty(_snapshot.CurrentValuePreview))
                {
                    trigger = "auto_defaults";
                    builtInKey = "guide.spotlight.gen.auto_defaults";
                }
                else if (_snapshot.IsHistoryPreservingTable)
                {
                    trigger = "history";
                    builtInKey = "guide.spotlight.ctx.history";
                }
                else if (_snapshot.FormatKind is ColumnFormatKind.Currency or ColumnFormatKind.CurrencyPerHour or ColumnFormatKind.WageDynamic)
                {
                    trigger = "currency";
                    builtInKey = "guide.spotlight.gen.total_row";
                }
                else if (_snapshot.IsLocked)
                {
                    trigger = "readonly";
                    builtInKey = null;
                }
                else if (_snapshot.IsRequired && string.IsNullOrEmpty(_snapshot.CurrentValuePreview))
                {
                    trigger = "required";
                    builtInKey = null;
                }
            }

            if (trigger != null)
            {
                var candidates = new List<string>(Loc.GetCustomContextTipKeys(trigger));
                if (builtInKey != null) candidates.Add(builtInKey);
                if (candidates.Count > 0)
                    return (candidates[Math.Abs(_tipIndex) % candidates.Count], true);
            }
        }

        var allTipKeys = Loc.GetAllGeneralTipKeys(GeneralTipIds);
        string generalKey = allTipKeys.Count > 0
            ? allTipKeys[Math.Abs(_tipIndex) % allTipKeys.Count]
            : $"guide.spotlight.gen.{GeneralTipIds[0]}";
        return (generalKey, false);
    }

    private Control BuildSpotlightTipCard(int width)
    {
        var (tipKey, isContextual) = ResolveSpotlightTip();

        var card = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = TipCardBack,
            Padding = new Padding(10, 7, 10, 7),
            Margin = new Padding(0, 0, 0, 8),
            MinimumSize = new Size(width, 0),
            MaximumSize = new Size(width, 0)
        };
        card.Paint += PaintTipCardBorder;

        var header = MakeLabel(
            Loc.Tip(isContextual ? "guide.spotlight.headerContext" : "guide.spotlight.headerTip").ToUpperInvariant(),
            CardInner(width), TipCardAccent, size: 8F, bold: true);
        header.Margin = new Padding(0, 0, 0, 3);
        card.Controls.Add(header);

        card.Controls.Add(MakeLabel(Loc.Tip(tipKey), CardInner(width), TextPrimary));

        var footer = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = false,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = TipCardBack,
            Margin = new Padding(0, 2, 0, 0),
            Padding = Padding.Empty
        };

        var nextLink = new LinkLabel
        {
            Text = Loc.Tip("guide.spotlight.next"),
            AutoSize = true,
            LinkColor = TipCardAccent,
            ActiveLinkColor = TipCardAccent,
            LinkBehavior = LinkBehavior.HoverUnderline,
            Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
            Margin = Padding.Empty,
            UseMnemonic = false
        };
        nextLink.LinkClicked += OnNextTipClicked;
        footer.Controls.Add(nextLink);

        var hideLink = new LinkLabel
        {
            Text = Loc.Tip("guide.spotlight.hide"),
            AutoSize = true,
            LinkColor = TextMuted,
            ActiveLinkColor = TextMuted,
            LinkBehavior = LinkBehavior.HoverUnderline,
            Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
            Margin = new Padding(10, 0, 0, 0),
            UseMnemonic = false
        };
        hideLink.LinkClicked += OnHideTipClicked;
        footer.Controls.Add(hideLink);

        card.Controls.Add(footer);
        return card;
    }

    private Control BuildShowTipLink(int width)
    {
        var showTip = new LinkLabel
        {
            Text = Loc.Tip("guide.spotlight.show"),
            AutoSize = true,
            MaximumSize = new Size(width, 0),
            LinkColor = TipCardAccent,
            ActiveLinkColor = TipCardAccent,
            LinkBehavior = LinkBehavior.HoverUnderline,
            Font = new Font("Segoe UI", 8.5F, FontStyle.Regular, GraphicsUnit.Point),
            Margin = new Padding(0, 0, 0, 6),
            UseMnemonic = false
        };
        showTip.LinkClicked += OnShowTipClicked;
        return showTip;
    }

    private void OnNextTipClicked(object? sender, LinkLabelLinkClickedEventArgs e)
    {
        MarkActive();
        _preferGeneralTipForCurrentCell = true;
        var allTipKeys = Loc.GetAllGeneralTipKeys(GeneralTipIds);
        if (allTipKeys.Count > 0)
            _tipIndex = (_tipIndex + 1) % allTipKeys.Count;
        try { BeginInvoke(new Action(Render)); } catch { }
    }

    private void OnHideTipClicked(object? sender, LinkLabelLinkClickedEventArgs e)
    {
        MarkActive();
        _tipDismissed = true;
        try { BeginInvoke(new Action(Render)); } catch { }
    }

    private void OnShowTipClicked(object? sender, LinkLabelLinkClickedEventArgs e)
    {
        MarkActive();
        _tipDismissed = false;
        try { BeginInvoke(new Action(Render)); } catch { }
    }

    private static Label MakeLabel(string text, int width, Color color, float size = 9F, bool bold = false, bool italic = false)
    {
        var style = (bold ? FontStyle.Bold : FontStyle.Regular) | (italic ? FontStyle.Italic : FontStyle.Regular);
        return new Label
        {
            Text = text,
            AutoSize = true,
            MaximumSize = new Size(width, 0),
            ForeColor = color,
            Font = new Font("Segoe UI", size, style, GraphicsUnit.Point),
            Margin = new Padding(0, 1, 0, 3),
            UseMnemonic = false
        };
    }

    private static Label MakeSectionHeader(string text, int width)
    {
        var label = MakeLabel(text.ToUpperInvariant(), width, Accent, size: 8F, bold: true);
        label.Margin = new Padding(0, 0, 0, 4);
        return label;
    }

    private Control BuildBannerCard(string text, int width)
    {
        var card = NewCard(width);
        card.BackColor = Color.FromArgb(232, 240, 254);
        card.Controls.Add(MakeLabel(text, CardInner(width), Color.FromArgb(23, 78, 166)));
        return card;
    }

    private Control BuildContextStrip(CellGuideSnapshot s, int width)
    {
        var strip = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = PaneBack,
            Margin = new Padding(0, 0, 0, 6),
            MaximumSize = new Size(width, 0)
        };

        var parts = new List<string>();
        if (!string.IsNullOrWhiteSpace(s.SheetName)) parts.Add(s.SheetName);
        if (!string.IsNullOrWhiteSpace(s.TableName)) parts.Add(Loc.TableDisplayName(s.TableName));
        if (!string.IsNullOrWhiteSpace(s.CellAddress)) parts.Add(s.CellAddress);
        string text = parts.Count == 0 ? Loc.Tip("guide.pane.title") : string.Join("  •  ", parts);
        strip.Controls.Add(MakeLabel(text, width, TextMuted, bold: true));

        if (s.TableNature.HasValue && s.Zone != CellZone.OutsideTable && s.Zone != CellZone.None)
        {
            string natureKey = s.TableNature.Value == TableNature.Additive ? "additive" : "absolute";
            var colors = s.TableNature.Value == TableNature.Additive ? BadgeSync : BadgeDropdown;
            strip.Controls.Add(MakeBadge(Loc.Tip($"guide.nature.{natureKey}.name"), colors));
        }
        return strip;
    }

    private static Control BuildBadges(CellGuideSnapshot s, int width)
    {
        var panel = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            WrapContents = true,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            BackColor = CardBack,
            Margin = new Padding(0, 4, 0, 0),
            MaximumSize = new Size(width, 0)
        };

        if (s.IsStrictlyCalculated) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.calculated"), BadgeCalculated));
        else if (s.IsReadOnly) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.readonly"), BadgeLocked));
        else if (s.HasRule)
        {
            if (s.IsRequired) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.required"), BadgeRequired));
            else if (s.IsWarningOnly) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.warning"), BadgeWarning));
            else panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.optional"), BadgeOptional));
        }

        if (!s.IsLocked && s.IsMultiSelect) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.multiselect"), BadgeDropdown));
        else if (!s.IsLocked && s.HasDropdown) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.dropdown"), BadgeDropdown));
        if (s.SyncedFromSources.Count > 0) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.autosync"), BadgeSync));
        if (s.HasDefaultValue && !s.IsLocked) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.default"), BadgeSync));
        if (s.IsDateColumn) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.date"), BadgeType));
        else if (s.IsNumericColumn) panel.Controls.Add(MakeBadge(Loc.Tip("guide.badge.number"), BadgeType));

        return panel;
    }

    private static Label MakeBadge(string text, (Color Back, Color Fore) colors)
    {
        return new Label
        {
            Text = text,
            AutoSize = true,
            BackColor = colors.Back,
            ForeColor = colors.Fore,
            Font = new Font("Segoe UI", 8F, FontStyle.Bold, GraphicsUnit.Point),
            Padding = new Padding(6, 2, 6, 2),
            Margin = new Padding(0, 0, 4, 4),
            UseMnemonic = false
        };
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing)
        {
            try
            {
                _fadeTimer.Stop();
                _fadeTimer.Tick -= OnFadeTimerTick;
                _fadeTimer.Dispose();
            }
            catch { }
            try { _scrollHost.Resize -= OnHostResize; } catch { }
            try { _scrollHost.MouseEnter -= OnScrollHostMouseEnter; } catch { }
            try { RestoreOpacity(); } catch { }
            _lastFloatingRootHwnd = IntPtr.Zero;
        }
        base.Dispose(disposing);
    }
}
