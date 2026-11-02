namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Environment;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;

public static class Arrow
{
    private static ArrowHostForm? _arrowHost;
    private static Button? _arrowButton;
    private static PopupHostForm? _popupHost;
    private static ListBox? _listBox;
    private static dynamic? _activeCell;
    private static Action<string>? _onItemSelected;
    private static bool _isEventsHooked = false;

    private class DropdownItem
    {
        public string Display { get; set; } = string.Empty;
        public string Raw { get; set; } = string.Empty;
        public int Section { get; set; } = 0;
        public bool IsValid { get; set; } = true;
    }

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr")]
    private static extern IntPtr SetWindowLongPtr64(IntPtr hWnd, int nIndex, IntPtr dwNewLong);

    [DllImport("user32.dll", EntryPoint = "SetWindowLong")]
    private static extern int SetWindowLong32(IntPtr hWnd, int nIndex, int dwNewLong);

    private static IntPtr SetWindowLongPtr(IntPtr hWnd, int nIndex, IntPtr dwNewLong)
    {
        bool is64Bit = ExcelRegistry.OfficeBitness switch
        {
            Bitness.Win64 => true,
            Bitness.Win32 => false,
            _ => IntPtr.Size == 8
        };

        if (is64Bit) return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
        return new IntPtr(SetWindowLong32(hWnd, nIndex, dwNewLong.ToInt32()));
    }

    public static IntPtr GetExcelHwnd(dynamic xlApp)
    {
        try
        {
            object rawHwnd = xlApp.Hwnd;
            if (rawHwnd is long l) return new IntPtr(l);
            if (rawHwnd is int i) return new IntPtr(i);
            if (rawHwnd is IntPtr p) return p;
            return new IntPtr(Convert.ToInt64(rawHwnd));
        }
        catch { return IntPtr.Zero; }
    }

    private const int GWL_HWNDPARENT = -8;
    private static readonly IntPtr HWND_TOP = new(0);
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;

    private class ArrowHostForm : Form
    {
        public ArrowHostForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x00000080;
                cp.ExStyle |= 0x08000000;
                return cp;
            }
        }
        protected override bool ShowWithoutActivation => true;
    }

    private class PopupHostForm : Form
    {
        public PopupHostForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.Manual;
            AutoScaleMode = AutoScaleMode.None;
        }
        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= 0x00000080;
                cp.ExStyle |= 0x08000000;
                return cp;
            }
        }
        protected override bool ShowWithoutActivation => true;
    }

    static Arrow()
    {
        InitializeArrowHost();
        InitializePopupHost();
    }

    private static void InitializeArrowHost()
    {
        _arrowHost = new ArrowHostForm { BackColor = Color.FromArgb(218, 220, 224) };
        _arrowButton = new Button
        {
            Dock = DockStyle.Fill,
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(242, 244, 246),
            Cursor = Cursors.Hand,
            Margin = Padding.Empty,
            Padding = Padding.Empty
        };
        _arrowButton.FlatAppearance.BorderSize = 1;
        _arrowButton.FlatAppearance.BorderColor = Color.FromArgb(170, 175, 182);
        _arrowButton.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            int cx = _arrowButton.ClientSize.Width / 2;
            int cy = _arrowButton.ClientSize.Height / 2;
            Point[] triangle =
            {
                new(cx - 4, cy - 2),
                new(cx + 4, cy - 2),
                new(cx, cy + 3)
            };
            using var brush = new SolidBrush(Color.FromArgb(70, 70, 70));
            e.Graphics.FillPolygon(brush, triangle);
        };
        _arrowButton.Click += (_, _) => ToggleDropdown();
        _arrowHost.Controls.Add(_arrowButton);
    }

    private static void InitializePopupHost()
    {
        _popupHost = new PopupHostForm
        {
            BackColor = Color.FromArgb(160, 165, 175),
            Padding = new Padding(1)
        };
        _listBox = new ListBox
        {
            Dock = DockStyle.Fill,
            BorderStyle = BorderStyle.None,
            Font = new Font("Segoe UI", 9.5f),
            DrawMode = DrawMode.OwnerDrawFixed,
            ItemHeight = 22,
            IntegralHeight = false
        };
        _listBox.DrawItem += OnListBoxDrawItem;
        _listBox.Click += (s, e) =>
        {
            if (e is MouseEventArgs me)
            {
                int index = _listBox.IndexFromPoint(me.Location);
                if (index != ListBox.NoMatches && index >= 0)
                {
                    CommitSelection();
                }
            }
            else
            {
                CommitSelection();
            }
        };
        _listBox.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter) CommitSelection();
            else if (e.KeyCode == Keys.Escape) DismissAndReturnFocus();
        };
        _popupHost.Controls.Add(_listBox);
        _popupHost.Deactivate += (_, _) => DismissAndReturnFocus();
    }

    public static void EnsureExcelEventsHooked()
    {
        if (_isEventsHooked) return;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return;
            xlApp.WindowDeactivate += new Action<dynamic, dynamic>((wb, wn) => HideAll());
            xlApp.SheetDeactivate += new Action<dynamic>(ws => HideAll());
            _isEventsHooked = true;
            if (_activeCell != null && _arrowHost != null && _arrowHost.Visible)
            {
                ShowArrowForCell(_activeCell);
            }
        }
        catch { }
    }

    private static void OnListBoxDrawItem(object? sender, DrawItemEventArgs e)
    {
        if (e.Index < 0 || e.Index >= _listBox!.Items.Count) return;
        object rawObj = _listBox.Items[e.Index];
        if (rawObj is not DropdownItem item)
        {
            e.DrawBackground();
            using var defBrush = new SolidBrush(Color.Gray);
            e.Graphics.DrawString(rawObj.ToString(), e.Font!, defBrush, e.Bounds.X + 6, e.Bounds.Y + 3);
            return;
        }

        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        Color backColor = !item.IsValid
            ? (isSelected ? Color.FromArgb(242, 242, 242) : Color.White)
            : (isSelected ? Color.FromArgb(217, 235, 255) : Color.White);
        Color foreColor = !item.IsValid ? Color.FromArgb(170, 170, 175) : Color.Black;

        using var bgBrush = new SolidBrush(backColor);
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        using var fgBrush = new SolidBrush(foreColor);
        var textBounds = new Rectangle(e.Bounds.X + 6, e.Bounds.Y + 2, e.Bounds.Width - 8, e.Bounds.Height - 4);
        var sf = new StringFormat { LineAlignment = StringAlignment.Center };
        e.Graphics.DrawString(item.Display, e.Font!, fgBrush, textBounds, sf);
    }

    public static void ShowArrowForCell(dynamic cell)
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            EnsureExcelEventsHooked();
            if (xlApp.WindowState == -4140)
            {
                HideAll();
                return;
            }

            _activeCell = cell.Cells[1, 1];
            dynamic window = xlApp.ActiveWindow;
            IntPtr excelHwnd = GetExcelHwnd(xlApp);
            if (excelHwnd != IntPtr.Zero)
            {
                SetWindowLongPtr(_arrowHost!.Handle, GWL_HWNDPARENT, excelHwnd);
                SetWindowLongPtr(_popupHost!.Handle, GWL_HWNDPARENT, excelHwnd);
            }

            int targetRow = (int)_activeCell.Row;
            int targetCol = (int)_activeCell.Column;
            dynamic targetPane = ResolveTargetPane(window, targetRow, targetCol);

            int cellTopPx = targetPane.PointsToScreenPixelsY((double)_activeCell.Top);
            int cellRightPx = targetPane.PointsToScreenPixelsX((double)_activeCell.Left + (double)_activeCell.Width);
            int cellBottomPx = targetPane.PointsToScreenPixelsY((double)_activeCell.Top + (double)_activeCell.Height);

            if (cellRightPx <= 0 || cellBottomPx <= 0)
            {
                HideAll();
                return;
            }

            int cellHeightPx = cellBottomPx - cellTopPx;
            int arrowWidthPx = 18;
            int arrowHeightPx = cellHeightPx;
            int arrowXPx = cellRightPx + 1;
            int arrowYPx = cellTopPx;

            if (!_arrowHost!.Visible) _arrowHost.Show();
            SetWindowPos(_arrowHost.Handle, HWND_TOP, arrowXPx, arrowYPx, arrowWidthPx, arrowHeightPx, SWP_SHOWWINDOW | SWP_NOACTIVATE);
            HidePopup();
        }
        catch
        {
            HideAll();
        }
    }

    public static void HideAll()
    {
        HidePopup();
        if (_arrowHost != null && _arrowHost.Visible) _arrowHost.Hide();
        _activeCell = null;
    }

    public static void HidePopup()
    {
        if (_popupHost != null && _popupHost.Visible) _popupHost.Hide();
    }

    public static void DismissAndReturnFocus()
    {
        ReleaseCapture();
        HidePopup();
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            IntPtr excelHwnd = GetExcelHwnd(xlApp);
            if (excelHwnd != IntPtr.Zero) SetForegroundWindow(excelHwnd);
        }
        catch { }
    }

    public static void ToggleDropdown()
    {
        if (_popupHost!.Visible)
        {
            DismissAndReturnFocus();
            return;
        }
        if (_activeCell == null) return;
        PopulateAndOpenMenu(_activeCell);
    }

    private static void PopulateAndOpenMenu(dynamic singleCell)
    {
        var swTotal = Stopwatch.StartNew();
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic wb = xlApp.ActiveWorkbook;
            dynamic sheet = singleCell.Worksheet;
            int row = (int)singleCell.Row;
            int col = (int)singleCell.Column;

            string inCellText = string.Empty;
            try
            {
                inCellText = singleCell.Text?.ToString()?.Trim() ?? string.Empty;
                if (string.IsNullOrWhiteSpace(inCellText))
                    inCellText = singleCell.Value2?.ToString()?.Trim() ?? string.Empty;
            }
            catch { }

            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }

            var rule = RuleRegistry.GetRuleForCell(singleCell);
            var processedItems = new List<DropdownItem>();

            // =========================================================================
            // PATH A: Unified C# Attribute Rule Engine (Zero COM formula evaluations)
            // =========================================================================
            if (rule != null && rule.HasDropdown)
            {
                var evaluatedOptions = RuleEvaluator.GetEvaluatedDropdownOptions(wb, table, row, rule);
                foreach (var opt in evaluatedOptions)
                {
                    processedItems.Add(new DropdownItem
                    {
                        Display = opt.Display,
                        Raw = opt.Raw,
                        Section = opt.Section,
                        IsValid = opt.IsValid
                    });
                }

                _onItemSelected = chosenDisplay =>
                {
                    var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                    if (targetItem == null) return;

                    bool origEvents = true;
                    try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                    try
                    {
                        WriteSelectedValueToCell(singleCell, targetItem, rule, inCellText);
                    }
                    finally
                    {
                        try { xlApp.EnableEvents = origEvents; } catch { }
                    }

                    // Immediately run self-healing row pass (executes [ColumnSync], [ColumnDefault], [ColumnStateGuard])
                    if (table != null)
                    {
                        RuleEvaluator.EvaluateAndHealRow(sheet, table, row, col);
                    }
                };
            }
            // =========================================================================
            // PATH B: Legacy Fallback (For top sheet parameter cells outside tables)
            // =========================================================================
            else
            {
                var config = Metadata.ReadConfig(singleCell);
                if (!string.IsNullOrWhiteSpace(config.Source))
                {
                    var items = RuleDataCache.GetSourceItems(wb, config.Source, preserveDuplicates: false);
                    foreach (var itm in items)
                    {
                        string disp = FormatIfDate(itm);
                        processedItems.Add(new DropdownItem { Display = disp, Raw = itm, Section = 0, IsValid = true });
                    }
                }

                _onItemSelected = chosenDisplay =>
                {
                    var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                    if (targetItem == null) return;

                    bool origEvents = true;
                    try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                    try
                    {
                        singleCell.Value2 = targetItem.Display;
                    }
                    finally
                    {
                        try { xlApp.EnableEvents = origEvents; } catch { }
                    }

                    if (!string.IsNullOrWhiteSpace(config.SyncLeft) && !string.IsNullOrEmpty(targetItem.Raw))
                    {
                        SyncEngine.ExecuteSync(singleCell, config.SyncLeft, targetItem.Raw, targetItem.Display);
                    }
                };
            }

            // =========================================================================
            // RENDER & SORT MENU ITEMS
            // =========================================================================
            _listBox!.Items.Clear();
            if (processedItems.Count > 0)
            {
                var validItems = processedItems.Where(i => i.IsValid);
                var invalidItems = processedItems.Where(i => !i.IsValid);

                var sortedValid = validItems.OrderBy(i => i.Section).ThenBy(i => i.Display);
                var sortedInvalid = invalidItems.OrderBy(i => i.Section).ThenBy(i => i.Display);

                if (!string.IsNullOrWhiteSpace(inCellText) && (rule == null || !rule.IsMultiSelect))
                {
                    string q = inCellText.Trim();
                    Func<DropdownItem, int> ranker = i =>
                    {
                        if (i.Display.Equals(q, StringComparison.OrdinalIgnoreCase)) return 0;
                        if (i.Display.StartsWith(q, StringComparison.OrdinalIgnoreCase)) return 1;
                        if (i.Display.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) return 2;
                        return 3;
                    };
                    sortedValid = validItems.OrderBy(ranker).ThenBy(i => i.Section).ThenBy(i => i.Display);
                    sortedInvalid = invalidItems.OrderBy(ranker).ThenBy(i => i.Section).ThenBy(i => i.Display);
                }

                foreach (var item in sortedValid) _listBox.Items.Add(item);
                foreach (var item in sortedInvalid) _listBox.Items.Add(item);

                int firstValid = _listBox.Items.Cast<object>().ToList().FindIndex(o => o is DropdownItem d && d.IsValid);
                if (firstValid >= 0) _listBox.SelectedIndex = firstValid;
            }
            else
            {
                _listBox.Items.Add(new DropdownItem { Display = "(No options available)", IsValid = false });
            }

            dynamic window = xlApp.ActiveWindow;
            dynamic targetPane = ResolveTargetPane(window, (int)singleCell.Row, (int)singleCell.Column);

            int cellLeftPx = targetPane.PointsToScreenPixelsX((double)singleCell.Left);
            int cellRightPx = targetPane.PointsToScreenPixelsX((double)singleCell.Left + (double)singleCell.Width);
            int cellBottomPx = targetPane.PointsToScreenPixelsY((double)singleCell.Top + (double)singleCell.Height);

            int popupWidthPx = Math.Max((cellRightPx - cellLeftPx) + 19, 160);
            int calculatedHeightPx = Math.Min(240, Math.Max(38, (_listBox.Items.Count * _listBox.ItemHeight) + 4));

            if (!_popupHost!.Visible) _popupHost.Show();
            SetWindowPos(_popupHost.Handle, HWND_TOP, cellLeftPx, cellBottomPx, popupWidthPx, calculatedHeightPx, SWP_SHOWWINDOW | SWP_NOACTIVATE);

            swTotal.Stop();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Arrow.PopulateAndOpenMenu Error] {ex.Message}");
        }
    }

    private static void WriteSelectedValueToCell(
        dynamic singleCell,
        DropdownItem targetItem,
        PropertyRuleDescriptor rule,
        string existingCellText)
    {
        if (rule.IsMultiSelect && !string.IsNullOrWhiteSpace(existingCellText))
        {
            var tokens = existingCellText
                .Split(new[] { ',' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(t => t.Trim())
                .Where(t => !string.IsNullOrEmpty(t))
                .ToList();

            int existingIdx = tokens.FindIndex(t => t.Equals(targetItem.Display, StringComparison.OrdinalIgnoreCase));
            if (existingIdx >= 0)
                tokens.RemoveAt(existingIdx);
            else
                tokens.Add(targetItem.Display);

            singleCell.Value2 = string.Join(", ", tokens);
            return;
        }

        if (rule.IsDateColumn && RuleDataCache.TryParseDate(targetItem.Raw, out DateOnly dOnly))
        {
            singleCell.Value2 = dOnly.ToDateTime(TimeOnly.MinValue).ToOADate();
            return;
        }

        if (rule.IsNumericColumn &&
            double.TryParse(targetItem.Raw, NumberStyles.Any, CultureInfo.InvariantCulture, out double dNum))
        {
            singleCell.Value2 = dNum;
            return;
        }

        singleCell.Value2 = targetItem.Display;
    }

    private static dynamic ResolveTargetPane(dynamic window, int targetRow, int targetCol)
    {
        dynamic? targetPane = null;
        int panesCount = 0;
        try { panesCount = (int)window.Panes.Count; } catch { }

        if (panesCount > 1)
        {
            for (int i = 1; i <= panesCount; i++)
            {
                try
                {
                    dynamic p = window.Panes[i];
                    dynamic vr = p.VisibleRange;
                    int r1 = (int)vr.Row;
                    int r2 = r1 + (int)vr.Rows.Count - 1;
                    int c1 = (int)vr.Column;
                    int c2 = c1 + (int)vr.Columns.Count - 1;
                    if (targetRow >= r1 && targetRow <= r2 && targetCol >= c1 && targetCol <= c2)
                    {
                        targetPane = p;
                        break;
                    }
                }
                catch { }
            }
        }
        return targetPane ?? window.ActivePane;
    }

    private static readonly string[] RecognizedDatePatterns =
    {
        "yyyy-MM-dd", "yyyy/MM/dd", "dd/MM/yyyy", "d/M/yyyy", "dd-MM-yyyy", "d-M-yyyy"
    };

    public static string FormatIfDate(string input)
    {
        if (string.IsNullOrWhiteSpace(input)) return input;
        if (DateTime.TryParseExact(input.Trim(), RecognizedDatePatterns, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime exactDt))
        {
            return exactDt.ToString("dd/MM/yyyy");
        }
        if (double.TryParse(input, NumberStyles.Any, CultureInfo.InvariantCulture, out double d))
        {
            if (d >= 30000 && d <= 80000)
            {
                return DateTime.FromOADate(d).ToString("dd/MM/yyyy");
            }
        }
        return input;
    }

    private static void CommitSelection()
    {
        if (_listBox!.SelectedItem is DropdownItem item && _onItemSelected != null)
        {
            _onItemSelected(item.Display);
        }
        DismissAndReturnFocus();
    }

    public static void Teardown()
    {
        try
        {
            HideAll();
            _arrowHost?.Dispose();
            _arrowHost = null;
            _popupHost?.Dispose();
            _popupHost = null;
        }
        catch { }
    }
}