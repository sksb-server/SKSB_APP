namespace SKSB_App.Core.DropDown;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.ExcelBitManager;

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
        bool is64Bit = Context.OfficeBitness switch
        {
            Bitness.Win64 => true,
            Bitness.Win32 => false,
            _ => IntPtr.Size == 8
        };

        if (is64Bit)
        {
            return SetWindowLongPtr64(hWnd, nIndex, dwNewLong);
        }
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
        catch
        {
            return IntPtr.Zero;
        }
    }

    private const int GWL_HWNDPARENT = -8;
    // Z-Order: Use HWND_NOTOPMOST so the windows stay linked strictly to Excel's Z-order level
    private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
    private static readonly IntPtr HWND_TOP = new IntPtr(0);
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
                cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW: Excluded from Alt+Tab
                cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE: Does not steal activation
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
                cp.ExStyle |= 0x00000080; // WS_EX_TOOLWINDOW: Excluded from Alt+Tab
                cp.ExStyle |= 0x08000000; // WS_EX_NOACTIVATE: Does not steal activation
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

            Point[] triangle = new Point[]
            {
                new Point(cx - 4, cy - 2),
                new Point(cx + 4, cy - 2),
                new Point(cx, cy + 3)
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
        _listBox.Click += (_, _) => CommitSelection();
        _listBox.KeyDown += (s, e) =>
        {
            if (e.KeyCode == Keys.Enter) CommitSelection();
            else if (e.KeyCode == Keys.Escape) DismissAndReturnFocus();
        };

        _popupHost.Controls.Add(_listBox);
        _popupHost.Deactivate += (_, _) => DismissAndReturnFocus();
    }

    /// <summary>
    /// Explicitly attaches event hooks to Excel to auto-hide the controls when Excel deactivates.
    /// </summary>
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

            if (xlApp.WindowState == -4140 /* xlMinimized */)
            {
                HideAll();
                return;
            }

            _activeCell = cell.Cells[1, 1];
            dynamic window = xlApp.ActiveWindow;

            IntPtr excelHwnd = GetExcelHwnd(xlApp);
            if (excelHwnd != IntPtr.Zero)
            {
                // Assign Excel main window as owner so the arrow/popup follows Excel's Z-order directly
                SetWindowLongPtr(_arrowHost!.Handle, GWL_HWNDPARENT, excelHwnd);
                SetWindowLongPtr(_popupHost!.Handle, GWL_HWNDPARENT, excelHwnd);
            }

            dynamic targetPane = window.ActivePane;
            try
            {
                foreach (dynamic p in window.Panes)
                {
                    dynamic vr = p.VisibleRange;
                    int r1 = vr.Row;
                    int r2 = r1 + vr.Rows.Count - 1;
                    int c1 = vr.Column;
                    int c2 = c1 + vr.Columns.Count - 1;

                    if (_activeCell.Row >= r1 && _activeCell.Row <= r2 &&
                        _activeCell.Column >= c1 && _activeCell.Column <= c2)
                    {
                        targetPane = p;
                        break;
                    }
                }
            }
            catch { }

            double offsetX = (double)_activeCell.Left - (double)targetPane.VisibleRange.Left;
            double offsetY = (double)_activeCell.Top - (double)targetPane.VisibleRange.Top;

            int cellLeftPx = targetPane.PointsToScreenPixelsX(offsetX);
            int cellTopPx = targetPane.PointsToScreenPixelsY(offsetY);
            int cellRightPx = targetPane.PointsToScreenPixelsX(offsetX + (double)_activeCell.Width);
            int cellBottomPx = targetPane.PointsToScreenPixelsY(offsetY + (double)_activeCell.Height);

            int cellWidthPx = cellRightPx - cellLeftPx;
            int cellHeightPx = cellBottomPx - cellTopPx;

            int arrowWidthPx = Math.Min(17, cellWidthPx);
            int arrowHeightPx = cellHeightPx;
            int arrowXPx = cellRightPx - arrowWidthPx;
            int arrowYPx = cellTopPx;

            if (!_arrowHost.Visible)
            {
                _arrowHost.Show();
            }

            // Placed at HWND_TOP relative to its owner (Excel) - not system-wide HWND_TOPMOST
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
        if (_arrowHost != null && _arrowHost.Visible)
        {
            _arrowHost.Hide();
        }
        _activeCell = null;
    }

    public static void HidePopup()
    {
        if (_popupHost != null && _popupHost.Visible)
        {
            _popupHost.Hide();
        }
    }

    public static void DismissAndReturnFocus()
    {
        ReleaseCapture();
        HidePopup();

        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            IntPtr excelHwnd = GetExcelHwnd(xlApp);
            if (excelHwnd != IntPtr.Zero)
            {
                SetForegroundWindow(excelHwnd);
            }
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
        try
        {
            var config = Metadata.ReadConfig(singleCell);
            dynamic xlApp = ExcelDnaUtil.Application;
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

            string leftVal = string.Empty;
            string rightVal = string.Empty;
            try { leftVal = sheet.Cells[row, col - 1].Value2?.ToString()?.Trim() ?? string.Empty; } catch { }
            try { rightVal = sheet.Cells[row, col + 1].Value2?.ToString()?.Trim() ?? string.Empty; } catch { }

            var baseItems = new List<string>();

            // 1. Source Evaluation
            if (!string.IsNullOrWhiteSpace(config.Source))
            {
                dynamic? rawSource = null;
                string cleanSource = config.Source.Trim();
                if (cleanSource.StartsWith("=", StringComparison.Ordinal))
                    cleanSource = cleanSource.Substring(1).Trim();

                if (cleanSource.Contains("[") && cleanSource.Contains("]"))
                {
                    try { rawSource = xlApp.Range[cleanSource]; } catch { }
                }

                if (rawSource == null)
                {
                    try { rawSource = xlApp.Evaluate(cleanSource); } catch { }
                }

                if (rawSource != null)
                {
                    try
                    {
                        object val2 = rawSource.Value2;
                        if (val2 != null) rawSource = val2;
                    }
                    catch { }
                }

                if (rawSource is object[,] matrix)
                {
                    int rCount = matrix.GetLength(0);
                    int cCount = matrix.GetLength(1);
                    for (int r = 1; r <= rCount; r++)
                    {
                        for (int c = 1; c <= cCount; c++)
                        {
                            string val = matrix[r, c]?.ToString()?.Trim() ?? string.Empty;
                            bool isComError = long.TryParse(val, out long errNum) && errNum < -2000000000;
                            if (!string.IsNullOrEmpty(val) && !val.StartsWith("#") && !isComError && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                                baseItems.Add(val);
                        }
                    }
                }
                else if (rawSource is object[] vector)
                {
                    foreach (var item in vector)
                    {
                        string val = item?.ToString()?.Trim() ?? string.Empty;
                        bool isComError = long.TryParse(val, out long errNum) && errNum < -2000000000;
                        if (!string.IsNullOrEmpty(val) && !val.StartsWith("#") && !isComError && !val.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                            baseItems.Add(val);
                    }
                }
                else if (rawSource != null)
                {
                    string s = rawSource.ToString()?.Trim() ?? string.Empty;
                    bool isComError = long.TryParse(s, out long errNum) && errNum < -2000000000;
                    if (!string.IsNullOrWhiteSpace(s) && !s.StartsWith("#") && !isComError && !s.Equals("System.__ComObject", StringComparison.OrdinalIgnoreCase))
                    {
                        baseItems.Add(s);
                    }
                }
            }

            // 2. Transform & Filter
            var processedItems = new List<DropdownItem>();

            foreach (string rawItem in baseItems)
            {
                string safeItem = rawItem.Replace("\"", "\"\"");
                bool isNumeric = double.TryParse(rawItem, out _);
                string itemToken = isNumeric ? rawItem : $"\"{safeItem}\"";
                string leftToken = double.TryParse(leftVal, out _) ? leftVal : $"\"{leftVal.Replace("\"", "\"\"")}\"";
                string rightToken = double.TryParse(rightVal, out _) ? rightVal : $"\"{rightVal.Replace("\"", "\"\"")}\"";

                string displayLabel = rawItem;
                if (!string.IsNullOrWhiteSpace(config.Transform))
                {
                    string transformExpr = config.Transform
                        .Replace("@item", itemToken)
                        .Replace("@left", leftToken)
                        .Replace("@right", rightToken);

                    object? transformed = Metadata.EvaluateFormulaSafely(xlApp, transformExpr);
                    if (transformed != null && !(transformed is ExcelError))
                    {
                        string outStr = transformed.ToString()?.Trim() ?? string.Empty;
                        if (!string.IsNullOrEmpty(outStr)) displayLabel = outStr;
                    }
                }

                bool isValid = true;
                if (!string.IsNullOrWhiteSpace(config.Filter))
                {
                    string safeDisplay = displayLabel.Replace("\"", "\"\"");
                    string filterExpr = config.Filter
                        .Replace("@item", itemToken)
                        .Replace("@display", $"\"{safeDisplay}\"")
                        .Replace("@left", leftToken)
                        .Replace("@right", rightToken);

                    object? filterResult = Metadata.EvaluateFormulaSafely(xlApp, filterExpr);
                    if (filterResult is bool b)
                    {
                        isValid = b;
                    }
                    else if (filterResult != null)
                    {
                        string sRes = filterResult.ToString()!.Trim();
                        isValid = sRes.Equals("TRUE", StringComparison.OrdinalIgnoreCase) || sRes.Equals("1");
                    }
                    else
                    {
                        isValid = false;
                    }
                }

                processedItems.Add(new DropdownItem
                {
                    Display = displayLabel,
                    Raw = rawItem,
                    IsValid = isValid
                });
            }

            // 3. Selection Commit & Sync Writeback
            _onItemSelected = (chosenDisplay) =>
            {
                var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                if (targetItem == null) return;

                singleCell.Value2 = targetItem.Display;

                if (!string.IsNullOrWhiteSpace(config.SyncLeft) && !string.IsNullOrEmpty(targetItem.Raw))
                {
                    bool isNum = double.TryParse(targetItem.Raw, out _);
                    string rawToken = isNum ? targetItem.Raw : $"\"{targetItem.Raw.Replace("\"", "\"\"")}\"";
                    string displayToken = $"\"{targetItem.Display.Replace("\"", "\"\"")}\"";

                    int targetCol;
                    if (config.SyncLeft.Contains("@to_right", StringComparison.OrdinalIgnoreCase) ||
                        config.SyncLeft.Contains("@right", StringComparison.OrdinalIgnoreCase))
                    {
                        targetCol = col + 1;
                    }
                    else if (config.SyncLeft.Contains("@to_left", StringComparison.OrdinalIgnoreCase) ||
                             config.SyncLeft.Contains("@left", StringComparison.OrdinalIgnoreCase))
                    {
                        targetCol = col - 1;
                    }
                    else
                    {
                        string header = string.Empty;
                        for (int hRow = 1; hRow <= 3; hRow++)
                        {
                            try { header = sheet.Cells[hRow, col].Value2?.ToString()?.Trim() ?? string.Empty; } catch { }
                            if (!string.IsNullOrEmpty(header)) break;
                        }
                        targetCol = header.IndexOf("ID", StringComparison.OrdinalIgnoreCase) >= 0 ? col + 1 : col - 1;
                    }

                    string syncExpr = config.SyncLeft
                        .Replace("@to_right", "", StringComparison.OrdinalIgnoreCase)
                        .Replace("@to_left", "", StringComparison.OrdinalIgnoreCase)
                        .Replace("@right", "", StringComparison.OrdinalIgnoreCase)
                        .Replace("@left", "", StringComparison.OrdinalIgnoreCase)
                        .Replace("@item", rawToken)
                        .Replace("@display", displayToken)
                        .Trim();

                    if (syncExpr.StartsWith("=", StringComparison.Ordinal))
                        syncExpr = syncExpr.Substring(1).Trim();

                    object? resolvedVal = Metadata.EvaluateFormulaSafely(xlApp, syncExpr);
                    if (resolvedVal != null && !(resolvedVal is ExcelError))
                    {
                        string resStr = resolvedVal.ToString() ?? string.Empty;
                        bool isComError = long.TryParse(resStr, out long valNum) && valNum < -2000000000;
                        if (!resStr.StartsWith("#") && !isComError)
                        {
                            sheet.Cells[row, targetCol].Value2 = resolvedVal;
                        }
                    }
                }
            };

            // 4. Populate ListBox
            _listBox!.Items.Clear();
            if (processedItems.Count > 0)
            {
                var validItems = processedItems.Where(i => i.IsValid);
                var invalidItems = processedItems.Where(i => !i.IsValid);

                IEnumerable<DropdownItem> sortedValid;
                IEnumerable<DropdownItem> sortedInvalid;

                if (!string.IsNullOrWhiteSpace(inCellText))
                {
                    string q = inCellText.Trim();
                    Func<DropdownItem, int> ranker = i =>
                    {
                        if (i.Display.Equals(q, StringComparison.OrdinalIgnoreCase)) return 0;
                        if (i.Display.StartsWith(q, StringComparison.OrdinalIgnoreCase)) return 1;
                        var words = i.Display.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries);
                        if (words.Any(w => w.StartsWith(q, StringComparison.OrdinalIgnoreCase))) return 2;
                        if (i.Display.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) return 3;
                        return 4;
                    };
                    sortedValid = validItems.OrderBy(ranker).ThenBy(i => i.Display);
                    sortedInvalid = invalidItems.OrderBy(ranker).ThenBy(i => i.Display);
                }
                else
                {
                    sortedValid = validItems.OrderBy(i => i.Display);
                    sortedInvalid = invalidItems.OrderBy(i => i.Display);
                }

                foreach (var item in sortedValid) _listBox.Items.Add(item);
                foreach (var item in invalidItems) _listBox.Items.Add(item);

                int firstValid = _listBox.Items.Cast<object>().ToList().FindIndex(o => o is DropdownItem d && d.IsValid);
                if (firstValid >= 0) _listBox.SelectedIndex = firstValid;
            }
            else
            {
                _listBox.Items.Add(new DropdownItem { Display = "(No options available)", IsValid = false });
            }

            // Position Popup
            dynamic window = xlApp.ActiveWindow;
            dynamic targetPane = window.ActivePane;
            try
            {
                foreach (dynamic p in window.Panes)
                {
                    dynamic vr = p.VisibleRange;
                    if (singleCell.Row >= vr.Row && singleCell.Row <= vr.Row + vr.Rows.Count - 1 &&
                        singleCell.Column >= vr.Column && singleCell.Column <= vr.Column + vr.Columns.Count - 1)
                    {
                        targetPane = p;
                        break;
                    }
                }
            }
            catch { }

            double offsetX = (double)singleCell.Left - (double)targetPane.VisibleRange.Left;
            double offsetY = (double)singleCell.Top - (double)targetPane.VisibleRange.Top;

            int cellLeftPx = targetPane.PointsToScreenPixelsX(offsetX);
            int cellRightPx = targetPane.PointsToScreenPixelsX(offsetX + (double)singleCell.Width);
            int cellBottomPx = targetPane.PointsToScreenPixelsY(offsetY + (double)singleCell.Height);

            int popupWidthPx = Math.Max(cellRightPx - cellLeftPx, 160);
            int calculatedHeightPx = Math.Min(240, Math.Max(38, (_listBox.Items.Count * _listBox.ItemHeight) + 4));

            if (!_popupHost!.Visible)
            {
                _popupHost.Show();
            }

            SetWindowPos(_popupHost.Handle, HWND_TOP, cellLeftPx, cellBottomPx, popupWidthPx, calculatedHeightPx, SWP_SHOWWINDOW | SWP_NOACTIVATE);
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Arrow.PopulateAndOpenMenu Error] {ex.Message}");
        }
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
            if (_arrowHost != null && !_arrowHost.IsDisposed)
            {
                _arrowHost.Dispose();
                _arrowHost = null;
            }
            if (_popupHost != null && !_popupHost.IsDisposed)
            {
                _popupHost.Dispose();
                _popupHost = null;
            }
        }
        catch { }
    }
}