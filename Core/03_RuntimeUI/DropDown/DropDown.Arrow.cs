namespace SKSB_App.Core.RuntimeUI.DropDown;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Bootstrap.Environment;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.RuntimeUI.Undo;
using SKSB_App.Core.WorkbookSave.Validation;

public static class Arrow
{
    private static NoActivateOverlayForm? _arrowHost;
    private static Button? _arrowButton;
    private static NoActivateOverlayForm? _popupHost;
    private static ListBox? _listBox;
    private static dynamic? _activeCell;
    private static Action<string>? _onItemSelected;
    private static bool _isEventsHooked = false;

    // Multi-select session state: the popup stays open while items are toggled, and checked items are drawn with a tick.
    private static bool _isMultiSelectSession = false;
    private static Action? _pendingMultiSelectRowCheck;
    private static HashSet<string> _checkedTokens = new(StringComparer.OrdinalIgnoreCase);

    [DllImport("user32.dll")]
    private static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int X, int Y, int cx, int cy, uint uFlags);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool SetForegroundWindow(IntPtr hWnd);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ReleaseCapture();

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hWnd);

    // DPI scale (1.0 = 96 DPI) of the monitor hosting the Excel window; refreshed on every show/open.
    private static double _currentDpiScale = 1.0;

    private static double GetDpiScale(IntPtr hwnd)
    {
        if (hwnd == IntPtr.Zero) return 1.0;
        try
        {
            uint dpi = GetDpiForWindow(hwnd);
            return dpi == 0 ? 1.0 : dpi / 96.0;
        }
        catch { return 1.0; }
    }

    private static int ScalePx(int px, double dpiScale) => Math.Max(1, (int)Math.Round(px * dpiScale));

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
    private const uint SWP_NOSIZE = 0x0001;
    private const uint SWP_NOMOVE = 0x0002;
    private const uint SWP_NOZORDER = 0x0004;
    private const uint SWP_NOACTIVATE = 0x0010;
    private const uint SWP_SHOWWINDOW = 0x0040;
    private const uint SWP_HIDEWINDOW = 0x0080;
    private const int WM_MOUSEACTIVATE = 0x0021;
    private const int MA_NOACTIVATE = 3;
    private const int SW_HIDE = 0;
    private const int SW_SHOWNOACTIVATE = 4;
    private const int VK_LBUTTON = 0x01;
    private const int VK_RBUTTON = 0x02;
    private const int VK_MBUTTON = 0x04;

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

    [DllImport("user32.dll")]
    private static extern IntPtr GetForegroundWindow();

    [DllImport("user32.dll")]
    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("user32.dll")]
    [return: MarshalAs(UnmanagedType.Bool)]
    private static extern bool IsWindowVisible(IntPtr hWnd);

    [DllImport("user32.dll")]
    private static extern short GetAsyncKeyState(int vKey);

    private delegate void WindowDeactivateDelegate(dynamic wb, dynamic wn);
    private delegate void SheetDeactivateDelegate(dynamic sh);
    private static WindowDeactivateDelegate? _windowDeactivateHandler;
    private static SheetDeactivateDelegate? _sheetDeactivateHandler;

    private static bool _isArrowVisible;
    private static bool _isPopupVisible;
    private static bool _wasMouseDown;
    private static Rectangle _arrowBounds;
    private static Rectangle _popupBounds;
    private static System.Windows.Forms.Timer? _visibilityTimer;
    private static readonly PopupMessageFilter _messageFilter = new();
    private static bool _isMessageFilterAdded;

    /// <summary>Borderless, non-activating, tool-window overlay shared by the arrow button and the popup list.</summary>
    private sealed class NoActivateOverlayForm : Form
    {
        public NoActivateOverlayForm()
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

        protected override void WndProc(ref Message m)
        {
            if (m.Msg == WM_MOUSEACTIVATE)
            {
                m.Result = new IntPtr(MA_NOACTIVATE);
                return;
            }
            base.WndProc(ref m);
        }
    }

    static Arrow()
    {
        InitializeArrowHost();
        InitializePopupHost();
    }

    private static void InitializeArrowHost()
    {
        _arrowHost = new NoActivateOverlayForm { BackColor = Color.FromArgb(218, 220, 224) };
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
            int dx = ScalePx(4, _currentDpiScale);
            int dyTop = ScalePx(2, _currentDpiScale);
            int dyBot = ScalePx(3, _currentDpiScale);
            Point[] triangle =
            {
                new(cx - dx, cy - dyTop),
                new(cx + dx, cy - dyTop),
                new(cx, cy + dyBot)
            };
            using var brush = new SolidBrush(Color.FromArgb(70, 70, 70));
            e.Graphics.FillPolygon(brush, triangle);
        };
        _arrowButton.Click += (_, _) => ToggleDropdown();
        _arrowHost.Controls.Add(_arrowButton);
    }

    private static void InitializePopupHost()
    {
        _popupHost = new NoActivateOverlayForm
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
        _listBox.MouseDown += (s, e) =>
        {
            if (e.Button != MouseButtons.Left) return;
            int index = _listBox.IndexFromPoint(e.Location);
            if (index != ListBox.NoMatches && index >= 0)
            {
                _listBox.SelectedIndex = index;
                CommitSelection();
            }
        };
        _popupHost.Controls.Add(_listBox);
    }

    public static void EnsureExcelEventsHooked()
    {
        if (_isEventsHooked) return;
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            if (xlApp == null) return;
            _windowDeactivateHandler = new WindowDeactivateDelegate((wb, wn) => HideAll());
            _sheetDeactivateHandler = new SheetDeactivateDelegate(ws => HideAll());
            xlApp.WindowDeactivate += _windowDeactivateHandler;
            xlApp.SheetDeactivate += _sheetDeactivateHandler;
            _isEventsHooked = true;
            if (_activeCell != null && _isArrowVisible)
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
        var (label, isOption) = FormatItemLabel(rawObj);

        // Full item height (no vertical inset) so VerticalCenter never clips glyphs at high DPI.
        int padX = ScalePx(6, _currentDpiScale);
        var textBounds = new Rectangle(e.Bounds.X + padX, e.Bounds.Y, Math.Max(1, e.Bounds.Width - (padX * 2)), e.Bounds.Height);
        const TextFormatFlags flags = TextFormatFlags.VerticalCenter | TextFormatFlags.Left | TextFormatFlags.NoPrefix |
                                      TextFormatFlags.SingleLine | TextFormatFlags.EndEllipsis;

        if (!isOption || rawObj is not DropdownOption item)
        {
            e.DrawBackground();
            TextRenderer.DrawText(e.Graphics, label, e.Font!, textBounds, Color.Gray, flags);
            return;
        }

        bool isSelected = (e.State & DrawItemState.Selected) == DrawItemState.Selected;
        bool isChecked = _isMultiSelectSession && item.IsValid && (_checkedTokens.Contains(OptionValue(item)) || _checkedTokens.Contains(item.Display));
        Color backColor = !item.IsValid
            ? (isSelected ? Color.FromArgb(242, 242, 242) : Color.White)
            : isChecked
                ? (isSelected ? Color.FromArgb(196, 228, 200) : Color.FromArgb(226, 243, 229))
                : (isSelected ? Color.FromArgb(217, 235, 255) : Color.White);
        Color foreColor = !item.IsValid ? Color.FromArgb(170, 170, 175) : Color.Black;

        using var bgBrush = new SolidBrush(backColor);
        e.Graphics.FillRectangle(bgBrush, e.Bounds);

        // GDI text (not GDI+ DrawString) so Windows font-linking falls back to Segoe UI Emoji / Symbol for icons.
        TextRenderer.DrawText(e.Graphics, label, e.Font!, textBounds, foreColor, flags);
    }

    /// <summary>Single source of truth for the rendered item text (incl. multi-select tick / indent) used by drawing and width measurement.</summary>
    // Enriched labels ("PLATE — Type", "0001 — Name") write the clean Raw value; all other options keep writing Display
    // so text transforms and date formatting are preserved.
    private static string OptionValue(DropdownOption item) =>
        !string.IsNullOrEmpty(item.Raw)
        && (item.Display.StartsWith(item.Raw + " — ", StringComparison.Ordinal)
            || (long.TryParse(item.Raw, out long id) && item.Display.StartsWith($"{id:D4} — ", StringComparison.Ordinal)))
            ? item.Raw
            : item.Display;

    private static (string Label, bool IsOption) FormatItemLabel(object rawObj)
    {
        if (rawObj is not DropdownOption item) return (rawObj?.ToString() ?? string.Empty, false);

        if (_isMultiSelectSession && item.IsValid)
        {
            bool isChecked = _checkedTokens.Contains(OptionValue(item)) || _checkedTokens.Contains(item.Display);
            return ((isChecked ? "\u2713  " : "     ") + item.Display, true);
        }
        return (item.Display, true);
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
            int targetRow = (int)_activeCell.Row;
            int targetCol = (int)_activeCell.Column;
            dynamic targetPane = ResolveTargetPane(window, targetRow, targetCol);

            int cellTopPx = targetPane.PointsToScreenPixelsY((double)_activeCell.Top);
            int cellRightPx = targetPane.PointsToScreenPixelsX((double)_activeCell.Left + (double)_activeCell.Width);
            int cellBottomPx = targetPane.PointsToScreenPixelsY((double)_activeCell.Top + (double)_activeCell.Height);

            // Screen coordinates can be negative on monitors left of / above the primary,
            // so validate against the full virtual desktop instead of 0.
            Rectangle virtualScreen = SystemInformation.VirtualScreen;
            if (cellRightPx <= virtualScreen.Left || cellBottomPx <= virtualScreen.Top ||
                cellRightPx >= virtualScreen.Right || cellTopPx >= virtualScreen.Bottom)
            {
                AppLogger.Warn($"[Arrow] ShowArrowForCell hidden: off-screen pixels (right={cellRightPx}, top={cellTopPx}, bottom={cellBottomPx}, virtual={virtualScreen}) for row {targetRow}, col {targetCol}.");
                HideAll();
                return;
            }

            double dpiScale = GetDpiScale(excelHwnd);
            _currentDpiScale = dpiScale;

            int cellHeightPx = cellBottomPx - cellTopPx;
            int arrowWidthPx = ScalePx(18, dpiScale);
            int arrowHeightPx = cellHeightPx;
            int arrowXPx = cellRightPx + ScalePx(1, dpiScale);
            int arrowYPx = cellTopPx;

            if (!_arrowHost!.Visible) _arrowHost.Show();
            EnsureOwner(_arrowHost, excelHwnd);
            SetWindowPos(_arrowHost.Handle, HWND_TOP, arrowXPx, arrowYPx, arrowWidthPx, arrowHeightPx, SWP_SHOWWINDOW | SWP_NOACTIVATE);
            _arrowBounds = new Rectangle(arrowXPx, arrowYPx, arrowWidthPx, arrowHeightPx);
            _isArrowVisible = true;
            _arrowButton?.Invalidate();
            HidePopup();
            EnsureVisibilityMonitor();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[Arrow] ShowArrowForCell failed: {ex.GetType().Name}: {ex.Message}");
            HideAll();
        }
    }

    private static void HideNoActivate(Form? form)
    {
        if (form == null || form.IsDisposed || !form.IsHandleCreated) return;
        // Clear the owner first so hiding never triggers owned-window activation changes on Excel.
        SetWindowLongPtr(form.Handle, GWL_HWNDPARENT, IntPtr.Zero);
        SetWindowPos(form.Handle, IntPtr.Zero, 0, 0, 0, 0,
            SWP_HIDEWINDOW | SWP_NOMOVE | SWP_NOSIZE | SWP_NOZORDER | SWP_NOACTIVATE);
        form.Visible = false;
    }

    /// <summary>Attaches the overlay to Excel as its owner; re-applied on every show.</summary>
    private static void EnsureOwner(Form? form, IntPtr excelHwnd)
    {
        if (form == null || excelHwnd == IntPtr.Zero) return;
        SetWindowLongPtr(form.Handle, GWL_HWNDPARENT, excelHwnd);
    }

    public static void HideAll()
    {
        HidePopup();
        if (_isArrowVisible)
        {
            HideNoActivate(_arrowHost);
            _isArrowVisible = false;
        }
        _activeCell = null;
        _visibilityTimer?.Stop();
    }

    public static void HidePopup()
    {
        _isMultiSelectSession = false;
        var pending = _pendingMultiSelectRowCheck;
        _pendingMultiSelectRowCheck = null;
        if (_isPopupVisible)
        {
            HideNoActivate(_popupHost);
            _isPopupVisible = false;
        }
        pending?.Invoke();
    }

    /// <summary>Hides the popup without touching foreground/activation (Excel stays active).</summary>
    public static void DismissPopup()
    {
        try { ReleaseCapture(); } catch { }
        HidePopup();
    }

    public static void ToggleDropdown()
    {
        if (_isPopupVisible)
        {
            DismissPopup();
            return;
        }
        if (_activeCell == null) return;
        PopulateAndOpenMenu(_activeCell);
    }

    private static void PopulateAndOpenMenu(dynamic singleCell)
    {
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

            PropertyRuleDescriptor? rule = RuleRegistry.GetRuleForCell(singleCell);
            var processedItems = new List<DropdownOption>();

            CustomDropdownResult? custom = CustomDropdowns.TryBuild(singleCell);
            _isMultiSelectSession = custom == null && rule != null && rule.HasDropdown && rule.IsMultiSelect;
            _checkedTokens = _isMultiSelectSession
                ? new HashSet<string>(RuleEvaluator.SplitMultiSelectTokens(inCellText), StringComparer.OrdinalIgnoreCase)
                : new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            // =========================================================================
            // PATH C: Module-supplied dropdown (e.g. Schedule board match-% ranking)
            // =========================================================================
            if (custom != null)
            {
                processedItems.AddRange(custom.Options);

                _onItemSelected = chosenDisplay =>
                {
                    var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                    if (targetItem == null) return;

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        UndoManager.CaptureCellSnapshot(sheet, singleCell);
                        bool origEvents = true;
                        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                        try { custom.Commit(singleCell, targetItem); }
                        finally
                        {
                            try { xlApp.EnableEvents = origEvents; } catch { }
                        }

                        ParameterCells.NotifyChanged(sheet, singleCell);
                        UndoManager.CommitEdit(sheet, singleCell, $"Dropdown {singleCell.Address?.ToString() ?? ""}");

                        if (table != null)
                        {
                            TableRowManager.HandleSheetChange(singleCell);
                            StackedEntryGuard.OnSheetChanged(sheet, singleCell);
                        }
                    });
                };
            }
            // =========================================================================
            // PATH A: Unified C# Attribute Rule Engine (Zero COM formula evaluations)
            // =========================================================================
            else if (rule != null && rule.HasDropdown)
            {
                var evaluatedOptions = RuleEvaluator.GetEvaluatedDropdownOptions(wb, table, row, rule);
                processedItems.AddRange(evaluatedOptions);

                _onItemSelected = chosenDisplay =>
                {
                    var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                    if (targetItem == null) return;

                    string prevText = inCellText;
                    if (rule.IsMultiSelect)
                    {
                        // Update in-memory state immediately so the tick shows without waiting for the COM write.
                        var tokens = RuleEvaluator.SplitMultiSelectTokens(prevText);
                        string val = OptionValue(targetItem);
                        int idx = tokens.FindIndex(t => t.Equals(val, StringComparison.OrdinalIgnoreCase));
                        if (idx >= 0) tokens.RemoveAt(idx);
                        else tokens.Add(val);
                        _checkedTokens = new HashSet<string>(tokens, StringComparer.OrdinalIgnoreCase);
                        inCellText = string.Join(", ", tokens);
                    }

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        UndoManager.CaptureCellSnapshot(sheet, singleCell);
                        bool origEvents = true;
                        try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
                        try
                        {
                            WriteSelectedValueToCell(singleCell, targetItem, rule, prevText);
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

                        // Events were suppressed for the write: let module observers (rollups etc.) know.
                        ParameterCells.NotifyChanged(sheet, singleCell);
                        UndoManager.CommitEdit(sheet, singleCell, $"Dropdown {singleCell.Address?.ToString() ?? ""}");

                        if (table != null)
                        {
                            if (rule.IsMultiSelect)
                                _pendingMultiSelectRowCheck = () => ExcelAsyncUtil.QueueAsMacro(() =>
                                {
                                    TableRowManager.HandleSheetChange(singleCell);
                                    StackedEntryGuard.OnSheetChanged(sheet, singleCell);
                                });
                            else
                            {
                                TableRowManager.HandleSheetChange(singleCell);
                                StackedEntryGuard.OnSheetChanged(sheet, singleCell);
                            }
                        }
                    });
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
                        processedItems.Add(new DropdownOption { Display = disp, Raw = itm, Section = 0, IsValid = true });
                    }
                }

                _onItemSelected = chosenDisplay =>
                {
                    var targetItem = processedItems.FirstOrDefault(i => i.Display.Equals(chosenDisplay, StringComparison.OrdinalIgnoreCase));
                    if (targetItem == null) return;

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        UndoManager.CaptureCellSnapshot(sheet, singleCell);
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

                        ParameterCells.NotifyChanged(sheet, singleCell);
                        UndoManager.CommitEdit(sheet, singleCell, $"Dropdown {singleCell.Address?.ToString() ?? ""}");

                        if (table != null)
                        {
                            TableRowManager.HandleSheetChange(singleCell);
                            StackedEntryGuard.OnSheetChanged(sheet, singleCell);
                        }
                    });
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

                var sortedValid = validItems.OrderBy(i => i.Section).ThenBy(i => i.Display, NaturalStringComparer.Instance);
                var sortedInvalid = invalidItems.OrderBy(i => i.Section).ThenBy(i => i.Display, NaturalStringComparer.Instance);

                if (custom != null)
                {
                    // Provider order (e.g. match-% rank) is authoritative.
                    sortedValid = validItems.OrderBy(i => i.Section);
                    sortedInvalid = invalidItems.OrderBy(i => i.Section);
                }
                else if (!string.IsNullOrWhiteSpace(inCellText) && (rule == null || !rule.IsMultiSelect))
                {
                    string q = inCellText.Trim();
                    Func<DropdownOption, int> ranker = i =>
                    {
                        if (i.Display.Equals(q, StringComparison.OrdinalIgnoreCase) ||
                            (!string.IsNullOrEmpty(i.Raw) && i.Raw.Equals(q, StringComparison.OrdinalIgnoreCase))) return 0;
                        if (i.Display.StartsWith(q, StringComparison.OrdinalIgnoreCase) ||
                            (!string.IsNullOrEmpty(i.Raw) && i.Raw.StartsWith(q, StringComparison.OrdinalIgnoreCase))) return 1;
                        if (i.Display.IndexOf(q, StringComparison.OrdinalIgnoreCase) >= 0) return 2;
                        return 3;
                    };
                    sortedValid = validItems.OrderBy(ranker).ThenBy(i => i.Section).ThenBy(i => i.Display, NaturalStringComparer.Instance);
                    sortedInvalid = invalidItems.OrderBy(ranker).ThenBy(i => i.Section).ThenBy(i => i.Display, NaturalStringComparer.Instance);
                }

                foreach (var item in sortedValid) _listBox.Items.Add(item);
                foreach (var item in sortedInvalid) _listBox.Items.Add(item);

                int firstValid = _listBox.Items.Cast<object>().ToList().FindIndex(o => o is DropdownOption d && d.IsValid);
                if (firstValid >= 0) _listBox.SelectedIndex = firstValid;
            }
            else
            {
                _listBox.Items.Add(new DropdownOption { Display = "(No options available)", IsValid = false });
            }

            dynamic window = xlApp.ActiveWindow;
            dynamic targetPane = ResolveTargetPane(window, (int)singleCell.Row, (int)singleCell.Column);

            int cellLeftPx = targetPane.PointsToScreenPixelsX((double)singleCell.Left);
            int cellRightPx = targetPane.PointsToScreenPixelsX((double)singleCell.Left + (double)singleCell.Width);
            int cellTopPx = targetPane.PointsToScreenPixelsY((double)singleCell.Top);
            int cellBottomPx = targetPane.PointsToScreenPixelsY((double)singleCell.Top + (double)singleCell.Height);

            // =========================================================================
            // DPI-AWARE SIZING (measure rendered labels on the Excel window's monitor)
            // =========================================================================
            IntPtr excelHwnd = GetExcelHwnd(xlApp);
            double dpiScale = GetDpiScale(excelHwnd);
            _currentDpiScale = dpiScale;

            int maxTextWidth = 0;
            int maxTextHeight = 0;
            using (var g = Graphics.FromHwnd(excelHwnd != IntPtr.Zero ? excelHwnd : _listBox.Handle))
            {
                double gScale = g.DpiY > 0 ? g.DpiY / 96.0 : 1.0;
                // Mixed-DPI: if the measuring surface reports a lower DPI than the target monitor, scale up so we never underestimate.
                double measureAdjust = dpiScale > gScale ? dpiScale / gScale : 1.0;
                var unbounded = new Size(int.MaxValue, int.MaxValue);

                foreach (object rawObj in _listBox.Items)
                {
                    string label = FormatItemLabel(rawObj).Label;
                    Size measured = TextRenderer.MeasureText(g, label, _listBox.Font, unbounded, TextFormatFlags.NoPrefix | TextFormatFlags.SingleLine);
                    maxTextWidth = Math.Max(maxTextWidth, (int)Math.Ceiling(measured.Width * measureAdjust));
                    maxTextHeight = Math.Max(maxTextHeight, (int)Math.Ceiling(measured.Height * measureAdjust));
                }
            }

            int itemHeightPx = Math.Clamp(Math.Max(maxTextHeight + ScalePx(6, dpiScale), ScalePx(24, dpiScale)), 18, 255);
            if (_listBox.ItemHeight != itemHeightPx) _listBox.ItemHeight = itemHeightPx;

            // Fit to the working area of the monitor hosting the cell (coordinates may be negative
            // on secondary monitors); flip above the cell / shift left only when it would overflow.
            Rectangle workArea = Screen.FromPoint(new Point(cellLeftPx, cellBottomPx)).WorkingArea;

            int maxPopupHeightPx = Math.Min(workArea.Height / 2, (_listBox.ItemHeight * 10) + ScalePx(4, dpiScale));
            int contentHeightPx = (_listBox.Items.Count * _listBox.ItemHeight) + ScalePx(4, dpiScale);
            int calculatedHeightPx = Math.Min(maxPopupHeightPx, Math.Max(_listBox.ItemHeight + ScalePx(8, dpiScale), contentHeightPx));

            bool needsVScroll = contentHeightPx > calculatedHeightPx;
            int scrollBarPx = needsVScroll ? Math.Max(SystemInformation.VerticalScrollBarWidth, ScalePx(18, dpiScale)) : 0;
            int hPadPx = ScalePx(16, dpiScale);
            int cellWidthPx = (cellRightPx - cellLeftPx) + ScalePx(19, dpiScale);
            int desiredWidthPx = Math.Max(cellWidthPx, Math.Max(ScalePx(160, dpiScale), maxTextWidth + hPadPx + scrollBarPx));
            int popupWidthPx = Math.Min(desiredWidthPx, Math.Max(ScalePx(160, dpiScale), workArea.Width - ScalePx(16, dpiScale)));

            int popupXPx = cellLeftPx;
            int popupYPx = cellBottomPx;
            if (popupYPx + calculatedHeightPx > workArea.Bottom && cellTopPx - calculatedHeightPx >= workArea.Top)
                popupYPx = cellTopPx - calculatedHeightPx;
            popupXPx = Math.Max(workArea.Left, Math.Min(popupXPx, workArea.Right - popupWidthPx));
            popupYPx = Math.Max(workArea.Top, Math.Min(popupYPx, workArea.Bottom - calculatedHeightPx));

            if (!_popupHost!.Visible) _popupHost.Show();
            EnsureOwner(_popupHost, excelHwnd);
            SetWindowPos(_popupHost.Handle, HWND_TOP, popupXPx, popupYPx, popupWidthPx, calculatedHeightPx, SWP_SHOWWINDOW | SWP_NOACTIVATE);
            _popupBounds = new Rectangle(popupXPx, popupYPx, popupWidthPx, calculatedHeightPx);
            _isPopupVisible = true;
            _wasMouseDown = true;
            _listBox.Invalidate();
            EnsureVisibilityMonitor();
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Arrow.PopulateAndOpenMenu Error] {ex.Message}");
        }
    }

    private static void WriteSelectedValueToCell(
        dynamic singleCell,
        DropdownOption targetItem,
        PropertyRuleDescriptor rule,
        string existingCellText)
    {
        if (rule.IsMultiSelect)
        {
            var tokens = RuleEvaluator.SplitMultiSelectTokens(existingCellText);

            string val = OptionValue(targetItem);
            int existingIdx = tokens.FindIndex(t => t.Equals(val, StringComparison.OrdinalIgnoreCase));
            if (existingIdx >= 0)
                tokens.RemoveAt(existingIdx);
            else
                tokens.Add(val);

            // An emptied selection clears the cell instead of leaving an empty-string text value.
            if (tokens.Count == 0) singleCell.ClearContents();
            else singleCell.Value2 = string.Join(", ", tokens);
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

        singleCell.Value2 = OptionValue(targetItem);
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
        if (_listBox!.SelectedItem is not DropdownOption item || _onItemSelected == null) return;
        var handler = _onItemSelected;

        // Multi-select keeps the popup open so several items can be toggled in one session.
        if (_isMultiSelectSession)
        {
            handler(item.Display);
            _listBox.Invalidate();
            return;
        }

        // Single-select: hide the popup first so it is gone before the queued COM write runs.
        DismissPopup();
        handler(item.Display);
    }

    // =========================================================================
    // 8. CLICK-AWAY / KEYBOARD DISMISSAL
    // =========================================================================
    private sealed class PopupMessageFilter : IMessageFilter
    {
        private const int WM_KEYDOWN = 0x0100;
        private const int WM_LBUTTONDOWN = 0x0201;
        private const int WM_RBUTTONDOWN = 0x0204;
        private const int WM_MBUTTONDOWN = 0x0207;
        private const int WM_MOUSEWHEEL = 0x020A;
        private const int WM_NCLBUTTONDOWN = 0x00A1;
        private const int WM_NCRBUTTONDOWN = 0x00A4;
        private const int WM_NCMBUTTONDOWN = 0x00A7;

        public bool PreFilterMessage(ref Message m)
        {
            if (!_isPopupVisible || _popupHost == null || _listBox == null) return false;

            switch (m.Msg)
            {
                case WM_MOUSEWHEEL:
                {
                    if (_popupBounds.Contains(Cursor.Position))
                    {
                        int delta = (short)((m.WParam.ToInt64() >> 16) & 0xFFFF);
                        int step = delta > 0 ? -1 : 1;
                        int top = Math.Max(0, Math.Min(_listBox.Items.Count - 1, _listBox.TopIndex + step));
                        if (_listBox.Items.Count > 0) _listBox.TopIndex = top;
                        return true;
                    }
                    DismissPopup();
                    return false;
                }
                case WM_LBUTTONDOWN:
                case WM_RBUTTONDOWN:
                case WM_MBUTTONDOWN:
                case WM_NCLBUTTONDOWN:
                case WM_NCRBUTTONDOWN:
                case WM_NCMBUTTONDOWN:
                {
                    if (!IsInsideDropdownUi(Cursor.Position)) DismissPopup();
                    return false;
                }
                case WM_KEYDOWN:
                {
                    var key = (Keys)(m.WParam.ToInt64() & 0xFFFF);
                    switch (key)
                    {
                        case Keys.Escape:
                            DismissPopup();
                            return true;
                        case Keys.Up:
                            if (_listBox.SelectedIndex > 0) _listBox.SelectedIndex--;
                            return true;
                        case Keys.Down:
                            if (_listBox.SelectedIndex < _listBox.Items.Count - 1) _listBox.SelectedIndex++;
                            return true;
                        case Keys.Enter:
                            if (_isMultiSelectSession) DismissPopup();
                            else CommitSelection();
                            return true;
                        case Keys.Space when _isMultiSelectSession:
                            CommitSelection();
                            return true;
                        case Keys.ShiftKey:
                        case Keys.ControlKey:
                        case Keys.Menu:
                            return false;
                        default:
                            DismissPopup();
                            return false;
                    }
                }
            }
            return false;
        }
    }

    private static bool IsInsideDropdownUi(Point pt)
    {
        if (_popupHost != null && _isPopupVisible && _popupBounds.Contains(pt)) return true;
        if (_arrowHost != null && _isArrowVisible && _arrowBounds.Contains(pt)) return true;
        return false;
    }

    private static void EnsureVisibilityMonitor()
    {
        if (!_isMessageFilterAdded)
        {
            try { Application.AddMessageFilter(_messageFilter); _isMessageFilterAdded = true; } catch { }
        }
        if (_visibilityTimer == null)
        {
            _visibilityTimer = new System.Windows.Forms.Timer { Interval = 100 };
            _visibilityTimer.Tick += (_, _) => OnVisibilityTimerTick();
        }
        if (!_visibilityTimer.Enabled) _visibilityTimer.Start();
    }

    private static void OnVisibilityTimerTick()
    {
        try
        {
            if (!_isArrowVisible && !_isPopupVisible)
            {
                _visibilityTimer?.Stop();
                return;
            }

            IntPtr fg = GetForegroundWindow();
            bool foreignForeground = fg == IntPtr.Zero;
            if (!foreignForeground)
            {
                GetWindowThreadProcessId(fg, out uint pid);
                foreignForeground = pid != (uint)Environment.ProcessId;
            }
            if (foreignForeground)
            {
                HideAll();
                return;
            }

            // The host message loop may not route mouse messages through Application filters,
            // so also poll the physical buttons for a click outside the open popup.
            bool down = (GetAsyncKeyState(VK_LBUTTON) & 0x8000) != 0
                        || (GetAsyncKeyState(VK_RBUTTON) & 0x8000) != 0
                        || (GetAsyncKeyState(VK_MBUTTON) & 0x8000) != 0;
            if (down && !_wasMouseDown && _isPopupVisible && !IsInsideDropdownUi(Cursor.Position))
                DismissPopup();
            _wasMouseDown = down;
        }
        catch { }
    }

    public static void Teardown()
    {
        try
        {
            _pendingMultiSelectRowCheck = null;
            HideAll();
            _visibilityTimer?.Stop();
            _visibilityTimer?.Dispose();
            _visibilityTimer = null;
            if (_isMessageFilterAdded)
            {
                try { Application.RemoveMessageFilter(_messageFilter); } catch { }
                _isMessageFilterAdded = false;
            }

            if (_isEventsHooked)
            {
                try
                {
                    dynamic xlApp = ExcelDnaUtil.Application;
                    if (_windowDeactivateHandler != null) xlApp.WindowDeactivate -= _windowDeactivateHandler;
                    if (_sheetDeactivateHandler != null) xlApp.SheetDeactivate -= _sheetDeactivateHandler;
                }
                catch { }
                _windowDeactivateHandler = null;
                _sheetDeactivateHandler = null;
                _isEventsHooked = false;
            }

            _arrowHost?.Dispose();
            _arrowHost = null;
            _popupHost?.Dispose();
            _popupHost = null;
        }
        catch { }
    }
}