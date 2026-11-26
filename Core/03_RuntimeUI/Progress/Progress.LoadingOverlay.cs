namespace SKSB_App.Core.RuntimeUI.Progress;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Localization;
using SKSB_App.Core.RuntimeUI.DropDown;

/// <summary>
/// Non-blocking animated loading window shown while Excel is frozen (ScreenUpdating = false).
/// It lives on its own background STA thread so its spinner and clock keep animating while Excel's
/// main thread is busy with COM writes, calculation or synchronous database calls.
/// The background thread NEVER touches Excel COM objects; only the raw HWND is handed over.
/// </summary>
public static class LoadingOverlay
{
    // =========================================================================
    // 1. STATE
    // =========================================================================
    private static readonly object _gate = new();
    private static readonly Stack<(string Title, string Step, string? Detail)> _parents = new();
    private static LoadingOverlayForm? _form;
    private static Thread? _uiThread;
    private static int _scopeDepth;
    private static int _suspendDepth;
    private static IntPtr _excelHwnd = IntPtr.Zero;

    private static string _title = string.Empty;
    private static string _step = string.Empty;
    private static string? _detail;

    private static readonly Stopwatch _rootStopwatch = new();
    private static string _rootOperationKey = string.Empty;
    private static double? _rootEstimateMs;

    // =========================================================================
    // 2. PUBLIC API
    // =========================================================================
    public static IDisposable BeginScope(string title, string? initialStep = null, string? detail = null, string? operationKey = null)
    {
        initialStep ??= Loc.T("loading.processing");

        IntPtr hwnd = IntPtr.Zero;
        try { hwnd = Arrow.GetExcelHwnd(ExcelDnaUtil.Application); } catch { }

        bool showNow = false;
        lock (_gate)
        {
            _scopeDepth++;
            if (_scopeDepth == 1)
            {
                _parents.Clear();
                _excelHwnd = hwnd;
                showNow = _suspendDepth == 0;
                _rootOperationKey = LoadingTimeEstimator.NormalizeKey(title, operationKey);
                _rootEstimateMs = LoadingTimeEstimator.GetEstimateMs(_rootOperationKey);
                _rootStopwatch.Restart();
                if (_suspendDepth > 0) _rootStopwatch.Stop(); // started while suspended: don't count modal wait time
            }
            else
            {
                _parents.Push((_title, _step, _detail));
            }
            _title = title;
            _step = initialStep;
            _detail = detail;
        }

        if (showNow) ShowForm(restartClock: true);
        else RequestRepaint();
        return new Scope();
    }

    public static void Report(string step, string? detail = null)
    {
        lock (_gate)
        {
            if (_scopeDepth <= 0) return;
            _step = step;
            if (detail != null) _detail = detail;
        }
        RequestRepaint();
    }

    public static IDisposable Suspend()
    {
        bool hide;
        lock (_gate)
        {
            _suspendDepth++;
            hide = _suspendDepth == 1;
            if (_suspendDepth == 1 && _scopeDepth > 0)
                _rootStopwatch.Stop();
        }
        if (hide) HideForm();
        return new SuspendToken();
    }

    // =========================================================================
    // 3. SCOPE TOKENS
    // =========================================================================
    private sealed class Scope : IDisposable
    {
        private int _disposed;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;

            bool hide = false;
            double elapsedMs = 0;
            string rootKey = string.Empty;
            lock (_gate)
            {
                if (_scopeDepth <= 0) return;
                _scopeDepth--;
                if (_scopeDepth == 0)
                {
                    _parents.Clear();
                    hide = true;
                    _rootStopwatch.Stop();
                    elapsedMs = _rootStopwatch.Elapsed.TotalMilliseconds;
                    rootKey = _rootOperationKey;
                    _rootOperationKey = string.Empty;
                    _rootEstimateMs = null;
                }
                else if (_parents.Count > 0)
                {
                    (_title, _step, _detail) = _parents.Pop();
                }
            }

            if (hide)
            {
                HideForm();
                if (!string.IsNullOrEmpty(rootKey))
                    LoadingTimeEstimator.RecordDuration(rootKey, elapsedMs);
            }
            else RequestRepaint();
        }
    }

    private sealed class SuspendToken : IDisposable
    {
        private int _disposed;

        public void Dispose()
        {
            if (Interlocked.Exchange(ref _disposed, 1) != 0) return;

            bool show;
            lock (_gate)
            {
                if (_suspendDepth > 0) _suspendDepth--;
                if (_suspendDepth == 0 && _scopeDepth > 0)
                    _rootStopwatch.Start();
                show =_scopeDepth > 0 && _suspendDepth == 0;
            }
            if (show) ShowForm(restartClock: false);
        }
    }

    // =========================================================================
    // 4. UI THREAD MARSHALING
    // =========================================================================
    private static (string Title, string Step, string? Detail, double? EstimateMs) Snapshot()
    {
        string? pendingKey = null;
        lock (_gate)
        {
            if (_rootEstimateMs == null && !string.IsNullOrEmpty(_rootOperationKey))
                pendingKey = _rootOperationKey;
        }

        if (pendingKey != null)
        {
            double? est = LoadingTimeEstimator.GetEstimateMs(pendingKey);
            if (est != null)
            {
                lock (_gate)
                {
                    if (_rootEstimateMs == null && _rootOperationKey == pendingKey)
                        _rootEstimateMs = est;
                }
            }
        }

        lock (_gate) return (_title, _step, _detail, _rootEstimateMs);
    }

    [DllImport("user32.dll")]
    private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr dpiContext);
    private static readonly IntPtr DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2 = new(-4);

    [DllImport("user32.dll")]
    private static extern uint GetDpiForWindow(IntPtr hwnd);

    private static LoadingOverlayForm? EnsureForm()
    {
        if (_form != null) return _form;

        lock (_gate)
        {
            if (_form != null) return _form;
            using var ready = new ManualResetEventSlim(false);
            LoadingOverlayForm? created = null;

            var thread = new Thread(() =>
            {
                try
                {
                    try { SetThreadDpiAwarenessContext(DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2); } catch { }
                    created = new LoadingOverlayForm();
                    _ = created.Handle;
                    ready.Set();
                    Application.Run(new ApplicationContext());
                }
                catch
                {
                    ready.Set();
                }
            })
            {
                IsBackground = true,
                Name = "SKSB LoadingOverlay UI"
            };
            thread.SetApartmentState(ApartmentState.STA);
            thread.Start();

            if (!ready.Wait(3000) || created == null) return null;
            // Wait for the handle to exist so Invoke/BeginInvoke are safe.
            for (int i = 0; i < 100 && !created.IsHandleCreated; i++) Thread.Sleep(10);
            if (!created.IsHandleCreated) return null;

            _uiThread = thread;
            _form = created;
            return _form;
        }
    }

    private static void ShowForm(bool restartClock)
    {
        try
        {
            var form = EnsureForm();
            if (form == null || form.IsDisposed) return;
            IntPtr hwnd;
            lock (_gate) hwnd = _excelHwnd;
            form.Invoke(new Action(() => form.ShowOverlay(hwnd, restartClock))); // synchronous: visible before heavy work begins
        }
        catch { }
    }

    private static void HideForm()
    {
        try
        {
            var form = _form;
            if (form == null || form.IsDisposed || !form.IsHandleCreated) return;
            form.Invoke(new Action(form.HideOverlay)); // synchronous: modal dialogs must never collide with the overlay
        }
        catch { }
    }

    private static void RequestRepaint()
    {
        try
        {
            var form = _form;
            if (form == null || form.IsDisposed || !form.IsHandleCreated) return;
            form.BeginInvoke(new Action(form.Invalidate));
        }
        catch { }
    }

    // =========================================================================
    // 5. OVERLAY FORM (runs on the background STA thread)
    // =========================================================================
    private sealed class LoadingOverlayForm : Form
    {
        [StructLayout(LayoutKind.Sequential)]
        private struct RECT { public int Left, Top, Right, Bottom; }

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool GetWindowRect(IntPtr hWnd, out RECT lpRect);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        private static extern bool IsIconic(IntPtr hWnd);

        private readonly Stopwatch _stopwatch = new();
        private readonly Stopwatch _spinClock = Stopwatch.StartNew();
        private readonly System.Windows.Forms.Timer _timer;

        private const int BaseWidth = 440;
        private const int BaseHeight = 108;
        private float _dpiScale = 1f;
        private float _cachedFontScale = -1f;
        private Font? _titleFont;
        private Font? _clockFont;
        private Font? _stepFont;
        private Font? _detailFont;

        public LoadingOverlayForm()
        {
            AutoScaleMode = AutoScaleMode.None;
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            DoubleBuffered = true;
            BackColor = Color.White;
            ApplyDpiScale((uint)(DeviceDpi > 0 ? DeviceDpi : 96));

            _timer = new System.Windows.Forms.Timer { Interval = 30 };
            _timer.Tick += (_, _) => Invalidate();

            // Force handle creation so cross-thread Invoke works before the form is ever shown.
            CreateHandle();
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ClassStyle |= 0x00020000;            // CS_DROPSHADOW
                cp.ExStyle |= 0x00000080 | 0x08000000;  // WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                return cp;
            }
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (e.CloseReason == CloseReason.UserClosing)
            {
                e.Cancel = true;
                Hide();
                return;
            }
            base.OnFormClosing(e);
        }

        public void ShowOverlay(IntPtr excelHwnd, bool restartClock)
        {
            if (restartClock) _stopwatch.Restart();
            else if (!_stopwatch.IsRunning) _stopwatch.Start();

            uint dpi = 0;
            if (excelHwnd != IntPtr.Zero)
            {
                try { dpi = GetDpiForWindow(excelHwnd); } catch { }
            }
            if (dpi == 0) dpi = (uint)(DeviceDpi > 0 ? DeviceDpi : 96);
            ApplyDpiScale(dpi);

            Location = ComputeLocation(excelHwnd);
            if (!Visible) Show();
            TopMost = true;
            _timer.Start();
            Invalidate();
            Update(); // paint now so the window is fully rendered before the caller starts heavy work
        }

        private void ApplyDpiScale(uint dpi)
        {
            float scale = dpi >= 96 ? dpi / 96f : 1f;
            _dpiScale = scale;
            Size = new Size((int)Math.Round(BaseWidth * scale), (int)Math.Round(BaseHeight * scale));

            if (Math.Abs(_cachedFontScale - scale) > 0.001f)
            {
                DisposeFonts();
                _titleFont = new Font("Segoe UI", 14f * scale, FontStyle.Bold, GraphicsUnit.Pixel);
                _clockFont = new Font("Segoe UI", 11.5f * scale, FontStyle.Regular, GraphicsUnit.Pixel);
                _stepFont = new Font("Segoe UI", 12.5f * scale, FontStyle.Regular, GraphicsUnit.Pixel);
                _detailFont = new Font("Segoe UI", 11f * scale, FontStyle.Regular, GraphicsUnit.Pixel);
                _cachedFontScale = scale;
            }
        }

        private void DisposeFonts()
        {
            _titleFont?.Dispose();
            _clockFont?.Dispose();
            _stepFont?.Dispose();
            _detailFont?.Dispose();
            _titleFont = _clockFont = _stepFont = _detailFont = null;
        }

        protected override void OnDpiChanged(DpiChangedEventArgs e)
        {
            ApplyDpiScale((uint)e.DeviceDpiNew);
            Invalidate();
        }

        public void HideOverlay()
        {
            _timer.Stop();
            _stopwatch.Stop();
            if (Visible) Hide();
        }

        private Point ComputeLocation(IntPtr excelHwnd)
        {
            try
            {
                if (excelHwnd != IntPtr.Zero && !IsIconic(excelHwnd) && GetWindowRect(excelHwnd, out RECT r)
                    && r.Right > r.Left && r.Bottom > r.Top)
                {
                    return new Point(
                        r.Left + ((r.Right - r.Left) - Width) / 2,
                        r.Top + ((r.Bottom - r.Top) - Height) / 2);
                }
            }
            catch { }

            var area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1280, 720);
            return new Point(area.Left + (area.Width - Width) / 2, area.Top + (area.Height - Height) / 2);
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            var g = e.Graphics;
            g.Clear(Color.White);
            g.SmoothingMode = SmoothingMode.HighQuality;
            g.PixelOffsetMode = PixelOffsetMode.HighQuality;
            g.CompositingQuality = CompositingQuality.HighQuality;

            float S(float v) => v * _dpiScale;
            int Si(float v) => (int)Math.Round(v * _dpiScale);

            var (title, step, detail, estimateMs) = Snapshot();
            if (string.IsNullOrWhiteSpace(detail)) detail = Loc.T("loading.detailDefault"); // Loc is thread-safe (immutable dictionaries)

            // Border + accent bar
            int bw = Math.Max(1, (int)Math.Round(_dpiScale));
            using (var borderPen = new Pen(Color.FromArgb(203, 213, 225), bw) { Alignment = PenAlignment.Inset })
                g.DrawRectangle(borderPen, 0, 0, Width - 1, Height - 1);
            using (var accent = new SolidBrush(Color.FromArgb(16, 149, 193)))
            {
                if (estimateMs is double est && est > 0)
                {
                    using (var trackBrush = new SolidBrush(Color.FromArgb(226, 232, 240)))
                        g.FillRectangle(trackBrush, 0, 0, Width, Si(4));
                    double ratio = _stopwatch.Elapsed.TotalMilliseconds / est;
                    double progress = ratio <= 0.95
                        ? Math.Clamp(ratio, 0.0, 0.95)
                        : 0.95 + 0.04 * (1.0 - Math.Exp(-(ratio - 0.95) * 2.0));
                    g.FillRectangle(accent, 0, 0, Math.Max(Si(4), (int)Math.Round(Width * progress)), Si(4));
                }
                else
                {
                    g.FillRectangle(accent, 0, 0, Width, Si(4));
                }
            }

            // Spinner
            var spinRect = new RectangleF(S(22), S(36), S(34), S(34));
            using (var track = new Pen(Color.FromArgb(226, 232, 240), S(3.5f)))
                g.DrawEllipse(track, spinRect);
            float start = (float)((_spinClock.ElapsedMilliseconds * 0.36) % 360.0);
            using (var sweep = new Pen(Color.FromArgb(16, 149, 193), S(3.5f)) { StartCap = LineCap.Round, EndCap = LineCap.Round })
                g.DrawArc(sweep, spinRect, start, 110f);

            // Text (GDI ClearType on a solid white background)
            const TextFormatFlags common = TextFormatFlags.SingleLine | TextFormatFlags.NoPrefix | TextFormatFlags.NoPadding;
            const TextFormatFlags left = common | TextFormatFlags.Left | TextFormatFlags.EndEllipsis;
            Rectangle clockRect, titleRect;
            string clock;
            if (estimateMs is double estClock && estClock > 0)
            {
                clock = $"{_stopwatch.Elapsed.TotalSeconds:0.0}s / ~{estClock / 1000.0:0.0}s";
                clockRect = new Rectangle(Si(300), Si(20), Si(120), Si(20));
                titleRect = new Rectangle(Si(72), Si(18), Si(224), Si(24));
            }
            else
            {
                clock = $"{_stopwatch.Elapsed.TotalSeconds:0.0}s";
                clockRect = new Rectangle(Si(340), Si(20), Si(80), Si(20));
                titleRect = new Rectangle(Si(72), Si(18), Si(264), Si(24));
            }

            TextRenderer.DrawText(g, clock, _clockFont, clockRect,
                Color.FromArgb(148, 163, 184), common | TextFormatFlags.Right);
            TextRenderer.DrawText(g, title, _titleFont, titleRect,
                Color.FromArgb(15, 23, 42), left);
            TextRenderer.DrawText(g, step, _stepFont, new Rectangle(Si(72), Si(45), Si(348), Si(22)),
                Color.FromArgb(51, 65, 85), left);
            TextRenderer.DrawText(g, detail, _detailFont, new Rectangle(Si(72), Si(70), Si(348), Si(20)),
                Color.FromArgb(100, 116, 139), left);
        }

        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                _timer.Dispose();
                DisposeFonts();
            }
            base.Dispose(disposing);
        }
    }
}
