namespace SKSB_App.Core.RuntimeUI.Dictionary;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Localization;

/// <summary>Outcome of one "Pick Text on Screen" session.</summary>
public sealed class ScreenPickResult
{
    public bool Cancelled { get; init; }
    public bool IsExcelGridCell { get; init; }
    public Point ScreenPoint { get; init; }
    public string? ControlTag { get; init; }
    public string? CapturedText { get; init; }
    public string? HelpText { get; init; }

    /// <summary>Text-bearing elements near the click (WinForms + UI Automation), ordered by ascending distance.</summary>
    public List<ScreenPickCandidate> NearbyCandidates { get; init; } = new();
}

/// <summary>One text-bearing element found near the picked point (proximity sampling).</summary>
public sealed class ScreenPickCandidate
{
    public string? DirectTag { get; init; }
    public string? ElementName { get; init; }
    public string? ElementHelpText { get; init; }
    public double DistancePx { get; init; }
}

/// <summary>
/// Modal "click any text on screen" picker. A click-through highlight follows the element under the cursor,
/// a top-center banner previews the detected text, and global low-level hooks swallow the picking click so
/// the underlying Ribbon button / control never fires. Esc or right-click cancels.
///
/// UI Automation is reached through a hand-written <c>UIAutomationCore.dll</c> COM interop (no NuGet) and is
/// always queried from thread-pool (MTA) threads, so providers that live on this UI thread (Excel's Ribbon,
/// our own WinForms windows) can answer while the banner's modal loop keeps pumping.
/// </summary>
public static class ScreenTextPickerSession
{
    public static ScreenPickResult Pick(string langCode)
    {
        try
        {
            using var session = new PickerSession(langCode);
            return session.Run();
        }
        catch (Exception ex)
        {
            Debug.WriteLine($"[ScreenPicker] Pick failed: {ex.Message}");
            return new ScreenPickResult { Cancelled = true };
        }
    }

    // =========================================================================
    // 1. SESSION
    // =========================================================================
    private sealed class PickerSession : IDisposable
    {
        private readonly PickerBannerForm _banner;
        private readonly HighlightOverlayForm _overlay;
        private readonly System.Windows.Forms.Timer _timer;

        // Pinned for the hook lifetime (never unpinned lambdas).
        private readonly NativeMethods.LowLevelProc _mouseProc;
        private readonly NativeMethods.LowLevelProc _keyboardProc;
        private IntPtr _mouseHook = IntPtr.Zero;
        private IntPtr _keyboardHook = IntPtr.Zero;

        private volatile bool _swallowNextLeftUp;
        private volatile bool _swallowNextRightUp;
        private volatile bool _completing;
        private int _inspectBusy;
        private bool _disposed;

        private ScreenPickResult _result = new() { Cancelled = true };

        public PickerSession(string langCode)
        {
            _ = langCode; // Banner texts follow Loc's active language; kept for future per-language previews.
            _banner = new PickerBannerForm();
            _overlay = new HighlightOverlayForm();
            _timer = new System.Windows.Forms.Timer { Interval = 60 };
            _timer.Tick += OnTimerTick;
            _mouseProc = MouseHookCallback;
            _keyboardProc = KeyboardHookCallback;

            _banner.Shown += OnBannerShown;
            _banner.FormClosed += (s, e) => StopAll();
        }

        public ScreenPickResult Run()
        {
            try
            {
                _banner.PositionAtTopCenter(Cursor.Position);
                _banner.ShowDialog();
            }
            finally
            {
                StopAll();
            }
            return _result;
        }

        private void OnBannerShown(object? sender, EventArgs e)
        {
            try { _overlay.Show(); _overlay.Visible = false; } catch { }
            InstallHooks();
            _timer.Start();
        }

        private void InstallHooks()
        {
            IntPtr hMod = NativeMethods.GetModuleHandle(null);
            _mouseHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_MOUSE_LL, _mouseProc, hMod, 0);
            _keyboardHook = NativeMethods.SetWindowsHookEx(NativeMethods.WH_KEYBOARD_LL, _keyboardProc, hMod, 0);
            if (_mouseHook == IntPtr.Zero)
            {
                Debug.WriteLine($"[ScreenPicker] Mouse hook failed (Win32 error {Marshal.GetLastWin32Error()}).");
                // Without the mouse hook a click would reach the underlying control: abort safely.
                SafeBeginInvoke(Cancel);
            }
        }

        private void Unhook()
        {
            if (_mouseHook != IntPtr.Zero)
            {
                try { NativeMethods.UnhookWindowsHookEx(_mouseHook); } catch { }
                _mouseHook = IntPtr.Zero;
            }
            if (_keyboardHook != IntPtr.Zero)
            {
                try { NativeMethods.UnhookWindowsHookEx(_keyboardHook); } catch { }
                _keyboardHook = IntPtr.Zero;
            }
        }

        private void StopAll()
        {
            try { _timer.Stop(); } catch { }
            Unhook();
            try { if (!_overlay.IsDisposed) _overlay.Hide(); } catch { }
        }

        public void Dispose()
        {
            if (_disposed) return;
            _disposed = true;
            StopAll();
            try { _timer.Dispose(); } catch { }
            try { _overlay.Close(); _overlay.Dispose(); } catch { }
            try { _banner.Dispose(); } catch { }
            UiaInspector.Release();
        }

        // =====================================================================
        // 2. LOW-LEVEL HOOKS (keep fast: only capture + marshal)
        // =====================================================================
        private IntPtr MouseHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                int msg = wParam.ToInt32();
                switch (msg)
                {
                    case NativeMethods.WM_LBUTTONDOWN:
                    {
                        var info = Marshal.PtrToStructure<NativeMethods.MSLLHOOKSTRUCT>(lParam);
                        var pt = new Point(info.pt.X, info.pt.Y);
                        _swallowNextLeftUp = true;
                        SafeBeginInvoke(() => BeginComplete(pt));
                        return (IntPtr)1;
                    }
                    case NativeMethods.WM_LBUTTONUP:
                        if (_swallowNextLeftUp)
                        {
                            _swallowNextLeftUp = false;
                            return (IntPtr)1;
                        }
                        break;
                    case NativeMethods.WM_RBUTTONDOWN:
                        // Cancel immediately; hooks stay installed until the matching up is swallowed too,
                        // so Excel never sees a stray WM_RBUTTONUP (context menu).
                        _swallowNextRightUp = true;
                        SafeBeginInvoke(Cancel);
                        return (IntPtr)1;
                    case NativeMethods.WM_RBUTTONUP:
                        if (_swallowNextRightUp)
                        {
                            _swallowNextRightUp = false;
                            return (IntPtr)1;
                        }
                        break;
                }
            }
            return NativeMethods.CallNextHookEx(_mouseHook, nCode, wParam, lParam);
        }

        private IntPtr KeyboardHookCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0 && wParam.ToInt32() == NativeMethods.WM_KEYDOWN)
            {
                var info = Marshal.PtrToStructure<NativeMethods.KBDLLHOOKSTRUCT>(lParam);
                if (info.vkCode == NativeMethods.VK_ESCAPE)
                {
                    SafeBeginInvoke(Cancel);
                    return (IntPtr)1;
                }
            }
            return NativeMethods.CallNextHookEx(_keyboardHook, nCode, wParam, lParam);
        }

        private void SafeBeginInvoke(Action action)
        {
            try
            {
                if (!_banner.IsDisposed && _banner.IsHandleCreated)
                    _banner.BeginInvoke(action);
            }
            catch { }
        }

        // =====================================================================
        // 3. HOVER TRACKING (WinForms timer, UI thread)
        // =====================================================================
        private void OnTimerTick(object? sender, EventArgs e)
        {
            if (_completing || _banner.IsDisposed) return;

            Point pt = Cursor.Position;
            _banner.PositionAtTopCenter(pt);

            if (_banner.Bounds.Contains(pt))
            {
                _overlay.HideHighlight();
                _banner.SetPreview(PreviewKind.Hint, null);
                return;
            }

            var local = InspectLocal(pt);
            if (local.IsOwnPickerWindow)
            {
                _overlay.HideHighlight();
                return;
            }

            if (Interlocked.CompareExchange(ref _inspectBusy, 1, 0) != 0) return;

            Task.Run(() => UiaInspector.Inspect(pt)).ContinueWith(t =>
            {
                Interlocked.Exchange(ref _inspectBusy, 0);
                var uia = t.Status == TaskStatus.RanToCompletion ? t.Result : default;
                SafeBeginInvoke(() => ApplyHover(pt, local, uia));
            }, TaskScheduler.Default);
        }

        private void ApplyHover(Point pt, LocalHit local, UiaHit uia)
        {
            if (_completing || _banner.IsDisposed) return;
            if (_banner.Bounds.Contains(Cursor.Position)) { _overlay.HideHighlight(); return; }

            Rectangle rect = uia.Bounds;
            if (rect.IsEmpty && local.ControlBounds.HasValue) rect = local.ControlBounds.Value;
            _overlay.ShowHighlight(rect);

            if (local.IsExcelGrid)
            {
                _banner.SetPreview(PreviewKind.ExcelCell, null);
                return;
            }

            string preview = FirstNonEmpty(uia.Name, local.Text, uia.HelpText);
            _banner.SetPreview(preview.Length > 0 ? PreviewKind.Found : PreviewKind.Hint, preview);
        }

        // =====================================================================
        // 4. PICK COMPLETION / CANCEL
        // =====================================================================
        private void BeginComplete(Point pt)
        {
            if (_completing || _banner.IsDisposed) return;

            // Clicks on the HUD banner itself are swallowed but ignored.
            if (_banner.Bounds.Contains(pt)) return;

            var local = InspectLocal(pt);
            if (local.IsOwnPickerWindow) return;

            _completing = true;
            _timer.Stop();
            _overlay.HideHighlight();

            // Proximity scans run once, on click only (never on hover).
            List<ScreenPickCandidate> winFormsNearby = ScanNearbyWinForms(pt);
            Rectangle bannerBounds = _banner.Bounds;

            var uiaTask = Task.Run(() =>
            {
                UiaHit exact = UiaInspector.Inspect(pt);
                List<ScreenPickCandidate> nearby = SampleNearbyUia(pt, bannerBounds);
                return (Exact: exact, Nearby: nearby);
            });
            uiaTask.ContinueWith(t =>
            {
                UiaHit uia = default;
                List<ScreenPickCandidate> uiaNearby = new();
                if (t.Status == TaskStatus.RanToCompletion)
                {
                    uia = t.Result.Exact;
                    uiaNearby = t.Result.Nearby;
                }
                SafeBeginInvoke(() => FinishPick(pt, local, uia, winFormsNearby, uiaNearby));
            }, TaskScheduler.Default);
        }

        private void FinishPick(Point pt, LocalHit local, UiaHit uia,
            List<ScreenPickCandidate> winFormsNearby, List<ScreenPickCandidate> uiaNearby)
        {
            string? captured = NullIfBlank(FirstNonEmpty(uia.Name, local.Text));
            string? help = NullIfBlank(uia.HelpText);

            var nearby = new List<ScreenPickCandidate>(winFormsNearby.Count + uiaNearby.Count);
            nearby.AddRange(winFormsNearby);
            nearby.AddRange(uiaNearby);
            nearby.Sort((a, b) => a.DistancePx.CompareTo(b.DistancePx));

            _result = new ScreenPickResult
            {
                Cancelled = false,
                IsExcelGridCell = local.IsExcelGrid,
                ScreenPoint = pt,
                ControlTag = local.Tag,
                CapturedText = captured,
                HelpText = help,
                NearbyCandidates = nearby
            };
            CloseBanner(DialogResult.OK);
        }

        // =====================================================================
        // 4b. PROXIMITY SAMPLING (click only)
        // =====================================================================
        private const double WinFormsProximityRadiusPx = 90.0;

        /// <summary>Ring offsets (radius 36px and 72px, 8 compass points each) sampled through UI Automation.</summary>
        private static readonly Point[] UiaSampleOffsets =
        {
            new(36, 0), new(-36, 0), new(0, 36), new(0, -36),
            new(25, 25), new(25, -25), new(-25, 25), new(-25, -25),
            new(72, 0), new(-72, 0), new(0, 72), new(0, -72),
            new(51, 51), new(51, -51), new(-51, 51), new(-51, -51)
        };

        /// <summary>UI thread: walks every visible control of the clicked in-process form within 90px of the click.</summary>
        private List<ScreenPickCandidate> ScanNearbyWinForms(Point pt)
        {
            var list = new List<ScreenPickCandidate>();
            try
            {
                IntPtr hwnd = NativeMethods.WindowFromPoint(new NativeMethods.POINT { X = pt.X, Y = pt.Y });
                if (hwnd == IntPtr.Zero) return list;

                Control? ctl = Control.FromHandle(hwnd);
                if (ctl == null) return list;

                Control root = (Control?)ctl.FindForm() ?? ctl;
                if (root is not Form)
                {
                    while (root.Parent != null) root = root.Parent;
                }
                if (ReferenceEquals(root, _banner) || ReferenceEquals(root, _overlay)) return list;

                var stack = new Stack<Control>();
                stack.Push(root);
                while (stack.Count > 0)
                {
                    Control parent = stack.Pop();
                    foreach (Control c in parent.Controls)
                    {
                        if (!c.Visible) continue;
                        stack.Push(c);

                        Rectangle r;
                        try { r = c.RectangleToScreen(c.ClientRectangle); }
                        catch { continue; }
                        if (r.Width <= 0 || r.Height <= 0) continue;

                        double dist = DistanceToRect(pt, r);
                        if (dist > WinFormsProximityRadiusPx) continue;

                        string? tag = c.Tag is string t
                            && (t.StartsWith("ui::", StringComparison.OrdinalIgnoreCase)
                                || t.StartsWith("tooltips::", StringComparison.OrdinalIgnoreCase))
                            ? t
                            : null;
                        string? text;
                        try { text = NullIfBlank(c.Text?.Trim()); } catch { text = null; }
                        if (tag == null && text == null) continue;

                        list.Add(new ScreenPickCandidate
                        {
                            DirectTag = tag,
                            ElementName = text,
                            ElementHelpText = null,
                            DistancePx = dist
                        });
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ScreenPicker] WinForms proximity scan failed: {ex.Message}");
            }
            return list;
        }

        /// <summary>Thread-pool (MTA): samples two rings of points around the click through UI Automation.</summary>
        private static List<ScreenPickCandidate> SampleNearbyUia(Point pt, Rectangle bannerBounds)
        {
            var best = new Dictionary<(string, string), ScreenPickCandidate>();
            foreach (Point offset in UiaSampleOffsets)
            {
                var sample = new Point(pt.X + offset.X, pt.Y + offset.Y);
                if (bannerBounds.Contains(sample)) continue;

                UiaHit hit = UiaInspector.Inspect(sample);
                if (string.IsNullOrWhiteSpace(hit.Name) && string.IsNullOrWhiteSpace(hit.HelpText)) continue;

                double dist = !hit.Bounds.IsEmpty
                    ? DistanceToRect(pt, hit.Bounds)
                    : Math.Sqrt(offset.X * offset.X + offset.Y * offset.Y);

                var key = ((hit.Name ?? string.Empty).Trim().ToLowerInvariant(),
                           (hit.HelpText ?? string.Empty).Trim().ToLowerInvariant());
                if (best.TryGetValue(key, out var existing) && existing.DistancePx <= dist) continue;

                best[key] = new ScreenPickCandidate
                {
                    DirectTag = null,
                    ElementName = NullIfBlank(hit.Name?.Trim()),
                    ElementHelpText = NullIfBlank(hit.HelpText?.Trim()),
                    DistancePx = dist
                };
            }
            return new List<ScreenPickCandidate>(best.Values);
        }

        private static double DistanceToRect(Point pt, Rectangle r)
        {
            double dx = Math.Max(0, Math.Max(r.Left - pt.X, pt.X - r.Right));
            double dy = Math.Max(0, Math.Max(r.Top - pt.Y, pt.Y - r.Bottom));
            return Math.Sqrt(dx * dx + dy * dy);
        }

        private void Cancel()
        {
            if (_completing) return;
            _completing = true;
            _result = new ScreenPickResult { Cancelled = true };
            CloseBanner(DialogResult.Cancel);
        }

        /// <summary>
        /// Closes the banner once any swallowed button-down has had its matching button-up swallowed as well
        /// (bounded by a short deadline), so no orphan up-click ever reaches the control underneath.
        /// </summary>
        private void CloseBanner(DialogResult result)
        {
            _timer.Stop();
            try { _overlay.HideHighlight(); } catch { }

            DateTime deadline = DateTime.UtcNow.AddMilliseconds(800);
            void TryClose()
            {
                if ((_swallowNextLeftUp || _swallowNextRightUp) && DateTime.UtcNow < deadline)
                {
                    var wait = new System.Windows.Forms.Timer { Interval = 25 };
                    wait.Tick += (s, e) => { wait.Stop(); wait.Dispose(); TryClose(); };
                    wait.Start();
                    return;
                }

                StopAll();
                try
                {
                    if (!_banner.IsDisposed)
                    {
                        _banner.DialogResult = result;
                        _banner.Close();
                    }
                }
                catch { }
            }
            TryClose();
        }

        // =====================================================================
        // 5. LOCAL (UI-THREAD) HIT TEST: WinForms tags/text + Excel grid detection
        // =====================================================================
        private LocalHit InspectLocal(Point pt)
        {
            var hit = new LocalHit();
            try
            {
                IntPtr hwnd = NativeMethods.WindowFromPoint(new NativeMethods.POINT { X = pt.X, Y = pt.Y });
                if (hwnd == IntPtr.Zero) return hit;

                if (hwnd == SafeHandle(_overlay) || hwnd == SafeHandle(_banner))
                {
                    hit.IsOwnPickerWindow = true;
                    return hit;
                }

                hit.IsExcelGrid = IsExcelGridWindow(hwnd);

                Control? ctl = Control.FromHandle(hwnd);
                if (ctl == null) return hit;

                Form? form = ctl.FindForm();
                if (ReferenceEquals(form, _banner) || ReferenceEquals(form, _overlay))
                {
                    hit.IsOwnPickerWindow = true;
                    return hit;
                }

                // Drill down to the deepest visible child under the point.
                Control deepest = ctl;
                for (int guard = 0; guard < 32; guard++)
                {
                    Control? child;
                    try { child = deepest.GetChildAtPoint(deepest.PointToClient(pt), GetChildAtPointSkip.Invisible); }
                    catch { child = null; }
                    if (child == null || ReferenceEquals(child, deepest)) break;
                    deepest = child;
                }

                try { hit.ControlBounds = deepest.RectangleToScreen(deepest.ClientRectangle); } catch { }

                for (Control? c = deepest; c != null; c = c.Parent)
                {
                    if (c.Tag is string tag
                        && (tag.StartsWith("ui::", StringComparison.OrdinalIgnoreCase)
                            || tag.StartsWith("tooltips::", StringComparison.OrdinalIgnoreCase)))
                    {
                        hit.Tag = tag;
                        break;
                    }
                }

                string text = deepest.Text ?? string.Empty;
                if (string.IsNullOrWhiteSpace(text) && deepest is not TextBoxBase && deepest.Parent != null && hit.Tag != null)
                {
                    // Tag came from an ancestor: prefer its text as the fallback.
                    for (Control? c = deepest.Parent; c != null; c = c.Parent)
                    {
                        if (c.Tag is string t && string.Equals(t, hit.Tag, StringComparison.Ordinal)) { text = c.Text ?? string.Empty; break; }
                    }
                }
                hit.Text = text;
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ScreenPicker] Local hit-test failed: {ex.Message}");
            }
            return hit;
        }

        private static IntPtr SafeHandle(Form f)
        {
            try { return !f.IsDisposed && f.IsHandleCreated ? f.Handle : IntPtr.Zero; }
            catch { return IntPtr.Zero; }
        }

        private static bool IsExcelGridWindow(IntPtr hwnd)
        {
            IntPtr current = hwnd;
            for (int level = 0; level <= 3 && current != IntPtr.Zero; level++)
            {
                var sb = new StringBuilder(64);
                if (NativeMethods.GetClassName(current, sb, sb.Capacity) > 0
                    && string.Equals(sb.ToString(), "EXCEL7", StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
                current = NativeMethods.GetParent(current);
            }
            return false;
        }

        private static string FirstNonEmpty(params string?[] values)
        {
            foreach (string? v in values)
                if (!string.IsNullOrWhiteSpace(v)) return v.Trim();
            return string.Empty;
        }

        private static string? NullIfBlank(string? s) => string.IsNullOrWhiteSpace(s) ? null : s;
    }

    private sealed class LocalHit
    {
        public bool IsOwnPickerWindow { get; set; }
        public bool IsExcelGrid { get; set; }
        public string? Tag { get; set; }
        public string? Text { get; set; }
        public Rectangle? ControlBounds { get; set; }
    }

    private readonly struct UiaHit
    {
        public UiaHit(string? name, string? helpText, Rectangle bounds)
        {
            Name = name;
            HelpText = helpText;
            Bounds = bounds;
        }

        public string? Name { get; }
        public string? HelpText { get; }
        public Rectangle Bounds { get; }
    }

    // =========================================================================
    // 6. UI AUTOMATION (UIAutomationCore.dll COM interop, MTA threads only)
    // =========================================================================
    private static class UiaInspector
    {
        private const int UIA_BoundingRectanglePropertyId = 30001;
        private const int UIA_NamePropertyId = 30005;
        private const int UIA_HelpTextPropertyId = 30013;

        private static readonly object _gate = new();
        private static IUIAutomation? _automation;
        private static bool _unavailable;

        private static IUIAutomation? GetAutomation()
        {
            lock (_gate)
            {
                if (_automation != null || _unavailable) return _automation;
                try
                {
                    _automation = (IUIAutomation)new CUIAutomationClass();
                }
                catch (Exception ex)
                {
                    _unavailable = true;
                    Debug.WriteLine($"[ScreenPicker] UI Automation unavailable: {ex.Message}");
                }
                return _automation;
            }
        }

        public static UiaHit Inspect(Point pt)
        {
            var automation = GetAutomation();
            if (automation == null) return default;

            IUIAutomationElement? element = null;
            try
            {
                int hr = automation.ElementFromPoint(new NativeMethods.POINT { X = pt.X, Y = pt.Y }, out element);
                if (hr < 0 || element == null) return default;

                string? name = GetString(element, UIA_NamePropertyId);
                string? help = GetString(element, UIA_HelpTextPropertyId);
                Rectangle bounds = GetBounds(element);
                return new UiaHit(name, help, bounds);
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"[ScreenPicker] UIA inspect failed: {ex.Message}");
                return default;
            }
            finally
            {
                if (element != null)
                {
                    try { Marshal.FinalReleaseComObject(element); } catch { }
                }
            }
        }

        public static void Release()
        {
            lock (_gate)
            {
                if (_automation == null) return;
                var a = _automation;
                _automation = null;
                try { Marshal.FinalReleaseComObject(a); } catch { }
            }
        }

        private static string? GetString(IUIAutomationElement element, int propertyId)
        {
            try
            {
                int hr = element.GetCurrentPropertyValue(propertyId, out object? value);
                return hr >= 0 ? value as string : null;
            }
            catch
            {
                return null;
            }
        }

        private static Rectangle GetBounds(IUIAutomationElement element)
        {
            try
            {
                int hr = element.GetCurrentPropertyValue(UIA_BoundingRectanglePropertyId, out object? value);
                if (hr < 0 || value is not double[] r || r.Length < 4) return Rectangle.Empty;
                if (double.IsNaN(r[0]) || double.IsInfinity(r[0]) || r[2] <= 0 || r[3] <= 0) return Rectangle.Empty;
                return new Rectangle((int)Math.Round(r[0]), (int)Math.Round(r[1]), (int)Math.Round(r[2]), (int)Math.Round(r[3]));
            }
            catch
            {
                return Rectangle.Empty;
            }
        }
    }

    [ComImport, Guid("FF48DBA4-60EF-4201-AA87-54103EEF594E")]
    private class CUIAutomationClass { }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("30CBE57D-D9D0-452A-AB13-7AC5AC4825EE")]
    private interface IUIAutomation
    {
        // Slots 3-6: declared only to keep the vtable layout exact (never called).
        void CompareElements();
        void CompareRuntimeIds();
        void GetRootElement();
        void ElementFromHandle();

        // Slot 7
        [PreserveSig]
        int ElementFromPoint(NativeMethods.POINT pt, [MarshalAs(UnmanagedType.Interface)] out IUIAutomationElement element);
    }

    [ComImport, InterfaceType(ComInterfaceType.InterfaceIsIUnknown), Guid("D22108AA-8AC5-49A5-837B-37BBB3D7591E")]
    private interface IUIAutomationElement
    {
        // Slots 3-9: declared only to keep the vtable layout exact (never called).
        void SetFocus();
        void GetRuntimeId();
        void FindFirst();
        void FindAll();
        void FindFirstBuildCache();
        void FindAllBuildCache();
        void BuildUpdatedCache();

        // Slot 10
        [PreserveSig]
        int GetCurrentPropertyValue(int propertyId, [MarshalAs(UnmanagedType.Struct)] out object retVal);
    }

    // =========================================================================
    // 7. HUD FORMS
    // =========================================================================
    private enum PreviewKind { Hint, Found, ExcelCell }

    /// <summary>Click-through, non-activating, top-most border drawn around the hovered element.</summary>
    private sealed class HighlightOverlayForm : Form
    {
        private const int BorderThickness = 3;
        private static readonly Color HighlightColor = Color.FromArgb(79, 70, 229);
        private Rectangle _target = Rectangle.Empty;

        public HighlightOverlayForm()
        {
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = Color.Magenta;
            TransparencyKey = Color.Magenta;
            Bounds = new Rectangle(-10000, -10000, 10, 10);
            SetStyle(ControlStyles.OptimizedDoubleBuffer | ControlStyles.AllPaintingInWmPaint | ControlStyles.UserPaint, true);
        }

        protected override bool ShowWithoutActivation => true;

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= NativeMethods.WS_EX_TRANSPARENT | NativeMethods.WS_EX_LAYERED | NativeMethods.WS_EX_TOPMOST
                            | NativeMethods.WS_EX_NOACTIVATE | NativeMethods.WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        public void ShowHighlight(Rectangle rect)
        {
            if (IsDisposed) return;
            if (rect.IsEmpty || rect.Width <= 0 || rect.Height <= 0)
            {
                HideHighlight();
                return;
            }

            var outer = Rectangle.Inflate(rect, BorderThickness, BorderThickness);
            if (outer != _target || !Visible)
            {
                _target = outer;
                Bounds = outer;
                if (!Visible) Visible = true;
                Invalidate();
            }
        }

        public void HideHighlight()
        {
            if (IsDisposed) return;
            _target = Rectangle.Empty;
            if (Visible) Visible = false;
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            e.Graphics.Clear(Color.Magenta);
            using var pen = new Pen(HighlightColor, BorderThickness) { Alignment = PenAlignment.Inset };
            e.Graphics.DrawRectangle(pen, 0, 0, ClientSize.Width - 1, ClientSize.Height - 1);
        }
    }

    /// <summary>Dark top-center pill with the instruction line and a live preview of the detected text.</summary>
    private sealed class PickerBannerForm : Form
    {
        private static readonly Color BgSlate = Color.FromArgb(15, 23, 42);
        private static readonly Color TextMutedLight = Color.FromArgb(148, 163, 184);
        private static readonly Color Emerald = Color.FromArgb(52, 211, 153);

        private readonly Label _lblTitle;
        private readonly Label _lblPreview;
        private PreviewKind _lastKind = (PreviewKind)(-1);
        private string? _lastPreview;
        private Rectangle _lastScreen = Rectangle.Empty;

        public PickerBannerForm()
        {
            SuspendLayout();
            FormBorderStyle = FormBorderStyle.None;
            ShowInTaskbar = false;
            TopMost = true;
            StartPosition = FormStartPosition.Manual;
            BackColor = BgSlate;
            Font = new Font("Segoe UI", 9f, FontStyle.Regular);
            ClientSize = new Size(560, 62);
            Padding = new Padding(18, 8, 18, 8);
            KeyPreview = true;

            _lblTitle = new Label
            {
                Dock = DockStyle.Top,
                Height = 24,
                ForeColor = Color.White,
                Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
                TextAlign = ContentAlignment.MiddleCenter,
                Text = Loc.T("dict.picker.bannerTitle"),
                AutoEllipsis = true
            };
            _lblPreview = new Label
            {
                Dock = DockStyle.Fill,
                ForeColor = TextMutedLight,
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                TextAlign = ContentAlignment.MiddleCenter,
                Text = Loc.T("dict.picker.bannerHint"),
                AutoEllipsis = true
            };
            Controls.Add(_lblPreview);
            Controls.Add(_lblTitle);

            AutoScaleDimensions = new SizeF(96f, 96f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ResumeLayout(false);
        }

        protected override CreateParams CreateParams
        {
            get
            {
                var cp = base.CreateParams;
                cp.ExStyle |= NativeMethods.WS_EX_TOPMOST | NativeMethods.WS_EX_TOOLWINDOW;
                return cp;
            }
        }

        protected override void OnKeyDown(KeyEventArgs e)
        {
            // Fallback when the keyboard hook is unavailable.
            if (e.KeyCode == Keys.Escape)
            {
                DialogResult = DialogResult.Cancel;
                Close();
                return;
            }
            base.OnKeyDown(e);
        }

        protected override void OnSizeChanged(EventArgs e)
        {
            base.OnSizeChanged(e);
            ApplyRoundedRegion();
        }

        private void ApplyRoundedRegion()
        {
            if (Width <= 0 || Height <= 0) return;
            int radius = Math.Min(Height, (int)Math.Round(28 * DeviceDpi / 96f));
            using var path = new GraphicsPath();
            path.AddArc(0, 0, radius, radius, 180, 90);
            path.AddArc(Width - radius - 1, 0, radius, radius, 270, 90);
            path.AddArc(Width - radius - 1, Height - radius - 1, radius, radius, 0, 90);
            path.AddArc(0, Height - radius - 1, radius, radius, 90, 90);
            path.CloseFigure();
            var old = Region;
            Region = new Region(path);
            old?.Dispose();
        }

        public void PositionAtTopCenter(Point cursor)
        {
            if (IsDisposed) return;
            Rectangle area;
            try { area = Screen.FromPoint(cursor).WorkingArea; }
            catch { area = Screen.PrimaryScreen?.WorkingArea ?? new Rectangle(0, 0, 1280, 720); }
            if (area == _lastScreen) return;
            _lastScreen = area;

            int top = area.Top + (int)Math.Round(16 * DeviceDpi / 96f);
            Location = new Point(area.Left + (area.Width - Width) / 2, top);
        }

        public void SetPreview(PreviewKind kind, string? preview)
        {
            if (IsDisposed) return;
            if (kind == _lastKind && string.Equals(preview, _lastPreview, StringComparison.Ordinal)) return;
            _lastKind = kind;
            _lastPreview = preview;

            switch (kind)
            {
                case PreviewKind.Found:
                    string single = LocalizationTextMatcher.Normalize(preview);
                    if (single.Length > 70) single = single.Substring(0, 67) + "...";
                    _lblPreview.Text = Loc.T("dict.picker.found", single);
                    _lblPreview.ForeColor = Emerald;
                    break;
                case PreviewKind.ExcelCell:
                    _lblPreview.Text = Loc.T("dict.picker.excelCell");
                    _lblPreview.ForeColor = Emerald;
                    break;
                default:
                    _lblPreview.Text = Loc.T("dict.picker.bannerHint");
                    _lblPreview.ForeColor = TextMutedLight;
                    break;
            }
        }
    }

    // =========================================================================
    // 8. NATIVE METHODS
    // =========================================================================
    private static class NativeMethods
    {
        public const int WH_KEYBOARD_LL = 13;
        public const int WH_MOUSE_LL = 14;

        public const int WM_KEYDOWN = 0x0100;
        public const int WM_LBUTTONDOWN = 0x0201;
        public const int WM_LBUTTONUP = 0x0202;
        public const int WM_RBUTTONDOWN = 0x0204;
        public const int WM_RBUTTONUP = 0x0205;
        public const int VK_ESCAPE = 0x1B;

        public const int WS_EX_TOPMOST = 0x00000008;
        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int WS_EX_TOOLWINDOW = 0x00000080;
        public const int WS_EX_LAYERED = 0x00080000;
        public const int WS_EX_NOACTIVATE = 0x08000000;

        public delegate IntPtr LowLevelProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct MSLLHOOKSTRUCT
        {
            public POINT pt;
            public uint mouseData;
            public uint flags;
            public uint time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [DllImport("user32.dll", SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int idHook, LowLevelProc lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hhk);

        [DllImport("user32.dll")]
        public static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetModuleHandle(string? lpModuleName);

        [DllImport("user32.dll")]
        public static extern IntPtr WindowFromPoint(POINT point);

        [DllImport("user32.dll")]
        public static extern IntPtr GetParent(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern int GetClassName(IntPtr hWnd, StringBuilder lpClassName, int nMaxCount);
    }
}
