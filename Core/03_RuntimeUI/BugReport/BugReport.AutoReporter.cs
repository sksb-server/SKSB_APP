namespace SKSB_App.Core.RuntimeUI.BugReport;

using System;
#if !DEBUG
using System.Collections.Concurrent;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;
#endif

/// <summary>
/// Release-only automatic bug reporter. Captures Excel-DNA runtime errors and unhandled .NET/WinForms
/// exceptions and submits them to <c>public.bug_reports</c>. All methods are no-ops in Debug builds.
/// Never touches Excel COM, so it is safe from any thread, during SyncMacro errors and at shutdown.
/// </summary>
public static class AutoBugReporter
{
    // =========================================================================
    // 1. PUBLIC API
    // =========================================================================
    public static void InstallForExcelAddIn(string? sourceModule = null)
    {
#if !DEBUG
        try
        {
            if (!string.IsNullOrWhiteSpace(sourceModule)) _sourceModule = sourceModule!;

            lock (_lock)
            {
                if (_listener == null)
                {
                    try
                    {
                        var listener = new ExcelDnaErrorTraceListener();
                        var listeners = GetIntegrationTraceSource()?.Listeners;
                        if (listeners != null)
                        {
                            listeners.Add(listener);
                            _listener = listener;
                        }
                    }
                    catch (Exception ex)
                    {
                        AppLogger.Warn($"[AutoBugReporter] Trace listener install failed: {ex.Message}");
                    }
                }
            }

            InstallGlobalHooks();
        }
        catch { }
#endif
    }

    public static void UninstallForExcelAddIn()
    {
#if !DEBUG
        lock (_lock)
        {
            try
            {
                if (_listener != null)
                {
                    GetIntegrationTraceSource()?.Listeners.Remove(_listener);
                    _listener = null;
                }
            }
            catch { }
        }
#endif
    }

    public static void InstallForPortal()
    {
#if !DEBUG
        try
        {
            _sourceModule = "Portal";
            try { Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException); } catch { }
            InstallGlobalHooks();
        }
        catch { }
#endif
    }

    public static void ReportException(string category, Exception ex, string? extraContext = null, bool waitBriefly = false)
    {
#if !DEBUG
        try
        {
            if (ex == null) return;
            var sb = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(extraContext)) sb.AppendLine(extraContext).AppendLine();
            sb.AppendLine(ex.ToString());
            for (var inner = ex.InnerException; inner != null; inner = inner.InnerException)
                sb.AppendLine().AppendLine($"--- Inner: {inner.GetType().FullName}: {inner.Message}");
            SubmitAutoReport(category, $"{ex.GetType().Name}: {ex.Message}", sb.ToString(), waitBriefly);
        }
        catch { }
#endif
    }

#if !DEBUG
    // =========================================================================
    // 2. STATE
    // =========================================================================
    private static readonly object _lock = new();
    private static string _sourceModule = "Core";
    private static ExcelDnaErrorTraceListener? _listener;
    private static bool _globalHooksInstalled;

    private static readonly ConcurrentDictionary<string, DateTime> _recent = new();
    private static readonly TimeSpan Cooldown = TimeSpan.FromMinutes(5);
    private const int MaxReportsPerSession = 10;
    private static int _reportCount;

    [ThreadStatic] private static bool _isReporting;

    private static string? _cachedVersion;
    private static string CachedVersion
    {
        get
        {
            if (_cachedVersion != null) return _cachedVersion;
            try { _cachedVersion = BugReportRepository.ResolveAppVersion(); }
            catch { _cachedVersion = string.Empty; }
            return _cachedVersion;
        }
    }

    // =========================================================================
    // 3. GLOBAL .NET HOOKS
    // =========================================================================
    private static void InstallGlobalHooks()
    {
        lock (_lock)
        {
            if (_globalHooksInstalled) return;
            _globalHooksInstalled = true;
        }

        try
        {
            AppDomain.CurrentDomain.UnhandledException += (_, e) =>
            {
                if (e.ExceptionObject is Exception ex)
                    ReportException("UnhandledException", ex, $"IsTerminating={e.IsTerminating}", waitBriefly: e.IsTerminating);
                else
                    SubmitAutoReport("UnhandledException", "Non-Exception object thrown",
                        $"IsTerminating={e.IsTerminating}{Environment.NewLine}{e.ExceptionObject}", e.IsTerminating);
            };
        }
        catch { }

        try
        {
            TaskScheduler.UnobservedTaskException += (_, e) =>
                ReportException("UnobservedTaskException", e.Exception, waitBriefly: false);
        }
        catch { }

        try
        {
            Application.ThreadException += (_, e) =>
                ReportException("UIThreadException", e.Exception, waitBriefly: false);
        }
        catch { }
    }

    // =========================================================================
    // 4. SUBMISSION
    // =========================================================================
    private static void SubmitAutoReport(string category, string summary, string details, bool waitBriefly)
    {
        if (_isReporting) return;
        _isReporting = true;
        try
        {
            summary ??= string.Empty;
            details ??= string.Empty;

            if (Mentions(summary) || Mentions(details)) return;
            if (ServerEnvironment.IsOffline) return;

            string singleLine = summary.Replace('\r', ' ').Replace('\n', ' ').Trim();
            string normalized = singleLine.Length > 120 ? singleLine[..120] : singleLine;
            string key = $"{category}|{normalized}";
            DateTime now = DateTime.UtcNow;

            if (_recent.TryGetValue(key, out DateTime last) && now - last < Cooldown) return;
            if (Volatile.Read(ref _reportCount) >= MaxReportsPerSession) return;

            _recent[key] = now;
            if (Interlocked.Increment(ref _reportCount) > MaxReportsPerSession) return;

            var ctx = new BugReportContext
            {
                SourceModule = _sourceModule,
                ActiveSheet = string.Empty,
                AppVersion = CachedVersion,
                IsOffline = false,
                RuntimeLogTail = SafeLogTail()
            };

            string title = $"[Auto] {category}: {singleLine}".Trim();
            if (title.Length > 200) title = title[..200];

            var desc = new StringBuilder();
            desc.AppendLine($"Automatic report at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
            desc.AppendLine($"Category: {category}");
            desc.AppendLine($"Process: {Environment.ProcessId}");
            try { desc.AppendLine($"XLL: {ExcelDnaUtil.XllPath}"); } catch { }
            desc.AppendLine().AppendLine(details);

            string description = desc.ToString();
            var task = Task.Run(() =>
            {
                try
                {
                    BugReportRepository.SubmitBugReport(title, description, ctx, Array.Empty<BugReportAttachmentInput>());
                }
                catch (Exception ex)
                {
                    try { AppLogger.Warn($"[AutoBugReporter] Submit failed: {ex.Message}"); } catch { }
                }
            });

            if (waitBriefly)
            {
                try { task.Wait(TimeSpan.FromSeconds(3)); } catch { }
            }
        }
        catch (Exception ex)
        {
            try { AppLogger.Warn($"[AutoBugReporter] Report failed: {ex.Message}"); } catch { }
        }
        finally
        {
            _isReporting = false;
        }
    }

    private static bool Mentions(string text) =>
        text.Contains("BugReportRepository", StringComparison.Ordinal) ||
        text.Contains("AutoBugReporter", StringComparison.Ordinal);

    /// <summary>Resolved by reflection: the accessor is not part of the public ExcelDna 1.9.0 surface.</summary>
    private static TraceSource? GetIntegrationTraceSource()
    {
        const System.Reflection.BindingFlags flags =
            System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.NonPublic |
            System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.DeclaredOnly;
        Type[] types;
        try { types = typeof(ExcelDnaUtil).Assembly.GetTypes(); }
        catch (System.Reflection.ReflectionTypeLoadException rtle) { types = rtle.Types.Where(t => t != null).ToArray()!; }

        foreach (Type t in types)
        {
            var m = t.GetMethod("GetIntegrationTraceSource", flags, null, Type.EmptyTypes, null);
            if (m != null && m.Invoke(null, null) is TraceSource ts) return ts;
        }
        return null;
    }

    private static string SafeLogTail()
    {
        try { return BugReportRepository.CaptureRuntimeLogTail(200); }
        catch { return string.Empty; }
    }

    // =========================================================================
    // 5. EXCEL-DNA TRACE LISTENER
    // =========================================================================
    private sealed class ExcelDnaErrorTraceListener : TraceListener
    {
        public override void TraceEvent(TraceEventCache? eventCache, string source, TraceEventType eventType, int id, string? message)
        {
            if (eventType is not (TraceEventType.Error or TraceEventType.Critical)) return;
            Handle(source, eventType, message ?? string.Empty);
        }

        public override void TraceEvent(TraceEventCache? eventCache, string source, TraceEventType eventType, int id, string? format, params object?[]? args)
        {
            if (eventType is not (TraceEventType.Error or TraceEventType.Critical)) return;
            string text;
            try { text = args != null && args.Length > 0 && format != null ? string.Format(format, args) : format ?? string.Empty; }
            catch { text = format ?? string.Empty; }
            Handle(source, eventType, text);
        }

        public override void Write(string? message) { }
        public override void WriteLine(string? message) { }

        private static void Handle(string source, TraceEventType eventType, string text)
        {
            try
            {
                string trimmed = text.Trim();
                int nl = trimmed.IndexOfAny(new[] { '\r', '\n' });
                string firstLine = nl >= 0 ? trimmed[..nl] : trimmed;
                string details = $"Source: {source}{Environment.NewLine}Level: {eventType}{Environment.NewLine}{Environment.NewLine}{trimmed}";
                SubmitAutoReport("Excel-DNA", firstLine, details, waitBriefly: false);
            }
            catch { }
        }
    }
#endif
}
