namespace SKSB_App.Core.Diagnostics;

using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading;
using SKSB_App.Core.Bootstrap.Storage;

public static class AppLogger
{
    private static readonly object _lock = new();
    private static string? _logFilePath;
    private static int _indentLevel = 0;

    public static string LogFilePath
    {
        get
        {
            if (_logFilePath == null)
            {
                try
                {
                    string dir = FolderRegistry.Data.EnsureExists();
                    _logFilePath = Path.Combine(dir, "SKSB_Runtime_Trace.log");
                }
                catch
                {
                    _logFilePath = Path.Combine(Path.GetTempPath(), "SKSB_Runtime_Trace.log");
                }
            }
            return _logFilePath;
        }
    }

    public static void ResetSessionLog(string sessionTitle)
    {
        lock (_lock)
        {
            try
            {
                _indentLevel = 0;
                var sb = new StringBuilder();
                sb.AppendLine("================================================================================");
                sb.AppendLine($" SKSB RUNTIME TRACE LOG - {sessionTitle}");
                sb.AppendLine($" Started : {DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}");
                sb.AppendLine($" Log Path: {LogFilePath}");
                sb.AppendLine("================================================================================");
                File.WriteAllText(LogFilePath, sb.ToString(), Encoding.UTF8);
            }
            catch { }
        }
    }

    public static void Info(
        string message,
        [CallerMemberName] string member = "",
        [CallerFilePath] string file = "",
        [CallerLineNumber] int line = 0)
    {
        WriteEntry("INFO ", message, member, file, line);
    }

    public static void Warn(
        string message,
        [CallerMemberName] string member = "",
        [CallerFilePath] string file = "",
        [CallerLineNumber] int line = 0)
    {
        WriteEntry("WARN ", message, member, file, line);
    }

    public static void Error(
        string message,
        Exception? ex = null,
        [CallerMemberName] string member = "",
        [CallerFilePath] string file = "",
        [CallerLineNumber] int line = 0)
    {
        string full = ex != null
            ? $"{message} | EX: {ex.GetType().Name}: {ex.Message}\n{ex.StackTrace}"
            : message;
        WriteEntry("ERROR", full, member, file, line);
    }

    public static IDisposable Scope(
        string scopeName,
        [CallerMemberName] string member = "",
        [CallerFilePath] string file = "",
        [CallerLineNumber] int line = 0)
    {
        return new LogScope(scopeName, member, file, line);
    }

    private static void WriteEntry(string level, string message, string member, string file, int line)
    {
        lock (_lock)
        {
            try
            {
                string shortFile = Path.GetFileName(file);
                string indent = new string(' ', Math.Max(0, _indentLevel * 2));
                int tid = Thread.CurrentThread.ManagedThreadId;
                string lineStr = $"[{DateTime.Now:HH:mm:ss.fff}] [T{tid:D2}] [{level}] [{shortFile}:{line} {member}] {indent}{message}{Environment.NewLine}";

                File.AppendAllText(LogFilePath, lineStr, Encoding.UTF8);
                Debug.Write(lineStr);
            }
            catch { }
        }
    }

    private sealed class LogScope : IDisposable
    {
        private readonly string _scopeName;
        private readonly string _member;
        private readonly string _file;
        private readonly int _line;
        private readonly Stopwatch _sw;

        public LogScope(string scopeName, string member, string file, int line)
        {
            _scopeName = scopeName;
            _member = member;
            _file = file;
            _line = line;
            _sw = Stopwatch.StartNew();

            WriteEntry("START", $"--> {_scopeName}", _member, _file, _line);
            Interlocked.Increment(ref _indentLevel);
        }

        public void Dispose()
        {
            Interlocked.Decrement(ref _indentLevel);
            _sw.Stop();
            WriteEntry("END  ", $"<-- {_scopeName} ({_sw.ElapsedMilliseconds} ms)", _member, _file, _line);
        }
    }
}