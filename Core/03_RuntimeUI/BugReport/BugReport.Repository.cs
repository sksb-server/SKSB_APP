namespace SKSB_App.Core.RuntimeUI.BugReport;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.Json;
using System.Windows.Forms;
using Dapper;
using Npgsql;
using SKSB_App.Core.Bootstrap.Storage;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Integration.Database;

/// <summary>
/// Persistence for bug reports. Always targets the <c>public</c> schema through
/// <see cref="PostgreSqlService.ReleaseConnectionString"/> so Release submissions and the
/// Debug-only Inbox viewer share the same <c>public.bug_reports</c> table.
/// </summary>
public static class BugReportRepository
{
    public const long MaxFileBytes = 25L * 1024 * 1024;
    public const long MaxTotalBytes = 50L * 1024 * 1024;

    private static readonly object _schemaLock = new();
    private static bool _schemaEnsured;

    static BugReportRepository()
    {
        DefaultTypeMap.MatchNamesWithUnderscores = true;
    }

    // =========================================================================
    // 1. CONNECTION & SCHEMA
    // =========================================================================
    private static NpgsqlConnection OpenConnection()
    {
#if DEBUG
        if (ServerEnvironment.SimulateOffline)
            throw new TimeoutException("[Simulated Offline Mode Active] Connection to PostgreSQL server blocked by user switch.");
#endif
        var csb = new NpgsqlConnectionStringBuilder(PostgreSqlService.ReleaseConnectionString)
        {
            Timeout = 5,
            CommandTimeout = 30,
            SearchPath = "public"
        };
        var conn = new NpgsqlConnection(csb.ConnectionString);
        try
        {
            conn.Open();
        }
        catch
        {
            conn.Dispose();
            throw;
        }
        return conn;
    }

    /// <summary>Runs the embedded <c>bug_report_schema.sql</c> once per process (idempotent DDL).</summary>
    public static void EnsureSchema()
    {
        if (_schemaEnsured) return;
        lock (_schemaLock)
        {
            if (_schemaEnsured) return;
            using Stream? stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("bug_report_schema.sql");
            if (stream == null) throw new InvalidOperationException("Embedded resource 'bug_report_schema.sql' not found.");
            using var reader = new StreamReader(stream, Encoding.UTF8);
            string sql = reader.ReadToEnd();

            using var conn = OpenConnection();
            conn.Execute(sql);
            _schemaEnsured = true;
        }
    }

    // =========================================================================
    // 2. SUBMIT
    // =========================================================================
    public static long SubmitBugReport(
        string title,
        string description,
        BugReportContext ctx,
        IReadOnlyList<BugReportAttachmentInput> attachments)
    {
        if (string.IsNullOrWhiteSpace(title)) throw new ArgumentException("A bug title is required.", nameof(title));
        EnsureSchema();

        using var conn = OpenConnection();
        using var trans = conn.BeginTransaction();

        const string insertBug = @"
INSERT INTO public.bug_reports
    (title, description, source_module, active_sheet, app_version, build_mode, is_offline,
     reported_by, machine_name, status, runtime_log_tail)
VALUES
    (@Title, @Description, @SourceModule, @ActiveSheet, @AppVersion, @BuildMode, @IsOffline,
     @ReportedBy, @MachineName, @Status, @RuntimeLogTail)
RETURNING bug_id;";

        long bugId = conn.ExecuteScalar<long>(insertBug, new
        {
            Title = Truncate(title.Trim(), 200),
            Description = description ?? string.Empty,
            SourceModule = Truncate(ctx.SourceModule, 60),
            ActiveSheet = Truncate(ctx.ActiveSheet, 100),
            AppVersion = Truncate(ctx.AppVersion, 40),
            BuildMode = Truncate(ctx.BuildMode, 20),
            ctx.IsOffline,
            ReportedBy = Truncate(ctx.ReportedBy, 120),
            MachineName = Truncate(ctx.MachineName, 120),
            Status = BugReportStatus.Open,
            RuntimeLogTail = ctx.RuntimeLogTail ?? string.Empty
        }, trans);

        const string insertAttachment = @"
INSERT INTO public.bug_report_attachments (bug_id, file_name, content_type, file_size, file_data)
VALUES (@BugId, @FileName, @ContentType, @FileSize, @FileData);";

        foreach (var att in attachments)
        {
            conn.Execute(insertAttachment, new
            {
                BugId = bugId,
                FileName = Truncate(att.FileName, 255),
                ContentType = Truncate(att.ContentType, 100),
                FileSize = att.FileSize,
                FileData = att.Data
            }, trans);
        }

        trans.Commit();
        AppLogger.Info($"[BugReport] Submitted bug #{bugId} '{title}' with {attachments.Count} attachment(s).");
        return bugId;
    }

    // =========================================================================
    // 3. INBOX QUERIES & TRIAGE
    // =========================================================================
    public static int GetOpenBugCount()
    {
        try
        {
            EnsureSchema();
            using var conn = OpenConnection();
            return conn.ExecuteScalar<int>(
                "SELECT COUNT(*) FROM public.bug_reports WHERE status IN (@Open, @InProgress);",
                new { Open = BugReportStatus.Open, InProgress = BugReportStatus.InProgress });
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[BugReport] GetOpenBugCount failed: {ex.Message}");
            return 0;
        }
    }

    /// <param name="statusFilter"><c>"Active"</c> (Open + In Progress), <c>"All"</c>/null, or an exact status.</param>
    public static List<BugReportSummary> ListBugReports(string? statusFilter)
    {
        EnsureSchema();
        string where;
        object param;
        if (string.IsNullOrWhiteSpace(statusFilter) || statusFilter.Equals("All", StringComparison.OrdinalIgnoreCase))
        {
            where = string.Empty;
            param = new { };
        }
        else if (statusFilter.Equals("Active", StringComparison.OrdinalIgnoreCase))
        {
            where = "WHERE b.status IN (@Open, @InProgress)";
            param = new { Open = BugReportStatus.Open, InProgress = BugReportStatus.InProgress };
        }
        else
        {
            where = "WHERE b.status = @Status";
            param = new { Status = statusFilter };
        }

        string sql = $@"
SELECT b.bug_id, b.title, b.description, b.source_module, b.active_sheet, b.app_version, b.build_mode,
       b.is_offline, b.reported_by, b.machine_name, b.status, b.dev_notes, b.runtime_log_tail,
       b.created_at, b.updated_at,
       (SELECT COUNT(*)::int FROM public.bug_report_attachments a WHERE a.bug_id = b.bug_id) AS attachment_count
FROM public.bug_reports b
{where}
ORDER BY b.created_at DESC;";

        using var conn = OpenConnection();
        return conn.Query<BugReportSummary>(sql, param).ToList();
    }

    public static List<BugReportAttachmentMeta> GetAttachmentsMeta(long bugId)
    {
        using var conn = OpenConnection();
        return conn.Query<BugReportAttachmentMeta>(@"
SELECT attachment_id, bug_id, file_name, content_type, file_size, created_at
FROM public.bug_report_attachments
WHERE bug_id = @BugId
ORDER BY attachment_id;", new { BugId = bugId }).ToList();
    }

    public static byte[]? GetAttachmentData(long attachmentId)
    {
        using var conn = OpenConnection();
        return conn.ExecuteScalar<byte[]?>(
            "SELECT file_data FROM public.bug_report_attachments WHERE attachment_id = @Id;",
            new { Id = attachmentId });
    }

    public static void UpdateBugTriage(long bugId, string status, string devNotes)
    {
        using var conn = OpenConnection();
        conn.Execute(@"
UPDATE public.bug_reports
SET status = @Status, dev_notes = @DevNotes, updated_at = CURRENT_TIMESTAMP
WHERE bug_id = @BugId;", new { BugId = bugId, Status = Truncate(status, 30), DevNotes = devNotes ?? string.Empty });
    }

    public static void DeleteBugReport(long bugId)
    {
        using var conn = OpenConnection();
        conn.Execute("DELETE FROM public.bug_reports WHERE bug_id = @BugId;", new { BugId = bugId });
    }

    // =========================================================================
    // 4. CAPTURE & ENVIRONMENT HELPERS
    // =========================================================================
    /// <summary>Captures the primary screen as PNG bytes; returns null on failure.</summary>
    public static byte[]? CaptureScreenPng()
    {
        try
        {
            Rectangle bounds = Screen.PrimaryScreen?.Bounds ?? Rectangle.Empty;
            if (bounds.Width <= 0 || bounds.Height <= 0) return null;

            using var bmp = new Bitmap(bounds.Width, bounds.Height, PixelFormat.Format32bppArgb);
            using (var g = Graphics.FromImage(bmp))
            {
                g.CopyFromScreen(bounds.Left, bounds.Top, 0, 0, bounds.Size, CopyPixelOperation.SourceCopy);
            }
            using var ms = new MemoryStream();
            bmp.Save(ms, ImageFormat.Png);
            return ms.ToArray();
        }
        catch (Exception ex)
        {
            AppLogger.Warn($"[BugReport] Screen capture failed: {ex.Message}");
            return null;
        }
    }

    /// <summary>Encodes any GDI+ image as PNG bytes (caller keeps ownership of <paramref name="image"/>).</summary>
    public static byte[] EncodePng(Image image)
    {
        using var ms = new MemoryStream();
        image.Save(ms, ImageFormat.Png);
        return ms.ToArray();
    }

    /// <summary>Reads the last <paramref name="maxLines"/> lines of the runtime trace log without locking it.</summary>
    public static string CaptureRuntimeLogTail(int maxLines = 150)
    {
        try
        {
            string path = AppLogger.LogFilePath;
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return string.Empty;

            const long maxTailBytes = 512 * 1024;
            using var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            long start = Math.Max(0, fs.Length - maxTailBytes);
            fs.Seek(start, SeekOrigin.Begin);
            using var reader = new StreamReader(fs, Encoding.UTF8);
            string text = reader.ReadToEnd();

            var lines = text.Replace("\r\n", "\n").Split('\n');
            int skipFirst = start > 0 ? 1 : 0; // first line may be partial when seeking mid-file
            var tail = lines.Skip(skipFirst).TakeLast(Math.Max(1, maxLines));
            return string.Join(Environment.NewLine, tail).TrimEnd();
        }
        catch (Exception ex)
        {
            return $"(Failed to read runtime log: {ex.Message})";
        }
    }

    public static string ResolveAppVersion()
    {
        foreach (string dir in new[] { SafeAppRoot(), AppContext.BaseDirectory })
        {
            try
            {
                if (string.IsNullOrWhiteSpace(dir)) continue;
                string file = Path.Combine(dir, "version.json");
                if (!File.Exists(file)) continue;
                using var doc = JsonDocument.Parse(File.ReadAllText(file));
                if (doc.RootElement.TryGetProperty("version", out var prop))
                {
                    string? ver = prop.GetString();
                    if (!string.IsNullOrWhiteSpace(ver)) return ver.Trim().TrimStart('v');
                }
            }
            catch { }
        }
        return "unknown";
    }

    public static string GetContentType(string fileName)
    {
        return Path.GetExtension(fileName ?? string.Empty).ToLowerInvariant() switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".webp" => "image/webp",
            ".mp4" => "video/mp4",
            ".mov" => "video/quicktime",
            ".webm" => "video/webm",
            ".mkv" => "video/x-matroska",
            ".avi" => "video/x-msvideo",
            ".txt" or ".log" => "text/plain",
            _ => "application/octet-stream"
        };
    }

    private static string SafeAppRoot()
    {
        try { return FolderRegistry.SKSB_App.GetPath(); }
        catch { return string.Empty; }
    }

    private static string Truncate(string? value, int max)
    {
        value ??= string.Empty;
        return value.Length <= max ? value : value[..max];
    }
}
