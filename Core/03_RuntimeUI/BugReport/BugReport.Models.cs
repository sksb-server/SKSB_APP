namespace SKSB_App.Core.RuntimeUI.BugReport;

using System;

/// <summary>Canonical status values stored in <c>public.bug_reports.status</c>.</summary>
public static class BugReportStatus
{
    public const string Open = "Open";
    public const string InProgress = "In Progress";
    public const string Resolved = "Resolved";
    public const string Closed = "Closed";

    public static readonly string[] AllStatuses = { Open, InProgress, Resolved, Closed };
}

/// <summary>Environment captured at the moment the user opened the Report Bug dialog.</summary>
public sealed class BugReportContext
{
    public string SourceModule { get; set; } = string.Empty;
    public string ActiveSheet { get; set; } = string.Empty;
    public string AppVersion { get; set; } = string.Empty;
#if DEBUG
    public string BuildMode { get; set; } = "Debug";
#else
    public string BuildMode { get; set; } = "Release";
#endif
    public bool IsOffline { get; set; }
    public string ReportedBy { get; set; } = Environment.UserName;
    public string MachineName { get; set; } = Environment.MachineName;

    /// <summary>Screenshot captured before the dialog was shown (pre-attached, removable by the user).</summary>
    public byte[]? InitialScreenshotPng { get; set; }

    /// <summary>Runtime trace log tail; filled by the submit dialog when the user opts in.</summary>
    public string RuntimeLogTail { get; set; } = string.Empty;
}

/// <summary>A file the user attached in the submit dialog (held in memory until submit).</summary>
public sealed class BugReportAttachmentInput
{
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = "application/octet-stream";
    public byte[] Data { get; set; } = Array.Empty<byte>();

    public long FileSize => Data.LongLength;
    public bool IsImage => ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);
}

/// <summary>Row projection of <c>public.bug_reports</c> used by the Inbox grid.</summary>
public sealed class BugReportSummary
{
    public long BugId { get; set; }
    public string Title { get; set; } = string.Empty;
    public string Description { get; set; } = string.Empty;
    public string SourceModule { get; set; } = string.Empty;
    public string ActiveSheet { get; set; } = string.Empty;
    public string AppVersion { get; set; } = string.Empty;
    public string BuildMode { get; set; } = string.Empty;
    public bool IsOffline { get; set; }
    public string ReportedBy { get; set; } = string.Empty;
    public string MachineName { get; set; } = string.Empty;
    public string Status { get; set; } = BugReportStatus.Open;
    public string DevNotes { get; set; } = string.Empty;
    public string RuntimeLogTail { get; set; } = string.Empty;
    public DateTime CreatedAt { get; set; }
    public DateTime UpdatedAt { get; set; }
    public int AttachmentCount { get; set; }
}

/// <summary>Attachment metadata without the BYTEA payload.</summary>
public sealed class BugReportAttachmentMeta
{
    public long AttachmentId { get; set; }
    public long BugId { get; set; }
    public string FileName { get; set; } = string.Empty;
    public string ContentType { get; set; } = "application/octet-stream";
    public long FileSize { get; set; }
    public DateTime CreatedAt { get; set; }

    public bool IsImage => ContentType.StartsWith("image/", StringComparison.OrdinalIgnoreCase);

    public override string ToString() => $"{FileName}  ({BugReportFormat.FormatSize(FileSize)})";
}

/// <summary>Small shared formatting helpers for the bug report UI.</summary>
public static class BugReportFormat
{
    public static string FormatSize(long bytes)
    {
        if (bytes < 1024) return $"{bytes} B";
        if (bytes < 1024 * 1024) return $"{bytes / 1024.0:0.#} KB";
        return $"{bytes / (1024.0 * 1024.0):0.##} MB";
    }
}
