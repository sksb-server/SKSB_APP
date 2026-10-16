namespace SKSB_App.Core.DialogBoxes;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;

/// <summary>
/// Represents a row skipped during import due to validation or formatting errors.
/// </summary>
public record ImportSkippedRecord
{
    public int ExcelRowNumber { get; init; }
    public string Identifier { get; init; } = string.Empty; // e.g. Worker Code, Project Code, Plate
    public string NameOrTitle { get; init; } = string.Empty; // e.g. Name, Description
    public string Reason { get; init; } = string.Empty;
}

/// <summary>
/// Domain-agnostic model containing the results and metrics of an ERP or file import operation.
/// </summary>
public class ImportSummaryReport
{
    public string SourceSystemName { get; init; } = "ERP"; // e.g. "SQL Payroll", "SQL Accounting"
    public string EntityLabel { get; init; } = "Record";   // e.g. "worker(s)", "project(s)"
    public int SuccessfullyImportedCount { get; init; } = 0;
    public int FilteredExclusionCount { get; init; } = 0;
    public string FilterRuleDescription { get; init; } = string.Empty;
    public List<ImportSkippedRecord> SkippedErrors { get; init; } = new();
}

public class GenericImportResult : Form
{
    private readonly ImportSummaryReport _report;

    public GenericImportResult(ImportSummaryReport report, string fileName)
    {
        _report = report;

        Text = $"{report.SourceSystemName} Import Summary";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24, 20, 24, 20);
        BackColor = Color.White;

        bool hasErrors = _report.SkippedErrors.Count > 0;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // Status Circle Indicator (Teal on clean success, Amber if warnings/skipped items)
        var pnlIndicator = new Panel
        {
            Size = new Size(36, 36),
            Margin = new Padding(0, 2, 16, 0)
        };
        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            Color iconColor = hasErrors ? Color.FromArgb(217, 119, 6) : Color.FromArgb(16, 149, 193);
            using var brush = new SolidBrush(iconColor);
            e.Graphics.FillEllipse(brush, 2, 2, 32, 32);

            using var pen = new Pen(Color.White, 2.5f);
            if (!hasErrors)
            {
                e.Graphics.DrawLines(pen, new Point[] { new(10, 18), new(16, 24), new(26, 12) });
            }
            else
            {
                e.Graphics.DrawLine(pen, 18, 10, 18, 20);
                e.Graphics.FillEllipse(Brushes.White, 16.5f, 23.5f, 3, 3);
            }
        };

        var pnlContent = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        var lblTitle = new Label
        {
            Text = hasErrors ? "Import Completed with Advisories" : "Import Completed Successfully",
            Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 8)
        };

        var sbMessage = new StringBuilder();
        if (!string.IsNullOrWhiteSpace(fileName))
        {
            sbMessage.AppendLine($"Source: {fileName}");
            sbMessage.AppendLine();
        }

        if (!string.IsNullOrWhiteSpace(_report.FilterRuleDescription))
        {
            sbMessage.AppendLine($"• Filter Rule: {_report.FilterRuleDescription}");
        }

        if (_report.FilteredExclusionCount > 0)
        {
            sbMessage.AppendLine($"• Excluded by Filter: {_report.FilteredExclusionCount} record(s).");
        }

        sbMessage.AppendLine($"• Successfully Imported to Table: {_report.SuccessfullyImportedCount} {_report.EntityLabel}.");
        sbMessage.AppendLine($"• Skipped Entries (Errors): {_report.SkippedErrors.Count} record(s).");
        sbMessage.AppendLine();
        sbMessage.AppendLine("Review the entries in Excel and click 'Save to System' when ready.");

        var lblMessage = new Label
        {
            Text = sbMessage.ToString(),
            Font = new Font("Segoe UI", 9.5f),
            ForeColor = Color.FromArgb(71, 85, 105),
            AutoSize = true,
            MaximumSize = new Size(460, 0),
            Margin = new Padding(0, 0, 0, 18)
        };

        var flowButtons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = Padding.Empty
        };

        var btnOk = new Button
        {
            Text = "OK",
            Font = new Font("Segoe UI", 9f),
            AutoSize = true,
            Padding = new Padding(16, 6, 16, 6),
            BackColor = Color.FromArgb(241, 245, 249),
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.OK
        };
        btnOk.FlatAppearance.BorderSize = 1;
        btnOk.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        btnOk.Click += (_, _) => Close();
        flowButtons.Controls.Add(btnOk);

        if (hasErrors)
        {
            var btnViewErrors = new Button
            {
                Text = $"View Skipped Errors ({_report.SkippedErrors.Count})...",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(14, 6, 14, 6),
                ForeColor = Color.FromArgb(185, 28, 28),
                BackColor = Color.FromArgb(254, 242, 242),
                FlatStyle = FlatStyle.Flat
            };
            btnViewErrors.FlatAppearance.BorderColor = Color.FromArgb(252, 165, 165);
            btnViewErrors.Click += (_, _) => ShowDetailedErrors();
            flowButtons.Controls.Add(btnViewErrors);
        }

        pnlContent.Controls.Add(lblTitle);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(flowButtons);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);

        Controls.Add(layout);
    }

    private void ShowDetailedErrors()
    {
        using var errForm = new Form
        {
            Text = $"{_report.SourceSystemName} - Skipped Entries Report",
            Size = new Size(680, 440),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.Sizable,
            MinimizeBox = false,
            MaximizeBox = true
        };

        var txtLog = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 9.5f),
            BackColor = Color.White
        };

        var sb = new StringBuilder();
        sb.AppendLine("==========================================================================");
        sb.AppendLine($" SKIPPED ENTRIES REPORT ({_report.SkippedErrors.Count} total)");
        sb.AppendLine("==========================================================================");
        sb.AppendLine();

        foreach (var err in _report.SkippedErrors)
        {
            sb.AppendLine($"[Excel Row {err.ExcelRowNumber}]");
            sb.AppendLine($"  • Identifier : {err.Identifier}");
            if (!string.IsNullOrWhiteSpace(err.NameOrTitle))
                sb.AppendLine($"  • Name/Title : {err.NameOrTitle}");
            sb.AppendLine($"  • Reason     : {err.Reason}");
            sb.AppendLine("--------------------------------------------------------------------------");
        }

        txtLog.Text = sb.ToString();
        errForm.Controls.Add(txtLog);
        errForm.ShowDialog(this);
    }
}