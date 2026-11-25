namespace SKSB_App.Core.Integration.Erp;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Text;
using System.Windows.Forms;

public record ImportSkippedRecord
{
    public int ExcelRowNumber { get; init; }
    public string Identifier { get; init; } = string.Empty;
    public string NameOrTitle { get; init; } = string.Empty;
    public string Reason { get; init; } = string.Empty;

    /// <summary>Optional grouping label (e.g. "Department", "Wages", "Thumbprint") shown when there is no Excel row.</summary>
    public string Category { get; init; } = string.Empty;
}

public class ImportSummaryReport
{
    public const string DefaultSkippedSectionTitle = "Skipped Entries (Errors)";

    public string SourceSystemName { get; init; } = "ERP";
    public string EntityLabel { get; init; } = "Record";
    public int SuccessfullyImportedCount { get; init; } = 0;
    public int FilteredExclusionCount { get; init; } = 0;
    public string FilterRuleDescription { get; init; } = string.Empty;
    public List<ImportSkippedRecord> SkippedErrors { get; init; } = new();

    /// <summary>
    /// Optional prominent banner (e.g. "fix these in SQL Payroll, otherwise the data shown is incomplete"). When set,
    /// the dialog switches to advisory wording ("View Advisories...").
    /// </summary>
    public string? AdvisoryBanner { get; init; }

    /// <summary>Label of the issue count line / detail report; defaults to "Skipped Entries (Errors)".</summary>
    public string? SkippedSectionTitle { get; init; }

    /// <summary>Optional dialog title override (defaults to "Import Completed ...").</summary>
    public string? Title { get; init; }

    /// <summary>Optional closing line (defaults to the "Review ... click 'Save to System'" hint).</summary>
    public string? FooterHint { get; init; }

    internal bool IsAdvisoryMode => !string.IsNullOrWhiteSpace(AdvisoryBanner);
    internal string EffectiveSkippedSectionTitle =>
        !string.IsNullOrWhiteSpace(SkippedSectionTitle) ? SkippedSectionTitle! : (IsAdvisoryMode ? "Advisories" : DefaultSkippedSectionTitle);
}

public class ImportResultDialog : Form
{
    private readonly ImportSummaryReport _report;

    public ImportResultDialog(ImportSummaryReport report, string fileName)
    {
        SuspendLayout();
        _report = report;
        Text = $"{report.SourceSystemName} Import Summary";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
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

        var pnlIndicator = new Panel
        {
            Size = new Size(36, 36),
            Margin = new Padding(0, 2, 16, 0)
        };

        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.ScaleTransform(pnlIndicator.ClientSize.Width / 36f, pnlIndicator.ClientSize.Height / 36f);
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
            Text = !string.IsNullOrWhiteSpace(_report.Title)
                ? _report.Title!
                : (hasErrors ? "Import Completed with Advisories" : "Import Completed Successfully"),
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
            sbMessage.AppendLine($"  Filter Rule: {_report.FilterRuleDescription}");
        }

        if (_report.FilteredExclusionCount > 0)
        {
            sbMessage.AppendLine($"  Excluded by Filter: {_report.FilteredExclusionCount} record(s).");
        }

        sbMessage.AppendLine($"  Successfully Imported to Table: {_report.SuccessfullyImportedCount} {_report.EntityLabel}.");
        sbMessage.AppendLine($"  {_report.EffectiveSkippedSectionTitle}: {_report.SkippedErrors.Count} {(_report.IsAdvisoryMode ? "item(s)" : "record(s)")}.");
        sbMessage.AppendLine();
        sbMessage.AppendLine(!string.IsNullOrWhiteSpace(_report.FooterHint)
            ? _report.FooterHint
            : "Review the entries in Excel and click 'Save to System' when ready.");

        Label? lblBanner = null;
        if (_report.IsAdvisoryMode)
        {
            lblBanner = new Label
            {
                Text = _report.AdvisoryBanner,
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                ForeColor = Color.FromArgb(146, 64, 14),
                BackColor = Color.FromArgb(254, 243, 199),
                BorderStyle = BorderStyle.FixedSingle,
                AutoSize = true,
                MaximumSize = new Size(460, 0),
                Padding = new Padding(10, 8, 10, 8),
                Margin = new Padding(0, 0, 0, 12)
            };
        }

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
                Text = _report.IsAdvisoryMode
                    ? $"View Advisories ({_report.SkippedErrors.Count})..."
                    : $"View Skipped Errors ({_report.SkippedErrors.Count})...",
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
        if (lblBanner != null) pnlContent.Controls.Add(lblBanner);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(flowButtons);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);
        Controls.Add(layout);

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    private void ShowDetailedErrors()
    {
        using var errForm = new Form
        {
            Text = $"{_report.SourceSystemName} - {(_report.IsAdvisoryMode ? "Advisories Report" : "Skipped Entries Report")}",
            Size = new Size(680, 440),
            StartPosition = FormStartPosition.CenterParent,
            FormBorderStyle = FormBorderStyle.Sizable,
            MinimizeBox = false,
            MaximizeBox = true
        };
        errForm.SuspendLayout();

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
        sb.AppendLine($" {_report.EffectiveSkippedSectionTitle.ToUpperInvariant()} REPORT ({_report.SkippedErrors.Count} total)");
        sb.AppendLine("==========================================================================");
        sb.AppendLine();

        if (_report.IsAdvisoryMode)
        {
            sb.AppendLine(_report.AdvisoryBanner);
            sb.AppendLine();
            sb.AppendLine("--------------------------------------------------------------------------");
        }

        foreach (var err in _report.SkippedErrors)
        {
            if (err.ExcelRowNumber > 0)
                sb.AppendLine($"[Excel Row {err.ExcelRowNumber}]");
            else if (!string.IsNullOrWhiteSpace(err.Category))
                sb.AppendLine($"[{err.Category}]");
            else if (!string.IsNullOrWhiteSpace(err.Identifier))
                sb.AppendLine($"[{err.Identifier}]");
            sb.AppendLine($"    Identifier : {err.Identifier}");
            if (!string.IsNullOrWhiteSpace(err.NameOrTitle))
                sb.AppendLine($"    Name/Title : {err.NameOrTitle}");
            sb.AppendLine($"    Reason     : {err.Reason}");
            sb.AppendLine("--------------------------------------------------------------------------");
        }

        txtLog.Text = sb.ToString();
        errForm.Controls.Add(txtLog);
        errForm.AutoScaleDimensions = new SizeF(96f, 96f);
        errForm.AutoScaleMode = AutoScaleMode.Dpi;
        errForm.ResumeLayout(false);
        errForm.PerformLayout();
        errForm.ShowDialog(this);
    }
}