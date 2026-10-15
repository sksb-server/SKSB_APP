namespace SKSB_App.Core.DialogBoxes;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

public enum ValidationDialogAction
{
    Cancel,
    ProceedAnyway
}

public partial class ValidationSummaryDialog : Form
{
    public ValidationDialogAction Action { get; private set; } = ValidationDialogAction.Cancel;

    public ValidationSummaryDialog(string moduleName, List<string> errors, List<string> warnings)
    {
        Text = $"{moduleName} - Data Integrity Check";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24, 20, 24, 20);
        BackColor = Color.White;

        bool hasErrors = errors.Count > 0;
        bool hasWarnings = warnings.Count > 0;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // Status Accent Circle (Red if errors, Amber if only warnings)
        var pnlIndicator = new Panel
        {
            Size = new Size(38, 38),
            Margin = new Padding(0, 2, 16, 0)
        };
        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            Color iconColor = hasErrors ? Color.FromArgb(220, 38, 38) : Color.FromArgb(217, 119, 6);
            using var brush = new SolidBrush(iconColor);
            e.Graphics.FillEllipse(brush, 2, 2, 34, 34);

            using var pen = new Pen(Color.White, 2.5f);
            e.Graphics.DrawLine(pen, 19, 10, 19, 21);
            e.Graphics.FillEllipse(Brushes.White, 17.5f, 25f, 3, 3);
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
            Text = hasErrors ? "Validation Issues Found" : "Validation Advisories",
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6)
        };

        var sbMessage = new StringBuilder();
        if (hasErrors)
        {
            sbMessage.AppendLine("Cannot save records to the system until critical data conflicts are resolved.");
            sbMessage.AppendLine();
            sbMessage.AppendLine($"• Critical Errors: {errors.Count} issue(s) detected.");
        }
        else
        {
            sbMessage.AppendLine("All critical checks passed, but some recommended fields are missing.");
            sbMessage.AppendLine();
        }

        if (hasWarnings)
        {
            sbMessage.AppendLine($"• Advisories / Warnings: {warnings.Count} item(s) to review.");
        }

        var lblMessage = new Label
        {
            Text = sbMessage.ToString(),
            Font = new Font("Segoe UI", 9.5f),
            ForeColor = Color.FromArgb(71, 85, 105),
            AutoSize = true,
            MaximumSize = new Size(440, 0),
            Margin = new Padding(0, 0, 0, 16)
        };

        // Drill-down button bar for details
        var flowDrillDown = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 18)
        };

        if (hasErrors)
        {
            var btnViewErrors = new Button
            {
                Text = $"View Errors ({errors.Count})...",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                ForeColor = Color.FromArgb(185, 28, 28),
                BackColor = Color.FromArgb(254, 242, 242),
                FlatStyle = FlatStyle.Flat,
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
                Margin = new Padding(0, 0, 8, 0)
            };
            btnViewErrors.FlatAppearance.BorderColor = Color.FromArgb(252, 165, 165);
            btnViewErrors.Click += (_, _) => ShowDetailsModal("Critical Validation Errors", errors, isError: true);
            flowDrillDown.Controls.Add(btnViewErrors);
        }

        if (hasWarnings)
        {
            var btnViewWarnings = new Button
            {
                Text = $"View Warnings ({warnings.Count})...",
                Font = new Font("Segoe UI", 9f, FontStyle.Regular),
                ForeColor = Color.FromArgb(180, 83, 9),
                BackColor = Color.FromArgb(254, 252, 232),
                FlatStyle = FlatStyle.Flat,
                AutoSize = true,
                Padding = new Padding(10, 5, 10, 5),
                Margin = new Padding(0, 0, 8, 0)
            };
            btnViewWarnings.FlatAppearance.BorderColor = Color.FromArgb(253, 224, 71);
            btnViewWarnings.Click += (_, _) => ShowDetailsModal("Validation Advisories & Warnings", warnings, isError: false);
            flowDrillDown.Controls.Add(btnViewWarnings);
        }

        // Action Buttons (Right to Left)
        var flowActions = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = Padding.Empty
        };

        var btnCancel = new Button
        {
            Text = hasErrors ? "Return to Sheet (Fix Errors)" : "Cancel",
            Font = new Font("Segoe UI", 9.5f, hasErrors ? FontStyle.Bold : FontStyle.Regular),
            AutoSize = true,
            Padding = new Padding(14, 6, 14, 6),
            DialogResult = DialogResult.Cancel
        };
        btnCancel.Click += (_, _) => { Action = ValidationDialogAction.Cancel; Close(); };
        flowActions.Controls.Add(btnCancel);

        // If only warnings exist, allow proceeding with save anyway
        if (!hasErrors && hasWarnings)
        {
            var btnProceed = new Button
            {
                Text = "Proceed with Save",
                Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(14, 6, 14, 6),
                BackColor = Color.FromArgb(241, 245, 249)
            };
            btnProceed.Click += (_, _) => { Action = ValidationDialogAction.ProceedAnyway; Close(); };
            flowActions.Controls.Add(btnProceed);
        }

        pnlContent.Controls.Add(lblTitle);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(flowDrillDown);
        pnlContent.Controls.Add(flowActions);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);
        Controls.Add(layout);
    }

    private void ShowDetailsModal(string title, List<string> items, bool isError)
    {
        using var detailsForm = new Form
        {
            Text = title,
            Size = new Size(650, 420),
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
        sb.AppendLine($"==========================================================================");
        sb.AppendLine($" {title.ToUpperInvariant()} ({items.Count} items total)");
        sb.AppendLine($"==========================================================================");
        sb.AppendLine();

        foreach (var item in items)
        {
            sb.AppendLine(item);
        }

        txtLog.Text = sb.ToString();
        detailsForm.Controls.Add(txtLog);
        detailsForm.ShowDialog(this);
    }
}