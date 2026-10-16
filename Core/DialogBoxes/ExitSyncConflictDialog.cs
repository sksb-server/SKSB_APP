namespace SKSB_App.Core.DialogBoxes;

using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Windows.Forms;

public enum ExitResolution
{
    Retry,
    SaveLocallyAndQuit,
    Cancel
}

public partial class ExitSyncConflictDialog : Form
{
    public ExitResolution Resolution { get; private set; } = ExitResolution.Cancel;

    public ExitSyncConflictDialog(string moduleName, string errorMessage, bool allowSaveOffline = true, bool isClosing = true)
    {
        Text = $"{moduleName} - Synchronization Alert";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24, 20, 24, 20);
        BackColor = Color.White;

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // Circular Warning Accent Icon
        var pnlIndicator = new Panel
        {
            Size = new Size(36, 36),
            Margin = new Padding(0, 2, 16, 0)
        };
        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new SolidBrush(Color.FromArgb(217, 119, 6)); // Amber
            e.Graphics.FillEllipse(brush, 2, 2, 32, 32);

            using var pen = new Pen(Color.White, 2.5f);
            e.Graphics.DrawLine(pen, 18, 10, 18, 20);
            e.Graphics.FillEllipse(Brushes.White, 16.5f, 23.5f, 3, 3);
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
            Text = "Server Database Offline",
            Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6)
        };

        string actionPrompt = allowSaveOffline
            ? (isClosing
                ? "Would you like to retry the server, save an offline stash and quit, or stay in Excel?"
                : "Would you like to retry connecting, or save an offline stash to upload later?")
            : "Data validation failed. Fix errors before saving. Saving offline is disabled for invalid data.";

        var lblMessage = new Label
        {
            Text = $"Unable to persist records to the server database.\n\n" +
                   $"Details:\n{errorMessage}\n\n" +
                   actionPrompt,
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

        // Cancel / Dismiss Button
        var btnCancel = new Button
        {
            Text = isClosing ? "Cancel (Stay in Excel)" : "Cancel",
            Font = new Font("Segoe UI", 9f),
            AutoSize = true,
            Padding = new Padding(14, 6, 14, 6),
            BackColor = Color.FromArgb(241, 245, 249),
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.Cancel
        };
        btnCancel.FlatAppearance.BorderSize = 1;
        btnCancel.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        btnCancel.Click += (_, _) => { Resolution = ExitResolution.Cancel; Close(); };
        flowButtons.Controls.Add(btnCancel);

        if (allowSaveOffline)
        {
            // Save Offline Button
            var btnSaveLocal = new Button
            {
                Text = isClosing ? "Save Locally & Quit" : "Save Locally",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(14, 6, 14, 6),
                BackColor = Color.FromArgb(79, 70, 229), // Indigo accent
                ForeColor = Color.White,
                FlatStyle = FlatStyle.Flat
            };
            btnSaveLocal.FlatAppearance.BorderSize = 0;
            btnSaveLocal.Click += (_, _) => { Resolution = ExitResolution.SaveLocallyAndQuit; Close(); };
            flowButtons.Controls.Add(btnSaveLocal);

            // Retry Button
            var btnRetry = new Button
            {
                Text = "Retry Server",
                Font = new Font("Segoe UI", 9f),
                AutoSize = true,
                Padding = new Padding(14, 6, 14, 6),
                BackColor = Color.FromArgb(241, 245, 249),
                FlatStyle = FlatStyle.Flat
            };
            btnRetry.FlatAppearance.BorderSize = 1;
            btnRetry.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
            btnRetry.Click += (_, _) => { Resolution = ExitResolution.Retry; Close(); };
            flowButtons.Controls.Add(btnRetry);
        }

        pnlContent.Controls.Add(lblTitle);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(flowButtons);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);

        Controls.Add(layout);
    }
}