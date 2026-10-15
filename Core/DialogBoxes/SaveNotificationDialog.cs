namespace SKSB_App.Core.DialogBoxes;

using System;
using System.Drawing;
using System.Windows.Forms;

public enum SaveNotificationType
{
    Success,
    Warning,
    Error,
    Offline
}

public partial class SaveNotificationDialog : Form
{
    public SaveNotificationDialog(string moduleName, string message, SaveNotificationType type = SaveNotificationType.Success)
    {
        Text = $"{moduleName} - Save Status";
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

        // Status Accent Icon / Indicator
        var pnlIndicator = new Panel
        {
            Size = new Size(36, 36),
            Margin = new Padding(0, 2, 16, 0)
        };
        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;
            Color iconColor = type switch
            {
                SaveNotificationType.Success => Color.FromArgb(16, 149, 193),   // Brand Teal
                SaveNotificationType.Offline => Color.FromArgb(79, 70, 229),    // Indigo
                SaveNotificationType.Warning => Color.FromArgb(217, 119, 6),    // Amber
                SaveNotificationType.Error => Color.FromArgb(220, 38, 38),      // Red
                _ => Color.Gray
            };

            using var brush = new SolidBrush(iconColor);
            e.Graphics.FillEllipse(brush, 2, 2, 32, 32);

            // Draw clean white checkmark / exclamation
            using var pen = new Pen(Color.White, 2.5f);
            if (type == SaveNotificationType.Success || type == SaveNotificationType.Offline)
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
            Text = type switch
            {
                SaveNotificationType.Success => "Saved Successfully",
                SaveNotificationType.Offline => "Saved Offline",
                SaveNotificationType.Warning => "Validation Issue",
                SaveNotificationType.Error => "Save Failed",
                _ => "Save Status"
            },
            Font = new Font("Segoe UI", 11.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6)
        };

        var lblMessage = new Label
        {
            Text = message,
            Font = new Font("Segoe UI", 9.5f),
            ForeColor = Color.FromArgb(71, 85, 105),
            AutoSize = true,
            MaximumSize = new Size(420, 0),
            Margin = new Padding(0, 0, 0, 18)
        };

        var flowButton = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = Padding.Empty
        };

        var btnOk = new Button
        {
            Text = "OK",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            AutoSize = true,
            Padding = new Padding(18, 6, 18, 6),
            BackColor = Color.FromArgb(241, 245, 249),
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.OK
        };
        btnOk.FlatAppearance.BorderSize = 1;
        btnOk.FlatAppearance.BorderColor = Color.FromArgb(203, 213, 225);
        btnOk.Click += (_, _) => Close();
        flowButton.Controls.Add(btnOk);

        pnlContent.Controls.Add(lblTitle);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(flowButton);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);

        Controls.Add(layout);
    }

    /// <summary>
    /// Helper helper method to show the unified save confirmation dialog.
    /// </summary>
    public static void ShowSuccess(string moduleName, string message = "All records have been safely synchronized to the system.")
    {
        using var dlg = new SaveNotificationDialog(moduleName, message, SaveNotificationType.Success);
        dlg.ShowDialog();
    }

    public static void ShowOffline(string moduleName, string message)
    {
        using var dlg = new SaveNotificationDialog(moduleName, message, SaveNotificationType.Offline);
        dlg.ShowDialog();
    }

    public static void ShowError(string moduleName, string message)
    {
        using var dlg = new SaveNotificationDialog(moduleName, message, SaveNotificationType.Error);
        dlg.ShowDialog();
    }
}