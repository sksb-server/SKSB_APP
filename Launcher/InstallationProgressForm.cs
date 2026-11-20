namespace SKSB_App.Launcher;

using System;
using System.Drawing;
using System.Windows.Forms;

public class InstallationProgressForm : Form
{
    private readonly TableLayoutPanel _layout;
    private readonly Label _lblTitle;
    private readonly Label _lblStatus;
    private readonly ProgressBar _progressBar;

    public InstallationProgressForm()
    {
        SuspendLayout();
        Text = "SKSB Operations Hub Setup";
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(24, 18, 24, 20);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        TopMost = true;
        BackColor = Color.White;

        _layout = new TableLayoutPanel
        {
            ColumnCount = 1,
            Dock = DockStyle.Fill,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = Padding.Empty
        };

        _lblTitle = new Label
        {
            Text = "Setting up SKSB App...",
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6)
        };

        _lblStatus = new Label
        {
            Text = "Connecting to server...",
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            MaximumSize = new Size(392, 0),
            Margin = new Padding(0, 0, 0, 14)
        };

        _progressBar = new ProgressBar
        {
            Size = new Size(392, 18),
            Dock = DockStyle.Top,
            Margin = Padding.Empty,
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30
        };

        _layout.Controls.Add(_lblTitle);
        _layout.Controls.Add(_lblStatus);
        _layout.Controls.Add(_progressBar);
        Controls.Add(_layout);

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        FitClientSizeToLayout();
        CenterToScreen();
    }

    protected override void OnDpiChanged(DpiChangedEventArgs e)
    {
        base.OnDpiChanged(e);
        FitClientSizeToLayout();
    }

    private void FitClientSizeToLayout()
    {
        int contentWidth = LogicalToDeviceUnits(392);
        int barHeight = LogicalToDeviceUnits(18);

        _lblStatus.MaximumSize = new Size(contentWidth, 0);
        _progressBar.MinimumSize = new Size(contentWidth, barHeight);
        _progressBar.Size = new Size(contentWidth, barHeight);
        _layout.PerformLayout();

        Size pref = _layout.GetPreferredSize(new Size(contentWidth, 0));
        ClientSize = new Size(Math.Max(contentWidth, pref.Width) + Padding.Horizontal, pref.Height + Padding.Vertical);
    }

    public void UpdateStatus(string message)
    {
        if (IsDisposed || !IsHandleCreated) return;

        if (InvokeRequired)
        {
            BeginInvoke(new Action(() => UpdateStatus(message)));
            return;
        }

        _lblStatus.Text = message;
        FitClientSizeToLayout();
    }
}