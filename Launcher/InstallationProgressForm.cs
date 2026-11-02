namespace SKSB_App.Launcher;

using System;
using System.Drawing;
using System.Windows.Forms;

public partial class InstallationProgressForm : Form
{
    private readonly Label _lblTitle;
    private readonly Label _lblStatus;
    private readonly ProgressBar _progressBar;

    public InstallationProgressForm()
    {
        Text = "SKSB Operations Hub Setup";
        Size = new Size(440, 150);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        ControlBox = false;
        TopMost = true;
        BackColor = Color.White;

        _lblTitle = new Label
        {
            Text = "Setting up SKSB App...",
            Font = new Font("Segoe UI", 10.5f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            Location = new Point(24, 16),
            AutoSize = true
        };

        _lblStatus = new Label
        {
            Text = "Connecting to server...",
            Font = new Font("Segoe UI", 9f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            Location = new Point(25, 44),
            AutoSize = true
        };

        _progressBar = new ProgressBar
        {
            Location = new Point(26, 72),
            Size = new Size(372, 18),
            Style = ProgressBarStyle.Marquee,
            MarqueeAnimationSpeed = 30
        };

        Controls.Add(_lblTitle);
        Controls.Add(_lblStatus);
        Controls.Add(_progressBar);
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
    }
}