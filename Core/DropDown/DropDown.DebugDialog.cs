namespace SKSB_App.Core.DropDown;

using System;
using System.Drawing;
using System.Windows.Forms;

public partial class DebugDialog : Form
{
    public static bool SuppressForSession { get; set; } = false;

    public DebugDialog(string title, string diagnosticLog)
    {
        Text = title;
        Size = new Size(680, 500);
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.Sizable;
        TopMost = true;

        var txtLog = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Both,
            Dock = DockStyle.Fill,
            Font = new Font("Consolas", 9.5f),
            BackColor = Color.White,
            Text = diagnosticLog
        };

        var pnlBottom = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            Height = 44,
            FlowDirection = FlowDirection.LeftToRight,
            Padding = new Padding(10, 8, 10, 8)
        };

        var chkSuppress = new CheckBox
        {
            Text = "Don't show debugging for the rest of this session",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f),
            Margin = new Padding(0, 4, 20, 0)
        };
        chkSuppress.CheckedChanged += (s, e) => SuppressForSession = chkSuppress.Checked;

        var btnClose = new Button
        {
            Text = "Continue",
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Width = 90,
            Height = 28,
            DialogResult = DialogResult.OK
        };
        btnClose.Click += (s, e) => Close();

        pnlBottom.Controls.Add(chkSuppress);
        pnlBottom.Controls.Add(btnClose);

        Controls.Add(txtLog);
        Controls.Add(pnlBottom);
        AcceptButton = btnClose;
    }

    public static void ShowLog(string title, string message)
    {
        if (SuppressForSession) return;
        using var dlg = new DebugDialog(title, message);
        dlg.ShowDialog();
    }
}