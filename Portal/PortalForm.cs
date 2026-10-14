namespace SKSB_App.Portal;

using SKSB_App.Core.FileManager.Constants;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

public partial class PortalForm : Form
{
    public PortalForm()
    {
        Text = "SKSB Operations Portal";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        BackColor = Color.FromArgb(248, 250, 252);

        // 1. Enable automated DPI scaling on the form
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);

        // Let the form expand proportionally to its contents
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24, 20, 24, 24);

        InitializeLayout();
    }

    private void InitializeLayout()
    {
        // Use a vertical TableLayoutPanel so rows stack cleanly without overlapping
        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Margin = new Padding(0),
            Padding = new Padding(0)
        };
        layout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));

        // 1. Header Label
        var lblHeader = new Label
        {
            Text = "Operations Hub",
            Font = new Font("Segoe UI", 16f, FontStyle.Bold),
            ForeColor = Color.FromArgb(15, 23, 42),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 4)
        };

        // 2. Subheader Label
        var lblSubheader = new Label
        {
            Text = "Select an operational module to open in Excel:",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        };

        // 3. Module Buttons (Stacked neatly with 10px spacing)
        var btnScheduler = CreateModuleButton(
            "Daily Work Allocation",
            "Biometric logs & shift scheduling",
            () => PortalRouter.Launch(Window.DailyScheduler));

        var btnProject = CreateModuleButton(
            "Project Manager",
            "Job sites, scope, priorities & timelines",
            () => PortalRouter.Launch(Window.ProjectManager));

        var btnWorker = CreateModuleButton(
            "Worker Manager",
            "Employee profiles, badges & privileges",
            () => PortalRouter.Launch(Window.WorkerManager));

        var btnVehicle = CreateModuleButton(
            "Vehicle Maintenance",
            "Fleet records, servicing & status",
            () => PortalRouter.Launch(Window.VehicleManager));

        // Add controls in order
        layout.Controls.Add(lblHeader);
        layout.Controls.Add(lblSubheader);
        layout.Controls.Add(btnScheduler);
        layout.Controls.Add(btnProject);
        layout.Controls.Add(btnWorker);
        layout.Controls.Add(btnVehicle);

        Controls.Add(layout);
    }

    private Button CreateModuleButton(string title, string subtitle, Action onClick)
    {
        var btn = new Button
        {
            Width = 340,
            Height = 62,
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 0, 10)
        };

        btn.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);

        btn.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // Render Title
            using (var titleFont = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(Color.FromArgb(30, 41, 59)))
            {
                e.Graphics.DrawString(title, titleFont, titleBrush, 14, 8);
            }

            // Render Subtitle
            using (var subFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (var subBrush = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                e.Graphics.DrawString(subtitle, subFont, subBrush, 14, 32);
            }
        };

        btn.Click += (s, e) => onClick();
        return btn;
    }
}