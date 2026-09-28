using SKSB_App.Core.FileManager.Constants;
using System;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Text;
using System.Windows.Forms;

namespace SKSB_App.Portal;

public partial class PortalForm : Form
{
    public PortalForm()
    {
        Text = "SKSB Operations Portal";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedSingle;
        MaximizeBox = false;
        ClientSize = new Size(420, 430);
        BackColor = Color.FromArgb(248, 250, 252);

        InitializeLayout();
    }

    private void InitializeLayout()
    {
        // Header Label
        var lblHeader = new Label
        {
            Text = "Operations Hub",
            Font = new Font("Segoe UI", 16, FontStyle.Bold),
            ForeColor = Color.FromArgb(15, 23, 42),
            Location = new Point(30, 25),
            AutoSize = true
        };

        var lblSubheader = new Label
        {
            Text = "Select an operational module to open in Excel:",
            Font = new Font("Segoe UI", 9, FontStyle.Regular),
            ForeColor = Color.FromArgb(100, 116, 139),
            Location = new Point(32, 58),
            AutoSize = true
        };

        // Module Buttons (70px vertical increments)
        var btnScheduler = CreateModuleButton(
            "Daily Work Allocation",
            "Biometric logs & shift scheduling",
            95,
            () => PortalRouter.Launch(Window.DailyScheduler));

        var btnProject = CreateModuleButton(
            "Project Manager",
            "Job sites, scope, priorities & timelines",
            165,
            () => PortalRouter.Launch(Window.ProjectManager));

        var btnStaff = CreateModuleButton(
            "Staff Manager",
            "Employee profiles, badges & privileges",
            235,
            () => PortalRouter.Launch(Window.StaffManager));

        var btnVehicle = CreateModuleButton(
            "Vehicle Maintenance",
            "Fleet records, servicing & status",
            305,
            () => PortalRouter.Launch(Window.VehicleManager));

        Controls.AddRange(new Control[] { lblHeader, lblSubheader, btnScheduler, btnProject, btnStaff, btnVehicle });
    }

    private Button CreateModuleButton(string title, string subtitle, int top, Action onClick)
    {
        var btn = new Button
        {
            Location = new Point(30, top),
            Size = new Size(360, 58),
            BackColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            TextAlign = ContentAlignment.TopLeft,
            Padding = new Padding(12, 8, 8, 8)
        };

        btn.FlatAppearance.BorderColor = Color.FromArgb(226, 232, 240);
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);

        btn.Paint += (s, e) =>
        {
            // 1. Enable smooth shapes and curves
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;

            // 2. Enable crisp sub-pixel ClearType text
            e.Graphics.TextRenderingHint = TextRenderingHint.ClearTypeGridFit;

            // Draw title
            using (var titleFont = new Font("Segoe UI", 10.5f, FontStyle.Bold))
            using (var titleBrush = new SolidBrush(Color.FromArgb(30, 41, 59)))
            {
                e.Graphics.DrawString(title, titleFont, titleBrush, 14, 9);
            }

            // Draw subtitle
            using (var subFont = new Font("Segoe UI", 8.5f, FontStyle.Regular))
            using (var subBrush = new SolidBrush(Color.FromArgb(100, 116, 139)))
            {
                e.Graphics.DrawString(subtitle, subFont, subBrush, 14, 31);
            }
        };

        btn.Click += (s, e) => onClick();
        return btn;
    }
}