namespace SKSB_App.Core.FileManager;

using System.Drawing;
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

    public ExitSyncConflictDialog(string moduleName, string errorMessage)
    {
        Text = $"{moduleName} - Synchronization Alert";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(20);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        var lblMsg = new Label
        {
            Text = $"Unable to persist changes to the server database.\n\n" +
                   $"Reason / Error: {errorMessage}\n\n" +
                   "What would you like to do?",
            Font = new Font("Segoe UI", 9.5f),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 16)
        };

        var flowButtons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill
        };

        var btnCancel = new Button
        {
            Text = "Cancel (Stay in Excel)",
            AutoSize = true,
            Padding = new Padding(10, 5, 10, 5),
            DialogResult = DialogResult.Cancel
        };
        btnCancel.Click += (_, _) => { Resolution = ExitResolution.Cancel; Close(); };

        var btnSaveLocal = new Button
        {
            Text = "Save Locally & Quit",
            AutoSize = true,
            Padding = new Padding(10, 5, 10, 5)
        };
        btnSaveLocal.Click += (_, _) => { Resolution = ExitResolution.SaveLocallyAndQuit; Close(); };

        var btnRetry = new Button
        {
            Text = "Retry Server",
            AutoSize = true,
            Padding = new Padding(10, 5, 10, 5)
        };
        btnRetry.Click += (_, _) => { Resolution = ExitResolution.Retry; Close(); };

        flowButtons.Controls.Add(btnCancel);
        flowButtons.Controls.Add(btnSaveLocal);
        flowButtons.Controls.Add(btnRetry);

        layout.Controls.Add(lblMsg, 0, 0);
        layout.Controls.Add(flowButtons, 0, 1);

        Controls.Add(layout);
    }
}