namespace SKSB_App.WorkerManager;

using System;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

public class ImportResultDialog : Form
{
    private readonly WorkerImportResult _result;

    public ImportResultDialog(WorkerImportResult result, string fileName)
    {
        _result = result;

        Text = "SQL Payroll Import Summary";
        StartPosition = FormStartPosition.CenterScreen;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        AutoScaleMode = AutoScaleMode.Dpi;
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(24);

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        var sbMessage = new StringBuilder();
        sbMessage.AppendLine("Import completed successfully!");
        sbMessage.AppendLine();
        sbMessage.AppendLine("• Filter Rule: Only synced personnel where Department = 'PROJECT'.");
        sbMessage.AppendLine($"• Office / Admin Staff Filtered Out: {_result.OfficeFilteredCount} record(s).");
        sbMessage.AppendLine($"• Successfully Imported to Table: {_result.ImportedWorkers.Count} worker(s).");
        sbMessage.AppendLine($"• Skipped Entries (Errors): {_result.SkippedErrors.Count} record(s).");
        sbMessage.AppendLine();
        sbMessage.AppendLine("Review the entries in Excel and click 'Save to System' when ready.");

        var lblMessage = new Label
        {
            Text = sbMessage.ToString(),
            Font = new Font("Segoe UI", 9.5f),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 18)
        };

        var flowButtons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill
        };

        var btnClose = new Button
        {
            Text = "OK",
            Font = new Font("Segoe UI", 9f),
            AutoSize = true,
            Padding = new Padding(16, 6, 16, 6),
            DialogResult = DialogResult.OK
        };
        btnClose.Click += (_, _) => Close();
        flowButtons.Controls.Add(btnClose);

        // Show "View Skipped Errors..." button only if entries were skipped due to errors
        if (_result.SkippedErrors.Count > 0)
        {
            var btnViewErrors = new Button
            {
                Text = $"View Skipped Errors ({_result.SkippedErrors.Count})...",
                Font = new Font("Segoe UI", 9f, FontStyle.Bold),
                AutoSize = true,
                Padding = new Padding(12, 6, 12, 6),
                ForeColor = Color.FromArgb(185, 28, 28)
            };
            btnViewErrors.Click += (_, _) => ShowDetailedErrors();
            flowButtons.Controls.Add(btnViewErrors);
        }

        layout.Controls.Add(lblMessage, 0, 0);
        layout.Controls.Add(flowButtons, 0, 1);
        Controls.Add(layout);
    }

    private void ShowDetailedErrors()
    {
        using var errForm = new Form
        {
            Text = "Skipped / Invalid Worker Entries Log",
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
        sb.AppendLine("==========================================================================");
        sb.AppendLine($" SKIPPED ENTRIES REPORT ({_result.SkippedErrors.Count} total)");
        sb.AppendLine(" (Note: OFFICE and ADMIN staff exclusions are not considered errors)");
        sb.AppendLine("==========================================================================");
        sb.AppendLine();

        foreach (var err in _result.SkippedErrors)
        {
            sb.AppendLine($"[Excel Row {err.ExcelRowNumber}]");
            sb.AppendLine($"  • Worker Code : {err.RawCode}");
            sb.AppendLine($"  • Worker Name : {err.Name}");
            sb.AppendLine($"  • Reason      : {err.Reason}");
            sb.AppendLine("--------------------------------------------------------------------------");
        }

        txtLog.Text = sb.ToString();
        errForm.Controls.Add(txtLog);
        errForm.ShowDialog(this);
    }
}