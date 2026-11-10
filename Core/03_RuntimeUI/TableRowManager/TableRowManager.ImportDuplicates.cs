namespace SKSB_App.Core.RuntimeUI.TableRowManager;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SKSB_App.Core.Diagnostics;

public static partial class TableRowManager
{
    /// <summary>
    /// Isolates and skips duplicate records while preserving and returning clean unique records.
    /// If duplicates are found, displays a warning dialog showing what was skipped.
    /// Returns the list of non-duplicate records to import (or null if the user clicks Cancel).
    /// </summary>
    public static List<T>? FilterDuplicates<T>(
        IEnumerable<T> incomingRecords,
        Func<T, string> keySelector,
        string entityLabel,
        out List<string> duplicateWarnings,
        Func<T, string>? displaySelector = null,
        IEnumerable<T>? existingRecords = null)
    {
        duplicateWarnings = new List<string>();
        if (incomingRecords == null) return new List<T>();

        var incomingList = incomingRecords.ToList();
        if (incomingList.Count == 0) return incomingList;

        displaySelector ??= keySelector;

        // 1. Build fast O(1) set of existing table signatures
        var existingKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        if (existingRecords != null)
        {
            foreach (var existing in existingRecords)
            {
                if (existing == null) continue;
                string k = keySelector(existing)?.Trim() ?? string.Empty;
                if (!string.IsNullOrEmpty(k))
                {
                    existingKeys.Add(k);
                }
            }
        }

        // 2. Separate records in RAM
        var seenBatchKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var cleanRecords = new List<T>();
        var duplicateRecords = new List<T>();

        foreach (var item in incomingList)
        {
            if (item == null) continue;
            string key = keySelector(item)?.Trim() ?? string.Empty;
            string display = displaySelector(item)?.Trim() ?? key;

            if (string.IsNullOrEmpty(key))
            {
                cleanRecords.Add(item);
                continue;
            }

            bool isExistingDuplicate = existingKeys.Contains(key);
            bool isFileDuplicate = !seenBatchKeys.Add(key);

            if (isExistingDuplicate || isFileDuplicate)
            {
                duplicateRecords.Add(item);
                string reason = isExistingDuplicate
                    ? "already exists in destination table"
                    : "appears more than once in import file";

                string notice = $"[Duplicate Skipped] {entityLabel} '{display}' {reason}.";
                duplicateWarnings.Add(notice);
                AppLogger.Warn(notice);
            }
            else
            {
                cleanRecords.Add(item);
            }
        }

        // 3. If zero duplicates found, return clean records immediately
        if (duplicateRecords.Count == 0)
        {
            return cleanRecords;
        }

        // 4. If all records are duplicates, notify and return empty list
        if (cleanRecords.Count == 0)
        {
            MessageBox.Show(
                $"All {duplicateRecords.Count} {entityLabel} records in this file are duplicates that already exist in the system.\n\nNo new records to import.",
                $"{entityLabel} Import - All Duplicates",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            return cleanRecords;
        }

        // 5. Present notification dialog allowing user to import clean records or cancel
        using var dialog = new DuplicateSkippedSummaryDialog(
            entityLabel: entityLabel,
            duplicateCount: duplicateRecords.Count,
            cleanCount: cleanRecords.Count,
            skippedWarnings: duplicateWarnings
        );

        var result = dialog.ShowDialog();
        return result == DialogResult.OK ? cleanRecords : null;
    }
}

public class DuplicateSkippedSummaryDialog : Form
{
    public DuplicateSkippedSummaryDialog(
        string entityLabel,
        int duplicateCount,
        int cleanCount,
        List<string> skippedWarnings)
    {
        Text = $"{entityLabel} Import - Duplicates Excluded";
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

        var pnlIndicator = new Panel
        {
            Size = new Size(38, 38),
            Margin = new Padding(0, 2, 16, 0)
        };
        pnlIndicator.Paint += (s, e) =>
        {
            e.Graphics.SmoothingMode = SmoothingMode.AntiAlias;
            using var brush = new SolidBrush(Color.FromArgb(217, 119, 6)); // Amber Warning
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
            Text = "Duplicates Skipped During Import",
            Font = new Font("Segoe UI", 12f, FontStyle.Bold),
            ForeColor = Color.FromArgb(30, 41, 59),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 6)
        };

        var sbMessage = new StringBuilder();
        sbMessage.AppendLine($"Duplicate entries were detected and excluded to protect existing data.");
        sbMessage.AppendLine();
        sbMessage.AppendLine($"  • Duplicate Records Excluded : {duplicateCount}");
        sbMessage.AppendLine($"  • Clean Records to Import    : {cleanCount}");
        sbMessage.AppendLine();
        sbMessage.AppendLine($"Would you like to proceed with importing the {cleanCount} clean record(s)?");

        var lblMessage = new Label
        {
            Text = sbMessage.ToString(),
            Font = new Font("Segoe UI", 9.5f),
            ForeColor = Color.FromArgb(71, 85, 105),
            AutoSize = true,
            MaximumSize = new Size(520, 0),
            Margin = new Padding(0, 0, 0, 8)
        };

        var txtLog = new TextBox
        {
            Multiline = true,
            ReadOnly = true,
            ScrollBars = ScrollBars.Vertical,
            Width = 520,
            Height = 130,
            Font = new Font("Consolas", 8.8f),
            BackColor = Color.FromArgb(254, 252, 232),
            ForeColor = Color.FromArgb(146, 64, 14),
            BorderStyle = BorderStyle.FixedSingle,
            Margin = new Padding(0, 0, 0, 16),
            Text = string.Join(Environment.NewLine, skippedWarnings)
        };

        var flowActions = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            Margin = Padding.Empty
        };

        var btnCancel = new Button
        {
            Text = "Cancel Import",
            Font = new Font("Segoe UI", 9f),
            AutoSize = true,
            Padding = new Padding(14, 6, 14, 6),
            DialogResult = DialogResult.Cancel
        };
        btnCancel.Click += (_, _) => { DialogResult = DialogResult.Cancel; Close(); };
        flowActions.Controls.Add(btnCancel);

        var btnProceed = new Button
        {
            Text = $"Import {cleanCount} Clean Record(s)",
            Font = new Font("Segoe UI", 9.5f, FontStyle.Bold),
            ForeColor = Color.White,
            BackColor = Color.FromArgb(79, 70, 229),
            FlatStyle = FlatStyle.Flat,
            AutoSize = true,
            Padding = new Padding(16, 6, 16, 6),
            DialogResult = DialogResult.OK
        };
        btnProceed.FlatAppearance.BorderSize = 0;
        btnProceed.Click += (_, _) => { DialogResult = DialogResult.OK; Close(); };
        flowActions.Controls.Add(btnProceed);

        pnlContent.Controls.Add(lblTitle);
        pnlContent.Controls.Add(lblMessage);
        pnlContent.Controls.Add(txtLog);
        pnlContent.Controls.Add(flowActions);

        layout.Controls.Add(pnlIndicator, 0, 0);
        layout.Controls.Add(pnlContent, 1, 0);
        Controls.Add(layout);

        AcceptButton = btnProceed;
        CancelButton = btnCancel;
    }
}