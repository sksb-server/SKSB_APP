namespace SKSB_App.Core.DropDown;

using System;
using System.Drawing;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.TemplateEditor;

public partial class ConfiguratorForm : Form
{
    private readonly dynamic _singleCell;
    private readonly TextBox _txtSource;
    private readonly TextBox _txtTransform;
    private readonly TextBox _txtFilter;
    private readonly TextBox _txtSync;
    private readonly CheckBox _chkMultiSelect;
    private readonly Button _btnConfirm;
    private readonly Button _btnCancel;

    public ConfiguratorForm(dynamic singleCell)
    {
        _singleCell = singleCell;
        string cellAddress = _singleCell.Address?.ToString() ?? "Selected Cell";

        Text = $"Dropdown Configurator - {cellAddress}";
        FormBorderStyle = FormBorderStyle.FixedDialog;
        StartPosition = FormStartPosition.CenterScreen;
        MaximizeBox = false;
        MinimizeBox = false;
        TopMost = true;
        AutoScaleMode = AutoScaleMode.Dpi;
        Font = new Font("Segoe UI", 9.5f, FontStyle.Regular);
        AutoSize = true;
        AutoSizeMode = AutoSizeMode.GrowAndShrink;
        Padding = new Padding(18);

        var config = Metadata.ReadConfig(_singleCell);
        if (string.IsNullOrWhiteSpace(config.Source))
        {
            config.Source = Commands.ResolveContextListFormula(_singleCell);
        }

        var layout = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink
        };

        var lblSource = new Label
        {
            Text = "1. Source List / Table Column:\n   (e.g. =AllWorkersTable[Worker Name], =DepartmentList)",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 4)
        };
        _txtSource = new TextBox
        {
            Text = config.Source,
            Width = 440,
            Font = new Font("Segoe UI", 9.5f),
            Margin = new Padding(0, 0, 0, 10)
        };

        var lblTransform = new Label
        {
            Text = "2. Item Transform Formula (Optional. Use @item as placeholder):\n   (e.g. LET(p, TEXTSPLIT(TRIM(@item), \" \"), INDEX(p,1) & \" \" & INDEX(p,COUNTA(p))))",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 4)
        };
        _txtTransform = new TextBox
        {
            Text = config.Transform,
            Width = 440,
            Font = new Font("Segoe UI", 9.5f),
            Margin = new Padding(0, 0, 0, 10)
        };

        var lblFilter = new Label
        {
            Text = "3. Extra Filter / Condition (Optional. Use @item, @left, @right):\n   (e.g. IF(@left<>\"\", XLOOKUP(VALUE(@left), AllWorkersTable[Worker ID], AllWorkersTable[Worker Name], \"\") = @item, TRUE))",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 4)
        };
        _txtFilter = new TextBox
        {
            Text = config.Filter,
            Width = 440,
            Font = new Font("Segoe UI", 9.5f),
            Margin = new Padding(0, 0, 0, 10)
        };

        var lblSync = new Label
        {
            Text = "4. Auto-Fill Adjacent Cell Formula (Optional. Resolves when item is picked):\n   (e.g. XLOOKUP(@item, AllWorkersTable[Worker Name], AllWorkersTable[Worker ID], \"\"))",
            AutoSize = true,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Margin = new Padding(0, 0, 0, 4)
        };
        _txtSync = new TextBox
        {
            Text = config.SyncLeft,
            Width = 440,
            Font = new Font("Segoe UI", 9.5f),
            Margin = new Padding(0, 0, 0, 10)
        };

        _chkMultiSelect = new CheckBox
        {
            Text = "Allow Multi-Selections (Checklist popup)",
            Checked = config.IsMultiSelect,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 14)
        };

        var pnlButtons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            AutoSizeMode = AutoSizeMode.GrowAndShrink,
            Dock = DockStyle.Fill,
            Margin = new Padding(0)
        };

        _btnCancel = new Button
        {
            Text = "Cancel",
            AutoSize = true,
            Padding = new Padding(16, 6, 16, 6),
            DialogResult = DialogResult.Cancel
        };

        _btnConfirm = new Button
        {
            Text = "Confirm / Apply",
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            AutoSize = true,
            Padding = new Padding(18, 6, 18, 6),
            DialogResult = DialogResult.OK,
            BackColor = Color.FromArgb(240, 244, 250)
        };

        _btnConfirm.Click += OnConfirmClicked;
        _btnCancel.Click += (_, _) => Close();

        pnlButtons.Controls.Add(_btnCancel);
        pnlButtons.Controls.Add(_btnConfirm);

        layout.Controls.Add(lblSource);
        layout.Controls.Add(_txtSource);
        layout.Controls.Add(lblTransform);
        layout.Controls.Add(_txtTransform);
        layout.Controls.Add(lblFilter);
        layout.Controls.Add(_txtFilter);
        layout.Controls.Add(lblSync);
        layout.Controls.Add(_txtSync);
        layout.Controls.Add(_chkMultiSelect);
        layout.Controls.Add(pnlButtons);

        Controls.Add(layout);
        AcceptButton = _btnConfirm;
        CancelButton = _btnCancel;
    }

    private void OnConfirmClicked(object? sender, EventArgs e)
    {
        string rawSrc = _txtSource.Text.Trim();
        if (string.IsNullOrEmpty(rawSrc))
        {
            MessageBox.Show("Source List cannot be empty.", "Validation", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var newConfig = new Config
        {
            Source = rawSrc,
            Transform = _txtTransform.Text.Trim(),
            Filter = _txtFilter.Text.Trim(),
            SyncLeft = _txtSync.Text.Trim(),
            IsMultiSelect = _chkMultiSelect.Checked
        };

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                Commands.EnsureNamedStyleExists(xlApp, Commands.DropDownStyleName);
                _singleCell.Style = Commands.DropDownStyleName;

                Metadata.WriteConfig(_singleCell, newConfig);

                try { _singleCell.Validation.Delete(); } catch { }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Excel rejected validation setup: {ex.Message}", "Config Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        });

        Close();
    }
}