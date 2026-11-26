namespace SKSB_App.Core.RuntimeUI.BugReport;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.Localization;

/// <summary>Wraps a raw HWND (e.g. <c>ExcelDnaUtil.WindowHandle</c>) so it can parent a WinForms dialog.</summary>
public sealed class BugReportOwnerWindow : IWin32Window
{
    public BugReportOwnerWindow(IntPtr handle) => Handle = handle;
    public IntPtr Handle { get; }
}

/// <summary>
/// Modal "Report a Bug" dialog shared by Portal.exe and every Manager workbook Ribbon.
/// Collects title, description, image/video attachments and (optionally) the runtime trace log tail,
/// then submits to <c>public.bug_reports</c> via <see cref="BugReportRepository"/>.
/// </summary>
public sealed class BugReportSubmitDialog : Form
{
    // =========================================================================
    // 1. PALETTE & STATE
    // =========================================================================
    private static readonly Color Slate900 = Color.FromArgb(15, 23, 42);
    private static readonly Color Slate800 = Color.FromArgb(30, 41, 59);
    private static readonly Color Slate500 = Color.FromArgb(100, 116, 139);
    private static readonly Color Slate400 = Color.FromArgb(148, 163, 184);
    private static readonly Color Slate200 = Color.FromArgb(226, 232, 240);
    private static readonly Color Slate50 = Color.FromArgb(248, 250, 252);
    private static readonly Color Indigo600 = Color.FromArgb(79, 70, 229);
    private static readonly Color Indigo50 = Color.FromArgb(238, 242, 255);

    private readonly BugReportContext _context;
    private readonly BindingList<AttachmentItem> _attachments = new();
    private bool _submitting;

    private TextBox _txtTitle = null!;
    private TextBox _txtDescription = null!;
    private ListBox _lstAttachments = null!;
    private PictureBox _picPreview = null!;
    private Label _lblPreviewInfo = null!;
    private Label _lblTotals = null!;
    private CheckBox _chkLog = null!;
    private Button _btnSubmit = null!;
    private Button _btnCancel = null!;
    private TableLayoutPanel _root = null!;

    private sealed class AttachmentItem
    {
        public AttachmentItem(BugReportAttachmentInput input) => Input = input;
        public BugReportAttachmentInput Input { get; }
        public override string ToString() => $"{Input.FileName}  ({BugReportFormat.FormatSize(Input.FileSize)})";
    }

    public BugReportSubmitDialog(BugReportContext context)
    {
        _context = context ?? new BugReportContext();
        if (string.IsNullOrWhiteSpace(_context.AppVersion)) _context.AppVersion = BugReportRepository.ResolveAppVersion();

        SuspendLayout();
        Text = Loc.T("bugReport.windowTitle");
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.Sizable;
        MinimizeBox = false;
        MaximizeBox = true;
        ShowInTaskbar = false;
        KeyPreview = true;
        BackColor = Slate50;
        Font = new Font("Segoe UI", 9f, FontStyle.Regular);
        Padding = new Padding(16, 14, 16, 14);
        ClientSize = new Size(760, 700);
        MinimumSize = new Size(620, 560);

        InitializeLayout();

        AutoScaleDimensions = new SizeF(96f, 96f);
        AutoScaleMode = AutoScaleMode.Dpi;
        ResumeLayout(false);
        PerformLayout();
    }

    protected override void OnLoad(EventArgs e)
    {
        base.OnLoad(e);
        if (_context.InitialScreenshotPng is { Length: > 0 } shot)
        {
            TryAddAttachment(new BugReportAttachmentInput
            {
                FileName = $"auto_screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png",
                ContentType = "image/png",
                Data = shot
            }, showWarnings: false);
        }
        else
        {
            UpdatePreview();
        }
    }

    // =========================================================================
    // 2. LAYOUT
    // =========================================================================
    private void InitializeLayout()
    {
        _root = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 10,
            Margin = new Padding(0),
            Padding = new Padding(0)
        };
        _root.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 0 header
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 1 title label
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 2 title box
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 3 description label
        _root.RowStyles.Add(new RowStyle(SizeType.Percent, 38f));      // 4 description box
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 5 attachments toolbar
        _root.RowStyles.Add(new RowStyle(SizeType.Percent, 62f));      // 6 list + preview
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 7 log checkbox
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 8 metadata
        _root.RowStyles.Add(new RowStyle(SizeType.AutoSize));          // 9 buttons

        var lblHeader = new Label
        {
            Text = Loc.T("bugReport.header"),
            Font = new Font("Segoe UI", 14f, FontStyle.Bold),
            ForeColor = Slate900,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 2)
        };
        var lblSub = new Label
        {
            Text = Loc.T("bugReport.subtitle"),
            ForeColor = Slate500,
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10)
        };
        var header = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false, Margin = new Padding(0) };
        header.Controls.Add(lblHeader);
        header.Controls.Add(lblSub);

        _txtTitle = new TextBox
        {
            Dock = DockStyle.Fill,
            MaxLength = 200,
            PlaceholderText = Loc.T("bugReport.titlePlaceholder"),
            Margin = new Padding(0, 0, 0, 10)
        };

        _txtDescription = new TextBox
        {
            Dock = DockStyle.Fill,
            Multiline = true,
            AcceptsReturn = true,
            ScrollBars = ScrollBars.Vertical,
            PlaceholderText = Loc.T("bugReport.descPlaceholder"),
            Margin = new Padding(0, 0, 0, 10)
        };

        // --- Attachments toolbar ---
        var toolbar = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.LeftToRight,
            AutoSize = true,
            WrapContents = true,
            Dock = DockStyle.Fill,
            Margin = new Padding(0, 0, 0, 4)
        };
        var lblAttachments = new Label
        {
            Text = Loc.T("bugReport.attachments"),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            ForeColor = Slate800,
            AutoSize = true,
            Margin = new Padding(0, 7, 10, 0)
        };
        var tips = new ToolTip { AutoPopDelay = 15000 };
        var btnShot = CreateToolButton(Loc.T("bugReport.takeScreenshot"));
        var btnPaste = CreateToolButton(Loc.T("bugReport.pasteImage"));
        var btnAdd = CreateToolButton(Loc.T("bugReport.addFile"));
        var btnSnip = CreateToolButton(Loc.T("bugReport.snip"));
        var btnRemove = CreateToolButton(Loc.T("bugReport.remove"));
        tips.SetToolTip(btnShot, Loc.T("bugReport.tip.screenshot"));
        tips.SetToolTip(btnPaste, Loc.T("bugReport.tip.paste"));
        tips.SetToolTip(btnAdd, Loc.T("bugReport.tip.addFile"));
        tips.SetToolTip(btnSnip, Loc.T("bugReport.tip.snip"));
        tips.SetToolTip(btnRemove, Loc.T("bugReport.tip.remove"));
        btnShot.Click += async (s, e) => await TakeScreenshotAsync();
        btnPaste.Click += (s, e) => PasteClipboardImage(showWarnings: true);
        btnAdd.Click += (s, e) => AddFiles();
        btnSnip.Click += (s, e) => LaunchSnippingTool();
        btnRemove.Click += (s, e) => RemoveSelected();
        _lblTotals = new Label { AutoSize = true, ForeColor = Slate500, Margin = new Padding(8, 7, 0, 0) };
        toolbar.Controls.AddRange(new Control[] { lblAttachments, btnShot, btnPaste, btnAdd, btnSnip, btnRemove, _lblTotals });

        // --- List + preview ---
        var split = new TableLayoutPanel
        {
            Dock = DockStyle.Fill,
            ColumnCount = 2,
            RowCount = 1,
            Margin = new Padding(0, 0, 0, 8)
        };
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 40f));
        split.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 60f));
        split.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));

        _lstAttachments = new ListBox
        {
            Dock = DockStyle.Fill,
            IntegralHeight = false,
            BorderStyle = BorderStyle.FixedSingle,
            DataSource = _attachments,
            Margin = new Padding(0, 0, 8, 0)
        };
        _lstAttachments.SelectedIndexChanged += (s, e) => UpdatePreview();
        _lstAttachments.KeyDown += (s, e) => { if (e.KeyCode == Keys.Delete) { RemoveSelected(); e.Handled = true; } };

        var previewHost = new Panel { Dock = DockStyle.Fill, BackColor = Color.White, BorderStyle = BorderStyle.FixedSingle, Margin = new Padding(0) };
        _picPreview = new PictureBox { Dock = DockStyle.Fill, SizeMode = PictureBoxSizeMode.Zoom, BackColor = Color.White };
        _lblPreviewInfo = new Label
        {
            Dock = DockStyle.Fill,
            TextAlign = ContentAlignment.MiddleCenter,
            ForeColor = Slate400,
            Text = Loc.T("bugReport.previewNone"),
            Visible = true
        };
        previewHost.Controls.Add(_picPreview);
        previewHost.Controls.Add(_lblPreviewInfo);
        _lblPreviewInfo.BringToFront();

        split.Controls.Add(_lstAttachments, 0, 0);
        split.Controls.Add(previewHost, 1, 0);

        _chkLog = new CheckBox
        {
            Text = Loc.T("bugReport.attachLog"),
            Checked = true,
            AutoSize = true,
            ForeColor = Slate800,
            Margin = new Padding(0, 0, 0, 6)
        };

        var meta = new List<string> { string.IsNullOrWhiteSpace(_context.SourceModule) ? Loc.T("bugReport.unknownModule") : Loc.ModuleDisplayName(_context.SourceModule) };
        if (!string.IsNullOrWhiteSpace(_context.ActiveSheet)) meta.Add(_context.ActiveSheet);
        meta.Add($"v{_context.AppVersion}");
        meta.Add($"{_context.ReportedBy}@{_context.MachineName}");
        if (_context.IsOffline) meta.Add(Loc.T("bugReport.offline"));
        var lblMeta = new Label
        {
            Text = string.Join("  |  ", meta),
            Font = new Font("Segoe UI", 8f, FontStyle.Regular),
            ForeColor = Indigo600,
            BackColor = Indigo50,
            Padding = new Padding(6, 3, 6, 3),
            AutoSize = true,
            Margin = new Padding(0, 0, 0, 10)
        };

        // --- Buttons ---
        var buttons = new FlowLayoutPanel
        {
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Dock = DockStyle.Fill,
            WrapContents = false,
            Margin = new Padding(0)
        };
        _btnSubmit = new Button
        {
            Text = Loc.T("bugReport.submit"),
            AutoSize = true,
            MinimumSize = new Size(150, 32),
            BackColor = Indigo600,
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
            Cursor = Cursors.Hand,
            Margin = new Padding(8, 0, 0, 0)
        };
        _btnSubmit.FlatAppearance.BorderSize = 0;
        _btnSubmit.Click += async (s, e) => await SubmitAsync();

        _btnCancel = new Button
        {
            Text = Loc.T("common.cancel"),
            AutoSize = true,
            MinimumSize = new Size(90, 32),
            BackColor = Color.White,
            ForeColor = Slate800,
            FlatStyle = FlatStyle.Flat,
            DialogResult = DialogResult.Cancel,
            Margin = new Padding(0)
        };
        _btnCancel.FlatAppearance.BorderColor = Slate200;
        buttons.Controls.Add(_btnSubmit);
        buttons.Controls.Add(_btnCancel);
        CancelButton = _btnCancel;

        _root.Controls.Add(header, 0, 0);
        _root.Controls.Add(CreateFieldLabel(Loc.T("bugReport.titleLabel")), 0, 1);
        _root.Controls.Add(_txtTitle, 0, 2);
        _root.Controls.Add(CreateFieldLabel(Loc.T("bugReport.descLabel")), 0, 3);
        _root.Controls.Add(_txtDescription, 0, 4);
        _root.Controls.Add(toolbar, 0, 5);
        _root.Controls.Add(split, 0, 6);
        _root.Controls.Add(_chkLog, 0, 7);
        _root.Controls.Add(lblMeta, 0, 8);
        _root.Controls.Add(buttons, 0, 9);
        Controls.Add(_root);

        UpdateTotals();
    }

    private static Label CreateFieldLabel(string text) => new()
    {
        Text = text,
        Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        ForeColor = Slate800,
        AutoSize = true,
        Margin = new Padding(0, 0, 0, 3)
    };

    private static Button CreateToolButton(string text)
    {
        var btn = new Button
        {
            Text = text,
            AutoSize = true,
            MinimumSize = new Size(0, 28),
            BackColor = Color.White,
            ForeColor = Slate800,
            FlatStyle = FlatStyle.Flat,
            Cursor = Cursors.Hand,
            Margin = new Padding(0, 0, 6, 4)
        };
        btn.FlatAppearance.BorderColor = Slate200;
        btn.FlatAppearance.MouseOverBackColor = Color.FromArgb(241, 245, 249);
        return btn;
    }

    // =========================================================================
    // 3. ATTACHMENT ACTIONS
    // =========================================================================
    private async Task TakeScreenshotAsync()
    {
        // Opacity (not Hide) keeps the modal loop alive while the dialog is invisible.
        double originalOpacity = Opacity;
        byte[]? png;
        try
        {
            Opacity = 0;
            await Task.Delay(250);
            png = BugReportRepository.CaptureScreenPng();
        }
        finally
        {
            Opacity = originalOpacity;
            Activate();
        }

        if (png == null)
        {
            MessageBox.Show(this, Loc.T("bugReport.captureFailed"), Loc.T("bugReport.takeScreenshot"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }
        TryAddAttachment(new BugReportAttachmentInput
        {
            FileName = $"screenshot_{DateTime.Now:yyyyMMdd_HHmmss}.png",
            ContentType = "image/png",
            Data = png
        }, showWarnings: true);
    }

    private bool PasteClipboardImage(bool showWarnings)
    {
        try
        {
            if (!Clipboard.ContainsImage())
            {
                if (showWarnings)
                    MessageBox.Show(this, Loc.T("bugReport.clipboardNoImage"),
                        Loc.T("bugReport.pasteImage"), MessageBoxButtons.OK, MessageBoxIcon.Information);
                return false;
            }
            using Image? img = Clipboard.GetImage();
            if (img == null) return false;
            return TryAddAttachment(new BugReportAttachmentInput
            {
                FileName = $"clipboard_{DateTime.Now:yyyyMMdd_HHmmss}.png",
                ContentType = "image/png",
                Data = BugReportRepository.EncodePng(img)
            }, showWarnings: true);
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("bugReport.pasteFailed", ex.Message), Loc.T("bugReport.pasteImage"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
    }

    private void AddFiles()
    {
        using var ofd = new OpenFileDialog
        {
            Title = Loc.T("bugReport.addFileTitle"),
            Multiselect = true,
            Filter = FilterName("bugReport.filter.imagesVideos") + "|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp;*.mp4;*.mov;*.webm;*.mkv;*.avi|" +
                     FilterName("bugReport.filter.images") + "|*.png;*.jpg;*.jpeg;*.gif;*.bmp;*.webp|" +
                     FilterName("bugReport.filter.videos") + "|*.mp4;*.mov;*.webm;*.mkv;*.avi|" +
                     FilterName("bugReport.filter.all") + "|*.*"
        };
        if (ofd.ShowDialog(this) != DialogResult.OK) return;

        foreach (string path in ofd.FileNames)
        {
            try
            {
                var info = new FileInfo(path);
                if (info.Length > BugReportRepository.MaxFileBytes)
                {
                    MessageBox.Show(this,
                        Loc.T("bugReport.fileTooLargeDetailed", info.Name, BugReportFormat.FormatSize(info.Length)),
                        Loc.T("bugReport.fileTooLargeTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }
                if (TotalBytes() + info.Length > BugReportRepository.MaxTotalBytes)
                {
                    MessageBox.Show(this,
                        Loc.T("bugReport.limitDetailed", info.Name),
                        Loc.T("bugReport.limitTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    continue;
                }

                byte[] data;
                using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
                using (var ms = new MemoryStream())
                {
                    fs.CopyTo(ms);
                    data = ms.ToArray();
                }
                TryAddAttachment(new BugReportAttachmentInput
                {
                    FileName = info.Name,
                    ContentType = BugReportRepository.GetContentType(info.Name),
                    Data = data
                }, showWarnings: true);
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, Loc.T("bugReport.readFailed", Path.GetFileName(path), ex.Message),
                    Loc.T("bugReport.addFile").TrimEnd('.', ' '), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }

    private void LaunchSnippingTool()
    {
        try
        {
            Process.Start(new ProcessStartInfo("ms-screenclip:") { UseShellExecute = true });
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, Loc.T("bugReport.snipFailed", ex.Message),
                Loc.T("bugReport.snip"), MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }

    private void RemoveSelected()
    {
        if (_lstAttachments.SelectedItem is not AttachmentItem item) return;
        _attachments.Remove(item);
        UpdateTotals();
        UpdatePreview();
    }

    private bool TryAddAttachment(BugReportAttachmentInput input, bool showWarnings)
    {
        if (input.FileSize > BugReportRepository.MaxFileBytes)
        {
            if (showWarnings)
                MessageBox.Show(this, Loc.T("bugReport.fileTooLarge", input.FileName), Loc.T("bugReport.fileTooLargeTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }
        if (TotalBytes() + input.FileSize > BugReportRepository.MaxTotalBytes)
        {
            if (showWarnings)
                MessageBox.Show(this, Loc.T("bugReport.limit", input.FileName), Loc.T("bugReport.limitTitle"), MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return false;
        }

        var item = new AttachmentItem(input);
        _attachments.Add(item);
        _lstAttachments.SelectedItem = item;
        UpdateTotals();
        UpdatePreview();
        return true;
    }

    private long TotalBytes() => _attachments.Sum(a => a.Input.FileSize);

    private void UpdateTotals()
    {
        if (_lblTotals == null) return;
        _lblTotals.Text = _attachments.Count == 0
            ? Loc.T("bugReport.totalsNone")
            : Loc.T("bugReport.totals", _attachments.Count, BugReportFormat.FormatSize(TotalBytes()));
    }

    // =========================================================================
    // 4. PREVIEW (GDI+ images disposed on every switch and on close)
    // =========================================================================
    private void UpdatePreview()
    {
        ClearPreviewImage();

        if (_lstAttachments.SelectedItem is not AttachmentItem item)
        {
            ShowPreviewMessage(_attachments.Count == 0
                ? Loc.T("bugReport.previewEmpty")
                : Loc.T("bugReport.previewSelect"));
            return;
        }

        if (item.Input.IsImage)
        {
            try
            {
                using var ms = new MemoryStream(item.Input.Data);
                using var decoded = Image.FromStream(ms);
                _picPreview.Image = new Bitmap(decoded); // detached copy: no dependency on the stream
                _lblPreviewInfo.Visible = false;
                return;
            }
            catch
            {
                ShowPreviewMessage(Loc.T("bugReport.previewUnavailable", item.Input.FileName));
                return;
            }
        }

        string kindKey = item.Input.ContentType.StartsWith("video/", StringComparison.OrdinalIgnoreCase)
            ? "bugReport.previewVideo"
            : "bugReport.previewFile";
        ShowPreviewMessage(Loc.T(kindKey, item.Input.FileName, BugReportFormat.FormatSize(item.Input.FileSize)));
    }

    private void ShowPreviewMessage(string text)
    {
        _lblPreviewInfo.Text = text;
        _lblPreviewInfo.Visible = true;
        _lblPreviewInfo.BringToFront();
    }

    private void ClearPreviewImage()
    {
        Image? old = _picPreview.Image;
        _picPreview.Image = null;
        old?.Dispose();
    }

    // =========================================================================
    // 5. KEYBOARD (Ctrl+V pastes an image from anywhere in the dialog)
    // =========================================================================
    protected override void OnKeyDown(KeyEventArgs e)
    {
        if (!_submitting && e.Control && e.KeyCode == Keys.V)
        {
            try
            {
                bool textTarget = ActiveControl is TextBoxBase && Clipboard.ContainsText();
                if (!textTarget && Clipboard.ContainsImage())
                {
                    PasteClipboardImage(showWarnings: true);
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    return;
                }
            }
            catch { }
        }
        base.OnKeyDown(e);
    }

    // =========================================================================
    // 6. SUBMIT
    // =========================================================================
    private async Task SubmitAsync()
    {
        if (_submitting) return;
        string title = _txtTitle.Text.Trim();
        if (title.Length == 0)
        {
            MessageBox.Show(this, Loc.T("bugReport.titleRequired"), Loc.T("bugReport.titleRequiredTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            _txtTitle.Focus();
            return;
        }

        string description = _txtDescription.Text;
        var attachments = _attachments.Select(a => a.Input).ToList();
        _context.RuntimeLogTail = _chkLog.Checked ? BugReportRepository.CaptureRuntimeLogTail() : string.Empty;

        SetBusy(true);
        try
        {
            long bugId = await Task.Run(() => BugReportRepository.SubmitBugReport(title, description, _context, attachments));
            SetBusy(false);
            MessageBox.Show(this, Loc.T("bugReport.submitted", bugId), Loc.T("bugReport.submittedTitle"), MessageBoxButtons.OK, MessageBoxIcon.Information);
            DialogResult = DialogResult.OK;
            Close();
        }
        catch (Exception ex)
        {
            SetBusy(false);
            AppLogger.Error("[BugReport] Submit failed", ex);
            MessageBox.Show(this,
                Loc.T("bugReport.submitFailed", ex.Message),
                Loc.T("bugReport.submitFailedTitle"), MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void SetBusy(bool busy)
    {
        _submitting = busy;
        _root.Enabled = !busy;
        UseWaitCursor = busy;
        Cursor = busy ? Cursors.WaitCursor : Cursors.Default;
        _btnSubmit.Text = busy ? Loc.T("bugReport.submitting") : Loc.T("bugReport.submit");
    }

    /// <summary>OpenFileDialog filter names must never contain the '|' separator.</summary>
    private static string FilterName(string key) => Loc.T(key).Replace("|", "/");

    protected override void OnFormClosing(FormClosingEventArgs e)
    {
        if (_submitting)
        {
            e.Cancel = true;
            return;
        }
        base.OnFormClosing(e);
    }

    protected override void OnFormClosed(FormClosedEventArgs e)
    {
        ClearPreviewImage();
        base.OnFormClosed(e);
    }

    protected override void OnShown(EventArgs e)
    {
        base.OnShown(e);
        _txtTitle.Focus();
    }
}
