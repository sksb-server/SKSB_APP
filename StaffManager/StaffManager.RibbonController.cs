namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using ExcelDna.Integration;
using ExcelDna.Integration.CustomUI;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.Ribbon;
using SKSB_App.Core.ThumbprintSync;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookSave.Entry;

[ComVisible(true)]
public class RibbonController : ManagerRibbonBase
{
    // =========================================================================
    // 1. MODULE DESCRIPTOR
    // =========================================================================
    protected override ISaveContract Contract => Events.StaffSyncDescriptor;
    protected override string TabId => "tabStaffManager";
    protected override string TabLabel => "Staff Manager";

    protected override string PersistenceGroupXml => BuildPersistenceGroupXml(
        groupId: "grpStaffPersistence",
        groupLabel: "Staff Database Sync",
        loadId: "btnLoadStaff",
        loadScreentip: "Load Staff Records",
        loadSupertip: "Resyncs staff profiles and salary history from SQL Payroll, loads licenses from PostgreSQL, and checks thumbprint controller mismatches.",
        saveId: "btnSaveStaff",
        saveScreentip: "Save Staff Records",
        saveSupertip: "Validates and commits employee profiles and licenses.");

    protected override string GetCustomGroupsXml() => @"
                <group id='grpThumbprintTerminal' label='Thumbprint Terminal'>
                  <button id='btnPushStaffToTerminal'
                          label='Push ID &amp; Name&#10;to Controller'
                          size='large'
                          imageMso='AccessRefreshAllLists'
                          screentip='Push Staff ID &amp; Name to Controller'
                          supertip='Compares Staff IDs on the thumbprint controller and updates matching users names (preserving fingerprint templates and roles) after confirmation.'
                          onAction='OnPushStaffToTerminalClicked' />
                </group>";

    protected override string EmptyLoadMessage => "No staff records found in database or local snapshot.";

    // =========================================================================
    // 2. EXCEL COMMANDS (must live on a non-generic type)
    // =========================================================================
    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro() => Events.RunDelayedAutoLoad();

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave() => Events.RunHotKeySave();

    // =========================================================================
    // 3. THUMBPRINT TERMINAL PUSH CALLBACK
    // =========================================================================
    public void OnPushStaffToTerminalClicked(IRibbonControl control)
    {
        const string Title = "Push ID & Name to Controller";
        const int MaxNameLen = 23;

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            Dictionary<string, string> staffById;
            try
            {
                var (staffs, _, _) = SheetMapper.ExtractStaffRecordsWithDiagnostics();
                staffById = new Dictionary<string, string>();
                foreach (var s in staffs ?? new List<StaffRecord>())
                {
                    if (s.StaffID <= 0) continue;
                    string name = !string.IsNullOrWhiteSpace(s.StaffDetails?.StaffShortName)
                        ? s.StaffDetails!.StaffShortName.Trim()
                        : (s.StaffDetails?.StaffName ?? "").Trim();
                    if (name.Length == 0) continue;
                    if (name.Length > MaxNameLen) name = name.Substring(0, MaxNameLen);
                    staffById[s.StaffID.ToString()] = name;
                }
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed to read staff records for terminal push", ex);
                MessageBox.Show($"Failed to read staff records: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Task.Run(async () =>
            {
                try
                {
                    var terminalUsers = await DeviceSyncManager.FetchTerminalUsersAsync();

                    var updates = new List<TerminalUserUpdate>();
                    var preview = new List<(string EnrollNumber, string OldName, string NewName)>();
                    int alreadySynced = 0;
                    int unmatched = 0;
                    foreach (var u in terminalUsers)
                    {
                        if (!staffById.TryGetValue(u.EnrollNumber, out var targetName))
                            unmatched++;
                        else if (string.Equals((u.Name ?? "").Trim(), targetName, StringComparison.Ordinal))
                            alreadySynced++;
                        else
                        {
                            updates.Add(new TerminalUserUpdate(u.Uid, u.EnrollNumber, targetName));
                            preview.Add((u.EnrollNumber, (u.Name ?? "").Trim(), targetName));
                        }
                    }

                    preview = preview
                        .OrderBy(p => long.TryParse(p.EnrollNumber, out _) ? 0 : 1)
                        .ThenBy(p => long.TryParse(p.EnrollNumber, out var n) ? n : 0)
                        .ThenBy(p => p.EnrollNumber, StringComparer.Ordinal)
                        .ToList();

                    ExcelAsyncUtil.QueueAsMacro(() =>
                    {
                        if (updates.Count == 0)
                        {
                            MessageBox.Show(
                                $"Nothing to update.\n\nAlready synced: {alreadySynced}\nNot in staff list: {unmatched}",
                                Title, MessageBoxButtons.OK, MessageBoxIcon.Information);
                            return;
                        }

                        if (!ConfirmTerminalNameUpdates(Title, DeviceSyncManager.DefaultIp, preview, alreadySynced, unmatched)) return;

                        Task.Run(async () =>
                        {
                            try
                            {
                                int updated = await DeviceSyncManager.ApplyTerminalUsersAsync(updates);
                                ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                                    $"Updated {updated} user name(s) on the controller.",
                                    Title, MessageBoxButtons.OK, MessageBoxIcon.Information));
                            }
                            catch (Exception ex)
                            {
                                AppLogger.Error("Terminal name push failed", ex);
                                ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                                    $"Push failed: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error));
                            }
                        });
                    });
                }
                catch (Exception ex)
                {
                    AppLogger.Error("Terminal fetch failed", ex);
                    ExcelAsyncUtil.QueueAsMacro(() => MessageBox.Show(
                        $"Could not read users from controller: {ex.Message}", Title, MessageBoxButtons.OK, MessageBoxIcon.Error));
                }
            });
        });
    }

    private static bool ConfirmTerminalNameUpdates(
        string title,
        string hostIp,
        List<(string EnrollNumber, string OldName, string NewName)> preview,
        int alreadySynced,
        int unmatched)
    {
        using var form = new Form
        {
            Text = title,
            StartPosition = FormStartPosition.CenterScreen,
            Size = new Size(580, 440),
            MinimumSize = new Size(460, 300),
            MinimizeBox = false,
            ShowIcon = false,
            Font = new Font("Segoe UI", 9f),
            Padding = new Padding(12)
        };

        var header = new Label
        {
            Dock = DockStyle.Top,
            AutoSize = true,
            Padding = new Padding(0, 0, 0, 8),
            Text = $"Names to update: {preview.Count}   |   Already synced: {alreadySynced}   |   Not in staff list: {unmatched}\n" +
                   $"Controller: {hostIp} — fingerprint templates, cards, PINs and roles are preserved."
        };

        var list = new ListView
        {
            Dock = DockStyle.Fill,
            View = View.Details,
            FullRowSelect = true,
            GridLines = true,
            HeaderStyle = ColumnHeaderStyle.Nonclickable
        };
        list.Columns.Add("ID", 70);
        list.Columns.Add("Current Name (Controller)", 225);
        list.Columns.Add("New Name (Staff Sheet)", 225);
        foreach (var (id, oldName, newName) in preview)
            list.Items.Add(new ListViewItem(new[] { id, string.IsNullOrEmpty(oldName) ? "(blank)" : oldName, newName }));

        var btnCancel = new Button
        {
            Text = "Cancel",
            DialogResult = DialogResult.Cancel,
            AutoSize = true,
            Padding = new Padding(10, 4, 10, 4)
        };
        var btnPush = new Button
        {
            Text = "Push Updates",
            DialogResult = DialogResult.OK,
            AutoSize = true,
            Padding = new Padding(10, 4, 10, 4)
        };
        var footer = new FlowLayoutPanel
        {
            Dock = DockStyle.Bottom,
            FlowDirection = FlowDirection.RightToLeft,
            AutoSize = true,
            Padding = new Padding(0, 10, 0, 0)
        };
        footer.Controls.Add(btnCancel);
        footer.Controls.Add(btnPush);

        form.AcceptButton = btnPush;
        form.CancelButton = btnCancel;
        form.Controls.Add(list);
        form.Controls.Add(header);
        form.Controls.Add(footer);

        return form.ShowDialog() == DialogResult.OK;
    }
}
