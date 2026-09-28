namespace SKSB_App.StaffManager;

using ExcelDna.Integration;
using SKSB_App.Core.FileManager;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookManager;
using System;
using System.Windows.Forms;

public class Events : IExcelAddIn
{
    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private static dynamic? _appEvents;
    private static WorkbookBeforeCloseDelegate? _closeHandler;

    public static bool IsTemplateEditMode { get; set; } = false;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookApplicationClose(); //[cite: 10]

            // 1. Synchronise Staff Manager specific enums into Lookups tab
            SyncStaffEnums();
            SyncStaffLookupsData();

            // 2. Delayed auto-load live records[cite: 10]
            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadMacro"); //[cite: 10]

            // 3. Hotkeys[cite: 10]
            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeySave"); //[cite: 10]
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeySave"); //[cite: 10]
        });
    }

    public void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            UnhookApplicationClose();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
            XlCall.Excel(XlCall.xlcMessage, false);
        });
    }

    public static void SyncStaffEnums()
    {
        TypesToExcel.Sync(
            typeof(Department),
            typeof(Shift),
            typeof(PayType),
            typeof(ActiveStatus),
            typeof(LeaveType),
            typeof(Session),
            typeof(LicenseClass),
            typeof(LicenseStatus), // Generates =LicenseStatusList for Column E
            typeof(VehicleType)
            );
    }

    public static void SyncStaffLookupsData()
    {
        var vehicles = Services.Load<List<Vehicle>>();
        if (vehicles != null && vehicles.Count > 0)
        {
            ServerLookups.SyncValues("VehicleID", vehicles, v => v.VehicleID);
            ServerLookups.SyncValues("VehiclePlate", vehicles, v => v.Master.RegistrationNumber);
        }
    }

    private static void HookApplicationClose()
    {
        try
        {
            // Dynamic binding via ExcelDnaUtil (requires zero COM references in .csproj)
            dynamic xlApp = ExcelDnaUtil.Application;
            _appEvents = xlApp;
            _closeHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            xlApp.WorkbookBeforeClose += _closeHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Events] Failed to hook WorkbookBeforeClose: {ex.Message}");
        }
    }

    private static void UnhookApplicationClose()
    {
        try
        {
            if (_appEvents != null && _closeHandler != null)
            {
                _appEvents.WorkbookBeforeClose -= _closeHandler;
                _closeHandler = null;
                _appEvents = null;
            }
        }
        catch { }
    }

    private static void OnWorkbookBeforeClose(dynamic wb, ref bool cancel)
    {
        try
        {
            if (IsTemplateEditMode || WorkbookSession.IsMasterTemplateFile()) return;

            while (true)
            {
                bool saved = ServerSyncCommands.SaveStaffSilent();

                if (saved)
                {
                    MessageBox.Show(
                        "Auto-save completed successfully. Your changes have been synchronised to the system.",
                        "SKSB Staff Manager",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information
                    );

                    wb.Saved = true; // Tell Excel it's clean so it doesn't prompt[cite: 37]
                    return;
                }

                // Save failed -> Prompt user
                var choice = MessageBox.Show(
                    "Unable to synchronise staff records to the database.\n\n" +
                    "• Click 'Retry' to attempt saving to the database again.\n" +
                    "• Click 'Cancel' to return to Excel without closing.",
                    "System Synchronisation Warning",
                    MessageBoxButtons.RetryCancel,
                    MessageBoxIcon.Warning
                );

                if (choice == DialogResult.Cancel)
                {
                    cancel = true; // Keep workbook open so data isn't lost[cite: 37]
                    return;
                }

                // If user clicked 'Retry', loop continues and calls SaveStaffSilent() again[cite: 34]
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[OnWorkbookBeforeClose] Error: {ex.Message}");
        }
    }

    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro()
    {
        try
        {
            if (WorkbookSession.IsMasterTemplateFile())
            {
                IsTemplateEditMode = true;
                WorkbookSession.ShowStatusConfirmation("Template Design Mode: Auto-load bypassed.");
                return;
            }

            if (IsTemplateEditMode)
            {
                WorkbookSession.ShowStatusConfirmation("Edit Mode Active: Auto-load bypassed.");
                return;
            }

            bool success = ServerSyncCommands.LoadStaffSilent();
            if (success)
            {
                WorkbookSession.ShowStatusConfirmation("Staff records loaded successfully from system.");
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[DelayedAutoLoad] Error: {ex.Message}");
        }
    }

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave()
    {
        if (IsTemplateEditMode || WorkbookSession.IsMasterTemplateFile())
        {
            XlCall.Excel(XlCall.xlcSave);
            return;
        }

        bool success = ServerSyncCommands.SaveStaffSilent();
        if (success)
        {
            MessageBox.Show("Save completed successfully.", "SKSB Staff Manager",
                            MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}