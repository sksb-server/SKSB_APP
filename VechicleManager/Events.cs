namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Services;
using SKSB_App.Core.Types;
using System;
using System.Windows.Forms;
using SKSB_App.Core.WorkbookManager;

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
            HookApplicationClose();

            // Sync strictly vehicle-related enums
            TypesToExcel.Sync(
                typeof(VehicleType),
                typeof(OperationalStatus),
                typeof(ServiceType),
                typeof(MaintenanceStatus)
            );

            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedAutoLoadMacro");

            XlCall.Excel(XlCall.xlcOnKey, "^s", "OnHotKeySave");
            XlCall.Excel(XlCall.xlcOnKey, "^+s", "OnHotKeySave");
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

    private static void HookApplicationClose()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            _appEvents = xlApp;
            _closeHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            xlApp.WorkbookBeforeClose += _closeHandler;
        }
        catch { }
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

            bool saved = ServerSyncCommands.SaveVehiclesSilent();

            if (saved)
            {
                MessageBox.Show("Auto-save completed successfully. Your changes have been synchronised to the system.",
                                "SKSB Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
                wb.Saved = true;
            }
            else
            {
                var choice = MessageBox.Show("Unable to synchronise vehicle records to the database.\n\nClick 'Retry' to return to Excel, or 'Cancel' to close without saving.",
                                             "System Synchronisation Warning", MessageBoxButtons.RetryCancel, MessageBoxIcon.Warning);

                if (choice == DialogResult.Cancel) wb.Saved = true;
                else cancel = true;
            }
        }
        catch { }
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

            if (ServerSyncCommands.LoadVehiclesSilent())
            {
                WorkbookSession.ShowStatusConfirmation("Vehicle records loaded successfully from system.");
            }
        }
        catch { }
    }

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave()
    {
        if (IsTemplateEditMode || WorkbookSession.IsMasterTemplateFile())
        {
            XlCall.Excel(XlCall.xlcSave);
            return;
        }

        if (ServerSyncCommands.SaveVehiclesSilent())
        {
            MessageBox.Show("Save completed successfully.", "SKSB Vehicle Manager", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}