namespace SKSB_App.Core.Format;

using System;
using System.Windows.Forms;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookManager;

public class Events : IExcelAddIn
{
    private static dynamic? _appEvents;
    private delegate void SheetDoubleClickDelegate(dynamic sh, dynamic target, ref bool cancel);
    private static SheetDoubleClickDelegate? _doubleClickHandler;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookDoubleClick();

            try
            {
                XlCall.Excel(XlCall.xlcOnWindow, Type.Missing, "OnWindowActivatedMacro");
            }
            catch { }

            try
            {
                double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
                XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedTemplateCheckMacro");
            }
            catch { }
        });
    }

    public void AutoClose()
    {
        UnhookDoubleClick();
        try
        {
            XlCall.Excel(XlCall.xlcOnWindow, Type.Missing);
        }
        catch { }
    }

    [ExcelCommand(Name = "DelayedTemplateCheckMacro")]
    public static void DelayedTemplateCheckMacro()
    {
        try
        {
            if (WorkbookSession.EvaluateTemplateMode(showHeadsUp: true))
            {
                RibbonController.RefreshRibbonUI();
            }
        }
        catch { }
    }

    [ExcelCommand(Name = "OnWindowActivatedMacro")]
    public static void OnWindowActivatedMacro()
    {
        try
        {
            if (WorkbookSession.EvaluateTemplateMode(showHeadsUp: true))
            {
                RibbonController.RefreshRibbonUI();
            }
        }
        catch { }
    }

    private static void HookDoubleClick()
    {
        try
        {
            dynamic xlApp = ExcelDnaUtil.Application;
            _appEvents = xlApp;
            _doubleClickHandler = new SheetDoubleClickDelegate(OnSheetBeforeDoubleClick);
            xlApp.SheetBeforeDoubleClick += _doubleClickHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[HookDoubleClick] Error: {ex.Message}");
        }
    }

    private static void UnhookDoubleClick()
    {
        try
        {
            if (_appEvents != null && _doubleClickHandler != null)
            {
                _appEvents.SheetBeforeDoubleClick -= _doubleClickHandler;
                _doubleClickHandler = null;
                _appEvents = null;
            }
        }
        catch { }
    }

    private static void OnSheetBeforeDoubleClick(dynamic sh, dynamic target, ref bool cancel)
    {
        try
        {
            // Re-evaluate template status right on click
            WorkbookSession.EvaluateTemplateMode(showHeadsUp: false);

            if (WorkbookSession.IsTemplateEditMode)
            {
                // If in Template Edit Mode, cancel in-cell editing and launch inspector
                if (Commands.IsCustomDropdownCell(target))
                {
                    cancel = true; // Stop the blinking text cursor
                    Commands.InspectAndEditCell(target);
                    return;
                }
            }
            else
            {
                // Operational mode
                if (Commands.IsCustomDropdownCell(target))
                {
                    cancel = true;
                    Commands.ShowChecklistForActiveCell(target);
                }
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[Core DoubleClick Error] {ex.Message}");
        }
    }
}