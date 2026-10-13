namespace SKSB_App.Core.TemplateEditor;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.WorkbookManager;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;
    private delegate void SheetDoubleClickDelegate(dynamic sh, dynamic target, ref bool cancel);
    private delegate void SheetSelectionChangeDelegate(dynamic sh, dynamic target);

    private static SheetDoubleClickDelegate? _doubleClickHandler;
    private static SheetSelectionChangeDelegate? _selectionHandler;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookEvents();

            // 1. Hook window switch / open events
            try
            {
                XlCall.Excel(XlCall.xlcOnWindow, Type.Missing, "OnWindowActivatedMacro");
            }
            catch { }

            // 2. Delayed cold-boot timer (allows Excel window to initialize fully)
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
        UnhookEvents();
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
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - DelayedTemplateCheckMacro] Error: {ex.Message}");
        }
    }

    [ExcelCommand(Name = "OnWindowActivatedMacro")]
    public static void OnWindowActivatedMacro()
    {
        try
        {
            if (WorkbookSession.EvaluateTemplateMode(showHeadsUp: false))
            {
                RibbonController.RefreshRibbonUI();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - OnWindowActivatedMacro] Error: {ex.Message}");
        }
    }

    private static void HookEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;

            dynamic xlApp = ExcelDnaUtil.Application;
            _xlAppRoot = xlApp;

            _doubleClickHandler = new SheetDoubleClickDelegate(OnSheetBeforeDoubleClick);
            _xlAppRoot.SheetBeforeDoubleClick += _doubleClickHandler;

            _selectionHandler = new SheetSelectionChangeDelegate(OnSheetSelectionChange);
            _xlAppRoot.SheetSelectionChange += _selectionHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - HookEvents Error] {ex.Message}");
        }
    }

    private static void UnhookEvents()
    {
        try
        {
            if (_xlAppRoot != null)
            {
                if (_doubleClickHandler != null)
                {
                    _xlAppRoot.SheetBeforeDoubleClick -= _doubleClickHandler;
                    _doubleClickHandler = null;
                }
                if (_selectionHandler != null)
                {
                    _xlAppRoot.SheetSelectionChange -= _selectionHandler;
                    _selectionHandler = null;
                }
                _xlAppRoot = null;
            }
        }
        catch { }
    }

    private static void OnSheetSelectionChange(dynamic sh, dynamic target)
    {
        try
        {
            if (target == null) return;
            if (target.Cells.Count > 100) return;

            dynamic singleCell = target.Cells[1, 1];
            if (Commands.IsCustomDropdownCell(singleCell))
            {
                Commands.EnsureDropdownValidationAttached(singleCell);
            }
        }
        catch { }
    }

    private static void OnSheetBeforeDoubleClick(dynamic sh, dynamic target, ref bool cancel)
    {
        try
        {
            if (!Commands.IsCustomDropdownCell(target))
                return;

            cancel = true;

            // Uses current state determined either by auto-detection or the Ribbon toggle
            if (WorkbookSession.IsTemplateEditMode)
            {
                Commands.InspectAndEditCell(target);
            }
            else
            {
                Commands.ShowChecklistForActiveCell(target);
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - OnSheetBeforeDoubleClick Error] {ex.Message}");
        }
    }
}