namespace SKSB_App.Core.TemplateEditor;

using ExcelDna.Integration;
using SKSB_App.Core.DropDown;
using SKSB_App.Core.Services;
using SKSB_App.Core.WorkbookHelper;
using System;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;
    private delegate void SheetDoubleClickDelegate(dynamic sh, dynamic target, ref bool cancel);
    private delegate void SheetSelectionChangeDelegate(dynamic sh, dynamic target);
    private delegate void SheetChangeDelegate(dynamic sh, dynamic target);

    private static SheetDoubleClickDelegate? _doubleClickHandler;
    private static SheetSelectionChangeDelegate? _selectionHandler;
    private static SheetChangeDelegate? _changeHandler;

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            EventHooks();

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
        Arrow.HideAll();
        UnEventHooks();
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
            // NEW (using what you already have):
            bool isTemplate = SheetOps.IsMasterTemplateFile();
            SheetOps.IsTemplateEditMode = isTemplate;
            if (isTemplate)
            {
                SheetOps.SetAuxiliarySheetsVisibility(visible: true);
                TypesToExcel.SyncAllCoreEnums();
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
            // NEW (using what you already have):
            bool isTemplate = SheetOps.IsMasterTemplateFile();
            SheetOps.IsTemplateEditMode = isTemplate;
            if (isTemplate)
            {
                SheetOps.SetAuxiliarySheetsVisibility(visible: true);
                TypesToExcel.SyncAllCoreEnums();
                RibbonController.RefreshRibbonUI();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - OnWindowActivatedMacro] Error: {ex.Message}");
        }
    }

    private static void EventHooks()
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

            _changeHandler = new SheetChangeDelegate(OnSheetChange);
            _xlAppRoot.SheetChange += _changeHandler;
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - EventHooks Error] {ex.Message}");
        }
    }

    private static void UnEventHooks()
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
                if (_changeHandler != null)
                {
                    _xlAppRoot.SheetChange -= _changeHandler;
                    _changeHandler = null;
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
            if (target == null || target.Cells.Count > 1)
            {
                Arrow.HideAll();
                return;
            }

            dynamic singleCell = target.Cells[1, 1];
            if (Commands.IsCustomDropdownCell(singleCell))
            {
                if (!SheetOps.IsTemplateEditMode)
                {
                    Arrow.ShowArrowForCell(singleCell);
                }
                else
                {
                    Arrow.HideAll();
                }
            }
            else
            {
                Arrow.HideAll();
            }
        }
        catch
        {
            Arrow.HideAll();
        }
    }

    private static void OnSheetChange(dynamic sh, dynamic target)
    {
        try
        {
            if (target != null && target.Cells.Count == 1)
            {
                if (Commands.IsCustomDropdownCell(target) && !SheetOps.IsTemplateEditMode)
                {
                    Arrow.ShowArrowForCell(target);
                }
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

            if (SheetOps.IsTemplateEditMode)
            {
                cancel = true; // Open template configurator
                Arrow.HideAll();
                Commands.InspectAndEditCell(target);
            }
            else
            {
                // In Operational Mode: hide the arrow to type cleanly
                Arrow.HideAll();
            }
        }
        catch (Exception ex)
        {
            System.Diagnostics.Debug.WriteLine($"[TemplateEditor - OnSheetBeforeDoubleClick Error] {ex.Message}");
        }
    }
}