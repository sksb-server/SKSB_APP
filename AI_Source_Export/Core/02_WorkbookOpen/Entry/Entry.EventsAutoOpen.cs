namespace SKSB_App.Core.WorkbookOpen.Entry;

using System;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.RuntimeUI.RuleEngine;
using SKSB_App.Core.RuntimeUI.TableRowManager;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.SheetMapper;
using DropDownCommands = SKSB_App.Core.RuntimeUI.DropDown.Commands;

public class EventsAutoOpen : IExcelAddIn
{
    private static dynamic? _xlAppRoot;

    private delegate void SheetDoubleClickDelegate(dynamic sh, dynamic target, ref bool cancel);
    private delegate void SheetSelectionChangeDelegate(dynamic sh, dynamic target);
    private delegate void SheetChangeDelegate(dynamic sh, dynamic target);

    private static SheetDoubleClickDelegate? _doubleClickHandler;
    private static SheetSelectionChangeDelegate? _selectionHandler;
    private static SheetChangeDelegate? _changeHandler;

    // Cross-ALC process-wide state keys so Core-AddIn and Manager-AddIn never duplicate work
    private static string HookedPidKey => $"SKSB_EVENTS_HOOKED_{Environment.ProcessId}";
    private static string LastCheckedWorkbookKey => $"SKSB_LAST_CHECKED_WB_{Environment.ProcessId}";
    private static string LastTemplateModeKey => $"SKSB_TEMPLATE_MODE_{Environment.ProcessId}";

    public void AutoOpen()
    {
        RuleRegistry.EnsureInitialized();

        // If another add-in (e.g. Core-AddIn-packed.xll) already hooked COM sheet events in this process, do not double-hook!
        if (Environment.GetEnvironmentVariable(HookedPidKey) == "1")
        {
            AppLogger.Info($"[Coexistence] Core COM events already active in PID {Environment.ProcessId} (XLL: {ExcelDnaUtil.XllPath}).");
            return;
        }

        Environment.SetEnvironmentVariable(HookedPidKey, "1");
        AppLogger.ResetSessionLog("Workbook Startup");
        AppLogger.Info($"EventsAutoOpen.AutoOpen primary hook in PID {Environment.ProcessId} (XLL: {ExcelDnaUtil.XllPath}).");

        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            using var _ = AppLogger.Scope("AutoOpen Queued Macro");
            EventHooks();

            try
            {
                dynamic xlApp = ExcelDnaUtil.Application;
                xlApp.WindowActivate += new Action<dynamic, dynamic>((wb, wn) => OnWindowActivatedMacro());
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed hooking WindowActivate", ex);
            }

            try
            {
                double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
                XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), "DelayedTemplateCheckMacro");
            }
            catch (Exception ex)
            {
                AppLogger.Error("Failed scheduling DelayedTemplateCheckMacro", ex);
            }
        });
    }

    public void AutoClose()
    {
        AppLogger.Info("EventsAutoOpen.AutoClose triggered.");
        Environment.SetEnvironmentVariable(HookedPidKey, null);
        Environment.SetEnvironmentVariable(LastCheckedWorkbookKey, null);
        Arrow.HideAll();
        UnEventHooks();
    }

    [ExcelCommand(Name = "DelayedTemplateCheckMacro")]
    public static void DelayedTemplateCheckMacro()
    {
        try
        {
            RuleRegistry.EnsureInitialized();

            dynamic xlApp = ExcelDnaUtil.Application;
            dynamic? wb = null;
            try { wb = xlApp?.ActiveWorkbook; } catch { }
            if (wb == null) return;

            string wbName = wb.Name?.ToString() ?? "Workbook";
            string? alreadyCheckedWb = Environment.GetEnvironmentVariable(LastCheckedWorkbookKey);

            // Cross-ALC guard: If DelayedTemplateCheckMacro already ran for this workbook, skip!
            if (string.Equals(alreadyCheckedWb, wbName, StringComparison.OrdinalIgnoreCase))
            {
                AppLogger.Info($"[Coexistence] DelayedTemplateCheckMacro already completed for '{wbName}', skipping duplicate run.");
                return;
            }
            Environment.SetEnvironmentVariable(LastCheckedWorkbookKey, wbName);

            using var _ = AppLogger.Scope($"DelayedTemplateCheckMacro({wbName})");

            bool isTemplate = SheetOperations.IsMasterTemplateFile();
            SheetOperations.IsTemplateEditMode = isTemplate;
            Environment.SetEnvironmentVariable(LastTemplateModeKey, isTemplate ? "1" : "0");

            AppLogger.Info($"Workbook='{wbName}', IsMasterTemplateFile={isTemplate}");

            // Clamp any runaway 1000-row tables before styling
            SheetMapperHelper.ClampAllTablesToActualData(wb);
            RuleRegistry.ApplyDropdownStylesToWorkbook(wb);

            if (isTemplate)
            {
                SheetOperations.SetAuxiliarySheetsVisibility(visible: true);
                EnumList.SyncAllCoreEnums();
                DataList.SyncAllTemplateData();
                SKSB_App.Core.RuntimeUI.TemplateEditor.RibbonController.RefreshRibbonUI();
            }
            else
            {
                SheetOperations.SetAuxiliarySheetsVisibility(visible: false);
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("DelayedTemplateCheckMacro failed", ex);
        }
    }

    [ExcelCommand(Name = "OnWindowActivatedMacro")]
    public static void OnWindowActivatedMacro()
    {
        try
        {
            // Do not run until initial DelayedTemplateCheckMacro has completed
            if (string.IsNullOrEmpty(Environment.GetEnvironmentVariable(LastCheckedWorkbookKey)))
                return;

            bool isTemplate = SheetOperations.IsMasterTemplateFile();
            bool prevMode = Environment.GetEnvironmentVariable(LastTemplateModeKey) == "1";

            SheetOperations.IsTemplateEditMode = isTemplate;

            if (prevMode != isTemplate)
            {
                Environment.SetEnvironmentVariable(LastTemplateModeKey, isTemplate ? "1" : "0");
                AppLogger.Info($"WindowActivate mode switch: IsTemplateEditMode -> {isTemplate}");
                SheetOperations.SetAuxiliarySheetsVisibility(visible: isTemplate);
                SKSB_App.Core.RuntimeUI.TemplateEditor.RibbonController.RefreshRibbonUI();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("OnWindowActivatedMacro error", ex);
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

            AppLogger.Info("Application Sheet Events hooked successfully.");
        }
        catch (Exception ex)
        {
            AppLogger.Error("EventHooks failed", ex);
        }
    }

    private static void UnEventHooks()
    {
        try
        {
            if (_xlAppRoot != null)
            {
                if (_doubleClickHandler != null) { _xlAppRoot.SheetBeforeDoubleClick -= _doubleClickHandler; _doubleClickHandler = null; }
                if (_selectionHandler != null) { _xlAppRoot.SheetSelectionChange -= _selectionHandler; _selectionHandler = null; }
                if (_changeHandler != null) { _xlAppRoot.SheetChange -= _changeHandler; _changeHandler = null; }
                _xlAppRoot = null;
            }
        }
        catch { }
    }

    private static void OnSheetSelectionChange(dynamic sh, dynamic target)
    {
        if (target == null)
        {
            Arrow.HideAll();
            return;
        }

        dynamic singleCell = target.Cells[1, 1];
        if (DropDownCommands.IsDropdownCell(singleCell))
        {
            Arrow.ShowArrowForCell(singleCell);
            return;
        }

        Arrow.HideAll();
    }

    private static void OnSheetChange(dynamic sh, dynamic target)
    {
        try
        {
            if (target == null || target.Cells.Count > 1) return;

            dynamic singleCell = target.Cells[1, 1];
            string sheetName = sh?.Name?.ToString() ?? "Sheet";
            string addr = singleCell.Address?.ToString() ?? "";
            AppLogger.Info($"OnSheetChange: {sheetName}!{addr}");

            dynamic? table = null;
            try { table = singleCell.ListObject; } catch { }

            if (table != null)
            {
                RuleEvaluator.EvaluateAndHealRow(
                    worksheet: sh,
                    table: table,
                    excelRow: (int)singleCell.Row,
                    editedAbsoluteCol: (int)singleCell.Column
                );
            }
            else
            {
                ExecuteLegacyCellSyncOnType(singleCell);
            }

            if (DropDownCommands.IsDropdownCell(singleCell))
            {
                Arrow.ShowArrowForCell(singleCell);
            }

            TableRowManager.HandleSheetChange(singleCell);
        }
        catch (Exception ex)
        {
            AppLogger.Error("OnSheetChange error", ex);
        }
    }

    private static void ExecuteLegacyCellSyncOnType(dynamic singleCell)
    {
        var config = Metadata.ReadConfig(singleCell);
        if (string.IsNullOrWhiteSpace(config.SyncLeft)) return;

        string cellVal = singleCell.Value2?.ToString()?.Trim() ?? string.Empty;
        if (string.IsNullOrEmpty(cellVal)) return;

        SyncEngine.ExecuteSync(singleCell, config.SyncLeft, cellVal, cellVal);
    }

    private static void OnSheetBeforeDoubleClick(dynamic sh, dynamic target, ref bool cancel)
    {
        try
        {
            if (!DropDownCommands.IsDropdownCell(target))
                return;

            bool isTemplate = SheetOperations.IsTemplateEditMode ||
                              Environment.GetEnvironmentVariable(LastTemplateModeKey) == "1";

            if (isTemplate)
            {
                cancel = true;
                Arrow.HideAll();
                DropDownCommands.InspectAndConfigure(target);
            }
            else
            {
                Arrow.HideAll();
            }
        }
        catch (Exception ex)
        {
            AppLogger.Error("OnSheetBeforeDoubleClick error", ex);
        }
    }
}