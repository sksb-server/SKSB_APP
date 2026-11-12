namespace SKSB_App.Core.WorkbookOpen.Entry;

using System;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.Diagnostics;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;

// COM event signatures. Declared at namespace level (not nested in the generic base) so the
// delegate types themselves stay non-generic for late-bound COM event subscription.
public delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
public delegate void WorkbookBeforeSaveDelegate(dynamic wb, bool saveAsUi, ref bool cancel);

/// <summary>
/// Shared add-in lifecycle for every Manager module: Core event chaining, pinned COM event hooks,
/// Ctrl+S interception, delayed startup load and lookup synchronisation.
/// A module supplies its <see cref="SaveContract{T}"/> and (optionally) <see cref="SyncCustomLookups"/>,
/// plus two thin [ExcelCommand] entry points that delegate to <see cref="ExecuteHotKeySave"/> and
/// <see cref="ExecuteDelayedAutoLoad"/> (Excel-DNA only registers commands declared on non-generic types).
/// </summary>
public abstract class ManagerEventsBase<T> : IExcelAddIn
{
    // Static state is per closed generic type, i.e. one set per Manager module.
    private static dynamic? _xlAppRoot;
    private static bool _isExiting = false;
    private static bool _isInternalSaving = false;
    private static ManagerEventsBase<T>? _instance;

    private static readonly EventsAutoOpen _coreEvents = new();

    // Pinned so the GC can never collect a delegate Excel still holds.
    private static WorkbookBeforeCloseDelegate? _beforeCloseHandler;
    private static WorkbookBeforeSaveDelegate? _beforeSaveHandler;

    // =========================================================================
    // 1. MODULE CUSTOMIZATION HOOKS
    // =========================================================================
    protected abstract SaveContract<T> Contract { get; }

    /// <summary>Registered name of the module's Ctrl+S [ExcelCommand].</summary>
    protected virtual string HotKeySaveMacroName => "OnHotKeySave";

    /// <summary>Registered name of the module's delayed startup [ExcelCommand].</summary>
    protected virtual string DelayedAutoLoadMacroName => "DelayedAutoLoadMacro";

    /// <summary>Module-specific DataList / auxiliary table syncs. Runs on the Excel macro thread.</summary>
    protected virtual void SyncCustomLookups(bool isInteractive) { }

    protected ManagerEventsBase()
    {
        _instance = this;
        Contract.SyncLookups ??= SyncLookups;
    }

    // =========================================================================
    // 2. ADD-IN LIFECYCLE
    // =========================================================================
    public virtual void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _coreEvents.AutoOpen();

            HookApplicationEvents();
            Arrow.EnsureExcelEventsHooked();

            double nowSerial = Convert.ToDouble(XlCall.Excel(XlCall.xlfNow));
            XlCall.Excel(XlCall.xlcOnTime, nowSerial + (1.0 / 86400.0), DelayedAutoLoadMacroName);

            XlCall.Excel(XlCall.xlcOnKey, "^s", HotKeySaveMacroName);
            XlCall.Excel(XlCall.xlcOnKey, "^+s", HotKeySaveMacroName);
        });
    }

    public virtual void AutoClose()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _coreEvents.AutoClose();
            Arrow.Teardown();
            UnhookApplicationEvents();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    // =========================================================================
    // 3. LOOKUP SYNCHRONISATION
    // =========================================================================
    public virtual void SyncLookups(bool isInteractive)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                Type[] enums = Contract.EnumsToSync?.Distinct().ToArray() ?? Array.Empty<Type>();
                if (enums.Length > 0) EnumList.Sync(enums);

                SyncCustomLookups(isInteractive);
            }
            catch (Exception ex)
            {
                AppLogger.Warn($"[{Contract.ModuleName}] SyncLookups warning: {ex.Message}");
            }
        });
    }

    // =========================================================================
    // 4. APPLICATION EVENT HOOKS
    // =========================================================================
    protected static void HookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;
            _xlAppRoot = ExcelDnaUtil.Application;

            _beforeCloseHandler = new WorkbookBeforeCloseDelegate(OnWorkbookBeforeClose);
            _xlAppRoot.WorkbookBeforeClose += _beforeCloseHandler;

            _beforeSaveHandler = new WorkbookBeforeSaveDelegate(OnWorkbookBeforeSave);
            _xlAppRoot.WorkbookBeforeSave += _beforeSaveHandler;
        }
        catch (Exception ex)
        {
            AppLogger.Error("ManagerEventsBase.HookApplicationEvents failed", ex);
        }
    }

    protected static void UnhookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null)
            {
                if (_beforeCloseHandler != null) _xlAppRoot.WorkbookBeforeClose -= _beforeCloseHandler;
                if (_beforeSaveHandler != null) _xlAppRoot.WorkbookBeforeSave -= _beforeSaveHandler;
                _beforeCloseHandler = null;
                _beforeSaveHandler = null;
                _xlAppRoot = null;
            }
        }
        catch { }
    }

    private static void OnWorkbookBeforeClose(dynamic wb, ref bool cancel)
    {
        var contract = _instance?.Contract;
        if (contract == null) return;

        // Workbooks that do not belong to this module are left entirely to Excel.
        if (!SheetOperations.IsModuleWorkbook((object?)wb, contract)) return;

        SaveEntry.InterceptClose(
            (object)wb,
            ref cancel,
            contract,
            ref _isExiting,
            new Action(UnhookApplicationEvents)
        );
    }

    private static void OnWorkbookBeforeSave(dynamic wb, bool saveAsUi, ref bool cancel)
    {
        var contract = _instance?.Contract;
        if (contract == null) return;

        if (!SheetOperations.IsModuleWorkbook((object?)wb, contract)) return;
        if (_isInternalSaving || SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;

        // Cancel the native save and re-run it as a queued macro once this COM event has returned.
        cancel = true;
        ExcelAsyncUtil.QueueAsMacro(() => SaveEntry.Save(contract));
    }

    // =========================================================================
    // 5. COMMAND BODIES (called by each module's [ExcelCommand] entry points)
    // =========================================================================
    protected static void ExecuteHotKeySave()
    {
        var contract = _instance?.Contract;

        // Templates, design mode and foreign workbooks get Excel's native save.
        if (contract == null ||
            SheetOperations.IsMasterTemplateFile() ||
            SheetOperations.IsTemplateEditMode ||
            !SheetOperations.IsModuleWorkbook(null, contract))
        {
            _isInternalSaving = true;
            try { XlCall.Excel(XlCall.xlcSave); } finally { _isInternalSaving = false; }
            return;
        }

        SaveEntry.Save(contract);
    }

    protected static void ExecuteDelayedAutoLoad()
    {
        var contract = _instance?.Contract;
        if (contract == null) return;
        if (!SheetOperations.IsModuleWorkbook(null, contract)) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        try
        {
            if (xlApp.ActiveWorkbook != null &&
                !SheetOperations.IsMasterTemplateFile() &&
                !string.IsNullOrWhiteSpace(contract.WindowCaption))
            {
                xlApp.ActiveWindow.Caption = contract.WindowCaption;
            }
        }
        catch { }

        OfflineSave.HandleStartup(contract);

        object? wbObj = xlApp.ActiveWorkbook;
        if (wbObj == null || SheetOperations.IsTemplateEditMode) return;

        SheetOperations.SetAuxiliarySheetsVisibility(visible: false);

        try
        {
            contract.OnAfterDelayedLoad?.Invoke(wbObj);
        }
        catch (Exception ex)
        {
            AppLogger.Error($"[{contract.ModuleName}] OnAfterDelayedLoad failed", ex);
        }
    }
}
