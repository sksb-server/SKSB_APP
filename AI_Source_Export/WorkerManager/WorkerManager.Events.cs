namespace SKSB_App.WorkerManager;

using System;
using System.Collections.Generic;
using ExcelDna.Integration;
using SKSB_App.Core.ServerManager;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookHelper;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;
    private static bool _isExiting = false;
    private static bool _isInternalSaving = false;

    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private delegate void WorkbookBeforeSaveDelegate(dynamic wb, bool saveAsUi, ref bool cancel);

    private static WorkbookBeforeCloseDelegate? _beforeCloseHandler;
    private static WorkbookBeforeSaveDelegate? _beforeSaveHandler;

    public static readonly HookProperties<List<WorkerRecord>> WorkerSyncDescriptor = new()
    {
        ModuleName = "Worker Manager",
        ServerTableName = "worker_master",
        ExtractAndValidate = () => SheetMapper.ExtractWorkerRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        SaveToServer = list => (ServerSyncCommands.SaveWorkerSilent(out string err), err),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        FetchFromServer = () => ServerSyncCommands.FetchWorkerRecords(),
        SyncLookups = isInteractive => SyncWorkerLookupsData(isInteractive),
        ValidateOfflinePayload = list => list.Count > 0 && list.TrueForAll(w => w.WorkerID > 0)
    };

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            HookApplicationEvents();
            SKSB_App.Core.DropDown.Arrow.EnsureExcelEventsHooked();

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
            SKSB_App.Core.DropDown.Arrow.Teardown();
            UnhookApplicationEvents();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    public static void SyncWorkerLookupsData(bool isInteractive = false)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try { DataToExcel.SyncData<VehicleRecord, string>("VehiclePlate", v => v.VehiclePlate); } catch { }
            SheetMapper.PopulateSkillLicenseTypesTable(showDialogIfFailed: isInteractive);
        });
    }

    private static void HookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;
            dynamic xlApp = ExcelDnaUtil.Application;
            _xlAppRoot = xlApp;

            _beforeCloseHandler = new WorkbookBeforeCloseDelegate((dynamic wb, ref bool cancel) =>
            {
                EventHooks.InterceptClose(
                    (object)wb,
                    ref cancel,
                    WorkerSyncDescriptor,
                    ref _isExiting,
                    new Action(UnhookApplicationEvents)
                );
            });
            _xlAppRoot.WorkbookBeforeClose += _beforeCloseHandler;

            _beforeSaveHandler = new WorkbookBeforeSaveDelegate((dynamic wb, bool saveAsUi, ref bool cancel) =>
            {
                if (_isInternalSaving || SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode) return;
                cancel = true;
                ExcelAsyncUtil.QueueAsMacro(() => EventHooks.Save(WorkerSyncDescriptor));
            });
            _xlAppRoot.WorkbookBeforeSave += _beforeSaveHandler;
        }
        catch { }
    }

    private static void UnhookApplicationEvents()
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

    [ExcelCommand(Name = "OnHotKeySave")]
    public static void OnHotKeySave()
    {
        if (SheetOps.IsMasterTemplateFile() || SheetOps.IsTemplateEditMode)
        {
            _isInternalSaving = true;
            try { XlCall.Excel(XlCall.xlcSave); } finally { _isInternalSaving = false; }
            return;
        }
        EventHooks.Save(WorkerSyncDescriptor);
    }

    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro()
    {
        SyncWorkerLookupsData(isInteractive: false);
        SheetMapper.InitializeSheetRules();
        EventHooks.HandleStartup(WorkerSyncDescriptor);
    }
}