namespace SKSB_App.VehicleManager;

using ExcelDna.Integration;
using SKSB_App.Core.Integration.Database;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;
using SKSB_App.Core.WorkbookSave.Storage;
using System;
using System.Collections.Generic;

public class Events : IExcelAddIn
{
    private static dynamic? _xlAppRoot;
    private static bool _isExiting = false;
    private static bool _isInternalSaving = false;

    private static readonly EventsAutoOpen _coreEvents = new();

    private delegate void WorkbookBeforeCloseDelegate(dynamic wb, ref bool cancel);
    private delegate void WorkbookBeforeSaveDelegate(dynamic wb, bool saveAsUi, ref bool cancel);
    private static WorkbookBeforeCloseDelegate? _beforeCloseHandler;
    private static WorkbookBeforeSaveDelegate? _beforeSaveHandler;

    public static readonly SaveContract<List<VehicleRecord>> VehicleSyncDescriptor = new()
    {
        ModuleName = "Vehicle Manager",
        ServerTableName = "vehicle_master",
        ExtractAndValidate = () => SheetMapper.ExtractVehicleRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        SaveToServer = list => (ServerSyncCommands.SaveVehiclesSilent(list, out string err), err),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        FetchFromServer = () => ServerSyncCommands.FetchVehicleRecords(),
        SyncLookups = isInteractive => SyncVehicleLookupsData(isInteractive),
        ValidateOfflinePayload = list => list.Count > 0 && list.TrueForAll(v => !string.IsNullOrWhiteSpace(v.VehiclePlate))
    };

    public void AutoOpen()
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            _coreEvents.AutoOpen();

            HookApplicationEvents();
            Arrow.EnsureExcelEventsHooked();

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
            _coreEvents.AutoClose();
            Arrow.Teardown();
            UnhookApplicationEvents();
            XlCall.Excel(XlCall.xlcOnKey, "^s");
            XlCall.Excel(XlCall.xlcOnKey, "^+s");
        });
    }

    public static void SyncVehicleLookupsData(bool isInteractive = false)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                // Sync ONLY the specific enums used across Vehicle Manager sheets
                EnumList.Sync(
                    typeof(DateFilterScope),
                    typeof(VehicleType),
                    typeof(ServiceType),
                    typeof(MaintenanceStatus),
                    typeof(VehicleStatus),
                    typeof(VehiclePermitType),
                    typeof(VehicleType),
                    typeof(InsuranceType)
                );

                // 1. Sync Staff IDs (ordered numerically)
                DataList.SyncProperty<StaffRecord, long>(
                    s => s.StaffID,
                    customListName: "StaffIDList"
                );

                // 2. Push StaffShortName into StaffNameList (with fallback to StaffName if blank)
                DataList.SyncProperty<StaffRecord, string>(
                    s => !string.IsNullOrWhiteSpace(s.StaffDetails.StaffShortName)
                        ? s.StaffDetails.StaffShortName
                        : s.StaffDetails.StaffName,
                    customListName: "StaffNameList"
                );

                // Optional: Also sync dedicated StaffShortNameList if needed elsewhere
                DataList.SyncProperty<StaffRecord, string>(
                    s => s.StaffDetails.StaffShortName,
                    customListName: "StaffShortNameList"
                );
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SyncVehicleLookupsData Warning] {ex.Message}");
            }
        });
    }

    private static void HookApplicationEvents()
    {
        try
        {
            if (_xlAppRoot != null) return;
            dynamic xlApp = ExcelDnaUtil.Application;
            _xlAppRoot = xlApp;

            dynamic ourWorkbook = xlApp.ActiveWorkbook;

            _beforeCloseHandler = new WorkbookBeforeCloseDelegate((dynamic wb, ref bool cancel) =>
            {
                if (!Guard.IsVehicleWorkbook(wb))
                {
                    wb.Saved = true;
                    cancel = false;
                    return;
                }

                SaveEntry.InterceptClose(
                    (object)wb,
                    ref cancel,
                    VehicleSyncDescriptor,
                    ref _isExiting,
                    new Action(UnhookApplicationEvents)
                );
            });
            _xlAppRoot.WorkbookBeforeClose += _beforeCloseHandler;

            _beforeSaveHandler = new WorkbookBeforeSaveDelegate((dynamic wb, bool saveAsUi, ref bool cancel) =>
            {
                try
                {
                    if (ourWorkbook != null && wb != null && wb != ourWorkbook) return;
                }
                catch { }

                if (_isInternalSaving || SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode) return;

                cancel = true;
                ExcelAsyncUtil.QueueAsMacro(() => SaveEntry.Save(VehicleSyncDescriptor));
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
        if (SheetOperations.IsMasterTemplateFile() || SheetOperations.IsTemplateEditMode)
        {
            _isInternalSaving = true;
            try { XlCall.Excel(XlCall.xlcSave); } finally { _isInternalSaving = false; }
            return;
        }

        SaveEntry.Save(VehicleSyncDescriptor);
    }

    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro()
    {
        if (!Guard.IsVehicleWorkbook()) return;

        dynamic xlApp = ExcelDnaUtil.Application;
        try
        {
            if (xlApp.ActiveWorkbook != null && !SheetOperations.IsMasterTemplateFile())
            {
                xlApp.ActiveWindow.Caption = "Vehicle_Manager";
            }
        }
        catch { }

        OfflineSave.HandleStartup(VehicleSyncDescriptor);

        dynamic wb = xlApp.ActiveWorkbook;
        if (wb != null && !SheetOperations.IsTemplateEditMode)
        {
            foreach (dynamic sheet in wb.Worksheets)
            {
                SheetProtections.ApplyStyleDrivenProtection(sheet);
            }

            SheetOperations.SetAuxiliarySheetsVisibility(visible: false);
        }
    }
}