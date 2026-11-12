namespace SKSB_App.StaffManager;

using System;
using System.Collections.Generic;
using System.Linq;
using ExcelDna.Integration;
using SKSB_App.Core.RuntimeUI.DateFilter;
using SKSB_App.Core.RuntimeUI.DropDown;
using SKSB_App.Core.Types;
using SKSB_App.Core.WorkbookOpen.Entry;
using SKSB_App.Core.WorkbookOpen.Sheet;
using SKSB_App.Core.WorkbookOpen.Sync;
using SKSB_App.Core.WorkbookSave.Entry;

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

    public static readonly SaveContract<List<StaffRecord>> StaffSyncDescriptor = new()
    {
        ModuleName = "Staff Manager",
        ServerTableName = "staff_master",
        ExtractAndValidate = () => SheetMapper.ExtractStaffRecordsWithDiagnostics(),
        IsDatasetEmpty = list => list == null || list.Count == 0,
        SaveToServer = list => (ServerSyncCommands.SaveStaffSilent(list, out string err), err),
        PopulateWorksheet = list => SheetMapper.PopulateWorksheet(list),
        FetchFromServer = () => ServerSyncCommands.FetchStaffRecords(),
        SyncLookups = isInteractive => SyncStaffLookupsData(isInteractive),
        ValidateOfflinePayload = list => list != null && list.Count > 0 && list.TrueForAll(s => s.StaffID > 0),
        FormatConflictSummary = (local, server) =>
        {
            int localVLic = local.Sum(s => s.VehicleLicenses?.Count ?? 0);
            int localSLic = local.Sum(s => s.SkillLicenses?.Count ?? 0);
            int serverVLic = server.Sum(s => s.VehicleLicenses?.Count ?? 0);
            int serverSLic = server.Sum(s => s.SkillLicenses?.Count ?? 0);

            return $"  Local Offline : {local.Count} staff(s), {localVLic} vehicle license(s), {localSLic} skill license(s)\n" +
                   $"  Server Copy   : {server.Count} staff(s), {serverVLic} vehicle license(s), {serverSLic} skill license(s)";
        }
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

    public static void SyncStaffLookupsData(bool isInteractive = false)
    {
        ExcelAsyncUtil.QueueAsMacro(() =>
        {
            try
            {
                // 1. Sync Staff Manager Enums to 'Lookups' sheet
                EnumList.Sync(
                    typeof(Department),
                    typeof(ContractStatus),
                    typeof(VehicleLicenseClass),
                    typeof(VehicleLicenseStatus),
                    typeof(SkillLicenseStatus),
                    typeof(DateFilterScope)
                );

                // 2. Sync external Vehicle Plates to 'Data' sheet
                DataList.SyncProperty<VehicleRecord, string>(
                    v => v.VehiclePlate,
                    customListName: "VehiclePlateList"
                );

                // 3. Populate and sync SkillLicenseTypesTable & SkillLicenseTypeList
                SheetMapper.PopulateSkillLicenseTypesTable(showDialogIfFailed: isInteractive);
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"[SyncStaffLookupsData Warning] {ex.Message}");
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
                try
                {
                    if (ourWorkbook != null && wb != null && wb != ourWorkbook)
                    {
                        wb.Saved = true;
                        cancel = false;
                        return;
                    }
                }
                catch { }

                SaveEntry.InterceptClose(
                    (object)wb,
                    ref cancel,
                    StaffSyncDescriptor,
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
                ExcelAsyncUtil.QueueAsMacro(() => SaveEntry.Save(StaffSyncDescriptor));
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

        SaveEntry.Save(StaffSyncDescriptor);
    }

    [ExcelCommand(Name = "DelayedAutoLoadMacro")]
    public static void DelayedAutoLoadMacro()
    {
        dynamic xlApp = ExcelDnaUtil.Application;
        try
        {
            if (xlApp.ActiveWorkbook != null && !SheetOperations.IsMasterTemplateFile())
            {
                xlApp.ActiveWindow.Caption = "Staff_Manager";
            }
        }
        catch { }

        OfflineSave.HandleStartup(StaffSyncDescriptor);

        dynamic wb = xlApp.ActiveWorkbook;
        if (wb != null && !SheetOperations.IsTemplateEditMode)
        {
            SheetOperations.SetAuxiliarySheetsVisibility(visible: false);
        }
    }
}