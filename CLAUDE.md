CLAUDE.md — SKSB_App Architecture & Developer Guide

This document is the authoritative architectural reference and operational guide for AI coding assistants and developers working in the SKSB_App repository. All changes must strictly adhere to the tech stack, build rules, database conventions, and Excel-DNA thread-safety protocols documented below.

1. Tech Stack, Target Framework & Libraries
Runtime & Framework
Target Framework: net8.0-windows across all projects.
Language Features: C# 12 with <ImplicitUsings>enable</ImplicitUsings> and <Nullable>enable</Nullable>.
UI Framework: Windows Forms (<UseWindowsForms>true</UseWindowsForms>) with PerMonitorV2 High-DPI mode on standalone executables (Launcher, Portal) and modal dialogs in Core.
Platform Target: AnyCPU (with win-x64 self-contained single-file publishing for Launcher).
Solution Format: Visual Studio XML Solution (SKSB_App.slnx) supporting Any CPU, x64, and x86 platform configurations.
External NuGet Packages & Libraries
Package / Library	Exact Version	Project(s)	Purpose
ExcelDna.AddIn	1.9.0	Core, StaffManager, VehicleManager, ProjectManager, ScheduleManager	Generates 32-bit (*-AddIn.xll) and 64-bit (*-AddIn64.xll) Excel XLL add-ins (both unpacked and -packed.xll with RuntimeVersion="v8.0").
ExcelDna.Integration	(Transitive via ExcelDna.AddIn)	Core & Add-In Modules	Provides IExcelAddIn, ExcelRibbon, ExcelDnaUtil.Application, ExcelAsyncUtil.QueueAsMacro, XlCall, and [ExcelCommand].
Dapper	2.1.35	Core	Micro-ORM for parameterized SQL queries, bulk IEnumerable<T> execution, and snake_case to PascalCase mapping (Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true).
Npgsql	8.0.3	Core	ADO.NET data provider for PostgreSQL (PostgreSqlService).
PdfPig	0.1.16	VehicleManager	PDF text extraction for parsing external fuel statement documents (VehicleManager.FuelImporter.cs).
System.Text.Json	.NET 8 Built-in	Core	Local offline cache (*_data.json) and pending offline upload stashing (*_PendingUpload.json) via OfflineSnapshot and TolerantEnumConverterFactory.
Solution & Project Topology
Core/ (Core.csproj): Shared foundation library and base XLL add-in. Organized into 4 lifecycle phases plus shared subsystems:
01_Bootstrap/: Binary inspection, Excel path discovery, registry configuration, and local folder paths (%USERPROFILE%\Documents\SKSB_App\Data).
02_WorkbookOpen/: EventsAutoOpen (IExcelAddIn), SheetOperations, SheetProtections, EnumList, DataList, and OfflineSave.
03_RuntimeUI/: Interactive cell dropdowns (Arrow, Commands), live self-healing RuleEngine (RuleRegistry, RuleEvaluator, RuleDataCache), TableRowManager, DateFilter, and TemplateEditor.
04_WorkbookSave/: SaveEntry, SaveContract<T>, TableMapper, SheetMapperHelper, OfflineSnapshot, RulesExtractor, and conflict/validation WinForms dialogs.
Diagnostics/: Scoped session logging (AppLogger) and debug inspection dialogs (DebugDialog).
Integration/: Database/ (PostgreSqlService, ServerDispatcher, ServerEnvironment) and Erp/ (SQL Payroll / ERP importers & exporters).
Types/: Domain models (StaffRecord, VehicleRecord, ProjectRecord, ScheduleRecord) and declarative rule attributes (Attributes/).
Note: Core/ThumbprintSync/** is explicitly excluded from compilation in Core.csproj.
StaffManager/ (StaffManager.csproj): Excel-DNA Add-In (StaffManager-AddIn64.xll) for employee profiles, vehicle licenses, skill licenses, and SQL Payroll imports.
VehicleManager/ (VehicleManager.csproj): Excel-DNA Add-In (VehicleManager-AddIn64.xll) for fleet master records, maintenance, road tax, Puspakom inspections, permits, insurance, odometer logs, and PDF/Excel fuel card imports.
ProjectManager/ (ProjectManager.csproj) & ScheduleManager/ (ScheduleManager.csproj): Domain add-in modules for project master data and daily scheduling (present in repository; add to SKSB_App.slnx if building via solution).
Portal/ (Portal.csproj): Standalone WinForms desktop hub (Portal.exe) that launches domain Excel templates with their corresponding .xll add-ins.
Launcher/ (Launcher.csproj): Self-contained single-file WinForms bootstrapper (SKSB_App.exe) that checks version.json, downloads updates from the production server, and launches Portal.exe.
Templates/: Master Excel templates (*.xltx).
2. Build, Clean, Publish & Test Commands
MSBuild Automation Hooks (Directory.Build.props & Portal.csproj)

Before running any build or clean command, understand what MSBuild executes automatically:

Process Termination (Portal.csproj -> KillExistingExcelAndPortalOnBuild): Building Portal (or the full solution) runs taskkill /f /fi "IMAGENAME eq excel.exe" and taskkill /f /fi "IMAGENAME eq Portal.exe" before compilation so locked .xll and .dll files can be overwritten.
Add-In Artifact Sync (SyncAddInsToCodesFolderOnBuild): Every Build of Core or any ExcelDna.AddIn project copies .dll, .pdb, .deps.json, .runtimeconfig.json, *.dna, *-AddIn*.xll, and publish/*-packed.xll into ./AddIns/.
AI Code Bundle Generation (GenerateAIChatCodeBundle): Building Launcher or SKSB_App exports all clean source files into ./AI_Source_Export/ and atomically consolidates them into ./AI_Source_Export/AI_Context_AllCode.txt.
Deep Clean (DeepCleanAllArtifactsOnClean): Running a Clean or Rebuild target wipes ./AddIns/ and ./AI_Source_Export/. Warning: Running Clean with -c Release also deletes all non-Data files in %USERPROFILE%\Documents\SKSB_App.
Standard CLI Commands (PowerShell)
Restore Dependencies
powershell
dotnet restore SKSB_App.slnx
Build Solution (Debug — Enables Debug Ribbon Group & Audit Logs)
powershell
dotnet build SKSB_App.slnx -c Debug
Build Solution (Release)
powershell
dotnet build SKSB_App.slnx -c Release
Build Standalone Modules (Outside SKSB_App.slnx)
powershell
dotnet build ProjectManager/ProjectManager.csproj -c Debug
dotnet build ScheduleManager/ScheduleManager.csproj -c Debug
Clean Solution
powershell
# Safe Debug Clean (wipes bin/, obj/, ./AddIns/, and ./AI_Source_Export/)
dotnet clean SKSB_App.slnx -c Debug
# Release Clean (ALSO wipes deployed binaries in $env:USERPROFILE\Documents\SKSB_App except \Data\)
dotnet clean SKSB_App.slnx -c Release
Rebuild Solution
powershell
dotnet build SKSB_App.slnx -c Debug --no-incremental
Package & Publish Release to Server

Publishing Launcher in Release mode triggers AutoDeployReleaseToServer, which invokes publish.ps1. You can also run publish.ps1 directly after building in Release mode:

powershell
# Option A: Via dotnet publish (builds Launcher single-file exe + runs publish.ps1)
dotnet publish Launcher/Launcher.csproj -c Release
# Option B: Direct PowerShell packaging & upload (auto-increments patch version in version.json)
powershell -NoProfile -ExecutionPolicy Bypass -File .\publish.ps1
# Option C: Explicit version override & custom server host
powershell -NoProfile -ExecutionPolicy Bypass -File .\publish.ps1 -ReleaseVersion "1.2.0" -ServerHost "http://192.168.0.20:5000"
Testing & Verification Workflow

This solution does not use an external xUnit/NUnit test runner project; verification is performed in-process within Excel via Debug Build Instrumentation and Startup Diagnostics:

Build in Debug mode (dotnet build SKSB_App.slnx -c Debug).
Every Excel-DNA module exposes a Debug Controls Ribbon group in DEBUG builds containing:
Simulate Server Offline (ServerEnvironment.SimulateOffline): Forces ServerEnvironment.GuardConnection() to throw a TimeoutException on any PostgreSqlService call, testing local JSON fallback (OfflineSnapshot) and ExitConflictDialog workflows.
Debug Logs Active (DebugDialog.DebugEnabled): Displays detailed modal audit summaries (DebugDialog.ShowLog) after worksheet population and PostgreSQL upload/download batches.
Inspect runtime execution traces and COM event stack traces in %USERPROFILE%\Documents\SKSB_App\Data\ or via AppLogger (Core/Diagnostics/Diagnoestics.AppLogger.cs).
3. Architectural Code Style & Conventions
Namespaces, File Naming & Structure
File-Scoped Namespaces: Always use file-scoped namespaces at line 1 (e.g., namespace SKSB_App.StaffManager;, namespace SKSB_App.Core.WorkbookSave.SheetMapper;), followed by using directives placed inside the namespace declaration.
Dot-Prefixed File Naming: Source files follow a <DomainOrFolder>.<Component>.cs naming convention:
Domain modules: StaffManager.Events.cs, StaffManager.RibbonController.cs, StaffManager.ServerSyncCommands.cs, StaffManager.SheetMapper.cs, StaffManager.Importer.cs, StaffManager.Guard.cs.
Core subsystems: Entry.EventsAutoOpen.cs, Sheet.Protections.cs, SheetMapper.TableMapper.cs, RuleEngine.Evaluator.cs, Database.ServerDispatcher.cs.
Visual Section Headers: Divide large static classes into logical sections using 3-line banner comments:
csharp
// =========================================================================
// 1. MASTER EXTRACTION & VALIDATION PIPELINE
// =========================================================================
Domain Module Pattern (SaveContract<T>)

Every Excel-DNA domain add-in (StaffManager, VehicleManager, ProjectManager, ScheduleManager) follows a strict 4-file contract architecture wired into Core:

Events.cs (IExcelAddIn): Defines a static SaveContract<List<TRecord>> descriptor (ModuleName, ServerTableName, ExtractAndValidate, SaveToServer, PopulateWorksheet, FetchFromServer, SyncLookups, FormatConflictSummary) and delegates workbook save/close interception to SaveEntry.Save and SaveEntry.InterceptClose.
RibbonController.cs ([ComVisible(true)] public class RibbonController : ExcelRibbon): Defines Ribbon XML and routes button actions (OnLoadClicked, OnSaveClicked, row management buttons) to ServerSyncCommands, SheetMapper, and TableRowManager.
SheetMapper.cs: Implements Extract*WithDiagnostics() and PopulateWorksheet() using TableMapper.ExtractFromTable<T> and TableMapper.PopulateTableSafe<T>.
ServerSyncCommands.cs: Coordinates PostgreSQL queries (PostgreSqlService), Dapper persistence (OfflineSnapshot.Save(list, writeToDatabase: true) -> ServerDispatcher.Dispatch), and offline JSON fallback (OfflineSnapshot.Load<T>()).
Declarative Rule Engine & Model Attributes (Core/Types/)

Domain models are C# record types decorated with custom attributes in SKSB_App.Core.Types.Attributes that drive both Excel UI behavior (dropdowns, auto-sync, auto-defaults, state guards) and save-time validation (RuleRegistry & RuleEvaluator):

[ExcelTable("TableName")]: Binds a C# record to an Excel ListObject table name.
[ExcelColumn("Header Name", isRequired: true)]: Maps a property to an Excel column header. Properties without [ExcelColumn] are ignored during storage validation.
[DataDetail]: Placed on nested composition records (e.g., StaffDetails, VehicleDetail, ProjectDetails) so TableMapper and RuleRegistry recursively flatten and bind their inner properties to the same master table.
[DataList("ListName")]: Automatically syncs distinct property values to a named range on the hidden Data worksheet.
[ColumnRule("SourceList", IsRequired = ..., IsWarningOnly = ..., Transform = TransformKind...)]: Governs custom cell dropdown sources (Lookups enum lists, Data lists, or "TableName[ColumnHeader]"), text transformations (UpperCase, Trim, ShortNameFirstLast), and validation severity.
[ColumnSync("Target Header", "LookupSource", "ResultSource")]: Configures live two-way lookup filling between partner columns (e.g., selecting Staff ID fills Staff Name, and selecting Staff Name fills Staff ID).
[ColumnDefault(DefaultKind.AutoIncrement | Today | Literal | DateOffsetDays | DateOffsetMonths)]: Automatically populates blank cells once a row receives user input.
[ColumnStateGuard(priority, StateConditionKind..., forcedValueWhenTrue, fallbackValueWhenFalse, referenceHeaders)]: Self-healing state machine that forces status columns (e.g., forcing "Expired" when an expiry date is past, or "Resigned" when Resign Date is non-empty).
UI-Only Helper Columns: Properties that exist in an Excel table for two-way lookup convenience (like StaffName on OdometerLog or VehicleType on Maintenance) must be marked with [JsonIgnore] so they are not duplicated in JSON snapshots while remaining bound via [ExcelColumn] and [ColumnSync].
4. Database Mapping & Persistence Rules (Dapper + PostgreSQL)

All database operations flow through PostgreSqlService (Core/Integration/Database/Database.PostgreSqlService.cs) and ServerDispatcher (Core/Integration/Database/Database.ServerDispatcher.cs).

Underscore Name Matching:
Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true; is initialized in EventsAutoOpen.AutoOpen(). PostgreSQL columns use snake_case (staff_id, vehicle_plate, road_tax_expiry), while C# properties and Dapper parameters use PascalCase (@StaffId, @VehiclePlate, @RoadTaxExpiry).
Type Conversion Boundaries (DateOnly & Enum vs. Npgsql/Dapper):
Dates: Domain models use DateOnly and DateOnly?. Dapper does not natively bind DateOnly in anonymous parameter objects out-of-the-box; always project DateOnly to DateTime when constructing Dapper parameters:
csharp
JoinDate = s.StaffDetails.JoinDate.ToDateTime(TimeOnly.MinValue),
ResignDate = s.StaffDetails.ResignDate.HasValue
    ? (DateTime?)s.StaffDetails.ResignDate.Value.ToDateTime(TimeOnly.MinValue)
    : null
When reading dynamic rows back from Dapper, convert DateTime columns back via DateOnly.FromDateTime((DateTime)row.join_date).
Enums: PostgreSQL stores statuses and types as text/varchar. Always project C# enums via .ToString() when saving, and parse them back with Enum.TryParse<TEnum>(..., ignoreCase: true, out var res) with a safe fallback when loading.
Primary Keys: Normalize string primary keys (such as VehiclePlate or ProjectID) via .Trim().ToUpperInvariant() before executing queries. Filter out uninitialized records (StaffID > 0, !string.IsNullOrWhiteSpace(VehiclePlate)) prior to persistence.
Transactional Full-Batch Reconciliation (ExecuteInTransaction):
Multi-table saves must run inside Db.ExecuteInTransaction((conn, trans) => { ... }).
When isFullBatch == true (saving the entire workbook state), delete orphaned records in child tables first, then the master table, using PostgreSQL != ALL(@ActiveIds) / != ALL(@ActivePlates) array syntax before executing INSERT ... ON CONFLICT (...) DO UPDATE:
sql
DELETE FROM staff_vehicle_licenses WHERE staff_id != ALL(@ActiveIds);
DELETE FROM staff_skill_licenses WHERE staff_id != ALL(@ActiveIds);
DELETE FROM staff_master WHERE staff_id != ALL(@ActiveIds);
Every ON CONFLICT DO UPDATE clause must update updated_at = CURRENT_TIMESTAMP.
Dual-Write Local Cache (OfflineSnapshot):
Saving via OfflineSnapshot.Save(records, writeToDatabase: true) first serializes the dataset to %USERPROFILE%\Documents\SKSB_App\Data\<typename>_data.json and then calls ServerDispatcher.Dispatch(data).
Always check ServerEnvironment.SimulateOffline (or let PostgreSqlService call ServerEnvironment.GuardConnection()) and fall back to OfflineSnapshot.Load<List<TRecord>>() if the PostgreSQL server (192.168.0.20:5432) is unreachable.
5. Excel-DNA Thread Safety & COM Interop Guardrails

Excel's COM Object Model (ExcelDnaUtil.Application) is strictly Single-Threaded Apartment (STA) and bound to Excel's main calculation/UI thread. Violating any of the rules below will corrupt the Excel process, trigger RPC_E_SERVERCALL_RETRYLATER / 0x800AC472 COM exceptions, cause infinite event loops, or freeze the Excel viewport.

Rule 1: Marshal All COM Mutations via ExcelAsyncUtil.QueueAsMacro
Never mutate worksheets, ListObject tables, cells, or named ranges directly inside IExcelAddIn.AutoOpen(), AutoClose(), Ribbon callbacks (IRibbonControl), WinForms event handlers, or background Task threads.
Always wrap COM-mutating work in ExcelAsyncUtil.QueueAsMacro(() => { ... }); so Excel-DNA executes it safely on the main Excel macro thread when Excel is idle.
When scheduling delayed startup routines or intercepting WorkbookBeforeSave / WorkbookBeforeClose (e.g., QuitExcelIfLastWorkbook), defer execution via ExcelAsyncUtil.QueueAsMacro or XlCall.Excel(XlCall.xlcOnTime, ...) so the active COM event returns cleanly to Excel first.
Rule 2: Freeze EnableEvents, ScreenUpdating, and Calculation in try / finally
Any method that writes to cells, resizes tables, or modifies styles must capture and suppress Excel's state before touching the sheet and restore it in a finally block (wrapping each COM property assignment in try { ... } catch { }):
csharp
dynamic xlApp = ExcelDnaUtil.Application;
bool origEvents = true;
bool origUpdating = true;
int origCalculation = -4105; // xlCalculationAutomatic
try { origEvents = (bool)xlApp.EnableEvents; xlApp.EnableEvents = false; } catch { }
try { origUpdating = (bool)xlApp.ScreenUpdating; xlApp.ScreenUpdating = false; } catch { }
try { origCalculation = (int)xlApp.Calculation; xlApp.Calculation = -4135; /* xlCalculationManual */ } catch { }
try
{
    // Perform table population or sheet mutations here
}
finally
{
    try { xlApp.Calculation = origCalculation; } catch { }
    try { xlApp.EnableEvents = origEvents; } catch { }
    try { xlApp.ScreenUpdating = origUpdating; } catch { }
}
Critical Modal Dialog Exception (Save.Entry.cs): If you must display a blocking WinForms modal dialog (MessageBox.Show, SummaryDialog, ExitConflictDialog) while ScreenUpdating is false, temporarily set xlApp.ScreenUpdating = true and xlApp.EnableEvents = true before .ShowDialog(), and hide floating UI widgets (Arrow.HideAll()) to prevent Z-order deadlocks and frozen Excel repaints.
Rule 3: Bulk 2D Array Reads & Writes (Never Loop Cell-by-Cell Over COM)
Reading or writing cell-by-cell across COM takes milliseconds per call and freezes Excel on large tables.
Reads: Always read entire table ranges in a single COM call using TableMapper.GetRange2DValues(dataBody, rowCount, colCount), which handles the 1-based object[,] matrix (and the single-cell scalar edge case when a table has 1 row and 1 column).
Writes: Use TableMapper.PopulateTableSafe<T>, which builds a 0-based object[effectiveRows, 1] column array in memory, preserves the archetype row's (headerRow + 1) Style and NumberFormat (skipping formula columns starting with "="), and assigns colRange.Value2 = colData in a single COM call per column.
Bulk Healing: Use RuleEvaluator.EvaluateAndHealTable(worksheet, table) (3 COM reads total, O(1) in-memory dictionary lookups, and targeted writes only for dirty cells) rather than calling EvaluateAndHealRow in a loop.
Rule 4: Multi-AddIn Process Coexistence & Static Delegate Pinning
Multiple SKSB .xll add-ins (Core, StaffManager, VehicleManager) can be loaded simultaneously in the same excel.exe process.
Use process-scoped environment variables ($"SKSB_EVENTS_HOOKED_{Environment.ProcessId}", $"SKSB_LAST_CHECKED_WB_{Environment.ProcessId}") to prevent duplicate AutoOpen COM event hooks and redundant template checks.
Pin COM Event Delegates: Always store COM event delegates (SheetSelectionChangeDelegate, SheetChangeDelegate, WorkbookBeforeCloseDelegate, WorkbookBeforeSaveDelegate) in private static fields on the class. Never subscribe anonymous lambdas directly to _xlAppRoot events unless pinned, or the .NET Garbage Collector will collect the delegate and crash the Excel process on the next sheet event. Always unhook them in AutoClose() / UnEventHooks().
Re-entrancy Guards: Always guard recursive handlers like RuleEvaluator (_isEvaluating), Events (_isInternalSaving, _isExiting), and TableRowManager against re-entrancy.
Rule 5: Protect the User's Viewport & Table Structure
Never Switch Active Tabs Unnecessarily: Operate directly on target Worksheet and ListObject references (targetTable.Parent) rather than calling .Select() or .Activate(). If an operation must inspect or modify multiple sheets, capture dynamic? activeSheetBefore = xlApp.ActiveSheet before starting and restore activeSheetBefore.Activate() in the finally block (never leave hidden Lookups or Data tabs active).
Safe Table Resizing (SheetMapperHelper.ResizeTable): Always disable table.ShowTotals = false before calling table.Resize(newRange), clear any stale trailing rows when shrinking a table, and restore table.ShowTotals in the finally block.
Table Protection (SheetProtections): Native Excel sheet protection (worksheet.Protect / Unprotect, with or without password) is completely deprecated and forbidden because it breaks table sorting, filtering, and row insertion. Never call ws.Unprotect("SKSB_Internal_Guard") or set cell.Locked (it has no effect without protection). Cell styles (such as Cost or Comma) must never be used to infer editable vs. locked status; styles are formatting only. Cell locking is authored purely via Excel Custom Data Validation (=FALSE, Stop alert) in the master templates (*.xltx). Code-driven table protection is strictly limited to: (1) header modification prevention via xlApp.Undo() in the SheetChange guard in Entry.EventsAutoOpen.cs; (2) formula self-healing via SheetProtections.HealFormulas(ws, table), which restores archetype column formulas; (3) header self-healing via SheetProtections.HealHeaders(table, headers), which restores canonical header names. SheetProtections.ExecuteWithBypass is a legacy pass-through that simply invokes the action.