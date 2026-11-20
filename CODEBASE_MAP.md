# SKSB_App Codebase Map (Antigravity Lookup Index)

Quick feature, Excel table, C# record, and PostgreSQL table index for translating user requests into exact files.
*(For Antigravity's use — not auto-loaded by Claude Code.)*

## 1. Excel Table <-> C# Record <-> PostgreSQL Table Crosswalk

### StaffManager (`StaffManager/StaffManager.*.cs`, `Core/Types/Types.StaffRecord.cs`)
- `AllStaffsTable` <-> `StaffRecord` (`StaffDetails`) <-> `staff_master`
- `VehicleLicensesTable` <-> `VehicleLicense` <-> `staff_vehicle_licenses`
- `SkillLicensesTable` <-> `SkillLicense` <-> `staff_skill_licenses`
- `SkillLicenseTypesTable` (Lookups) <-> `SkillLicenseType` <-> `skill_license_types`

### VehicleManager (`VehicleManager/VehicleManager.*.cs`, `Core/Types/Types.VehicleRecord.cs`)
- `AllVehiclesTable` <-> `VehicleRecord` (`VehicleDetail`) <-> `vehicle_master`
- `MaintenancesTable` <-> `Maintenance` <-> `vehicle_maintenance`
- `InsurancesTable` <-> `VehicleInsurance` <-> `vehicle_insurances`
- `InsuranceClaimsTable` <-> `InsuranceClaim` <-> `vehicle_insurance_claims`
- `PermitsTable` <-> `VehiclePermit` <-> `vehicle_permits`
- `RoadTaxesTable` <-> `RoadTax` <-> `vehicle_road_taxes`
- `InspectionsTable` <-> `Inspection` <-> `vehicle_inspections`
- `OdometersTable` <-> `OdometerLog` <-> `vehicle_odometers`
- `WorkshopsTable` / `InsuranceCompaniesTable` / `FuelCardsTable` / `FuelCompanyNamesTable` <-> `Workshop` / `InsuranceCompanyRecord` / `FuelCardRecord` / `FuelCompanyNameRecord` <-> `workshops_master` / `insurance_companies` / `fuel_cards` / `fuel_companies`

### ProjectManager (`ProjectManager/ProjectManager.*.cs`, `Core/Types/Types.ProjectRecord.cs`)
- `iJobProjectsTable`, `MWOProjectsTable`, `PVTProjectsTable`, `AllProjectsTable` <-> `ProjectRecord` (`ProjectDetails`) <-> `project_master`
- `CompaniesTable`, `AllCompaniesTable` <-> `CompanyRecord` <-> `company_master`
- `ProjectSummaryTable` <-> `ProjectSummaryRow` (computed summary view)

### ScheduleManager (`ScheduleManager/ScheduleManager.*.cs`, `Core/Types/Types.ScheduleRecord.cs`)
- Daily Schedule Board (2D grid mapped by `ScheduleBoardMapper.cs`, ranked by `MatchEngine.cs`) <-> `DailyProjectSchedule` <-> `daily_project_schedules`
- `DailyAttendanceTable` (rolled up by `AttendanceRollup.cs`) <-> `DailyWorkerAttendance` <-> `daily_attendances`
- `AttendanceEntriesTable` (append-only punches + `Core/ThumbprintSync/DeviceSyncManager.cs`) <-> `AttendanceEntry` <-> `attendance_entries`
- `WorkerPreferencesTable` <-> `WorkerPreference` <-> `worker_preferences`
- `TeamTable` <-> `TeamRecord` <-> `teams_master`
- `ProjectTypeRequirementsTable` / `SkillIconsTable` (Lookups-only) <-> `ProjectTypeRequirement` / `SkillIcon` (workbook-only, not in DB)

## 2. Core Subsystem & Symptom Index (`Core/`)
- **Auto-updater (`SKSB_App.exe`) & Hub (`Portal.exe`)**: `Launcher/Program.cs`, `Portal/PortalForm.cs`, `Portal/Router.cs`, `publish.ps1`, `version.json`
- **Add-in Startup & COM Hooks**: `Core/02_WorkbookOpen/Entry/Entry.EventsAutoOpen.cs`, `Entry.ManagerEventsBase.cs`
- **Hidden `Lookups` / `Data` Sheets & Startup Sync**: `Core/02_WorkbookOpen/Sheet/Sheet.Operations.cs`, `Core/02_WorkbookOpen/Sync/Sync.DataList.cs`, `Sync.EnumList.cs`, `Sync.OfflineSave.cs`
- **Cell Dropdowns, Floating Arrow, Two-Way Column Sync**: `Core/03_RuntimeUI/DropDown/DropDown.Arrow.cs`, `DropDown.Commands.cs`, `DropDown.SyncEngine.cs`, `DropDown.CustomProviders.cs`, `DropDown.Metadata.cs`
- **Top Bar Date / Day Parameter Cells & Filtering**: `Core/03_RuntimeUI/DropDown/DropDown.ParameterCells.cs`, `Core/03_RuntimeUI/DateFilter/DateFilter.Helper.cs`, `DateFilter.Enum.cs`
- **Rule Engine, Number Formatting (`0000`, `#`, `—` em-dash), Auto-Defaults, State Guards**: `Core/03_RuntimeUI/RuleEngine/RuleEngine.Registry.cs`, `RuleEngine.Evaluator.cs`, `RuleEngine.Formats.cs`, `RuleEngine.DataCache.cs`, `RuleEngine.Calculator.cs`
- **Row Add/Delete & Duplicate Import Guard**: `Core/03_RuntimeUI/TableRowManager/TableRowManager.cs`, `TableRowManager.ImportDuplicates.cs`
- **Animated Loading Progress Overlay (`ScreenUpdating = false`)**: `Core/03_RuntimeUI/Progress/Progress.LoadingOverlay.cs`
- **Shared Ribbon & Offline Mode Status / Reconnect Button**: `Core/03_RuntimeUI/Ribbon/Ribbon.ManagerRibbonBase.cs`, `Core/Integration/Database/Database.ServerEnvironment.cs`
- **Template Rules Audit / Exporter / Importer**: `Core/03_RuntimeUI/TemplateEditor/RulesAudit/RulesAudit.*.cs` (auto-exports `*_Rules.md` silently on template save via `Entry.EventsAutoOpen.cs` `WorkbookAfterSave`)
- **Save Pipeline, DEBUG Save Intercept, Offline Snapshot**: `Core/04_WorkbookSave/Entry/Save.Contract.cs`, `Save.Entry.cs`, `Save.ExitConflictDialog.cs`, `Storage/Storage.OfflineSnapshot.cs`, `Storage.SaveNotificationDialog.cs`
- **2D Array Table Read/Write & Lookups**: `Core/04_WorkbookSave/SheetMapper/SheetMapper.TableMapper.cs`, `SheetMapperHelper.cs`, `WorkbookMapper.cs`, `StaffLookupHelper.cs`, `VehiclePlateHelper.cs`
- **Save Validation & Historical/Past-Record Guard**: `Core/04_WorkbookSave/Validation/Validation.RulesExtractor.cs`, `Validation.SummaryDialog.cs`, `Validation.PastRecordGuard.cs`
- **PostgreSQL / Dapper & Server-Authoritative Fetch**: `Core/Integration/Database/Database.PostgreSqlService.cs`, `Database.ServerDispatcher.cs`, `Database.ServerDispatcher.Schedule.cs`, `Database.ServerSyncHelper.cs`, `Schema/schedule_schema.sql`
- **ERP / SQL Payroll & Fuel PDF Import**: `Core/Integration/Erp/Erp.*.cs`, `StaffManager/StaffManager.Importer.cs`, `VehicleManager/VehicleManager.SheetMapper.FuelImport.cs`
- **Declarative Attributes**: `Core/Types/Attributes/Attributes.ColumnRules.cs`, `Attributes.DataAttributes.cs`, `Attributes.ExcelColumn.cs`
- **Runtime Log (`SKSB_Runtime_Trace.log`)**: `Core/Diagnostics/Diagnoestics.AppLogger.cs`

## 3. Core Domain Invariants (From Past Architecture Decisions)
- **Online vs. Offline Mode:** Online mode is strictly server-authoritative (`ServerSyncHelper.cs`) and never mixes offline snapshot rows with server rows. Both online and offline saves update `OfflineSnapshot` so offline data is ready if the server drops.
- **Historical Dates (`ScheduleDate`):** Staff active status and skill/vehicle license validity on schedules must be evaluated as of the selected `ScheduleDate` (so resigned staff or later-expired licenses remain valid on historical dates), gated by `PastRecordGuard` when editing dates before today.
- **Number Formatting & Computed Placeholders (`—`):** C# (`RuleEngine.Formats.cs`) owns column number formats (`0000` for Staff ID, integer `#` for Company ID, etc.); computed/read-only columns displaying `—` must not fail numeric validation on save.
