# CLAUDE.md — SKSB_App Coding & Build Guardrails

Authoritative implementation rules for `SKSB_App`. Follow all stack, build, database, and Excel-DNA COM rules below.

## 1. Stack & Packages
- **Target:** `net8.0-windows`, C# 12 (`ImplicitUsings`, `Nullable`), WinForms (`PerMonitorV2` High-DPI in `Launcher`, `Portal`, and `Core` modal dialogs), `AnyCPU` (`win-x64` single-file for `Launcher`), solution `SKSB_App.slnx`.
- **Packages:** `ExcelDna.AddIn` `1.9.0`, `Dapper` `2.1.35`, `Npgsql` `8.0.3`, `PdfPig` `0.1.16` (`VehicleManager`), `System.Text.Json` (.NET 8).

## 2. Build, Clean & Publish (PowerShell)
**MSBuild Hooks (`Directory.Build.props`, `Portal.csproj`):**
- Building `Portal` or `SKSB_App.slnx` runs `KillExistingExcelAndPortalOnBuild` (`taskkill /f` on `excel.exe` and `Portal.exe`).
- Every build copies binaries/`*-AddIn*.xll` to `./AddIns/`.
- `Clean` wipes `./AddIns/`; `dotnet clean -c Release` also wipes non-`Data` files in `%USERPROFILE%\Documents\SKSB_App`.

**Commands:**
- Build single module (fast iteration): `dotnet build <Project>/<Project>.csproj -c Debug`
- Build full solution: `dotnet build SKSB_App.slnx -c Debug` (or `-c Release`)
- Clean (Debug-safe): `dotnet clean SKSB_App.slnx -c Debug`
- Publish release: `dotnet publish Launcher/Launcher.csproj -c Release` or `powershell -NoProfile -ExecutionPolicy Bypass -File .\publish.ps1`
- **Verification:** No xUnit/NUnit project; verify via `dotnet build` and in-Excel DEBUG Ribbon (`ServerEnvironment.SimulateOffline`, `AppLogger` trace at `%USERPROFILE%\Documents\SKSB_App\Data\SKSB_Runtime_Trace.log`, DEBUG `ExitConflictDialog` save intercept).

## 3. Code Conventions
- **File-scoped namespaces** at line 1; `using` directives **inside** the namespace.
- **Dot-prefixed file names:** `<DomainOrFolder>.<Component>.cs` (e.g. `StaffManager.Events.cs`, `SheetMapper.TableMapper.cs`).
- **3-line section banners** for large static classes:
  ```csharp
  // =========================================================================
  // 1. SECTION TITLE
  // =========================================================================
  ```
- **Domain Modules (`SaveContract<T>`):** Each XLL module uses 4 core files (`Events.cs`, `RibbonController.cs`, `SheetMapper.cs`, `ServerSyncCommands.cs`) wired into `SaveEntry.Save` / `SaveEntry.InterceptClose`.
- **UI-Only Helper Columns:** Mark lookup-only display properties (e.g., `StaffName` on `OdometerLog`) with `[JsonIgnore]` so they bind via `[ExcelColumn]`/`[ColumnSync]` without duplicating in JSON snapshots.
- **Staged Code (DO NOT DELETE):** Preserve `ProjectTypeRequirement`/`SkillIcon` (`Types.ScheduleRecord.cs`), `Erp.Exporter.cs` (`ErpExporter`) + `Erp.Helper.cs` (`EscapeDelimitedCell`, `ParseDecimal`) + `Erp.Importer.cs` (`ERPTableDocument.Headers`, `HasColumn`, `ERPTableRow.GetDecimal`, `GetBool`), `RuleEngine.Calculator.cs` (`WageProvider`), and `ColumnRuleAttribute.CountTargetHeader`.

## 4. Database (Dapper + PostgreSQL)
- `Dapper.DefaultTypeMap.MatchNamesWithUnderscores = true;` (snake_case DB columns <-> PascalCase C# properties).
- **Dates:** Project `DateOnly`/`DateOnly?` to `DateTime`/`DateTime?` (`d.ToDateTime(TimeOnly.MinValue)`) in Dapper parameter objects; read back via `DateOnly.FromDateTime((DateTime)row.col)`.
- **Enums:** Store as text via `.ToString()`; parse with `Enum.TryParse<TEnum>(..., ignoreCase: true, out var res)` + safe fallback.
- **Keys:** Normalize string PKs (`VehiclePlate`, `ProjectID`) with `.Trim().ToUpperInvariant()`; filter uninitialized rows (`StaffID > 0`, `!string.IsNullOrWhiteSpace(VehiclePlate)`).
- **Transactions:** Multi-table saves run in `Db.ExecuteInTransaction((conn, trans) => { ... })`. On `isFullBatch == true`, delete child tables then master using `!= ALL(@ActiveIds)` / `!= ALL(@ActivePlates)` before `INSERT ... ON CONFLICT (...) DO UPDATE` (always update `updated_at = CURRENT_TIMESTAMP`).
- **Offline Fallback:** Respect `ServerEnvironment.SimulateOffline` / `GuardConnection()` and fall back to `OfflineSnapshot.Load<List<TRecord>>()` (`%USERPROFILE%\Documents\SKSB_App\Data\<typename>_data.json`).

## 5. Excel-DNA STA & COM Guardrails
1. **Marshal COM via `ExcelAsyncUtil.QueueAsMacro`:** Never mutate sheets, tables, cells, or named ranges directly from `AutoOpen`, `AutoClose`, Ribbon callbacks, WinForms handlers, or background `Task`s. Defer COM event continuations (`WorkbookBeforeSave`/`Close`) via `QueueAsMacro` or `xlcOnTime`.
2. **Freeze & Restore Excel State:** Wrap sheet/table mutations in `try { ... } finally { ... }` capturing and restoring `EnableEvents`, `ScreenUpdating`, and `Calculation` (`-4135` manual / `-4105` auto), with each COM property access in its own `try { } catch { }`.
   - *Modal Dialog Exception:* Before `.ShowDialog()` while `ScreenUpdating == false`, temporarily set `ScreenUpdating = true`, `EnableEvents = true`, and call `Arrow.HideAll()`.
3. **Bulk 2D Arrays Only (Never Cell-by-Cell):** Read with `TableMapper.GetRange2DValues(...)`; write with `TableMapper.PopulateTableSafe<T>` (preserves archetype row style/number format, skips `=` formulas); heal with `RuleEvaluator.EvaluateAndHealTable(...)`.
4. **Multi-AddIn Coexistence & Delegate Pinning:** Use process-scoped env vars (`$"SKSB_EVENTS_HOOKED_{Environment.ProcessId}"`, `$"SKSB_LAST_CHECKED_WB_{Environment.ProcessId}"`). Store COM event delegates in `private static` fields (never unpinned lambdas) and unhook in `AutoClose()`/`UnEventHooks()`. Maintain re-entrancy guards (`_isEvaluating`, `_isInternalSaving`, `_isExiting`).
5. **Viewport & Table Protection:** Operate on `targetTable.Parent` without `.Select()`/`.Activate()` (or restore `ActiveSheet` in `finally`). In `ResizeTable`, disable `ShowTotals`, resize, clear trailing rows, restore `ShowTotals`. **Never** call `worksheet.Protect`/`Unprotect` or infer locking from cell styles—locking uses Excel Data Validation (`=FALSE`) in `*.xltx` + header undo guard in `Entry.EventsAutoOpen.cs`.