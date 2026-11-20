---
trigger: model_decision
description: >-
  SKSB_App Excel template colour scheme standard based on the Microsoft Office
  Paper theme. Load whenever designing, styling, auditing, or documenting
  `.xltx` workbook themes, manager table headers, or cross-manager derived
  column header colours.
---

# SKSB_App Template Colour Scheme Standard (`Paper` Theme)

Authoritative colour standard for all SKSB_App Excel templates (`Templates/*.xltx` and reserved `Templates/Unuse/*.xltx`), built on the official Microsoft Office **Paper** theme colour palette (`Paper.xml`). Follow this standard whenever creating new templates, adding tables or columns, styling secondary/sub-topic sheets, colouring cross-manager derived headers, or auditing existing workbooks.

---

## 1. Official Microsoft Office `Paper` Colour Palette

All manager templates standardize on Excel's built-in **Paper** colour scheme (**Page Layout → Colors → Paper**, defined in Office `Document Themes 16\Theme Colors\Paper.xml`). Because Paper uses muted, low-glare editorial tones, **`+60%` (Lighter 60%)** provides the ideal visual weight for primary table headers with crisp `#000000` Bold text (WCAG AAA contrast ≥ `13.9:1`), **`+40%` (Lighter 40%)** provides the standard **darker shade** for secondary sheets / sub-topics of the same manager (WCAG AAA contrast ≥ `11.0:1`), and **Base (`0%`)** serves sheet tabs and UI badges.

| Theme Slot | Paper Name | Base Hex (`0%`) | `+80%` *(Soft Band)* | **`+60%` *(Primary Header)*** | **`+40%` *(Sub-Topic / Secondary Sheet)*** | `-25%` *(Deep Accent)* | `-50%` *(Darkest)* |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Dark 1 (`dk1`)** | Black *(Text)* | `#000000` | `#CCCCCC` | `#999999` | `#666666` | `#000000` | `#000000` |
| **Light 1 (`lt1`)** | White *(Background)* | `#FFFFFF` | `#FFFFFF` | `#FFFFFF` | `#FFFFFF` | `#BFBFBF` | `#808080` |
| **Dark 2 (`dk2`)** | Dark Olive | `#444D26` | `#EFF2E5` *(+90%)* | **`#D7DEBD`** *(+75%)* | **`#AEBE7B`** *(+50%)* | `#333A1C` | `#222613` |
| **Light 2 (`lt2`)** | Pale Cream | `#FEFAC9` | `#FFFEF4` | `#FFFDE9` | `#FEFCDF` | `#FCF059` | `#DFCF04` |
| **Accent 1** | Sage Green | `#A5B592` | `#EDF0E9` | **`#DBE1D3`** | **`#C9D3BE`** | `#7D9263` | `#536142` |
| **Accent 2** | Warm Amber | `#F3A447` | `#FDEDDA` | **`#FADBB5`** | **`#F8C891`** | `#DD7E0E` | `#93540A` |
| **Accent 3** | Mustard Gold | `#E7BC29` | `#FAF2D4` | **`#F5E4A9`** | **`#F1D77F`** | `#B79315` | `#7A620E` |
| **Accent 4** | Dusty Rose | `#D092A7` | `#F6E9ED` | **`#ECD3DC`** | **`#E3BECA`** | `#B55475` | `#7C354D` |
| **Accent 5** | Muted Lavender | `#9C85C0` | `#EBE7F2` | **`#D7CEE6`** | **`#C4B6D9`** | `#7153A1` | `#4C376B` |
| **Accent 6** | Steel Blue | `#809EC2` | `#E6ECF3` | **`#CCD8E7`** | **`#B3C5DA`** | `#4E75A3` | `#344E6D` |
| **Hyperlink (`hlink`)** | Violet | `#8E58B6` | — | — | — | — | — |
| **Followed (`folHlink`)** | Warm Gray | `#7F6F6F` | — | — | — | — | — |

> **Note on `Dark 2` (`dk2`) tint stops:** Because `#444D26` is darker than the six accent slots, Excel's Theme Colors picker generates `+90%`, `+75%`, `+50%`, `-25%`, `-50%` stops instead of `+80%`, `+60%`, `+40%`. Its **`+75%` (`#D7DEBD`)** and **`+50%` (`#AEBE7B`)** swatches match the exact lightness and contrast of the accents' `+60%` and `+40%` stops.

---

## 2. Manager Colour Assignments (One Paper Colour per Manager)

Each SKSB_App manager owns **one distinct hue** from the `Paper` palette — matching the established colour identities across the suite (Blue for Staff, Yellow for Vehicles, Green for Projects, Orange for Schedules, Purple for Costing, Dusty Rose for Salary, and Dark Olive for Inventory).

### Active Managers

| Manager | Template | Paper Slot & Hue | Primary Header Fill (`SKSBStyleTableSimple`) | Sub-Topic / Secondary Sheet Fill *(Darker)* | Tab / Badge Accent | Header Font |
| :--- | :--- | :--- | :---: | :---: | :---: | :--- |
| **StaffManager** | `Staff_Manager_Template.xltx` | **Accent 6** — Steel Blue | `Accent 6, +60%` **`#CCD8E7`** | `Accent 6, +40%` **`#B3C5DA`** | `#809EC2` | `#000000` Bold *(14.5:1 / 11.9:1)* |
| **VehicleManager** | `Vehicle_Manager_Template.xltx` | **Accent 3** — Mustard Gold | `Accent 3, +60%` **`#F5E4A9`** | `Accent 3, +40%` **`#F1D77F`** | `#E7BC29` | `#000000` Bold *(16.6:1 / 14.8:1)* |
| **ProjectManager** | `Project_Manager_Template.xltx` | **Accent 1** — Sage Green | `Accent 1, +60%` **`#DBE1D3`** | `Accent 1, +40%` **`#C9D3BE`** | `#A5B592` | `#000000` Bold *(15.7:1 / 13.5:1)* |
| **ScheduleManager** | `Schedule_Manager_Template.xltx` | **Accent 2** — Warm Amber | `Accent 2, +60%` **`#FADBB5`** | `Accent 2, +40%` **`#F8C891`** | `#F3A447` | `#000000` Bold *(15.9:1 / 13.7:1)* |

### Reserved Managers (`Templates/Unuse/`)

| Manager | Template | Paper Slot & Hue | Primary Header Fill (`SKSBStyleTableSimple`) | Sub-Topic / Secondary Sheet Fill *(Darker)* | Tab / Badge Accent | Header Font |
| :--- | :--- | :--- | :---: | :---: | :---: | :--- |
| **CostManager** *(Costing)* | `Unuse/Cost_Manager_Template.xltx` | **Accent 5** — Muted Lavender *(Purple)* | `Accent 5, +60%` **`#D7CEE6`** | `Accent 5, +40%` **`#C4B6D9`** | `#9C85C0` | `#000000` Bold *(13.9:1 / 11.0:1)* |
| **SalaryManager** *(Reserved)* | `Unuse/Salary_Manager_Template.xltx` | **Accent 4** — Dusty Rose | `Accent 4, +60%` **`#ECD3DC`** | `Accent 4, +40%` **`#E3BECA`** | `#D092A7` | `#000000` Bold *(14.9:1 / 12.5:1)* |
| **InventoryManager** *(Reserved)* | `Unuse/Inventory_Manager_Template.xltx` | **Dark 2** — Dark Olive | `Dark 2, +75%` **`#D7DEBD`** | `Dark 2, +50%` **`#AEBE7B`** | `#444D26` | `#000000` Bold *(15.1:1 / 10.5:1)* |

---

## 3. Core Principles & Styling Rules

### Rule 1 — Workbook Table Baseline (`SKSBStyleTableSimple`)
Every manager workbook sets its custom `SKSBStyleTableSimple` table style using its own assigned Paper colour:
- **Header Row (`headerRow`):** Solid fill = Host Manager's **Primary Header Fill (`+60%` Lighter; `+75%` for `Dark 2`)**; Font = `#000000` Bold.
- **First Row Stripe (`firstRowStripe`):** Solid fill = `White, Darker 5%` (`#F2F2F2`).
- **Second Row Stripe (`secondRowStripe`):** `White` (`#FFFFFF` / automatic).
- **Default Header Cells:** Leave host-owned master columns with **No Fill** on the cell itself so they inherit `SKSBStyleTableSimple` automatically.

### Rule 2 — Darker Shade for New Sheets & Sub-Topics of the Same Manager (Mandatory)
Within any manager workbook, distinguish the primary/master sheet from secondary sheets, child tables, and sub-topics by stepping one shade deeper **within the same Paper hue**:
- **Primary / Master Sheet (`+60%`, or `+75%` for `Dark 2`):** The manager's main master sheet(s) (`All Staffs`, `All Vehicles`, `iJob/PVT/MWO Projects`, `Daily Schedule`, `Monthly Report`, `All Inventories`) use the **Primary Header Fill** (inheriting `SKSBStyleTableSimple`).
- **Secondary Sheets, Child Tables & Sub-Topics (`+40%`, or `+50%` for `Dark 2`):** Every new sheet, secondary sheet, child table, or sub-topic within the same manager uses the manager's **Darker Sub-Topic Fill (`+40%` / `+50%`)** on all its host-owned sub-topic columns:
  - **Inherited Same-Manager Master Keys:** When a child table repeats the host manager's primary master key columns (e.g., `Staff ID`/`Staff Name` in `StaffManager` child sheets, `Vehicle Plate`/`Vehicle Type` in `VehicleManager` child sheets, `Item ID`/`Item Name` in `InventoryManager` child sheets), keep those inherited master keys at **Primary (`+60%` / `+75%`, `No Fill`)**, while all sub-topic/child-specific columns take **Darker (`+40%` / `+50%`)**.
  - **Independent Sub-Topic Sheets & Supporting Tables:** For supporting directories or sub-topic sheets/tables (`Workshops`, `Companies`, `Project Summary`, `Daily Attendances`, `Attendance Entries`, `Preference`, `SkillsTable`, `TeamTable`, `FuelCardsTable`, `InsuranceCompaniesTable`, `Claims`, `Accounts`, `Leaves`, `Borrowed Money`, `Allowance`, `Holiday and Settings`), all host-owned columns take **Darker (`+40%` / `+50%`)**.
- Never switch hues or introduce ad-hoc standard yellows/oranges/pinks (`#FFFF00`, `#FF9900`, `#FFCCCC`) for sub-tables — always stay strictly inside the manager's assigned Paper hue.

### Rule 3 — Cross-Manager Derived Column Header Rule (Mandatory)
**Whenever a table displays a column whose value, lookup list, foreign key, two-way sync, or rolled-up metric derives from another manager, colour that column's header cell with the *source manager's* canonical Primary Header Fill (`+60%` Lighter; `+75%` for `Dark 2`) and `#000000` Bold font.**

- This applies to:
  1. **Foreign-key & synced display columns** referencing another manager's `[DataList]` (e.g., `Staff ID`, `Staff Name`, `PIC`, `Team Lead`, `Project`, `Project ID`, `Vehicle Plate`, `Vehicle Type`).
  2. **Cross-manager read-only/derived attributes** synced from another manager on load (e.g., `WorkerPreferencesTable[Licensed Skills]` from `StaffManager`; `AllVehiclesTable[Driven By]` / `Used At`).
  3. **Cross-manager rolled-up metrics** aggregated from another manager's tables (e.g., `Days Worked`, `Total Regular Shifts`, `Total Hours` in `ProjectManager` from `ScheduleManager`; `Total Odometer`, `Fuel Cost`, `Vehicles Used` from `VehicleManager`).
- **Always use the source manager's canonical `+60%` Primary Header Fill** — even on secondary/sub-topic sheets whose local columns are `+40%` — so a `Staff ID` header looks identical (`#CCD8E7`) everywhere in the entire suite.
- **Exception — Intra-Manager Lookups:** Columns synced from another table *inside the same manager* (e.g., `Company ID`/`Company Name` in `ProjectManager`, `Workshop Name`/`Vehicle Plate` in `VehicleManager`, `Team` in `ScheduleManager`) are **not** cross-manager columns — they follow Rule 2 (same manager).

### Rule 4 — Multi-Source & Composite Column Tie-Breaker
- **Host + External Source (e.g., `FuelCardsTable[Card Owner ID]` / `Card Owner Name` accepting either `AllVehiclesTable` or `StaffIDList`):** Keep the **host manager's colour** (`+40%` on a secondary table) because the host domain is a primary valid source.
- **Composite Formulas Across Multiple Managers (e.g., `Actual Project Cost = Fuel + Labour + Claim`, `Profit / Margin`):** Keep the **host manager's colour** because the metric is owned and calculated by the host manager.
- **Columns Waiting on a Reserved Manager (`Labour Cost`, `Claim Cost`, `Claims Made`, `Pending Claims`, `Is Holiday`, `Total Claim`):** Colour using the **reserved source manager's** canonical Primary swatch (**Costing Purple `#D7CEE6`** or **Salary Dusty Rose `#ECD3DC`**) so templates are future-proof and visually flag external inputs.

### Rule 5 — Neutral Chrome, Tabs & Non-Data Surfaces
- **Top-Bar Parameter / Filter Headers (`DateFilter`, `Date Start`, `Date End`, `Display Range`, `Daily Schedule!A1:B1`, `Project Summary` / `All Vehicles` / `Daily Attendances` top filter bars):** Always neutral **`White, Darker 15%` (`#D9D9D9`)** with `#000000` Bold font. Never tint parameter bars with a manager colour.
- **Sheet Tabs (`tabColor`):** Visible sheets use the Host Manager's **Base (`0%`)** Paper colour on primary sheets and **Darker (`-25%`)** on secondary/sub-topic sheets (or uniform Base `0%` across all visible sheets of that manager; hidden/technical `Lookups` & `Data` tabs remain uncoloured).
- **Hidden / Technical Sheets (`Lookups`, `Data`):** Unstyled (`No Fill`, `#000000` regular font).
- **Data Cells:** Always keep data rows `#FFFFFF` / `#F2F2F2` zebra stripes; never apply solid manager header fills to data rows so Conditional Formatting (`lightUp` status stripes, `ColorScale` heatmaps, validation borders, and `IconSet`s per `conditional-formatting.md`) remains fully legible.
- **C# Ownership Boundary:** C# add-ins never author or overwrite header fill colours — header styling lives strictly in `.xltx` templates (except `ScheduleBoardMapper.PrepareColumns`, which clones the prototype slot `B4:D18` / `B20:D34` formatting from `Schedule_Manager_Template.xltx`).

---

## 4. Complete Cross-Manager & Sub-Topic Header Colour Mapping

> **Canonical Cross-Manager Swatch Legend (Primary `+60%`, Bold `#000000`):**
> - **Staff** = Steel Blue **`#CCD8E7`** *(Darker sub-topic: `#B3C5DA`)*
> - **Vehicle** = Mustard Gold **`#F5E4A9`** *(Darker sub-topic: `#F1D77F`)*
> - **Project** = Sage Green **`#DBE1D3`** *(Darker sub-topic: `#C9D3BE`)*
> - **Schedule** = Warm Amber **`#FADBB5`** *(Darker sub-topic: `#F8C891`)*
> - **Costing** *(Reserved)* = Muted Lavender / Purple **`#D7CEE6`** *(Darker sub-topic: `#C4B6D9`)*
> - **Salary** *(Reserved)* = Dusty Rose **`#ECD3DC`** *(Darker sub-topic: `#E3BECA`)*
> - **Inventory** *(Reserved)* = Dark Olive **`#D7DEBD`** *(Darker sub-topic: `#AEBE7B`)*

### 4.1 `StaffManager` (`Staff_Manager_Template.xltx`) — Host: Steel Blue (`#CCD8E7` / `#B3C5DA`)

`StaffManager` is the root people directory and has **no cross-manager derived columns**.

| Sheet *(Role)* | Table | Column Header *(Col)* | Owner / Source | Header Fill Swatch |
| :--- | :--- | :--- | :--- | :--- |
| `All Staffs` *(Primary)* | `AllStaffsTable` (`A1:J2`) | All columns (`A:J`) | **StaffManager** *(Primary)* | Steel Blue **`#CCD8E7`** (`No Fill` via `SKSBStyleTableSimple`) |
| `Vehicle Licenses` *(Sub-Topic)* | `VehicleLicensesTable` (`A1:H3`) | `Staff ID` (`B`), `Staff Name` (`C`) | **StaffManager** *(Master Key)* | Steel Blue **`#CCD8E7`** (`No Fill`) |
| `Vehicle Licenses` *(Sub-Topic)* | `VehicleLicensesTable` (`A1:H3`) | `Vehicle License ID` (`A`), `Vehicle License Class` (`D`), `Renew Date` (`E`), `Expiry Date` (`F`), `Status` (`G`), `Cost` (`H`) | **StaffManager** *(Sub-Topic)* | Steel Blue Darker **`#B3C5DA`** (`Accent 6 +40%`) |
| `Skill Licenses` *(Sub-Topic)* | `SkillLicensesTable` (`A1:H3`) | `Staff ID` (`B`), `Staff Name` (`C`) | **StaffManager** *(Master Key)* | Steel Blue **`#CCD8E7`** (`No Fill`) |
| `Skill Licenses` *(Sub-Topic)* | `SkillLicensesTable` (`A1:H3`) & `SkillLicenseTypesTable` (`J1:K2`) | `Skill License ID` (`A`), `Skill License Type` (`D`), `Renew Date` (`E`), `Expiry Date` (`F`), `Status` (`G`), `Cost` (`H`), `Skill License Types` (`J`), `Display Icon` (`K`) | **StaffManager** *(Sub-Topic)* | Steel Blue Darker **`#B3C5DA`** (`Accent 6 +40%`) |

---

### 4.2 `VehicleManager` (`Vehicle_Manager_Template.xltx`) — Host: Mustard Gold (`#F5E4A9` / `#F1D77F`)

| Sheet *(Role)* | Table | Column Header *(Col)* | Owner / Source | Header Fill Swatch | Derivation / Rule |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `All Vehicles` *(Primary)* | `AllVehiclesTable` (`A4:Z5`) | `Driven By` (`Q`) | **StaffManager** | Steel Blue **`#CCD8E7`** | Displays assigned driver (`Staff` entity) |
| `All Vehicles` *(Primary)* | `AllVehiclesTable` (`A4:Z5`) | `Used At` (`R`) | **ProjectManager** | Sage Green **`#DBE1D3`** | Displays assigned project site (`Project` entity) |
| `All Vehicles` *(Primary)* | `AllVehiclesTable` (`A4:Z5`) | All other columns (`A:P`, `S:Z`) | **VehicleManager** *(Primary)* | Mustard Gold **`#F5E4A9`** | Own fields & intra-manager cost rollups (`No Fill` via `SKSBStyleTableSimple`) |
| `Odometers` *(Sub-Topic)* | `OdometersTable` (`A1:U2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | **VehicleManager** *(Master Key)* | Mustard Gold **`#F5E4A9`** | Inherited master keys (`No Fill`) |
| `Odometers` *(Sub-Topic)* | `OdometersTable` (`A1:U2`) | `Staff ID` (`K`), `Staff Name` (`L`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffIDList")]` / `StaffShortNameList` |
| `Odometers` *(Sub-Topic)* | `OdometersTable` (`A1:U2`) | `Project` (`N`) | **ProjectManager** | Sage Green **`#DBE1D3`** | `[ColumnRule("ProjectIDList")]` |
| `Odometers` *(Sub-Topic)* | `OdometersTable` (`A1:U2`) | `Date` (`A`), `Odometer Before`..`Card Owner` (`D:J`), `Station Name` (`M`), `Fuel Type`..`Remarks` (`O:U`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Sub-topic odometer & fuel log columns (`Accent 3 +40%`) |
| `Odometers` *(Sub-Topic)* | `FuelCardsTable` (`W1:Z2`), `FuelCompanyNamesTable` (`AB1:AB2`) | All columns (`W:Z`, `AB`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Rule 4 tie-breaker keeps Host sub-topic on `Card Owner ID`/`Name` |
| `Maintenances` *(Sub-Topic)* | `MaintenancesTable` (`A1:Y2`) | `Vehicle Plate` (`A`), `Vehicle Type` (`B`) | **VehicleManager** *(Master Key)* | Mustard Gold **`#F5E4A9`** | Inherited master keys (`No Fill`) |
| `Maintenances` *(Sub-Topic)* | `MaintenancesTable` (`A1:Y2`) | `Maintenance ID`..`Parts` (`C:Y`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Sub-topic maintenance columns |
| `Workshops` *(Sub-Topic)* | `WorkshopsTable` (`A4:I5`) | All columns (`A:I`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Supporting directory (`DateFilterTableWorkshops` `A1:C2` stays `#D9D9D9`) |
| `Permits`, `Road Taxes`, `Inspections` *(Sub-Topics)* | `PermitsTable`, `RoadTaxesTable`, `InspectionsTable` | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | **VehicleManager** *(Master Key)* | Mustard Gold **`#F5E4A9`** | Inherited master keys (`No Fill`) |
| `Permits`, `Road Taxes`, `Inspections` *(Sub-Topics)* | *(same 3 tables)* | All other columns (`A`, `D:...`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Sub-topic compliance columns |
| `Insurances` *(Sub-Topic)* | `InsurancesTable` (`A1:M2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | **VehicleManager** *(Master Key)* | Mustard Gold **`#F5E4A9`** | Inherited master keys (`No Fill`) |
| `Insurances` *(Sub-Topic)* | `InsurancesTable` (`A1:M2`) & `InsuranceCompaniesTable` (`O1:U2`) | `Renewal Date` (`A`), `Insurance Renewal ID`..`Remarks` (`D:M`), `Insurance Company`..`Vehicle Plates Under Agent` (`O:U`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Sub-topic insurance & directory columns |
| `Insurance Claims` *(Sub-Topic)* | `InsuranceClaimsTable` (`A1:L2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | **VehicleManager** *(Master Key)* | Mustard Gold **`#F5E4A9`** | Inherited master keys (`No Fill`) |
| `Insurance Claims` *(Sub-Topic)* | `InsuranceClaimsTable` (`A1:L2`) | `Staff ID` (`F`), `Staff Name` (`G`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffIDList")]` / `StaffShortNameList` |
| `Insurance Claims` *(Sub-Topic)* | `InsuranceClaimsTable` (`A1:L2`) | `Claim Date` (`A`), `Insurance Claim ID` (`D`), `Settlement Date` (`E`), `Maintenance ID`..`Remarks` (`H:L`) | **VehicleManager** *(Sub-Topic)* | Mustard Gold Darker **`#F1D77F`** | Sub-topic claim columns |

---

### 4.3 `ProjectManager` (`Project_Manager_Template.xltx`) — Host: Sage Green (`#DBE1D3` / `#C9D3BE`)

> **Column-Order Note:** `iJobProjectsTable` places `Labour Cost` at `X`, `Total Odometer` at `Y`, `Fuel Cost` at `Z`; `PVTProjectsTable` and `MWOProjectsTable` place `Total Odometer` at `X`, `Fuel Cost` at `Y`, `Labour Cost` at `Z`. Always match colours by **Header Name**.

| Sheet *(Role)* | Table | Column Header *(Col)* | Owner / Source | Header Fill Swatch | Derivation / Rule |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `iJob Projects`, `PVT Projects`, `MWO Projects` *(Primary)* | `iJobProjectsTable`, `PVTProjectsTable`, `MWOProjectsTable` (`A4:AC5`) | `Staff ID` (`H`), `PIC` (`I`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffIDList")]` / `StaffShortNameList` |
| *(same 3 sheets)* | *(same 3 tables)* | `Days Worked` (`S`), `Total Regular Shifts` (`T`), `Total Hours` (`U`), `Overtime Shifts` (`V`), `Total Overtime Hours` (`W`) | **ScheduleManager** | Warm Amber **`#FADBB5`** | Rolled up from daily schedules & attendance |
| *(same 3 sheets)* | *(same 3 tables)* | `Total Odometer` (`Y`/`X`), `Fuel Cost` (`Z`/`Y`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | Rolled up from `VehicleManager` (`OdometersTable`) |
| *(same 3 sheets)* | *(same 3 tables)* | `Labour Cost` (`X`/`Z`) | **SalaryManager** *(Reserved)* | Dusty Rose **`#ECD3DC`** | Sum of attendance labour cost *(Salary domain)* |
| *(same 3 sheets)* | *(same 3 tables)* | `Claim Cost` (`AA`) | **CostManager** *(Costing)* | Muted Lavender **`#D7CEE6`** | Sum of project expense claims *(Costing domain)* |
| *(same 3 sheets)* | *(same 3 tables)* | `Project Code`..`Company Name` (`A:G`), `District`..`Remarks` (`J:R`), `Actual Project Cost` (`AB`), `Profit / Margin` (`AC`) | **ProjectManager** *(Primary)* | Sage Green **`#DBE1D3`** | Own project/company fields & composite formulas (`No Fill`) |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Date` (`A`) | **ProjectManager** *(Sub-Topic)* | Sage Green Darker **`#C9D3BE`** | Summary row date key (`Accent 1 +40%`) |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Team Lead` (`B`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffNameList")]` |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Time Start` (`C`), `Time End` (`D`), `Total Hours` (`E`), `Workers` (`F`), `Regular Shifts` (`G`), `Overtime Shifts` (`H`), `Overtime Hours` (`I`), `Additional OT Hours` (`J`) | **ScheduleManager** | Warm Amber **`#FADBB5`** | Derived from daily schedule & `DailyAttendanceTable` |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Is Holiday` (`K`), `Labour Cost` (`L`) | **SalaryManager** *(Reserved)* | Dusty Rose **`#ECD3DC`** | Derived from holiday calendar & attendance salary |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Vehicles Used` (`M`), `Vehicle Types` (`N`), `Odometer Used` (`O`), `Fuel Cost` (`P`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | Derived from `AllVehiclesTable` / `OdometersTable` |
| `Project Summary` *(Sub-Topic)* | `ProjectSummaryTable` (`A4:R5`) | `Claims Made` (`Q`), `Pending Claims` (`R`) | **CostManager** *(Costing)* | Muted Lavender **`#D7CEE6`** | Derived from `CostManager` claims |
| `Companies` *(Sub-Topic)* | `CompaniesTable` (`A1:C2`) | `Company ID` (`A`), `Legal Company Name` (`B`), `Short Company Name` (`C`) | **ProjectManager** *(Sub-Topic)* | Sage Green Darker **`#C9D3BE`** | Supporting client directory (`Accent 1 +40%`) |

---

### 4.4 `ScheduleManager` (`Schedule_Manager_Template.xltx`) — Host: Warm Amber (`#FADBB5` / `#F8C891`)

| Sheet *(Role)* | Table / Region | Column Header *(Col / Ref)* | Owner / Source | Header Fill Swatch | Derivation / Rule |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `Daily Schedule` *(Primary)* | Top Bar (`A1:B1`) | `Date`, `Day` | *(Neutral Chrome)* | Gray **`#D9D9D9`** | Sheet parameter header (`White -15%`) |
| `Daily Schedule` *(Primary)* | Prototype (`B4`, `B20`) | `Project Code` | **ProjectManager** | Sage Green **`#DBE1D3`** | Active project slot key (`ProjectIDList`) |
| `Daily Schedule` *(Primary)* | Prototype (`C4:D4`, `C20:D20`) | `Premix Hole Amount`, `Project start time - end time` | **ScheduleManager** *(Primary)* | Warm Amber **`#FADBB5`** | Schedule-owned slot headers *(cloned by `PrepareColumns`)* |
| `Daily Schedule` *(Primary)* | Left Band Rail | `Day Work` (`A4:A14`), `Night Work` (`A20:A30`) | **ScheduleManager** *(Primary / Sub-Band)* | Day: **`#FADBB5`** (`+60%`) / Night: **`#F8C891`** (`+40%`) | Shift band labels (worker rows stay `#FFFFFF`/`#F2F2F2`) |
| `Daily Schedule` *(Primary)* | Left Vehicle Rail | `Vehicles` (`A15:A18`, `A31:A34`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | Vehicle sub-block rail (vehicle rows stay `#FFFFFF`/`#F2F2F2`) |
| `Daily Attendances` *(Sub-Topic)* | `DailyAttendanceTable` (`A4:L5`) | `Staff ID` (`A`), `Staff Name` (`B`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffIDList")]` / `StaffNameList` |
| `Daily Attendances` *(Sub-Topic)* | `DailyAttendanceTable` (`A4:L5`) | `Project Check-Ins` (`F`) | **ProjectManager** | Sage Green **`#DBE1D3`** | Distinct `Project Code`s checked into that day |
| `Daily Attendances` *(Sub-Topic)* | `DailyAttendanceTable` (`A4:L5`) | `Status` (`C`), `Clock-In Time` (`D`), `Clock-Out Time` (`E`), `Regular Shift`..`Is Allocated` (`G:L`) | **ScheduleManager** *(Sub-Topic)* | Warm Amber Darker **`#F8C891`** | Attendance rollup columns (`Accent 2 +40%`) |
| `Attendance Entries` *(Sub-Topic)* | `AttendanceEntriesTable` (`A1:J2`) | `Staff ID` (`C`), `Staff Name` (`D`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `[ColumnRule("StaffIDList")]` / `StaffNameList` |
| `Attendance Entries` *(Sub-Topic)* | `AttendanceEntriesTable` (`A1:J2`) | `Project ID` (`E`) | **ProjectManager** | Sage Green **`#DBE1D3`** | `[ColumnRule("ProjectIDList")]` |
| `Attendance Entries` *(Sub-Topic)* | `AttendanceEntriesTable` (`A1:J2`) | `Date` (`A`), `Time` (`B`), `Entry Type`..`WhatsApp Picture` (`F:J`) | **ScheduleManager** *(Sub-Topic)* | Warm Amber Darker **`#F8C891`** | Punch log columns (`Accent 2 +40%`) |
| `Preference` *(Sub-Topic)* | `WorkerPreferencesTable` (`A1:L2`) | `Staff ID` (`A`), `Staff Name` (`B`), `Licensed Skills` (`E`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `AllFieldStaffIDList`/`NameList`, and `Licensed Skills` derived from `StaffManager` licenses |
| `Preference` *(Sub-Topic)* | `WorkerPreferencesTable` (`A1:L2`) | `Preferred Vehicle Type` (`H`), `Dedicated Vehicle Plate` (`I`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | `VehicleTypeList` / `VehiclePlateList` |
| `Preference` *(Sub-Topic)* | `WorkerPreferencesTable` (`A1:L2`) | `Team` (`C`), `Primary Skill` (`D`), `Other Skills` (`F`), `Blacklisted Skills` (`G`), `Preferred Shift`..`Rest Day Preference` (`J:L`) | **ScheduleManager** *(Sub-Topic)* | Warm Amber Darker **`#F8C891`** | Own scheduling preference columns (`Accent 2 +40%`) |
| `Preference` *(Sub-Topic)* | `TeamTable` (`N1:P2`) | `Team Leader` (`O`) | **StaffManager** | Steel Blue **`#CCD8E7`** | `AllFieldStaffNameList` |
| `Preference` *(Sub-Topic)* | `TeamTable` (`N1:P2`) | `Default Vehicles` (`P`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | `VehiclePlateList` |
| `Preference` *(Sub-Topic)* | `TeamTable` (`N1:P2`) & `SkillsTable` (`R1:S2`) | `Team Name` (`N`), `Skills` (`R`), `Display Icons` (`S`) | **ScheduleManager** *(Sub-Topic)* | Warm Amber Darker **`#F8C891`** | Own team & skill catalog columns |
| `Preference` *(Sub-Topic)* | `ProjectRequirementsTable` (`U1:W2`) | `Project Type` (`U`) | **ProjectManager** | Sage Green **`#DBE1D3`** | `ProjectTypeList` |
| `Preference` *(Sub-Topic)* | `ProjectRequirementsTable` (`U1:W2`) | `Required Vehicles` (`W`) | **VehicleManager** | Mustard Gold **`#F5E4A9`** | `VehicleTypeList` |
| `Preference` *(Sub-Topic)* | `ProjectRequirementsTable` (`U1:W2`) | `Required Skills` (`V`) | **ScheduleManager** *(Sub-Topic)* | Warm Amber Darker **`#F8C891`** | Own schedule skill requirement |

---

### 4.5 Reserved Managers (`Templates/Unuse/*.xltx`)

| Manager *(Host Swatch)* | Primary Sheet (`+60%` / `+75%`) | Sub-Topic Sheets (`+40%` / `+50%`) & Cross-Manager Overrides |
| :--- | :--- | :--- |
| **CostManager / Costing** (`#D7CEE6` / `#C4B6D9`) | `Monthly Report` (`StaffMasterTable` `A:H` → **`#D7CEE6`**) | `Claims`: `Staff ID`, `Staff Name` → **Staff `#CCD8E7`**; own cols → **`#C4B6D9`**. `Additional Transactions` & `Accounts`: all cols → **`#C4B6D9`**. |
| **SalaryManager** (`#ECD3DC` / `#E3BECA`) | `Montly Report` (`StaffMasterTable`: `Staff ID`, `Staff Name`, `Department` → **Staff `#CCD8E7`**; `Total Day Hours`..`Total Flatten Hours` → **Schedule `#FADBB5`**; `Total Claim` → **Costing `#D7CEE6`**; own cols → **`#ECD3DC`**) | `Leaves` & `Borrowed Money`: `Staff ID`, `Staff Name` → **Staff `#CCD8E7`**; own cols → **`#E3BECA`**. `Holiday and Settings` & `Allowance`: all cols → **`#E3BECA`**. |
| **InventoryManager** (`#D7DEBD` / `#AEBE7B`) | `All Inventories` (`StaffMasterTable` `A:I` → **`#D7DEBD`**) | `Check In` & `Check Out`: `Item ID`, `Item Name` → **`#D7DEBD`**; `Project` → **Project `#DBE1D3`**; `Team` → **Schedule `#FADBB5`**; `Staff ID`, `Staff Name` → **Staff `#CCD8E7`**; own cols → **`#AEBE7B`**. `Purchases`: `Item ID`, `Item Name` → **`#D7DEBD`**; `Staff ID`, `Staff Name` → **Staff `#CCD8E7`**; own cols → **`#AEBE7B`**. `Lost Items`: `Item ID`, `Item Name` → **`#D7DEBD`**; `Last Checked Out Team` → **Schedule `#FADBB5`**; `Last Checked Out Person` → **Staff `#CCD8E7`**; own col → **`#AEBE7B`**. `Use Frequencies`: `Item ID`, `Item Name` → **`#D7DEBD`**; own cols → **`#AEBE7B`**. |

---

## 5. Applying `Paper` Theme to `.xltx` Templates (Migration Checklist)

Perform these steps in Excel (or via surgical OpenXML patching — **never** round-trip `.xltx` through `openpyxl`, which strips custom table styles and `extLst` CF rules):

1. **Switch Theme Colours to `Paper`:** **Page Layout → Colors → Paper** (`xl/theme/theme1.xml` `<a:clrScheme name="Paper">`).
2. **Update `SKSBStyleTableSimple`:** **Table Design → Modify `SKSBStyleTableSimple` → Header Row → Fill** = Host Manager's **Primary Header Fill (`+60%`; `+75%` for `Dark 2`)**, **Font** = Bold `#000000`. Keep `First Row Stripe` = `White, Darker 5%` (`#F2F2F2`) and `Second Row Stripe` = `White` (`#FFFFFF`).
3. **Style Primary vs. Sub-Topic Sheets:**
   - Primary sheet host columns: set **Fill → No Fill** so they inherit `SKSBStyleTableSimple` (`+60%` / `+75%`).
   - Secondary / sub-topic sheets: leave inherited same-manager master keys as **No Fill** (`+60%` / `+75%`), and set all sub-topic host columns to Host **Darker (`+40%` / `+50%`)**.
4. **Apply Cross-Manager Header Overrides:** For every cross-manager column in Section 4, apply the **Source Manager's Primary Header Fill (`+60%`)** directly to the header cell with `#000000` Bold text.
5. **Set Neutral Parameter Bars & Optional Sheet Tabs:** Set top-bar filter headers (`DateFilter`, `Daily Schedule!A1:B1`, etc.) to `White, Darker 15%` (`#D9D9D9`) Bold `#000000`. Optionally set visible sheet tabs to Host Base (`0%`) on primary sheets and Host `-25%` on secondary sheets.
6. **Re-Pin Theme-Linked Conditional Formatting Colours *(CRITICAL)*:** Legacy `Office`-themed `.xltx` workbooks store some CF fills as theme indices (`theme="9"` Accent 6 `+80%` for Soft Green `#D9F2D0` on `Completed` stripes and `ColorScale` midpoints, and `theme="2"` Light 2 for Light Gray `#E8E8E8` on `Inactive/Resigned` stripes). Switching to **Paper** turns `Accent 6` into **Steel Blue** (`#E6ECF3`) and `Light 2` into **Pale Cream** (`#FEFAC9`). Always re-pin:
   - `Completed` stripes & `ColorScale` green midpoints → direct RGB **`#D9F2D0`**.
   - `Inactive / Resigned / Cancelled` stripes → `White, Darker 5%` (`#F2F2F2`) or direct RGB **`#E8E8E8`** (never `Light 2`).

---

## 6. Quick Decision Matrix & Pitfalls to Avoid

### Quick Decision Matrix

| Header / Surface Role | Fill Swatch | Font |
| :--- | :--- | :--- |
| **Host manager primary/master sheet column** | `No Fill` *(inherits `SKSBStyleTableSimple` at Host `+60%` / `+75%`)* | `#000000` Bold |
| **Same-manager inherited master key on a child sheet** | `No Fill` *(Host Primary `+60%` / `+75%`)* | `#000000` Bold |
| **Host manager secondary sheet / sub-topic column** | **Host Darker (`+40%` / `+50%`)** *(direct header cell fill)* | `#000000` Bold |
| **Column derived from another manager (FK, lookup, sync, rollup)** | **Source Manager's Primary (`+60%` / `+75%`)** *(direct header cell fill)* | `#000000` Bold |
| **Multi-source column including Host + External** | Host tier (`+60%` on primary / `+40%` on sub-topic) | `#000000` Bold |
| **Host-calculated composite formula across managers** | Host tier (`+60%` on primary / `+40%` on sub-topic) | `#000000` Bold |
| **Top-bar filter / parameter headers (`DateFilter`, `A1:B1`)** | Neutral `White, Darker 15%` (`#D9D9D9`) | `#000000` Bold |
| **Visible sheet tab (`tabColor`)** | Primary: Host **Base (`0%`)** / Sub-Topic: Host **`-25%`** | *(Excel automatic)* |
| **Table data rows & hidden `Lookups` / `Data` sheets** | `#FFFFFF` / `#F2F2F2` stripes *(never manager fill)* | `#000000` Regular |

### Pitfalls to Avoid

1. **Forgetting the CF theme-shift when changing `Office` → `Paper`:** Always re-pin CF `Completed` greens (`#D9F2D0`) and `Inactive` grays (`#F2F2F2` / `#E8E8E8`) to direct RGB / `White` shades so switching to `Paper` does not turn green heatmaps Steel Blue or gray rows Pale Cream.
2. **Using `+80%` tints for Paper headers:** Unlike the dark `Office` theme, Paper accents at `+80%` (`#EDF0E9`, `#FAF2D4`, `#E6ECF3`) are nearly white and wash out next to `#F2F2F2` zebra stripes. Always use **`+60%`** for primary headers and **`+40%`** for secondary/sub-topic sheets (`+75%` / `+50%` for `Dark 2`).
3. **Mixing non-Paper standard colours (`#FFFF00`, `#FF9900`, `#FFCCCC`) on child tables:** Keep child/sub-tables strictly within their manager's Paper Darker shade (`+40%` / `+50%`); `#FFCCCC` and `#FFFFCC` are reserved for Conditional Formatting status stripes (`conditional-formatting.md`).
4. **Darkening cross-manager columns on secondary sheets:** Cross-manager columns always keep the **source manager's canonical Primary (`+60%`)** fill even when placed inside a `+40%` secondary sheet.
5. **Round-tripping `.xltx` through `openpyxl`:** Always edit `.xltx` templates either in Excel or via surgical OpenXML ZIP/XML patching so `SKSBStyleTableSimple`, `extLst` CF rules, and theme metadata remain intact.
