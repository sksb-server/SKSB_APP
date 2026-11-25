---
trigger: model_decision
description: >-
  SKSB_App Excel template colour scheme standard based on the Microsoft Office
  Paper theme. Load whenever designing, styling, auditing, or documenting
  `.xltx` workbook themes, manager table headers, or cross-manager derived
  column header colours.
---

# SKSB_App Template Colour Scheme Standard (`Paper` Theme)

Authoritative colour standard for all SKSB_App Excel templates (`Templates/*.xltx` and reserved `Templates/Unuse/*.xltx`), built on Microsoft Office's **Paper** theme (`Page Layout → Colors → Paper`, `Document Themes 16\Theme Colors\Paper.xml`).

---

## 1. Official Microsoft Office `Paper` Colour Palette

Paper uses muted, low-glare editorial tones: **`+60%`** is the primary table header fill with `#000000` Bold text (WCAG AAA ≥ `13.9:1`), **`+40%`** is the darker shade for secondary sheets / sub-topics of the same manager (WCAG AAA ≥ `11.0:1`), and **Base (`0%`)** serves sheet tabs and UI badges.

| Slot | Paper Name | Base (`0%`) | `+80%` *(Soft)* | **`+60%` *(Primary)*** | **`+40%` *(Sub-Topic)*** | `-25%` *(Deep)* | `-50%` *(Darkest)* |
| :--- | :--- | :---: | :---: | :---: | :---: | :---: | :---: |
| **Dark 1 (`dk1`)** | Black *(Text)* | `#000000` | `#CCCCCC` | `#999999` | `#666666` | `#000000` | `#000000` |
| **Light 1 (`lt1`)** | White *(Bg)* | `#FFFFFF` | `#FFFFFF` | `#FFFFFF` | `#FFFFFF` | `#BFBFBF` | `#808080` |
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

> **`Dark 2` (`dk2`) stops:** Because `#444D26` is darker than the accents, Excel generates `+90%`/`+75%`/`+50%`/`-25%`/`-50%`. Its **`+75%` (`#D7DEBD`)** and **`+50%` (`#AEBE7B`)** match the accents' `+60%` and `+40%` stops.

---

## 2. Manager Colour Assignments (One Paper Hue per Manager)

All header fonts are **`#000000` Bold** (contrast shown as Primary / Sub-Topic).

| Manager *(Status)* | Template (`Templates/`) | Paper Slot & Hue | Primary (`SKSBStyleTableSimple`) | Sub-Topic Darker | Tab (`0%`) | Contrast |
| :--- | :--- | :--- | :---: | :---: | :---: | :---: |
| **StaffManager** *(Active)* | `Staff_Manager_Template.xltx` | **Accent 6** — Steel Blue | `+60%` **`#CCD8E7`** | `+40%` **`#B3C5DA`** | `#809EC2` | `14.5:1 / 11.9:1` |
| **VehicleManager** *(Active)* | `Vehicle_Manager_Template.xltx` | **Accent 3** — Mustard Gold | `+60%` **`#F5E4A9`** | `+40%` **`#F1D77F`** | `#E7BC29` | `16.6:1 / 14.8:1` |
| **ProjectManager** *(Active)* | `Project_Manager_Template.xltx` | **Accent 1** — Sage Green | `+60%` **`#DBE1D3`** | `+40%` **`#C9D3BE`** | `#A5B592` | `15.7:1 / 13.5:1` |
| **ScheduleManager** *(Active)* | `Schedule_Manager_Template.xltx` | **Accent 2** — Warm Amber | `+60%` **`#FADBB5`** | `+40%` **`#F8C891`** | `#F3A447` | `15.9:1 / 13.7:1` |
| **CostManager** *(Reserved)* | `Unuse/Cost_Manager_Template.xltx` | **Accent 5** — Muted Lavender | `+60%` **`#D7CEE6`** | `+40%` **`#C4B6D9`** | `#9C85C0` | `13.9:1 / 11.0:1` |
| **SalaryManager** *(Reserved)* | `Unuse/Salary_Manager_Template.xltx` | **Accent 4** — Dusty Rose | `+60%` **`#ECD3DC`** | `+40%` **`#E3BECA`** | `#D092A7` | `14.9:1 / 12.5:1` |
| **InventoryManager** *(Reserved)* | `Unuse/Inventory_Manager_Template.xltx` | **Dark 2** — Dark Olive | `+75%` **`#D7DEBD`** | `+50%` **`#AEBE7B`** | `#444D26` | `15.1:1 / 10.5:1` |

---

## 3. Core Principles & Styling Rules

1. **Table Baseline (`SKSBStyleTableSimple`):**
   - `headerRow`: Host Primary (`+60%`; `+75%` for `Dark 2`), `#000000` Bold.
   - `firstRowStripe`: `White, Darker 5%` (`#F2F2F2`); `secondRowStripe`: `#FFFFFF`.
   - Leave host-owned master columns **No Fill** to inherit `SKSBStyleTableSimple`.

2. **Darker Shade for Secondary Sheets & Sub-Topics (Mandatory):**
   - **Primary / Master Sheets (`+60%` / `+75%`, `No Fill`):** `All Staffs`, `All Vehicles`, `iJob/PVT/MWO Projects`, `Daily Schedule`, `Monthly Report`, `All Inventories`.
   - **Secondary Sheets, Child Tables & Sub-Topics (`+40%` / `+50%`):** Apply Host Darker Fill to all host-owned sub-topic columns.
     - *Inherited same-manager master keys* (`Staff ID`/`Staff Name`, `Vehicle Plate`/`Vehicle Type`, `Item ID`/`Item Name`) stay **Primary (`No Fill`)**.
     - *Independent supporting directories / sub-topic tables* (`Workshops`, `Companies`, `Project Summary`, `Daily Attendances`, `Attendance Entries`, `Preference`, `SkillsTable`, `TeamTable`, `FuelCardsTable`, `InsuranceCompaniesTable`, `Claims`, `Accounts`, `Leaves`, `Borrowed Money`, `Allowance`, `Holiday and Settings`) take **Darker (`+40%` / `+50%`)** on all host columns.
   - Never switch hues or use ad-hoc standard colours (`#FFFF00`, `#FF9900`, `#FFCCCC`) on sub-tables.

3. **Cross-Manager Derived Column Headers (Mandatory):**
   - Any column whose value, lookup list, FK, two-way sync, or rolled-up metric derives from another manager takes the **source manager's canonical Primary Fill (`+60%` / `+75%`)** + `#000000` Bold — **even on `+40%` secondary sheets** (e.g., `Staff ID` is `#CCD8E7` everywhere).
   - Covers: (a) FKs & synced `[DataList]` columns (`Staff ID`, `Staff Name`, `PIC`, `Team Lead`, `Project`, `Project ID`, `Vehicle Plate`, `Vehicle Type`), (b) synced read-only attributes (`Licensed Skills`, `Driven By`, `Used At`), and (c) cross-manager rollups (`Days Worked`, `Total Regular Shifts`, `Total Hours`, `Total Odometer`, `Fuel Cost`, `Vehicles Used`).
   - *Intra-manager lookups* (`Company ID`/`Company Name` in ProjectManager, `Workshop Name`/`Vehicle Plate` in VehicleManager, `Team` in ScheduleManager) follow Rule 2.

4. **Multi-Source & Composite Tie-Breaker:**
   - **Host + External** (`FuelCardsTable[Card Owner ID/Name]` accepting `AllVehiclesTable` or `StaffIDList`): keep **Host tier** (`+40%` on secondary table).
   - **Composite cross-manager formulas** (`Actual Project Cost`, `Profit / Margin`): keep **Host tier**.
   - **Reserved-manager inputs** (`Labour Cost`, `Claim Cost`, `Claims Made`, `Pending Claims`, `Is Holiday`, `Total Claim`): use reserved source's Primary swatch (**Costing `#D7CEE6`** or **Salary `#ECD3DC`**).

5. **Neutral Chrome, Tabs, Data Rows & C# Boundary:**
   - **Top-bar filter / parameter headers** (`DateFilter`, `Date Start`, `Date End`, `Display Range`, `Daily Schedule!A1:B1`, `DateFilterTableWorkshops` `A1:C2`, etc.): `White, Darker 15%` (`#D9D9D9`), `#000000` Bold.
   - **Sheet tabs (`tabColor`):** Host Base (`0%`) on primary sheets, `-25%` on secondary sheets (or uniform `0%`); hidden `Lookups` & `Data` tabs uncoloured.
   - **Hidden / technical sheets (`Lookups`, `Data`):** Unstyled (`No Fill`, `#000000` regular).
   - **Data cells:** `#FFFFFF` / `#F2F2F2` stripes only — never solid manager fills so `conditional-formatting.md` rules remain legible.
   - **C# boundary:** C# never authors or overwrites header fills — styling lives in `.xltx` (except `ScheduleBoardMapper.PrepareColumns` cloning prototype slots `B4:D18` / `B20:D34` in `Schedule_Manager_Template.xltx`).

---

## 4. Complete Cross-Manager & Sub-Topic Header Colour Mapping

Canonical Primary swatches (`#000000` Bold): **Staff** `#CCD8E7` · **Vehicle** `#F5E4A9` · **Project** `#DBE1D3` · **Schedule** `#FADBB5` · **Costing** `#D7CEE6` · **Salary** `#ECD3DC` · **Inventory** `#D7DEBD`.

### 4.1 `StaffManager` (`Staff_Manager_Template.xltx`) — Host: `#CCD8E7` / `#B3C5DA`

Root people directory — **no cross-manager derived columns**.

| Sheet *(Role)* | Table *(Range)* | Primary `No Fill` (`#CCD8E7`) | Sub-Topic (`#B3C5DA`) |
| :--- | :--- | :--- | :--- |
| `All Staffs` *(Primary)* | `AllStaffsTable` (`A1:J2`) | All (`A:J`) | — |
| `Vehicle Licenses` *(Sub)* | `VehicleLicensesTable` (`A1:H3`) | `Staff ID` (`B`), `Staff Name` (`C`) | `Vehicle License ID` (`A`), `Vehicle License Class` (`D`), `Renew Date` (`E`), `Expiry Date` (`F`), `Status` (`G`), `Cost` (`H`) |
| `Skill Licenses` *(Sub)* | `SkillLicensesTable` (`A1:H3`), `SkillLicenseTypesTable` (`J1:K2`) | `Staff ID` (`B`), `Staff Name` (`C`) | `Skill License ID` (`A`), `Skill License Type` (`D`), `Renew Date` (`E`), `Expiry Date` (`F`), `Status` (`G`), `Cost` (`H`), `Skill License Types` (`J`), `Display Icon` (`K`) |

### 4.2 `VehicleManager` (`Vehicle_Manager_Template.xltx`) — Host: `#F5E4A9` / `#F1D77F`

| Sheet *(Role)* | Table *(Range)* | Primary `No Fill` (`#F5E4A9`) | Sub-Topic (`#F1D77F`) | Cross-Manager Overrides *(Rule)* |
| :--- | :--- | :--- | :--- | :--- |
| `All Vehicles` *(Primary)* | `AllVehiclesTable` (`A4:Z5`) | `A:P`, `S:Z` *(own & intra-manager rollups)* | — | **Staff**: `Driven By` (`Q`); **Project**: `Used At` (`R`) |
| `Odometers` *(Sub)* | `OdometersTable` (`A1:U2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | `Date` (`A`), `Odometer Before`..`Card Owner` (`D:J`), `Station Name` (`M`), `Fuel Type`..`Remarks` (`O:U`) | **Staff**: `Staff ID` (`K`, `StaffIDList`), `Staff Name` (`L`, `StaffShortNameList`); **Project**: `Project` (`N`, `ProjectIDList`) |
| `Odometers` *(Sub)* | `FuelCardsTable` (`W1:Z2`), `FuelCompanyNamesTable` (`AB1:AB2`) | — | All (`W:Z`, `AB`; Rule 4 keeps `Card Owner ID`/`Name` Host) | — |
| `Maintenances` *(Sub)* | `MaintenancesTable` (`A1:Y2`) | `Vehicle Plate` (`A`), `Vehicle Type` (`B`) | `Maintenance ID`..`Parts` (`C:Y`) | — |
| `Workshops` *(Sub)* | `WorkshopsTable` (`A4:I5`) | — | All (`A:I`; `DateFilterTableWorkshops` `A1:C2` = `#D9D9D9`) | — |
| `Permits`, `Road Taxes`, `Inspections` *(Sub)* | `PermitsTable`, `RoadTaxesTable`, `InspectionsTable` | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | All other (`A`, `D:...`) | — |
| `Insurances` *(Sub)* | `InsurancesTable` (`A1:M2`), `InsuranceCompaniesTable` (`O1:U2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | `Renewal Date` (`A`), `Insurance Renewal ID`..`Remarks` (`D:M`), `Insurance Company`..`Vehicle Plates Under Agent` (`O:U`) | — |
| `Insurance Claims` *(Sub)* | `InsuranceClaimsTable` (`A1:L2`) | `Vehicle Plate` (`B`), `Vehicle Type` (`C`) | `Claim Date` (`A`), `Insurance Claim ID` (`D`), `Settlement Date` (`E`), `Maintenance ID`..`Remarks` (`H:L`) | **Staff**: `Staff ID` (`F`, `StaffIDList`), `Staff Name` (`G`, `StaffShortNameList`) |

### 4.3 `ProjectManager` (`Project_Manager_Template.xltx`) — Host: `#DBE1D3` / `#C9D3BE`

> **Column-Order Note:** `iJobProjectsTable` puts `Labour Cost` (`X`), `Total Odometer` (`Y`), `Fuel Cost` (`Z`); `PVTProjectsTable` & `MWOProjectsTable` put `Total Odometer` (`X`), `Fuel Cost` (`Y`), `Labour Cost` (`Z`). Always match by **Header Name**.

| Sheet *(Role)* | Table *(Range)* | Host Columns | Cross-Manager Overrides *(Rule)* |
| :--- | :--- | :--- | :--- |
| `iJob Projects`, `PVT Projects`, `MWO Projects` *(Primary)* | `iJobProjectsTable`, `PVTProjectsTable`, `MWOProjectsTable` (`A4:AC5`) | **Primary `No Fill`**: `Project Code`..`Company Name` (`A:G`), `District`..`Remarks` (`J:R`), `Actual Project Cost` (`AB`), `Profit / Margin` (`AC`) | **Staff**: `Staff ID` (`H`, `StaffIDList`), `PIC` (`I`, `StaffShortNameList`)<br>**Schedule**: `Days Worked` (`S`), `Total Regular Shifts` (`T`), `Total Hours` (`U`), `Overtime Shifts` (`V`), `Total Overtime Hours` (`W`)<br>**Vehicle**: `Total Odometer` (`Y`/`X`), `Fuel Cost` (`Z`/`Y`) *(from `OdometersTable`)*<br>**Salary**: `Labour Cost` (`X`/`Z`) · **Costing**: `Claim Cost` (`AA`) |
| `Project Summary` *(Sub)* | `ProjectSummaryTable` (`A4:R5`) | **Sub-Topic (`#C9D3BE`)**: `Date` (`A`) | **Staff**: `Team Lead` (`B`, `StaffNameList`)<br>**Schedule**: `Time Start` (`C`), `Time End` (`D`), `Total Hours` (`E`), `Workers` (`F`), `Regular Shifts` (`G`), `Overtime Shifts` (`H`), `Overtime Hours` (`I`), `Additional OT Hours` (`J`)<br>**Salary**: `Is Holiday` (`K`), `Labour Cost` (`L`)<br>**Vehicle**: `Vehicles Used` (`M`), `Vehicle Types` (`N`), `Odometer Used` (`O`), `Fuel Cost` (`P`)<br>**Costing**: `Claims Made` (`Q`), `Pending Claims` (`R`) |
| `Companies` *(Sub)* | `CompaniesTable` (`A1:C2`) | **Sub-Topic (`#C9D3BE`)**: `Company ID` (`A`), `Legal Company Name` (`B`), `Short Company Name` (`C`) | — |

### 4.4 `ScheduleManager` (`Schedule_Manager_Template.xltx`) — Host: `#FADBB5` / `#F8C891`

| Sheet *(Role)* | Table / Region *(Range)* | Host Columns | Cross-Manager / Chrome Overrides *(Rule)* |
| :--- | :--- | :--- | :--- |
| `Daily Schedule` *(Primary)* | Top Bar (`A1:B1`), Prototype (`B4:D4`, `B20:D20`), Left Rails (`A4:A34`) | **Primary (`#FADBB5`)**: `Premix Hole Amount`, `Project start time - end time` (`C4:D4`, `C20:D20`, cloned by `PrepareColumns`); `Day Work` (`A4:A14`)<br>**Sub-Band (`#F8C891`)**: `Night Work` (`A20:A30`) | **Neutral `#D9D9D9`**: `Date`, `Day` (`A1:B1`)<br>**Project**: `Project Code` (`B4`, `B20`, `ProjectIDList`)<br>**Vehicle**: `Vehicles` rail (`A15:A18`, `A31:A34`; data rows stay `#FFFFFF`/`#F2F2F2`) |
| `Daily Attendances` *(Sub)* | `DailyAttendanceTable` (`A4:L5`) | **Sub-Topic (`#F8C891`)**: `Status` (`C`), `Clock-In Time` (`D`), `Clock-Out Time` (`E`), `Regular Shift`..`Is Allocated` (`G:L`) | **Staff**: `Staff ID` (`A`, `StaffIDList`), `Staff Name` (`B`, `StaffNameList`)<br>**Project**: `Project Check-Ins` (`F`) |
| `Attendance Entries` *(Sub)* | `AttendanceEntriesTable` (`A1:J2`) | **Sub-Topic (`#F8C891`)**: `Date` (`A`), `Time` (`B`), `Entry Type`..`WhatsApp Picture` (`F:J`) | **Staff**: `Staff ID` (`C`, `StaffIDList`), `Staff Name` (`D`, `StaffNameList`)<br>**Project**: `Project ID` (`E`, `ProjectIDList`) |
| `Preference` *(Sub)* | `WorkerPreferencesTable` (`A1:L2`) | **Sub-Topic (`#F8C891`)**: `Team` (`C`), `Primary Skill` (`D`), `Other Skills` (`F`), `Blacklisted Skills` (`G`), `Preferred Shift`..`Rest Day Preference` (`J:L`) | **Staff**: `Staff ID` (`A`), `Staff Name` (`B`, `AllFieldStaffIDList`/`NameList`), `Licensed Skills` (`E`, synced from `StaffManager`)<br>**Vehicle**: `Preferred Vehicle Type` (`H`, `VehicleTypeList`), `Dedicated Vehicle Plate` (`I`, `VehiclePlateList`) |
| `Preference` *(Sub)* | `TeamTable` (`N1:P2`), `SkillsTable` (`R1:S2`) | **Sub-Topic (`#F8C891`)**: `Team Name` (`N`), `Skills` (`R`), `Display Icons` (`S`) | **Staff**: `Team Leader` (`O`, `AllFieldStaffNameList`)<br>**Vehicle**: `Default Vehicles` (`P`, `VehiclePlateList`) |
| `Preference` *(Sub)* | `ProjectRequirementsTable` (`U1:W2`) | **Sub-Topic (`#F8C891`)**: `Required Skills` (`V`) | **Project**: `Project Type` (`U`, `ProjectTypeList`)<br>**Vehicle**: `Required Vehicles` (`W`, `VehicleTypeList`) |

### 4.5 Reserved Managers (`Templates/Unuse/*.xltx`)

| Manager *(Primary / Sub)* | Primary Sheet (`+60%` / `+75%`) | Sub-Topic Sheets (`+40%` / `+50%`) & Cross-Manager Overrides |
| :--- | :--- | :--- |
| **CostManager** (`#D7CEE6` / `#C4B6D9`) | `Monthly Report` (`StaffMasterTable` `A:H` → `#D7CEE6`) | `Claims`: `Staff ID`, `Staff Name` → **Staff**; own → `#C4B6D9`. `Additional Transactions` & `Accounts`: all → `#C4B6D9`. |
| **SalaryManager** (`#ECD3DC` / `#E3BECA`) | `Montly Report` (`StaffMasterTable`: `Staff ID`, `Staff Name`, `Department` → **Staff**; `Total Day Hours`..`Total Flatten Hours` → **Schedule**; `Total Claim` → **Costing**; own → `#ECD3DC`) | `Leaves` & `Borrowed Money`: `Staff ID`, `Staff Name` → **Staff**; own → `#E3BECA`. `Holiday and Settings` & `Allowance`: all → `#E3BECA`. |
| **InventoryManager** (`#D7DEBD` / `#AEBE7B`) | `All Inventories` (`StaffMasterTable` `A:I` → `#D7DEBD`) | `Check In` & `Check Out`: `Item ID`, `Item Name` → `#D7DEBD`; `Project` → **Project**; `Team` → **Schedule**; `Staff ID`, `Staff Name` → **Staff**; own → `#AEBE7B`. `Purchases`: `Item ID`, `Item Name` → `#D7DEBD`; `Staff ID`, `Staff Name` → **Staff**; own → `#AEBE7B`. `Lost Items`: `Item ID`, `Item Name` → `#D7DEBD`; `Last Checked Out Team` → **Schedule**; `Last Checked Out Person` → **Staff**; own → `#AEBE7B`. `Use Frequencies`: `Item ID`, `Item Name` → `#D7DEBD`; own → `#AEBE7B`. |

---

## 5. Applying `Paper` Theme to `.xltx` Templates (Migration Checklist)

Edit in Excel or via surgical OpenXML ZIP/XML patching (**never** round-trip `.xltx` through `openpyxl`, which strips `SKSBStyleTableSimple`, `extLst` CF rules, and theme metadata):

1. **Switch Theme Colours:** **Page Layout → Colors → Paper** (`xl/theme/theme1.xml` `<a:clrScheme name="Paper">`).
2. **Update `SKSBStyleTableSimple`:** `Header Row` = Host Primary (`+60%`; `+75%` for `Dark 2`), Bold `#000000`; `First Row Stripe` = `#F2F2F2` (`White -5%`); `Second Row Stripe` = `#FFFFFF`.
3. **Style Primary vs. Sub-Topic Sheets:** Primary host cols & inherited same-manager master keys → **No Fill**; secondary/sub-topic host cols → Host **Darker (`+40%` / `+50%`)**.
4. **Apply Cross-Manager Overrides:** Colour every Section 4 cross-manager header cell with the **Source Manager's Primary (`+60%`)** + `#000000` Bold.
5. **Neutral Bars & Tabs:** Top-bar filter headers → `#D9D9D9` (`White -15%`) Bold `#000000`; optional visible tabs → Host Base (`0%`) / Secondary (`-25%`).
6. **Re-Pin Theme-Linked CF Colours *(CRITICAL)*:** Legacy `Office`-themed `.xltx` files store some CF fills as theme indices (`theme="9"` Accent 6 `+80%` for `#D9F2D0` on `Completed` stripes / `ColorScale` midpoints; `theme="2"` Light 2 for `#E8E8E8` on `Inactive/Resigned` stripes). Switching to **Paper** turns `Accent 6` into **Steel Blue** (`#E6ECF3`) and `Light 2` into **Pale Cream** (`#FEFAC9`). Re-pin:
   - `Completed` stripes & `ColorScale` green midpoints → direct RGB **`#D9F2D0`**.
   - `Inactive / Resigned / Cancelled` stripes → `White -5%` (`#F2F2F2`) or direct RGB **`#E8E8E8`** (never `Light 2`).

---

## 6. Quick Decision Matrix & Pitfalls

| Surface / Header Role | Fill Swatch | Font |
| :--- | :--- | :--- |
| Host primary col / same-manager inherited master key | `No Fill` *(Host `+60%` / `+75%` via table style)* | `#000000` Bold |
| Host secondary sheet / sub-topic col | **Host Darker (`+40%` / `+50%`)** | `#000000` Bold |
| Cross-manager derived col (FK, lookup, sync, rollup) | **Source Primary (`+60%` / `+75%`)** | `#000000` Bold |
| Multi-source (Host + External) or Host composite formula | Host tier (`+60%` primary / `+40%` sub-topic) | `#000000` Bold |
| Top-bar filter / parameter headers (`DateFilter`, `A1:B1`) | Neutral `White -15%` (`#D9D9D9`) | `#000000` Bold |
| Visible sheet tab (`tabColor`) | Primary: Host `0%` / Sub-Topic: Host `-25%` | *(Auto)* |
| Table data rows & hidden `Lookups` / `Data` sheets | `#FFFFFF` / `#F2F2F2` stripes *(never manager fill)* | `#000000` Regular |

**Pitfalls to Avoid:**
1. **CF theme-shift (`Office` → `Paper`):** Always re-pin CF greens (`#D9F2D0`) and grays (`#F2F2F2` / `#E8E8E8`) per Step 6.
2. **Using `+80%` tints for Paper headers:** Paper `+80%` (`#EDF0E9`, `#FAF2D4`, `#E6ECF3`) washes out next to `#F2F2F2` stripes — use `+60%` / `+40%` (`+75%` / `+50%` for `Dark 2`).
3. **Ad-hoc non-Paper colours on child tables:** Never use `#FFFF00`, `#FF9900`, or `#FFCCCC` on headers (`#FFCCCC` / `#FFFFCC` are reserved for `conditional-formatting.md` status stripes).
4. **Darkening cross-manager headers on secondary sheets:** Always keep the source manager's canonical Primary (`+60%`) fill.
5. **Round-tripping `.xltx` via `openpyxl`:** Always use Excel or surgical OpenXML patching.
