---
trigger: model_decision
description: >-
  SKSB_App Excel Conditional Formatting design standard and pattern guide.
  Load whenever designing, auditing, documenting, or generating manual `.xltx`
  conditional formatting rules for any manager workbook, sheet, or table.
---

# SKSB_App Conditional Formatting Standard

Authoritative conditional formatting (CF) conventions derived from `Staff_Manager_Template.xltx`, `Project_Manager_Template.xltx`, `Vehicle_Manager_Template.xltx`, and `Core/03_RuntimeUI/RuleEngine/RuleEngine.Formats.cs`. Follow these patterns whenever designing new tables, adding columns, or specifying manual `.xltx` template steps.

---

## 1. Core Principles & C# Ownership Boundary

- **Archetype-Row Scope Only:** Define every CF rule on the **single archetype data row** of the Excel `ListObject` (Row `2` for standard tables with headers on Row `1`; Row `5` for sheets with a 3-row top parameter/filter bar and headers on Row `4`). Excel propagates rules downward automatically when `TableMapper.PopulateTableSafe<T>` resizes the table.
- **C# Ownership Boundary:**
  - **Baseline Number Formats (`[ColumnFormat]`):** C# (`RuleEngine.Formats.cs`) owns static column number formats (currency, dates, IDs, em-dash zero display). Use CF `NumberFormat` overrides **only** when a cell's display unit switches dynamically based on a sibling column (Pattern 6).
  - **Visual Rules:** C# never deletes or auto-generates CF rules — all visual status shading, validation borders, icons, and heatmaps are authored in `.xltx`.
- **Additive Composition (`Stop If True = OFF`):** Never enable `Stop If True`. Rules are designed to layer cleanly (e.g., a row can simultaneously display a diagonal status stripe, a red validation border, and an expiry icon).
- **Strict Priority Stack (Top = Highest Priority / Lowest Priority Number):**
  1. **Cell-Level Analytics & Format Overrides:** `ColorScale`, `IconSet`, and conditional `NumberFormat` overrides.
  2. **Validation & Data-Integrity Borders:** Date chronology errors, cross-column state mismatches, and missing required fields (Hard Red `#FF0000` above Soft Amber `#FFC000`).
  3. **Whole-Row Status Fills:** Pastel `lightUp` diagonal stripes sit at the bottom so borders, icons, and heatmaps remain readable on top.
- **Em-Dash (`"—"`) & Blank Guards:** Always guard date comparisons against both blank cells and SKSB's computed `"—"` placeholder (`$Col<>"—"`, plus `ISNUMBER($Start)` when start dates can be blank or text).

---

## 2. Canonical Patterns

### Pattern 1 — Whole-Row Status Shading (`lightUp` Diagonal Stripe)
- **Use For:** Any table with a `Status` or `Contract Status` column to make non-active rows visually scannable without obscuring text or heatmaps.
- **Configuration:** `Rule Type: Use a formula` | **Fill Pattern:** `lightUp` (thin diagonal stripe — **never solid fill**) | **`Stop If True`:** `OFF`
- **Applies To:**
  - Default: full table row (e.g., `=$A$2:$J$2`, `=$A$5:$AC$5`).
  - **Heatmap Exclusion:** If the table has a right-hand block of `ColorScale` cost/metric columns (e.g., `AllVehiclesTable` `Q:Z` or `MaintenancesTable` `O:Y`), end the status fill before the heatmap block (`$A$5:$P$5`, `$A$2:$N$2`).
- **Formula:** Lock the Status column with `$`, leave the row relative:
  ```excel
  =$<StatusCol><Row>="Expired"
  =OR($<StatusCol><Row>="On_Hold", $<StatusCol><Row>="Pending")
  ```
- **Standard Status Palette:**

| Semantic State | Example Status Values | Fill Pattern | Foreground Color |
| :--- | :--- | :--- | :--- |
| **Completed / Positive** | `Completed` | `lightUp` | Soft Green `#D9F2D0` *(Theme Accent 6, `+80%`)* |
| **In Progress / Warning / Paused** | `In_Shop`, `In_Progress`, `On_Hold`, `Pending`, `Suspended` | `lightUp` | Soft Yellow `#FFFFCC` *(or `#FFFF99` in StaffManager)* |
| **Expired / Rejected / Negative** | `Expired`, `Rejected`, `Cancelled` *(Maintenances)* | `lightUp` | Soft Red `#FFCCCC` *(or `#FF9999` in ProjectManager)* |
| **Inactive / Archived / Neutral** | `Decommissioned`, `Cancelled` *(Insurances/Claims/Projects)*, `Resigned`, `Invalid` | `lightUp` | Light Gray `#F2F2F2` *(Theme White `-5%`)*, `#E8E8E8`, or `#D1D1D1` |
| **Active / Valid / Normal** | `Active`, `Valid`, `Open` | *(None)* | Default unshaded table row |

---

### Pattern 2 — Validation & Data-Integrity Borders
- **Use For:** Highlighting missing required inputs, impossible date ranges, and cross-column state contradictions at a glance.
- **Configuration:** `Rule Type: Use a formula` | **Border:** All 4 sides = `thin` solid | **Fill:** `None` | **`Stop If True`:** `OFF`
  - **Hard Error (Required / Logical Violation):** Bright Red `#FF0000`
  - **Soft Warning (Recommended / Either-Or Field):** Amber/Gold `#FFC000`
  - *(Rare numeric anomaly exception: e.g., `Current Odometer < Odometer Before` may use **Bold Red Font `#FF0000`** instead of a border.)*
- **Variant 2A — Missing Required Fields (Row-Activated, Red `#FF0000`):**
  - Gate on a row-identity check (`$PK<>""`, `OR($PK1<>"", $PK2<>"")`, or `LEN(TRIM(CONCAT($A2:$G2)))>0`) so empty template rows never turn red.
  - Keep `<TopLeftCell>` **column-relative** (no `$` before the column letter) when targeting multiple columns (e.g., `A5:B5 F5 L5:M5`).
  ```excel
  =AND(<RowHasIdentity>, ISBLANK(<TopLeftCell>))
  ```
- **Variant 2B — Date / Time Chronology Violation (Red `#FF0000`):**
  - Apply only to the offending End/Expiry cell. Pairs with Data Validation Pattern 4A.
  ```excel
  =AND($<End><>"", $<End><>"—", ISNUMBER($<Start>), $<End><$<Start>)
  =AND($<Expiry><>"", $<Expiry><>"—", ISNUMBER($<Start>), $<Expiry><=$<Start>)
  ```
- **Variant 2C — Cross-Column State Consistency (Red `#FF0000`):**
  - Apply to both coupled columns (e.g., `I2:J2` or `F2:G2`):
  ```excel
  =AND($<Status>="Resigned", OR(ISBLANK(<Date>), <Date>="—"))
  =AND($<Expiry><>"", TODAY()>$<Expiry>, $<Status>="Valid")
  ```
- **Variant 2D — Either-Or Soft Warning (Amber `#FFC000`):**
  - Apply to both paired columns (e.g., `E2:F2`). Pairs with Data Validation Pattern 3.
  ```excel
  =AND($<PK><>"", $<ColA>="", $<ColB>="")
  ```

---

### Pattern 3 — Expiry & Deadline Traffic-Light Icons (`IconSet`)
- **Use For:** Every **Expiry Date** (`Insurance`, `Permit`, `Road Tax`, `Inspection`, `License`) and **Target End Date** column.
- **Configuration:** `Rule Type: Format all cells based on their values` | **Style:** 3-state `IconSet` (`3Symbols2`, or custom mix: `3TrafficLights1` red dot + `3Symbols2` yellow `!` + `3Symbols2` green `✓`) | **Show Icon Only:** `OFF` | **Reverse Icon Order:** `OFF`
- **Thresholds:**

| State | Icon | Operator | Type | Formula / Value | Meaning |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Safe** | Green `✓` | `>=` | `Formula` | `=EDATE(TODAY(), 1)` | More than 1 month remaining |
| **Upcoming** | Yellow `!` | `>=` | `Formula` | `=TODAY()+1` | Due within 1 month |
| **Overdue / Due Today** | Red `●` / `✕` | `<` | *(Base)* | `percent 0` | Expired or due today |

---

### Pattern 4 — Cost, Amount & Metric Heatmaps (`ColorScale`)
- **Use For:** Financial costs, claim amounts, returns, odometer totals, and numeric summary metrics.
- **Rule Granularity:** Create **one `ColorScale` rule per column** so each metric scales independently (exception: homogeneous cost-breakdown blocks such as `Maintenances` `O2:Y2 K2` may share one scale).
- **Sub-Variants:**

| Variant | Use For | Min (`min` / `percent 0`) | Midpoint | Max (`max` / `percent 100`) |
| :--- | :--- | :--- | :--- | :--- |
| **4A — Zero-Anchored (Default)** | Costs, amounts, returns, usage (keeps `0` visually clean) | White `#FFFFFF` | `percent 5` → Soft Green `#D9F2D0` | Soft Red `#FF9999` |
| **4B — Centered / Symmetric** | Unit costs, `Profit / Margin`, centered metrics | Soft Green `#D9F2D0` | `percentile 50` (or `percent 50`) → White `#FFFFFF` | Soft Red `#FF9999` |
| **4C — 2-Color Intensity** | Simple counts / shift intensity (e.g., `Regular Shifts`) | White `#FFFFFF` | *(None)* | Soft Red `#FF9999` |

---

### Pattern 5 — Count, Volume, Age & Duration Rating Icons (`IconSet`)
- **Use For:** Visual magnitude indicators on counts, deltas, asset ages, and work durations.
- **Configuration:** `Rule Type: Format all cells based on their values` | **Show Icon Only:** `OFF`
- **Sub-Variants:**
  - **Variant 5A — Relative Volume / Count (5-Level, 20% Steps `0 / 20 / 40 / 60 / 80`):**
    - `5Rating` (signal bars) for counts (`Number of Vehicle Fixed`, `Number of Vehicles Under Agent`).
    - `5Quarters` (pie quarters) for deltas (`Odometer Difference`).
  - **Variant 5B — Fixed-Threshold Domain Scales (`Type: Number`):**
    - **Asset Age (`Years Old`):** `5Rating` reversed (full bars = newest, empty bars = oldest) at `0`, `5`, `8`, `10`, `15` years.
    - **Duration (`Total Hours`):** 3-state custom icons (`3Flags` red / `3TrafficLights1` yellow / `3Symbols2` green) at `< 1`, `1–4`, `>= 4` hours.

---

### Pattern 6 — Conditional Number Format Override (Dynamic Unit Switching)
- **Use For:** Numeric columns whose display unit changes based on a sibling category column without altering the underlying stored number (e.g., heavy machinery hours vs. road vehicle km).
- **Configuration:** `Rule Type: Use a formula` | **Format:** Custom `NumberFormat` only | **`Stop If True`:** `OFF`
- **Formula & Format Example (`Current Odometer` / `Odometer Before` / `Odometer Difference`):**
  ```excel
  =OR($B5="Backhoe", $B5="Excavator", $B5="Roller", $B5="Compressor", $B5="Generator", $B5="Machinery")
  ```
  - **Number Format:** `#,##0.#\ "hrs"`

---

## 3. Quick Decision Matrix (New Table / Column Checklist)

| Column / Table Role | Action / Pattern |
| :--- | :--- |
| **Table has a `Status` / `Contract Status` column** | Pattern 1 (`lightUp` row fill; end before right-hand cost heatmap block if present) |
| **Required PK / FK / Date fields** | Pattern 2A (Red `#FF0000` thin border, row-activated, relative `ISBLANK(<TopLeftCell>)`) |
| **Recommended / either-or fields** | Pattern 2D (Amber `#FFC000` thin border, row-activated; pairs with DV Pattern 3) |
| **Start/Renewal + End/Expiry date pair** | Pattern 2B (Red `#FF0000` thin border on End/Expiry cell, with `"—"` guard; pairs with DV Pattern 4A) |
| **Status coupled to a Date (e.g., `Resigned` / `Valid`)** | Pattern 2C (Red `#FF0000` thin border on both columns) |
| **Expiry Date or Target End Date** | Pattern 3 (`3Symbols2` / traffic-light icons at `TODAY()+1` and `EDATE(TODAY(),1)`) |
| **Cost / Amount / Return / Usage metric** | Pattern 4A (3-color White → 5% Green `#D9F2D0` → 100% Red `#FF9999`, one rule per column) |
| **Unit cost / Profit / Centered metric** | Pattern 4B (3-color Green `#D9F2D0` → 50% White `#FFFFFF` → 100% Red `#FF9999`) |
| **Count / Delta / Age / Hours column** | Pattern 5 (`5Rating` / `5Quarters` at 20% steps, or domain-fixed `Number` thresholds) |
| **Unit varies by sibling category** | Pattern 6 (Formula-based `NumberFormat` override) |

---

## 4. Pitfalls to Avoid

1. **Locking the column inside `ISBLANK($A2)` on multi-column `sqref`:** Always use a relative column (`ISBLANK(A2)`) where `A2` is the top-left cell of `Applies To`, otherwise Excel evaluates only column `A` for every cell in the range.
2. **Forgetting the `"—"` em-dash guard on date comparisons:** Computed/placeholder cells hold `"—"` (text, which evaluates greater than any Excel date serial number). Always include `$Col<>"—"` (and `ISNUMBER($Start)` when start dates can be blank/text).
3. **Enabling `Stop If True`:** Leave `Stop If True = OFF` on every rule so status shading, validation borders, and icons compose cleanly on the same cell.
4. **Using solid fills for row status:** Always use `lightUp` diagonal pattern fills so status rows stay visually distinct from `ColorScale` heatmaps and preserve text contrast.
5. **Wrong column letters or duplicate rules after cloning sheets:** Verify every `$Col` reference against the target table's actual headers, remove rules pointing outside the table, and check for duplicate `ColorScale` or border rules.
