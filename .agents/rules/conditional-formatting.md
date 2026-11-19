---
trigger: model_decision
description: >-
  SKSB_App Excel Conditional Formatting design standard and pattern guide.
  Load whenever designing, auditing, documenting, or generating manual `.xltx`
  conditional formatting rules for any manager workbook, sheet, or table.
---

# SKSB_App Conditional Formatting Standard

Authoritative conditional formatting (CF) conventions derived from `Vehicle_Manager_Template.xltx`, `Staff_Manager_Template.xltx`, and `Project_Manager_Template.xltx`. Follow these patterns whenever designing new tables, adding columns, or specifying manual `.xltx` template steps.

---

## 1. Core Principles & Layering Order

- **Archetype-Row Scope Only:** Define every CF rule on the **single archetype data row** of the Excel `ListObject` (Row `2` for standard tables with headers on Row `1`; Row `5` for sheets with a 3-row top parameter/filter bar and headers on Row `4`). Excel propagates rules downward automatically when `TableMapper.PopulateTableSafe<T>` resizes the table.
- **Additive Composition (`Stop If True = OFF`):** Never enable `stopIfTrue`. Rules are designed to layer cleanly (e.g., a row can simultaneously display a diagonal status stripe, a red validation border, and an expiry icon).
- **Strict Priority Stack (Top = Highest Priority / Lowest Priority Number):**
  1. **Cell-Level Analytics & Format Overrides:** `ColorScale`, `IconSet`, and conditional `NumberFormat` overrides.
  2. **Validation & Data-Integrity Borders:** Date chronology errors, cross-column state mismatches, missing required fields (Hard Red `#FF0000` above Soft Amber `#FFC000`).
  3. **Whole-Row Status Fills:** Pastel `lightUp` diagonal stripes sit at the bottom so borders, icons, and heatmaps remain readable on top.

---

## 2. Canonical Patterns

### Pattern 1 — Whole-Row Status Shading (`lightUp` Diagonal Stripe)
- **Visual Style:** **Never use solid fills** for row status. Always use Excel **`lightUp` pattern fill** (thin diagonal stripes) with a soft pastel foreground color. Active/normal/valid rows have **no fill**.
- **Formula:** Lock the Status column with `$`, leave the row relative:
  - Single value: `=$<StatusCol><Row>="Expired"`
  - Grouped values: `=OR($D5="On_Hold",$D5="Pending")`
- **Scope (`Applies To`):**
  - Default: full table row (e.g., `A2:J2`, `A5:AC5`).
  - **Heatmap Exclusion:** If the table has a dedicated right-hand block of cost/metric `ColorScale` columns (e.g., `AllVehiclesTable` `Q:Z` or `MaintenancesTable` `O:Y`), end the status fill before the heatmap block (`A5:P5`, `A2:N2`) so stripes do not visually clash with gradient cells.
- **Standard Status Palette:**

| Semantic State | Example Status Values | Pattern | Foreground Color |
| :--- | :--- | :--- | :--- |
| **Completed / Positive** | `Completed` | `lightUp` | Soft Green `#D9F2D0` *(Theme Accent 6, `+80%`)* |
| **In Progress / Warning / Paused** | `In_Shop`, `In_Progress`, `On_Hold`, `Pending`, `Suspended` | `lightUp` | Soft Yellow `#FFFFCC` *(or `#FFFF99` in StaffManager)* |
| **Expired / Rejected / Negative** | `Expired`, `Rejected`, `Cancelled` *(Maintenances)* | `lightUp` | Soft Red `#FFCCCC` *(or `#FF9999` in ProjectManager)* |
| **Inactive / Archived / Neutral** | `Decommissioned`, `Cancelled` *(Insurances/Claims/Projects)*, `Resigned`, `Invalid` | `lightUp` | Light Gray `#F2F2F2` *(Theme White `-5%`)*, `#E8E8E8`, or `#D1D1D1` |
| **Active / Valid / Normal** | `Active`, `Valid`, `Open` | *(None)* | Default unshaded table row |

---

### Pattern 2 — Validation & Data-Integrity Borders (Thin Solid Outline, No Fill)
- **Visual Style:** All 4 sides (`top`, `bottom`, `left`, `right`) = **`thin` solid border**, **no fill**.
  - **Hard Error (Required / Logical Violation):** Bright Red `#FF0000`
  - **Soft Warning (Recommended / Missing Optional-but-Expected Field):** Amber/Gold `#FFC000`
  - *(Rare inline numeric exception: direct numeric anomaly like `Current Odometer < Odometer Before` may use **Bold Red Font `#FF0000`** instead of a border.)*
- **Sub-Pattern 2A — Missing Required Fields (Row-Activated):**
  - Always gate on a row-identity check so empty archetype/template rows never turn red.
  - Formula: `=AND(<RowHasIdentity>, ISBLANK(<TopLeftTargetCell>))` *(or `LEN(TRIM(<TopLeftTargetCell>))=0`)*
  - `<RowHasIdentity>` is `$PK<>""`, `OR($PK1<>"", $PK2<>"")`, or `LEN(TRIM(CONCAT($A2:$G2)))>0`.
  - **CRITICAL:** The cell inside `ISBLANK(...)` MUST use a **relative column reference** (e.g., `ISBLANK(A5)` for `sqref="A5:B5 F5 L5:M5"`). Locking `$A5` inside `ISBLANK` breaks multi-column highlighting.
- **Sub-Pattern 2B — Date / Time Chronology Violation:**
  - Always guard against blank **and** the app's computed em-dash placeholder `"—"` (and `ISNUMBER($Start)` when start can be text/empty):
    - End before Start: `=AND($End<>"", $End<>"—", ISNUMBER($Start), $End<$Start)`
    - Expiry/Target on or before Renewal/Start: `=AND($Expiry<>"", $Expiry<>"—", ISNUMBER($Start), $Expiry<=$Start)`
  - Apply only to the offending End/Expiry cell.
- **Sub-Pattern 2C — Cross-Column State Consistency:**
  - Highlight both coupled columns (e.g., `I2:J2` or `F2:G2`):
    - Status requires date: `=AND($J2="Resigned", OR(ISBLANK(I2), I2="—"))`
    - Past expiry still marked Valid: `=AND($F2<>"", TODAY()>$F2, $G2="Valid")`
- **Sub-Pattern 2D — Either-Or Soft Warning (`#FFC000`):**
  - Example (IC or Passport required): `=AND($A2<>"", $E2="", $F2="")` on `E2:F2`.

---

### Pattern 3 — Expiry & Deadline Traffic-Light Icons (`IconSet`)
- **Use For:** Every **Expiry Date** (`Insurance`, `Permit`, `Road Tax`, `Inspection`, `License`) and **Target End Date** column.
- **Configuration:** 3-state Icon Set (`3Symbols2`, or custom mix: `3TrafficLights1` red dot + `3Symbols2` yellow `!` + `3Symbols2` green `✓`), `showValue = true`, `reverse = false`:
  - **Green (`✓`):** `>= Formula` `=EDATE(TODAY(), 1)` *(more than 1 month remaining)*
  - **Yellow (`!`):** `>= Formula` `=TODAY()+1` *(due within 1 month)*
  - **Red (`●` / `✕`):** `< TODAY()+1` *(base `percent 0` — expired or due today)*

---

### Pattern 4 — Cost, Amount & Metric Heatmaps (3-Color `ColorScale`)
- **Use For:** Financial costs, claim amounts, returns, odometer totals, and numeric summary metrics.
- **Rule Granularity:** Create **one `ColorScale` rule per column** so each metric scales independently (exception: homogeneous cost-breakdown blocks like `Maintenances` `O2:Y2 K2` may share one scale).
- **Variant 4A — Zero-Anchored Cost/Amount Scale (Default):**
  - Keeps `0` and low values visually clean (white) while surfacing non-zero activity and highlighting high-cost outliers:
    - **Min (`min` or `percent 0`):** White `#FFFFFF`
    - **Midpoint (`percent 5`):** Soft Green `#D9F2D0` / `#DAF2D0` *(Theme Accent 6, `+80%`)*
    - **Max (`max` or `percent 100`):** Soft Red `#FF9999`
- **Variant 4B — Centered / Symmetric Scale (Median-Anchored):**
  - Used for unit costs or centered metrics (e.g., License `Cost`, `Profit / Margin`):
    - **Min (`min` or `percent 0`):** Soft Green `#D9F2D0`
    - **Midpoint (`percentile 50` or `percent 50`):** White `#FFFFFF`
    - **Max (`max` or `percent 100`):** Soft Red `#FF9999`
- **Variant 4C — 2-Color Intensity Scale:**
  - For simple count/shift intensity (e.g., `Regular Shifts`): `percent 0` White `#FFFFFF` → `percent 100` Soft Red `#FF9999`.

---

### Pattern 5 — Count, Volume, Age & Duration Rating Icons (`IconSet`)
- **Relative Volume / Count Columns (5-level, 20% increments `0 / 20 / 40 / 60 / 80`):**
  - `5Rating` (signal bars) for counts (`Number of Vehicle Fixed`, `Number of Vehicles Under Agent`).
  - `5Quarters` (pie quarters) for deltas (`Odometer Difference`).
- **Fixed-Threshold Domain Scales (Numeric `num` thresholds):**
  - **Asset Age (`Years Old`):** `5Rating` visually reversed (full bars = newest, empty bars = oldest) at `0`, `5`, `8`, `10`, `15` years.
  - **Duration (`Total Hours`):** 3-state custom icons (`3Flags` red / `3TrafficLights1` yellow / `3Symbols2` green) at `< 1`, `1–4`, `>= 4` hours.

---

### Pattern 6 — Conditional Number Format Override (Unit Switching)
- **Use For:** Columns whose display unit depends on a sibling type/category column without altering the underlying numeric value.
- **Example (Heavy Machinery Hours vs. Vehicle km):**
  - Formula: `=OR($B5="Backhoe",$B5="Excavator",$B5="Roller",$B5="Compressor",$B5="Generator",$B5="Machinery")`
  - Format: `#,##0.#\ "hrs"` (applied to `Current Odometer` / `Odometer Before` / `Odometer Difference`).

---

## 3. Quick Decision Matrix (New Table / Column Checklist)

| Column / Table Role | Apply Pattern |
| :--- | :--- |
| **Table has a `Status` / `Contract Status` column** | Pattern 1 (`lightUp` row fill; stop before right-hand cost heatmap block if present) |
| **Required PK / FK / Date fields** | Pattern 2A (Red `#FF0000` thin border, row-activated, relative `ISBLANK(ColRow)`) |
| **Recommended / either-or fields** | Pattern 2D (Amber `#FFC000` thin border, row-activated) |
| **Start/Renewal + End/Expiry date pair** | Pattern 2B (Red `#FF0000` thin border on End/Expiry cell, with `"—"` guard) |
| **Status coupled to a Date (e.g., Resigned/Valid)** | Pattern 2C (Red `#FF0000` thin border on both columns) |
| **Expiry Date or Target End Date** | Pattern 3 (`3Symbols2` / traffic-light icons at `TODAY()+1` and `EDATE(TODAY(),1)`) |
| **Cost / Amount / Return / Usage metric** | Pattern 4A (3-color White → 5% Green `#D9F2D0` → 100% Red `#FF9999`, one rule per column) |
| **Count / Delta / Age / Hours column** | Pattern 5 (`5Rating` / `5Quarters` at 20% steps, or domain-fixed `num` thresholds) |
| **Unit varies by sibling category** | Pattern 6 (Formula-based `NumberFormat` override) |

---

## 4. Pitfalls to Avoid

1. **Locking the column inside `ISBLANK($A2)` on multi-column `sqref`:** Always write `ISBLANK(A2)` (relative column) where `A2` is the top-left cell of `sqref`.
2. **Forgetting the `"—"` em-dash guard on date comparisons:** Computed/placeholder cells in `SKSB_App` display `"—"` (text, which evaluates greater than any Excel date serial number). Always include `$Col<>"—"` (and `ISNUMBER($Start)` when relevant).
3. **Wrong column letter after copy-pasting CF between sheets:** Verify every `$Col` reference against the target table's actual column letters and remove leftover rules targeting out-of-table columns.
4. **Duplicate rules on the same cell:** Check for duplicate `ColorScale` or border rules when cloning sheets.
5. **Using solid fills for row status:** Always use `lightUp` diagonal pattern fills so status rows stay visually distinct from `ColorScale` heatmaps and preserve text contrast.

