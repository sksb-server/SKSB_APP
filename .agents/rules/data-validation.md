---
trigger: model_decision
description: >-
  SKSB_App Excel Data Validation design standard and pattern guide.
  Load whenever designing, auditing, documenting, or generating manual `.xltx`
  data validation rules for any manager workbook, sheet, or table.
---

# SKSB_App Data Validation Standard

Authoritative data validation (DV) conventions derived from `Staff_Manager_Template.xltx`, `Project_Manager_Template.xltx`, `Vehicle_Manager_Template.xltx`, and `Core/03_RuntimeUI/RuleEngine/RuleEngine.Registry.cs`. Follow these patterns whenever designing new tables, adding columns, or specifying manual `.xltx` template steps.

---

## 1. Core Principles & C# Ownership Boundary

- **Archetype-Row Scope Only:** Define every table-column DV rule on the **single archetype data row** of the Excel `ListObject` (Row `2` for standard tables with headers on Row `1`; Row `5` for sheets with a 3-row top parameter/filter bar and headers on Row `4`). Excel propagates rules downward automatically when `TableMapper.PopulateTableSafe<T>` resizes the table.
- **Never Duplicate C#-Owned Validations:**
  - **Table Dropdowns (`[ColumnRule(sources)]`):** **Do NOT add native Excel `List` DV** to table columns governed by C# `[ColumnRule]`. On workbook open, `RuleRegistry.ApplyDropdownStylesToWorkbook` deletes native Excel DV (`targetRange.Validation.Delete()`) on every `rule.HasDropdown` column and replaces it with SKSB's floating dropdown + save-time validation (`RulesExtractor`).
  - **C#-Calculated Columns (`[ColumnCalculate(..., Always)]` / `[ColumnRollup(..., Always)]`):** **Do NOT author lock DVs manually.** `RuleRegistry.ApplyCalculatedColumnLock` automatically applies and maintains the `=FALSE` lock + formula tooltip on workbook open.
- **Alert Style Severity (`Stop` vs. `Warning`):**
  - **`Stop` (Hard Block — default):** Use for strict invariants that can never be valid — mandatory identity fields, impossible date chronology, locked/unused/formula cells (`=FALSE`), and top-bar parameter constraints.
  - **`Warning` (Soft Override):** Use for format conventions or either-or recommendations where legitimate real-world exceptions exist (e.g., foreign ID formats, partial entry order). Users see a Continue/Cancel prompt rather than a hard dead-end.
- **No Excel Sheet Protection:** Never lock cells with Worksheet Protection (it breaks `ListObject` auto-expansion and C# COM writes). Use `Custom` `=FALSE` Data Validation instead — it blocks interactive typing while allowing C# `.Value2` writes to pass through cleanly.
- **`Ignore Blank` & Em-Dash (`"—"`) Rules:**
  - Set **`Ignore Blank = OFF`** (`allowBlank=0`) on mandatory non-empty fields and locked calculated columns (`=FALSE`) so clearing a cell does not silently bypass validation.
  - Set **`Ignore Blank = ON`** (`allowBlank=1`) on optional fields, format checks, either-or checks, and date chronology guards.
  - Always guard table date formulas against SKSB's computed em-dash placeholder `"—"` (`OR(ISBLANK(<Cell>), <Cell>="—", ...)`).

---

## 2. Canonical Patterns

### Pattern 1 — Mandatory Non-Empty Text (Free-Text Identity Fields)
- **Use For:** Required human-readable name/title columns that do not use a `[ColumnRule]` dropdown (e.g., `Staff Name`).
- **Configuration:** `Allow: Custom` | **Alert:** `Stop` | **`Ignore Blank`:** `OFF`
- **Formula:**
  ```excel
  =LEN(TRIM(<Cell>))>0
  ```
- **Error Alert:**
  - **Title:** `Value not valid`
  - **Message:** `<Field Name> is mandatory and cannot be empty.`

---

### Pattern 2 — Structured Format Check (Soft Warning)
- **Use For:** Free-text identifiers with an expected format where rare real-world exceptions may require user override (e.g., 12-digit Malaysian NRIC, registration/reference codes).
- **Configuration:** `Allow: Custom` | **Alert:** `Warning` | **`Ignore Blank`:** `ON`
- **Formula:**
  ```excel
  =OR(ISBLANK(<Cell>), <FormatPredicate>)
  ```
  - *Example (12-digit numeric NRIC, no dashes):* `=OR(ISBLANK(E2), AND(ISNUMBER(VALUE(E2)), LEN(E2)=12))`
- **Error Alert:**
  - **Title:** `Value entered incorrectly`
  - **Message:** Concise description of expected format (e.g., `Malaysian NRIC should be 12 digits (numeric only, no dashes).`).

---

### Pattern 3 — Either-Or / Cross-Column Presence (Soft Warning)
- **Use For:** Paired optional columns where at least one must be populated (e.g., `IC Number` or `Passport Number`). Pairs visually with Conditional Formatting Pattern 2D (Amber `#FFC000` border).
- **Configuration:** `Allow: Custom` | **Alert:** `Warning` | **`Ignore Blank`:** `ON`
- **Formula:**
  ```excel
  =OR(<ColA><Row><>"", <ColB><Row><>"")
  ```
- **Error Alert:**
  - **Title:** `Value should not be empty`
  - **Message:** `<Entity> should have at least <Field A> or <Field B> recorded.`

---

### Pattern 4 — Date Chronology Guard
- **Use For:** End / Resign / Expiry date columns paired with a Start / Join / Renewal date column. Pairs visually with Conditional Formatting Pattern 2B (Red `#FF0000` border).
- **Variant 4A — Table Column (with `"—"` Em-Dash Guard):**
  - **Configuration:** `Allow: Custom` | **Alert:** `Stop` | **`Ignore Blank`:** `ON`
  - **Formula:**
    - Non-decreasing (`>=`, e.g., Resign Date vs. Join Date): `=OR(ISBLANK(<EndCell>), <EndCell>="—", <EndCell>>=<StartCell>)`
    - Strictly later (`>`, e.g., Expiry Date vs. Renewal Date): `=OR(ISBLANK(<EndCell>), <EndCell>="—", <EndCell>><StartCell>)`
  - **Error Alert:**
    - **Title:** `Value not valid`
    - **Message:** `<End Date> cannot be earlier than <Start Date>.`
- **Variant 4B — Top-Bar Date Parameter Cell (Outside `ListObject`, e.g., `B2` End Date > `A2` Start Date):**
  - **Configuration:** `Allow: Date` | **Data:** `greater than` *(or `greater than or equal to`)* | **Alert:** `Stop` | **`Ignore Blank`:** `ON`
  - **Start Date:** `=<StartParamCell>` (e.g., `=A2`)

---

### Pattern 5 — Read-Only Cell Lock (`=FALSE`, Without Sheet Protection)
- **Use For:** Blocking manual edits on calculated, reserved, or structural cells without enabling Sheet Protection. Always uses `Allow: Custom`, **Formula:** `=FALSE`, **Alert:** `Stop`.
- **Sub-Variants:**

| Variant | Ownership | `Ignore Blank` | Input Message (`Prompt`) | Error Title & Message |
| :--- | :--- | :--- | :--- | :--- |
| **5A — C#-Calculated Column** (`[ColumnCalculate]` / `[ColumnRollup]`, `Always`) | **Auto-applied by C#** (`RuleRegistry.ApplyCalculatedColumnLock`) | `OFF` | **Title:** `Formula Evaluated Cell (Locked)`<br>**Msg:** `<DisplayFormula>` *(<=220 chars)* | **Title:** `Locked Calculated Column`<br>**Msg:** `'<Header>' is automatically calculated (<DisplayFormula>) and cannot be edited manually.` *(<=255 chars)* |
| **5B — Legacy Excel-Formula Column** *(not yet migrated to C#)* | Manual `.xltx` | `ON` | *(Optional)* | **Title:** `Formula Cell`<br>**Msg:** `This field is generated through formula, and cannot be edited.` |
| **5C — Unused / Reserved Column** *(future tracking)* | Manual `.xltx` | `ON` | *(None)* | **Title:** `Field Unused`<br>**Msg:** `These fields are reserved for future tracking. Leave empty for now.` |
| **5D — Protected Header Cell** *(optional)* | Manual `.xltx` | `ON` | *(None)* | **Title:** `Header not editable`<br>**Msg:** *(Optional)* |

---

### Pattern 6 — Top-Bar Parameter Dropdown (Outside `ListObject`)
- **Use For:** Sheet-level filter/selector cells in Row 2 (e.g., `C2` `Display Range` backed by `DateFilterScopeList`) that live outside any `ListObject` and are not managed by `[ColumnRule]`.
- **Configuration:** `Allow: List` | **In-cell dropdown:** `ON` | **Alert:** `Stop` | **`Ignore Blank`:** `ON`
- **Source:** `=<NamedRange>` (e.g., `=DateFilterScopeList`)
- **Error Alert:**
  - **Title:** `Value not valid`
  - **Message:** `Please select an option from the Dropdown menu.`

---

## 3. Quick Decision Matrix (New Table / Column Checklist)

| Column / Cell Role | Action / Pattern |
| :--- | :--- |
| **Table column with C# `[ColumnRule(sources)]` dropdown** | **No `.xltx` DV** — C# `RuleRegistry` deletes native DV and owns the dropdown + save validation |
| **Table column with `[ColumnCalculate]` / `[ColumnRollup]` (`Always`)** | **No manual `.xltx` DV** — C# `ApplyCalculatedColumnLock` auto-applies Pattern 5A on open |
| **Mandatory free-text column (no dropdown)** | Pattern 1 (`=LEN(TRIM(<Cell>))>0`, `Stop`, `Ignore Blank = OFF`) |
| **Free-text column with convention/format (NRIC, code, etc.)** | Pattern 2 (`=OR(ISBLANK(<Cell>), <Predicate>)`, `Warning`, `Ignore Blank = ON`) |
| **Paired either-or optional columns** | Pattern 3 (`=OR(<A><Row><>"", <B><Row><>"")`, `Warning`, `Ignore Blank = ON`) |
| **Editable End / Resign / Expiry date paired with Start date** | Pattern 4A (`=OR(ISBLANK(<End>), <End>="—", <End>>=<Start>)`, `Stop`, `Ignore Blank = ON`) |
| **Top-bar End Date parameter cell (`B2` vs. `A2`)** | Pattern 4B (`Allow: Date`, `greater than`, `=A2`, `Stop`) |
| **Top-bar selector parameter cell (`C2` `Display Range`, etc.)** | Pattern 6 (`Allow: List`, `=<NamedRange>`, `Stop`) |
| **Legacy Excel-formula column (not yet in C#)** | Pattern 5B (`=FALSE`, `Stop`, `Formula Cell`) |
| **Unused / reserved future column** | Pattern 5C (`=FALSE`, `Stop`, `Field Unused`) |

---

## 4. Pitfalls to Avoid

1. **Adding native `List` DV to `[ColumnRule]` table columns:** `RuleRegistry.ApplyDropdownStylesToWorkbook` strips native DV on every `HasDropdown` column at startup. Define table dropdowns in C# `[ColumnRule("...")]` instead.
2. **Using `Custom` instead of `List` for parameter-cell dropdowns:** A named range inside `Allow: Custom` (e.g., `All Vehicles!C2` using `Custom` with `DateFilterScopeList`) does not render a native Excel dropdown arrow or validate list membership properly. Always use `Allow: List` (`type="list"`) for Pattern 6.
3. **Forgetting the `"—"` em-dash guard on table date rules:** Computed/placeholder cells hold `"—"`, which breaks raw numeric/date comparisons. Always include `OR(ISBLANK(<Cell>), <Cell>="—", ...)`.
4. **Overlapping `sqref` ranges on locked columns:** Keep each calculated/locked column's `sqref` strictly single-column (e.g., `W5` and `X5` separately, not `W5:X5` overlapping `X5`) so tooltip and error messages never bleed across adjacent columns.
5. **Leaving `Ignore Blank = ON` on mandatory fields:** Excel skips `Custom` validation when a user clears a cell if `Ignore Blank` is checked. Always uncheck `Ignore Blank` for Pattern 1 and Pattern 5A.
6. **Using Worksheet Protection instead of `=FALSE` DV:** Sheet Protection breaks dynamic `ListObject` resizing and C# COM writes; always use Pattern 5 (`=FALSE` DV) for read-only cells.
7. **Inconsistent spelling in alert dialogs:** Use standard wording (`Header not editable`, `cannot be edited`) rather than ad-hoc variants (`edittable`, `editted`).
