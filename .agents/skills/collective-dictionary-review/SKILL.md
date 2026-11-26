---
name: collective-dictionary-review
description: On-demand AI consistency & multi-language autofill reviewer for the SKSB_App Collective Dictionary. Reviews staff-approved wording and new community tips (`guide.spotlight.custom.*`) in PostgreSQL (`app.localization_overrides` / `app.localization_history`), finds inconsistent phrases across `Core/Localization/ui.*.json` and `tooltips.*.json`, autofills missing translations for new tips/overrides across all 5 languages (`en`, `ms`, `zh`, `id`, `ms-en`) using the team's collective vocabulary, and submits proposals into `app.localization_suggestions` as a regular user (`AI Consistency Helper`, 1/3 votes) for staff voting or Debug-mode admin approval.
---

# Collective Dictionary — AI Consistency & Multi-Language Autofill Review

Review staff-approved wording and new community tips in the Collective Dictionary, propagate consistent phrasing across `ui.*.json` and `tooltips.*.json`, and autofill missing translations across all 5 languages (`en`, `ms`, `zh`, `id`, `ms-en`) using the team's collective vocabulary. **Never write directly to `app.localization_overrides`** — always submit as a pending suggestion (`AI Consistency Helper`, 1/3 votes) so staff or the Debug-mode admin makes the final call.

## Workflow

### 1. Export Current Collective Dictionary State via Debug `Portal.exe`
*(The read-only `sksb-db` MCP role does not have access to schema `app`; always use Debug `Portal.exe` which connects via `CollectiveDictionaryRepository`.)*

```powershell
$portal = "C:\Users\sksb-\Documents\Codes\SKSB_App\Portal\bin\Debug\net8.0-windows\Portal.exe"
$stateOut = Join-Path $env:TEMP "sksb_dict_state.json"
Start-Process -FilePath $portal -ArgumentList "--export-dict-state", "`"$stateOut`"" -Wait -NoNewWindow
Get-Content $stateOut -Raw
```

*(If `Portal.exe` is missing in `Portal\bin\Debug\net8.0-windows\`, build `Portal/Portal.csproj -c Debug` first.)*

The exported JSON contains `Overrides`, `PendingSuggestions`, `RejectedSuggestions`, and `RecentHistory`.
If `Overrides` is empty (no approved staff wording or tips yet), report that cleanly and stop.

### 2. Inspect Target Dictionaries (`Core/Localization/`)
Read `Core/Localization/ui.*.json` and `Core/Localization/tooltips.*.json` for all 5 languages (`en`, `ms`, `zh`, `id`, `ms-en`) as needed, and use `Overrides` as the authoritative reference for the team's preferred vocabulary and tone in each language.

### 3. Build Proposals (Two Jobs)

#### Job A — Same-Language Consistency Propagation
For each `LangCode` that has staff-approved overrides on built-in keys:
- Identify **clear, high-confidence terminology or phrasing inconsistencies** where staff approved a preferred term/phrasing in one or more keys, and other keys in the same language still use the older/inconsistent word or stiff translation for the same concept.
- Set `note` to cite the staff-approved key(s), e.g., `"Consistency with approved wording for ribbon.save.label"`.

#### Job B — Multi-Language Autofill for New Tips (`guide.spotlight.custom.*`) & Approved Overrides
For every approved **new community tip** (`entryKey` starting with `guide.spotlight.custom.` in `Overrides`, including both general tips `guide.spotlight.custom.<slug>_<ts>` and contextual tips `guide.spotlight.custom.ctx.<trigger>.<slug>_<ts>`) — and any approved override where staff added/changed a key in one language while other languages are missing or clearly out of sync:
- Check all 5 languages: `en`, `ms`, `zh`, `id`, `ms-en`.
- Preserve the **exact `entryKey`** (including `.ctx.<trigger>.` for contextual tips — supported triggers: `offline`, `total`, `header`, `date_filter`, `multiselect`, `autosync`, `dropdown`, `auto_defaults`, `history`, `currency`, `readonly`, `required`) so the tip continues to trigger on the right Excel context.
- For every language that does **not** yet have an approved override (or pending suggestion) for that `entryKey`:
  - Translate/adapt the approved tip/phrase into that language **using the team's approved Collective Dictionary vocabulary and tone** (for `ms-en`, match the natural Malaysian office Manglish style used in `Overrides` and `tooltips.ms-en.json`), keeping the trigger context in mind when it is a `.ctx.<trigger>.` tip.
  - Set `note` to cite the source language (and trigger if contextual), e.g., `"Auto-translated from approved ms-en contextual tip (trigger: total) using collective dictionary wording"`.

### 4. Strict Quality & Safety Filters
Before proposing any suggestion, verify ALL of the following:
- **Match staff voice:** Ground all terminology in the team's actual approved `Overrides`. Do **not** rewrite unrelated strings.
- **No duplicates:** Skip any `(dictType, entryKey, langCode)` that already appears in `Overrides` OR `PendingSuggestions`.
- **Respect rejections:** Never re-submit a `suggestedValue` that already appears in `RejectedSuggestions` for that `(dictType, entryKey, langCode)`.
- **Preserve placeholders:** Every `{0}`, `{1}`, `{0:N0}` token in the source/English string must appear identically in `suggestedValue`.

### 5. Submit Pending Suggestions via Debug `Portal.exe`
Write the batch as a UTF-8 JSON array to `$env:TEMP\sksb_dict_submit_in.json`:

```json
[
  {
    "dictType": "tooltips",
    "entryKey": "guide.spotlight.custom.check_odometer",
    "langCode": "ms",
    "suggestedValue": "Semak bacaan odometer sebelum menyimpan rekod penyelenggaraan.",
    "note": "Auto-translated from approved en tip using collective dictionary wording"
  }
]
```

Then submit via `Start-Process -Wait` and read the JSON result file:

```powershell
$portal = "C:\Users\sksb-\Documents\Codes\SKSB_App\Portal\bin\Debug\net8.0-windows\Portal.exe"
$submitIn = Join-Path $env:TEMP "sksb_dict_submit_in.json"
$submitOut = Join-Path $env:TEMP "sksb_dict_submit_out.json"
Start-Process -FilePath $portal -ArgumentList "--submit-dict-suggestions", "`"$submitIn`"", "`"$submitOut`"" -Wait -NoNewWindow
Get-Content $submitOut -Raw
Remove-Item $submitIn, $submitOut, (Join-Path $env:TEMP "sksb_dict_state.json") -ErrorAction SilentlyContinue
```

### 6. Report Summary to User
Report:
- Which staff-approved overrides and new tips were used as the reference.
- How many suggestions were submitted (`Submitted`), skipped (`Skipped`), or failed (`Errors`), split cleanly by **Consistency Updates** and **New Tip / Multi-Language Autofills**.
- Remind the user they can open **Collective Dictionary → Community Votes** in Portal or Excel (Debug build) to **Approve**, **Edit & Approve**, or **Reject** each suggestion.
