---
trigger: model_decision
description: >-
  SKSB_App conversation organization standard: topic tag taxonomy, archiving
  criteria, and pinned/archived rule-sync guidelines. Load whenever organizing,
  tagging, archiving, or syncing SKSB_App conversations.
---

# SKSB_App Conversation Organization Standard

Authoritative rules for keeping `SKSB_App` conversations (`projectId: e70ace11-b138-4914-9c29-1ac5baa583fc`) consistently tagged, archived, and aligned with internal documentation.

---

## 1. Canonical Topic Tags

Assign 1 primary tag (or 2 for genuinely cross-module work) from this fixed set:

| Tag | Scope |
|---|---|
| `schedule-manager` | Schedule Manager, Daily Schedule Board, hover details, worker matching, date alignment, save speed, stacked entries |
| `attendance` | Daily Attendance, thumbprint sync (`ZkBridge.exe` / Vigilance Controller), late check-in, regular shifts, auto attendance |
| `vehicle-manager` | Vehicle Manager, fuel logs, odometer analysis, fuel card report sync, insurance claims, vehicle UI |
| `project-manager` | Project Manager templates, code attribute calculations, 0-holes handling, project skill requirements |
| `staff-manager` | Staff Manager, staff licenses & expiration, staff ID validation, salary, worker skills display, preference sheet |
| `core-ui` | `Core` library, loading window, launcher DPI, add-in loading, ribbon, auto-updater, sheet protection, offline mode, concurrent edits, bug report system, data validation, conditional formatting |
| `roadmap-erp` | SKSB App Roadmap, WhatsApp SKSB App plan, SQL Accounting & SDK, ERP server hardware, Report Generator AI architecture, Todoist integration |
| `ops-git` | Git commits/merges/worktrees, debug/release server, `CLAUDE.md`, Antigravity rules/agents/swarm, Postgres DB schema & OneDrive backups, network/VLAN/Omada, conversation cleanup |

---

## 2. Archiving Criteria

- **Always Keep Active:**
  - Any **pinned** conversation (`annotations.pinned == true`).
  - Any conversation updated within the **last 48 hours** (unless it is clearly a completed one-off git commit/push or ops task).
  - Any active `roadmap-erp` or multi-day architecture/planning conversation still in progress.
- **Archive Automatically:**
  - Completed bug fixes, UI tweaks, one-off git commit/push/merge chats, and resolved troubleshooting sessions.
  - Stale non-pinned conversations with no activity for more than 48 hours.

---

## 3. Internal Rules & Documentation Sync Policy

- **Auto-Updated File (Safe):**
  - `.agents/rules/pinned-context.md` is automatically refreshed during each cleanup run so it always lists the current pinned and active strategic/architecture conversations.
- **Protected Core Rules (Suggest Only — Never Auto-Edit):**
  - **Do NOT** auto-edit `CLAUDE.md`, `GEMINI.md`, `CODEBASE_MAP.md`, or other `.agents/rules/*.md` files during unattended cleanup runs.
  - Instead, inspect newly archived and pinned conversations for durable new conventions, architectural decisions, or recurring gotchas, and list concrete **Suggested Rule / Doc Updates** in the run summary for user approval.

