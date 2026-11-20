---
trigger: model_decision
description: >-
  Context-efficient multi-task delegation policy for SKSB_App. Load whenever
  splitting work across multiple tasks, spawning subagents (`invoke_subagent`),
  or coordinating workers with `swarm`.
---

# Multi-Task Delegation Efficiency

Follow these rules whenever a session delegates multiple tasks to subagents or swarm workers so total context stays lean and free of cross-talk.

## 1. Private by Default (`invoke_subagent` vs. `swarm`)

- **Unrelated / independent tasks → `invoke_subagent`:** Each subagent communicates privately 1-to-1 with the coordinator, avoiding shared message-board noise across unrelated tasks. Use `Workspace: "share"` when multiple workers edit or build concurrently.
- **Interdependent tasks → `swarm`:** Use `swarm` **only** when workers must coordinate with each other or share intermediate discoveries on a public board.

## 2. Map-Only Triage (Zero Pre-Read Duplication)

- Consult `CODEBASE_MAP.md` (or minimal path lookups) to identify target files, module boundaries, and ownership, then delegate immediately.
- Do **not** read full target source files just to summarize them into a worker's prompt—let the assigned worker read the files it owns.
- Do **not** `fork` the coordinator's conversation history into a worker unless it strictly requires prior turns; start workers fresh with a tight brief.

## 3. Trust-by-Default Verification

- Instruct every worker to end with a compact structured report:
  - **Files changed**
  - **Build / verification result**
  - **Risks / manual `.xltx` or DB steps**
- **Low-risk, localized tasks with a clean build:** Trust the worker's report—do **not** re-read modified files or re-run passing checks.
- **High-risk or flagged tasks:** Inspect `git diff` or spot-check specific files only when changes touch `Core/`, COM interop, SQL/Dapper, shared contracts/templates, or when the worker reports warnings or uncertainty.

