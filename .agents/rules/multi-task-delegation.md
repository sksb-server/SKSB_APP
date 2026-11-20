---
trigger: model_decision
description: >-
  Context-efficient multi-task delegation policy for SKSB_App (English-Only Hub
  + Subagent Brains + Claude Code Hands). Load whenever the user pastes a list
  of 2+ tasks, runs a batch of tasks, wants the main chat kept English-only /
  code-free, splits work across subagents (`invoke_subagent`), or coordinates
  workers with `swarm`.
---

# Multi-Task Delegation Efficiency (English-Only Hub Model)

Use a **3-tier Hub-and-Spoke** architecture so the Main Chat stays lightweight, plain-English, and free of code/diff bloat while preserving the **Antigravity (Brain) → Claude Code (Hands)** workflow:

- **Tier 1 — Main Chat (English-Only Hub):** Pure English planning, triage, dispatch, and user summary. **Never** reads application source code (`*.cs`, `*.sql`, `*.csproj`, etc.), never runs `claude` directly, and never dumps `git diff` or build logs into its own context.
- **Tier 2 — Task Subagent (`invoke_subagent`, `TypeName: "self"`):** Disposable per-task Antigravity Brain. Reads only its task's target files, builds the minimal-context prompt, runs **Claude Code** (`claude -p`) as Hands, verifies `git diff` + build in its own context, and returns a compact English-only receipt.
- **Tier 3 — Claude Code (`claude -p --no-session-persistence`):** Surgical file edits + targeted `dotnet build`.

## 1. Zero-Code Triage (Main Chat)

- Group tasks that touch the **same module/files** into a single Task Subagent so edits and builds do not collide; keep unrelated tasks separate.
- Consult `CODEBASE_MAP.md` **only** if module ownership or overlap is unclear. Never read source files just to summarize them into a subagent's prompt—let the Task Subagent read the files it owns.
- Never `fork` the Main Chat history into a worker; start every Task Subagent fresh with a tight English brief (task goal, module hint, and receipt format).

## 2. Private by Default & Workspace Safety

- **Independent tasks → `invoke_subagent` (`TypeName: "self"`):** Private 1-to-1 channel avoids shared board cross-talk.
  - **Concurrent edits/builds:** Use `Workspace: "share"` (isolated working copy so parallel `dotnet build` calls do not lock `Core/obj` or `.xll`). Instruct each subagent to commit its verified change on its worktree branch and report the branch/commit; merge cleanly into the main workspace once done.
  - **Direct working-tree edits (no branches) or overlapping modules:** Serialize tasks in `Workspace: "inherit"` (one subagent at a time) or group them into one subagent.
- **Interdependent tasks → `swarm`:** Use `swarm` **only** when workers must coordinate with each other or share intermediate discoveries on a public board.

## 3. Structured English-Only Receipts (No Code Leakage)

Require every Task Subagent to reply with **plain English only** (max ~8 lines—no code blocks, no diffs, no raw build logs):

```text
Task: <short name>
Status: DONE | BLOCKED | NEEDS REVIEW
Summary: <1-2 plain-English sentences of what changed>
Files: <comma-separated paths only>
Build: <PASS/FAIL + target project>
Manual/Risks: <.xltx or DB steps, warnings, or None>
```

## 4. Trust-by-Default Verification

- **Clean receipt (`DONE` + `PASS` + low risk):** Trust the Task Subagent's verification—do **not** re-read files, run `git diff`, or re-run passing builds in the Main Chat.
- **Flagged or high-risk (`BLOCKED`, `NEEDS REVIEW`, or touches `Core/`, COM, SQL/Dapper, shared contracts):** Do not pull raw diffs into the Main Chat; either send a follow-up to that Task Subagent (`send_message`) or spawn a dedicated read-only review subagent to verify and report back in English.
- After all subagents complete, present a single consolidated English summary table + combined **Manual `.xltx` / DB Steps** to the user.

