---
name: open-branch
description: >-
  Resolves the current git worktree/branch (or a specified worktree/branch) and gives
  the user the clickable solution link and 1-line command to open it in a new Visual
  Studio 2026 or VS Code window without disturbing their main IDE window. Use whenever
  the user says "open branch location", "open this branch in VS", "open in Visual Studio",
  or asks to open a worktree in an IDE.
---

# Open Branch Location in Visual Studio 2026

> **Important Windows Constraint:** Antigravity's `run_command` executes on an isolated, non-interactive Windows desktop, so launching `devenv.exe` or `code` from `run_command` will **not** appear on the user's desktop. **Never** launch `devenv.exe`, `code`, `Portal.exe`, or `excel.exe` from `run_command`.

## Steps

1. Resolve the target worktree path and branch name (read-only):
   - If the user did not specify another branch/worktree, use the current workspace directory (`$PWD`) and `git rev-parse --abbrev-ref HEAD`.
   - If the user named a branch or worktree, run `git worktree list` to find the matching path.
2. Reply concisely with:
   - The **worktree path** and **branch name**
   - Clickable `file://` links to the worktree's `SKSB_App.slnx` and folder
   - The 1-line PowerShell command to run in the user's own terminal (which runs on their interactive desktop):
     ```powershell
     & "C:\Users\sksb-\Documents\Codes\SKSB_App\test.ps1" -OpenVS -Worktree "<branch-or-worktree>"
     ```
     *(or `.\test.ps1 -OpenVS` if already in that worktree)*
   - Reminder that in Antigravity IDE / VS Code they can also run **Tasks: Run Task → `Open Branch in Visual Studio 2026`**.
