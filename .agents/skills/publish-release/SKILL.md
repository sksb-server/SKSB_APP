---
name: publish-release
description: >-
  Inspects current git changes in SKSB_App, drafts user-facing update details (for
  release_notes.json and Portal's NEW badge hover tooltip) and a git commit description,
  then publishes the release via publish.ps1, commits all changes (including version.json
  and release_notes.json), and pushes to the remote repository. Use whenever the user asks
  to "commit, push and publish", "publish release", "ship a new version", or invokes
  /publish-release.
---

# Publish Release (Commit, Push & Publish)

Automates the end-to-end SKSB_App release flow: generating user-facing update notes (displayed when hovering over the `NEW` version tag in `Portal.exe`), publishing the Release package to the update server, and committing + pushing everything cleanly to Git.

> **Authority Note:** When the user invokes this skill, they are explicitly requesting `publish.ps1`, `git commit`, and `git push`. Always present the drafted version, update details, and commit message for a quick confirmation before running `publish.ps1` and `git push` (unless the user explicitly says to skip confirmation).

## Steps

### 1. Inspect Repository State & Changes
Run read-only git checks in one command:
```powershell
git rev-parse --abbrev-ref HEAD; git status --short; git log -n 5 --oneline; git diff --stat
```
- Also inspect `git diff` (and any new untracked source files) to accurately summarize every user-visible and technical change since the last commit/release.
- Read `version.json` to get the current version (`X.Y.Z`).
- Determine the target version:
  - If the user specified a version (e.g. `1.3.0`), use it.
  - Otherwise auto-increment the patch version (`X.Y.(Z+1)`), matching `publish.ps1`.
- Warn the user if:
  - They are not on `main` / `master`.
  - There are unexpected untracked/temporary files (never commit `SKSB_Release_Package.zip`, `publish_staging/`, `*.log`, or `*.tmp`).

### 2. Draft Update Details & Git Commit Description
Prepare two clear outputs based on the actual diff:
1. **Update Details (`release_notes.json` for Portal hover tooltip):**
   - 2 to 6 short, plain-English bullet points describing what changed from a **user / staff** perspective (avoid internal filenames or code jargon).
   - Keep each bullet concise (~60–100 characters) so the WinForms hover tooltip stays clean and readable.
2. **Git Commit Description:**
   - Concise subject line (`v<nextVersion>: <short summary>`, ~50–72 chars).
   - Short bulleted body summarizing both user-facing and technical changes.

### 3. Confirm Before Publishing (Default)
Present the following summary to the user and ask for confirmation (skip only if the user explicitly asked to run without confirmation):
- **Branch:** `<branch>`
- **Version:** `v<current>` → `v<next>`
- **Update Details (`release_notes.json`):** bullet list
- **Git Commit Message:** subject + body
- **Files to Commit:** summary of modified/new files (plus `version.json` and `release_notes.json` updated by `publish.ps1`)

### 4. Publish via `publish.ps1`
Run `publish.ps1` **before** committing so the updated `version.json` and `release_notes.json` are included in the same release commit:
```powershell
$notes = @(
    "First user-facing highlight",
    "Second user-facing highlight"
)
powershell -NoProfile -ExecutionPolicy Bypass -File .\publish.ps1 -ReleaseVersion "<nextVersion>" -ReleaseNotes $notes
```
- If `publish.ps1` fails (e.g. Release build error, Excel locking `.xll`, or update server `192.168.0.20:5000` offline), **stop immediately** and report the error. Do **not** commit or push a failed release.

### 5. Commit & Push
Once `publish.ps1` reports `DEPLOYMENT SUCCESS!`:
1. Verify `git status --short` — confirm `version.json` and `release_notes.json` were updated, and make sure `SKSB_Release_Package.zip` and `publish_staging/` are **not** staged.
2. Stage the intended files (explicitly excluding `SKSB_Release_Package.zip` and `publish_staging/`):
   ```powershell
   git add -A -- . ':!SKSB_Release_Package.zip' ':!publish_staging'
   ```
3. Commit using a multi-line PowerShell here-string so quoting never breaks:
   ```powershell
   $msg = @'
   v<nextVersion>: <subject>

   - <detail 1>
   - <detail 2>
   '@
   git commit -m $msg
   ```
4. Push to the current branch's remote:
   ```powershell
   git push
   ```
   *(If the branch has no upstream yet, use `git push -u origin <branch>`.)*

### 6. Report Result
Reply concisely with:
- Published version (`v<nextVersion>`) and server deployment confirmation
- Commit hash & pushed branch
- Final `release_notes.json` bullets that users will see when hovering over the `NEW` badge in `Portal.exe`
