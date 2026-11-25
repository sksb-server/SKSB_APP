# SKSB_App — Project Overrides (Antigravity ↔ Claude Code)

Follow the global **Antigravity (Brain) ↔ Claude Code (Hands)** rule and `multi-task-delegation` skill with these project-specific overrides:

- **Scope:** Never directly edit `*.cs`, `*.csproj`, `*.slnx`, `*.sql`, `*.ps1`, or `*.xltx`, or run `dotnet build`, `dotnet clean`, or `publish.ps1`.
- **Trivial Edit Exception:** Direct edit + `dotnet build <Module>/<Module>.csproj -c Debug` only for single non-`Core` file (<20 lines) changes with **no** COM, Dapper/SQL, `*.csproj`/`*.slnx`, or `*.xltx` changes (e.g. Ribbon XML/icons, UI labels/text, formatting/constants).
- **Locate Context:** Check `CODEBASE_MAP.md`, `CLAUDE.md`, and `sksb-db` MCP tools as needed.
- **Claude Code Execution:**
  - `--model`: `sonnet` default; `opus` for complex multi-file `Core`/COM/Dapper work.
  - `--allowedTools`: `"Read,Edit,Write,Bash(dotnet build*)"`; verify with `dotnet build <Module>/<Module>.csproj -c Debug` (or `SKSB_App.slnx` when `Core`/multiple projects change).
  - Environment block: stop and report if Excel locks `.xll` files.
  - Report **Manual Template (`*.xltx`) / DB Steps** with the verified diff.
- **Multi-Task Overrides:** Use `Workspace: "share"` for concurrent builds so parallel `dotnet build` calls do not lock `Core/obj` or `.xll`. High-risk review triggers: `Core/`, COM, SQL/Dapper, shared contracts.
