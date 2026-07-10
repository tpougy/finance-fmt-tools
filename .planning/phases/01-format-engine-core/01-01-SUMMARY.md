---
phase: 01-format-engine-core
plan: 01
subsystem: format-engine
tags: [dotnet, csharp, xunit, net48, net8.0, scaffolding]

# Dependency graph
requires: []
provides:
  - "FinanceFmtTools.sln solution wiring FinanceFmtTools.Engine + FinanceFmtTools.Engine.Tests"
  - "FinanceFmtTools.Engine class library multi-targeting net48;net8.0, zero Excel/COM references"
  - "FinanceFmtTools.Engine.Tests net8.0-only xUnit test project referencing the engine via ProjectReference"
  - "FormatKeys (11 format-key constants ported verbatim from modConfig.bas)"
  - "FormatCategory enum (Numeric, Percent, Date, Text)"
  - "CellAlignment enum (General, Left, Right) — COM-free stand-in for XlHAlign"
  - "FormatDef immutable class (Key, DisplayName, NumberFormat, Category, Alignment)"
affects: [01-02-format-registry, 01-03-format-registry-tests]

# Tech tracking
tech-stack:
  added: [".NET 8 SDK 8.0.422", "Microsoft.NETFramework.ReferenceAssemblies 1.0.3", "xunit 2.9.3", "xunit.runner.visualstudio 3.1.5", "Microsoft.NET.Test.Sdk 17.14.1"]
  patterns: ["Multi-target logic library (net48;net8.0), single-target test project (net8.0) so dotnet test runs on Linux", "Plain immutable classes instead of C# records to stay net48-compilable (CS0518 avoidance)"]

key-files:
  created:
    - src/FinanceFmtTools.sln
    - src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj
    - src/FinanceFmtTools.Engine/FormatKeys.cs
    - src/FinanceFmtTools.Engine/FormatCategory.cs
    - src/FinanceFmtTools.Engine/CellAlignment.cs
    - src/FinanceFmtTools.Engine/FormatDef.cs
    - src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj
    - .gitignore
  modified: []

key-decisions:
  - "Added .gitignore (bin/, obj/, .vscode/, OS cruft) — not in the plan's file list but required to avoid committing build artifacts; the repo had none before this plan."
  - "Reworded a code comment in FormatDef.cs that originally used the literal word 'record' to explain why records are avoided — the acceptance criteria grep for the literal token 'record' anywhere in the file, so the explanatory comment itself would have failed the check. Rephrased without losing the rationale."

patterns-established:
  - "FormatDef.cs pattern: plain sealed class, constructor-assigned get-only auto-properties, no record/init — future FormatDef-shaped types in this codebase should follow the same shape for net48 compatibility."

requirements-completed: [DEV-01]

# Metrics
duration: 25min
completed: 2026-07-10
---

# Phase 1 Plan 1: Bootstrap Solution & Contract Types Summary

**Two-project dotnet CLI solution (net48;net8.0 multi-target Engine + net8.0-only xUnit Tests) with zero Excel/COM references, building 0-warning/0-error on Linux via a freshly-installed .NET 8 SDK.**

## Performance

- **Duration:** 25 min
- **Started:** 2026-07-10T20:50:00Z
- **Completed:** 2026-07-10T21:15:00Z
- **Tasks:** 2 completed
- **Files modified:** 8 (7 created for the plan + .gitignore)

## Accomplishments
- Installed .NET 8 SDK 8.0.422 to `$HOME/.dotnet` (the prior partial install in this shared environment had no `dotnet` binary — reinstalled clean via the official `dotnet-install.sh` script) and confirmed `dotnet --version` works.
- Bootstrapped `FinanceFmtTools.sln` with two wired projects: `FinanceFmtTools.Engine` (`net48;net8.0`) and `FinanceFmtTools.Engine.Tests` (`net8.0` only, via `ProjectReference`).
- Ported the 11 `FMT_*` format-key constants from `src/modConfig.bas:19-29` verbatim into `FormatKeys.cs`, plus the COM-free `FormatCategory`, `CellAlignment`, and `FormatDef` contract types that 01-02/01-03 will implement logic against.
- `dotnet build src/FinanceFmtTools.sln -c Release` succeeds with 0 Warning(s)/0 Error(s), producing both `net48` and `net8.0` DLLs; `dotnet test` on the Tests project exits 0 with the expected "no tests" outcome (0 tests exist yet — 01-02/01-03 add them).

## Task Commits

Each task was committed atomically:

1. **Task 1: Bootstrap solution and multi-target Engine class library** - `f4e0c17` (feat)
2. **Task 2: Bootstrap xUnit test project and define shared contract types** - `ff82841` (feat)

**Plan metadata:** (pending — committed alongside this SUMMARY)

## Files Created/Modified
- `.gitignore` - excludes `bin/`, `obj/`, `.vscode/`, OS cruft from git (repo had none before this plan)
- `src/FinanceFmtTools.sln` - solution wiring both projects
- `src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj` - multi-targets `net48;net8.0`, `LangVersion 9.0`, `Nullable disable`, conditional `Microsoft.NETFramework.ReferenceAssemblies` PackageReference for the `net48` leg only
- `src/FinanceFmtTools.Engine/FormatKeys.cs` - 11 `public const string` fields (`Integer`, `Fin2D`, `Fin4D`, `Fin8D`, `Pct4D`, `Pct2D`, `SpreadBps`, `DateIso`, `DateBr`, `DateBrLong`, `Text`) with values copied verbatim from `modConfig.bas`
- `src/FinanceFmtTools.Engine/FormatCategory.cs` - `enum { Numeric, Percent, Date, Text }`
- `src/FinanceFmtTools.Engine/CellAlignment.cs` - `enum { General, Left, Right }`
- `src/FinanceFmtTools.Engine/FormatDef.cs` - `sealed class` with 5 get-only properties (`Key`, `DisplayName`, `NumberFormat`, `Category`, `Alignment`) and a 5-arg constructor
- `src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj` - `net8.0`-only, `IsPackable false`, `Nullable disable`, pinned `Microsoft.NET.Test.Sdk 17.14.1` / `xunit 2.9.3` / `xunit.runner.visualstudio 3.1.5`, `ProjectReference` to the Engine project

## Decisions Made
- Added `.gitignore` for .NET build artifacts (`bin/`, `obj/`) since none existed in this VBA-era repo — necessary to avoid accidentally committing binary build output; not itself in the plan's `files_modified` list but required by the plan's own action steps (running `dotnet build` generates `bin/`/`obj/` immediately).
- Kept `coverlet.collector` (a template default, not mentioned in the plan) in the Tests csproj since the plan only specified pinning `Microsoft.NET.Test.Sdk`/`xunit`/`xunit.runner.visualstudio` versions and said nothing about removing the coverage collector; leaving it is harmless and consistent with "don't remove what the plan didn't ask to remove."

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] Reworded a FormatDef.cs comment that inadvertently contained the literal grep-checked token**
- **Found during:** Task 2 (contract type creation) — acceptance-criteria verification gate
- **Issue:** The acceptance criteria require `src/FinanceFmtTools.Engine/FormatDef.cs` to not contain the literal token `record` anywhere in the file (a plain string grep, not "no record type declaration"). The first draft's explanatory comment ("NOT a C# `record`/`init`-only shape — records fail to compile...") used the word "record" twice, purely as English prose explaining the design choice, which would have failed the literal grep check.
- **Fix:** Reworded the comment to convey the same rationale ("C# 9's alternate immutable-type syntax fails to compile on net48 with CS0518") without using the literal word.
- **Files modified:** `src/FinanceFmtTools.Engine/FormatDef.cs`
- **Verification:** `grep -c '\brecord\b' src/FinanceFmtTools.Engine/FormatDef.cs` returns 0 (exit code 1); rebuild and `dotnet test` re-run confirmed 0 warnings/0 errors and exit code 0 after the edit.
- **Committed in:** `ff82841` (Task 2 commit — comment was fixed before the task's single commit, not a separate follow-up commit)

---

**Total deviations:** 1 auto-fixed (1 bug — acceptance-criteria literal-string collision in a comment)
**Impact on plan:** Cosmetic only; no behavioral change. No scope creep.

## Issues Encountered
- The environment note flagged that a prior research-session .NET 8 SDK install might not have produced a persistent binary in this shared environment. Confirmed true: `$HOME/.dotnet` existed with sentinel files but no `dotnet` executable. Removed the stale partial directory and reinstalled clean via `dotnet-install.sh --channel 8.0 --install-dir "$HOME/.dotnet"`, which succeeded (SDK 8.0.422) and was verified working for the remainder of the plan.

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- `FormatKeys`, `FormatCategory`, `CellAlignment`, `FormatDef` are ready for 01-02 (`AccountingFormatBuilder` + `FormatRegistry.TryGetFormatDef`) to implement logic against without further scaffolding.
- The `net8.0`-only Tests project with `ProjectReference` already wired means 01-03 can add `[Theory]`/`[InlineData]` test files immediately — no project-file changes needed.
- No blockers. `dotnet build`/`dotnet test` both run cleanly on this Linux dev machine per DEV-01's requirement.

---
*Phase: 01-format-engine-core*
*Completed: 2026-07-10*
