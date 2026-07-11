---
phase: 04-installation-registration
plan: 01
subsystem: installer
tags: [powershell, hkcu, com-registration, installer]

# Dependency graph
requires:
  - phase: 03-com-entry-point-real-excel-integration (plan 01, plan 02)
    provides: "Fixed COM identity (GUID 881EFDF3-424C-4240-BCA0-714DAC2B9CD7, ProgId FinanceFmtTools.Connect, AssemblyName FinanceFmtTools.ComAddin, Version 1.0.0.0) declared in Connect.cs's header comment, reused verbatim"
provides:
  - "scripts/install.ps1 — self-contained, admin-free PowerShell 5.1+ installer: GitHub-Releases one-liner default flow (INST-01) + -Package/-Source local-testing escape hatch + 3-registry-tree HKCU registration including DoNotDisableAddinList (INST-03)"
affects: [phase-4-plan-02-uninstall, phase-4-plan-03-live-verification, phase-5-cicd-release]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Zip-slip mitigation: Expand-Archive always targets a freshly created %TEMP% subfolder, never $InstallDir directly"
    - "No $PSScriptRoot/$MyInvocation dependency anywhere — the documented one-liner (irm | iex) has no on-disk script location to resolve, so the default branch always downloads from GitHub Releases rather than looking beside the script"
    - "Identity constants ($Guid/$ProgId/$ClassName/$AssemblyStr/etc.) declared once at the top and referenced everywhere else via variable, never re-typed literally — enforced by a grep-count-of-1 acceptance check on the literal GUID string"

key-files:
  created:
    - scripts/install.ps1
  modified: []

key-decisions:
  - "Wrote install.ps1 directly in the main conversation instead of spawning a fresh executor subagent, after the prior subagent run for this same plan was cut off by an API session-limit interruption before writing any code. The plan's task instructions were explicit enough (exact constant names, exact registry key shapes, exact function behaviors) to implement directly and verify with the plan's own grep-based checks — avoiding a third interruption risk on the same plan."
  - "Assert-ExcelNotRunning is called once, positioned before both the informational bitness check and the source-resolution/download step — satisfies the plan's 'before touching any file' requirement (all Copy-Item calls happen later) without adding a redundant second call; the plan's acceptance criteria only required 'at least once, before the first Copy-Item', which this satisfies with a simpler structure than a double-check."
  - "Added an EXCEL.EXE PE-bitness check (via Test-PeMachine on the App Paths registry value) alongside the required ClickToRun\\Configuration read, mirroring the sibling project's Outlook-equivalent check for parity — both are strictly informational (Write-Warn2 only, never exit/throw)."

patterns-established:
  - "scripts/install.ps1 is the template Plan 02's uninstall.ps1 must mirror for identity constants and registry tree paths (must stay byte-identical in value, not necessarily in code structure)."

requirements-completed: [INST-01, INST-03]

# Metrics
duration: unknown (written directly in main context after a subagent session-limit interruption on the same plan; no standalone timer)
completed: 2026-07-11
---

# Phase 4 Plan 1: install.ps1 — HKCU Registration Summary

**`scripts/install.ps1` (447 lines) implements the full INST-01/INST-03 installer: a GitHub-Releases one-liner default flow plus a `-Package`/`-Source` local-testing escape hatch, staging 4 required files into `%LocalAppData%\FinanceFmtTools\` and registering all 3 HKCU registry trees (COM class, non-versioned Excel discovery with `LoadBehavior=3`, versioned Resiliency `DoNotDisableAddinList=1`) — verified via every grep-based structural check embedded in the plan's two tasks.**

## Performance

- **Duration:** unknown — see Issues Encountered
- **Completed:** 2026-07-11
- **Tasks:** 2 (both `type="auto"`, same file)
- **Files modified:** 1 (created)

## Accomplishments
- Declared the fixed COM identity constants exactly as documented in `Connect.cs`'s header comment (`$Guid`, `$ProgId`, `$ClassName`, `$AssemblyStr`, `$RuntimeVer`, `$Shim`, `$ThreadingMdl`, `$FriendlyName`, `$Description`, `$OfficeVerKey`) — the literal GUID string appears exactly once in the whole file, in the `$Guid` declaration.
- Forced TLS 1.2 via `ServicePointManager` before any network call (PS 5.1 defaults to TLS 1.0, rejected by GitHub since 2018).
- Implemented three mutually-exclusive source-resolution branches with zero `$PSScriptRoot`/`$MyInvocation` dependency anywhere in the file: `-Package <zip>` (extract to fresh `%TEMP%` subfolder), `-Source <folder>` (direct `Find-BinDir` lookup), and the default documented one-liner path (`Get-LatestReleaseTag` for display + `Invoke-WebRequest` download of the version-agnostic `releases/latest/download/FinanceFmtTools.zip` URL).
- `Find-BinDir` checks the given root directly, then a `bin\` subfolder, then recurses, in that priority order.
- `Assert-ExcelNotRunning` guards against a running Excel process (with `-Force` auto-close fallback) before any file is copied; a non-blocking, `Write-Warn2`-only bitness check reads both `HKLM:\...\ClickToRun\Configuration`'s `Platform` value and `EXCEL.EXE`'s own PE header (`Test-PeMachine`) — never exits/throws on a non-x64 result, per the 64-bit-only-scope-is-not-a-hard-block distinction in REQUIREMENTS.md/FUT-01.
- Registered all 3 HKCU registry trees in a `try { ... } finally { Remove-TempExtract }` block: (a) COM class (`HKCU:\Software\Classes\CLSID\{Guid}\InprocServer32` + `ProgId`/`CLSID` cross-links), (b) non-versioned Excel discovery key (`HKCU:\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect`, `LoadBehavior=3`), (c) versioned Resiliency key (`HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList`, `FinanceFmtTools.Connect=1`) — INST-03.
- Post-install validation re-reads `LoadBehavior`/`CodeBase` from the registry and prints a final report (identity, install dir, files copied, exact registry paths written); exits `0`/`1` based on validation result.
- All plan-embedded automated `<verify>` grep checks pass (brace-balance 117/117, GUID count 1, zero `PSScriptRoot`/`MyInvocation`, `Assert-ExcelNotRunning` before first `Copy-Item` by line number, non-versioned discovery key present with zero incorrectly-versioned variants, versioned Resiliency key present, `-Type DWord` ×2+, `CodeBase` ×2+, `finally` present, `Remove-TempExtract` ×2+).

## Task Commits

1. **Task 1 + Task 2 (both `scripts/install.ps1`, written together as one file):** `4430019` (feat)

## Files Created/Modified
- `scripts/install.ps1` - new file, 447 lines — full installer per INST-01/INST-03

## Decisions Made
- See `key-decisions` above (direct-write-instead-of-subagent, single `Assert-ExcelNotRunning` call placement, added EXCEL.EXE PE-bitness check for parity with the sibling project's Outlook check).

## Deviations from Plan

None requiring correction. One structural interpretation call: the plan's prose mentioned calling `Assert-ExcelNotRunning` in two places ("after Assert-ExcelNotRunning succeeds, perform... bitness check" and, separately, "call Assert-ExcelNotRunning before touching any file"). The plan's own automated verify script and acceptance criteria only require the function to be defined, called **at least once**, with that call site preceding the first `Copy-Item` by line number — satisfied with a single, earlier call. Documented as a key-decision rather than a deviation since it fully satisfies every stated criterion.

## Issues Encountered
- The subagent originally dispatched to execute this plan was cut off by an Anthropic API session-limit error (`You've hit your session limit`) at the very start of its work ("checking the scripts directory and bin folder state"), before writing any file or making any commit. Verified via `git status --short` (only an unrelated stray uncommitted `STATE.md` edit from that same interrupted run, discarded via `git checkout --`) and `ls scripts/` (directory did not exist) that zero durable progress survived. Recovered by writing `scripts/install.ps1` directly in the main conversation against the plan's explicit task instructions, rather than re-spawning a third subagent attempt on the same plan.
- Because of the above, no plan-execution start/end timestamps exist for a duration metric — recorded as "unknown" rather than fabricated.

## User Setup Required
None for this plan's own scope (script authored and structurally verified only). Actual execution — running `install.ps1` on a real Windows+Excel machine, registering the add-in, and confirming Excel loads it — remains `human_needed`, tracked as Plan 03's live-verification checkpoint (per 04-CONTEXT.md's non-discretionary no-Windows/PowerShell environment constraint, same pattern as Phase 3).

## Next Phase Readiness
- `scripts/install.ps1` is ready for Plan 02's `uninstall.ps1` to mirror its identity constants and the 3 registry tree paths it wrote (COM class, Excel discovery, Resiliency).
- No blockers for Plan 02.

---
*Phase: 04-installation-registration*
*Completed: 2026-07-11*
