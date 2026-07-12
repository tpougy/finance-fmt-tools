---
phase: quick-260712-hzj
plan: 01
subsystem: infra
tags: [powershell, installer, com-interop, excel, race-condition]

# Dependency graph
requires:
  - phase: quick-260712-hbu
    provides: "Remove-LegacyVbaAddin function and its finally block in scripts/install.ps1"
provides:
  - "Race-free handoff between Remove-LegacyVbaAddin's own Excel.Quit() and PASSO 2's Assert-ExcelNotRunning call"
affects: [installer, scripts/install.ps1]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Same-style bounded polling wait loop (Start-Sleep then Get-Process, break on gone) reused from Assert-ExcelNotRunning inside a function-internal finally block, with a shorter timeout and no user-facing logging (fail-open, since a downstream check is the real safety net)"

key-files:
  created: []
  modified:
    - "scripts/install.ps1 — Remove-LegacyVbaAddin's finally block: added second [GC]::Collect() and a 15s wait loop for EXCEL.EXE to exit"

key-decisions:
  - "Wait loop is silent (no Write-* messages, no exit) and fails open after ~15s — PASSO 2's existing Assert-ExcelNotRunning remains the actual safety net for a truly stuck process"
  - "New local variable name ($legacyExcelProc) used instead of reusing $excel/$wb/$foundAddin to avoid shadowing/confusion inside the same function scope"

requirements-completed: []

# Metrics
duration: 8min
completed: 2026-07-12
---

# Quick Task 260712-hzj: Corrigir race condition no Remove-LegacyVbaAddin Summary

**Segunda `[GC]::Collect()` + loop de espera (até 15s) pelo término do EXCEL.EXE dentro do `finally` de `Remove-LegacyVbaAddin`, eliminando a race condition com `Assert-ExcelNotRunning` do PASSO 2**

## Performance

- **Duration:** ~8 min
- **Completed:** 2026-07-12
- **Tasks:** 1 completed
- **Files modified:** 1

## Accomplishments
- `Remove-LegacyVbaAddin`'s `finally` block now waits (with bounded polling, same style as `Assert-ExcelNotRunning`) for the `EXCEL.EXE` process it opened via COM to actually exit before returning, instead of returning immediately after `$excel.Quit()`.
- Added a second `[GC]::Collect()` pass after `WaitForPendingFinalizers()` to finish releasing any RCWs the first pass left queued for finalization, which was allowing the Excel process to linger past the function's return.
- Eliminates an intermittent installer failure observed against a real Excel install, where PASSO 2's `Assert-ExcelNotRunning` would find the just-closed Excel process still in the process of shutting down and abort the entire install with "Excel ainda está aberto."

## Task Commits

Each task was committed atomically:

1. **Task 1: Esperar o processo EXCEL.EXE terminar antes de Remove-LegacyVbaAddin retornar** - `528082f` (fix)

## Files Created/Modified
- `scripts/install.ps1` - `Remove-LegacyVbaAddin`'s `finally` block: added second `[GC]::Collect()` and a `for ($i = 0; $i -lt 15; $i++)` wait loop (`Start-Sleep -Seconds 1` then `Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue` into `$legacyExcelProc`, `break` when gone), plus a short Portuguese comment explaining why polling is safe at this point (any `EXCEL.EXE` found here is necessarily the process this function itself opened and told to quit, since `Assert-ExcelNotRunning` already confirmed no user Excel was open before this function ran).

## Decisions Made
- Kept the fix strictly additive and scoped to `Remove-LegacyVbaAddin`'s existing `finally` block — no changes to `Assert-ExcelNotRunning`, the main flow (PASSO 0/1/2/4), or any other script, per the plan's explicit scope restriction.
- No new `Write-*` message or `exit` call added for the wait loop: it fails open silently after ~15s, deferring to PASSO 2's own `Assert-ExcelNotRunning` (30s wait + `CloseMainWindow()` + actionable error) as the real safety net for a genuinely stuck process.

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered
None.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- The race condition fix is code-complete and passed structural + syntax verification (`PSParser::Tokenize` via `powershell.exe` reports `SYNTAX OK`).
- Per the plan's `<verification>` section, live-Excel end-to-end validation of the actual race-condition fix (rerunning the real VBA→C# migration install scenario against a real Excel install) is explicitly out of scope for this plan — the orchestrator should perform that validation next, immediately after this plan completes.

## Self-Check: PASSED

- FOUND: scripts/install.ps1 (modified, diff confirmed scoped to `Remove-LegacyVbaAddin`'s `finally` block only)
- FOUND: commit 528082f (`git log --oneline -3` confirms it exists in history)
- Structural checks confirmed: 2x `[GC]::Collect()`, 3x `Get-Process -Name 'EXCEL'`, 1x `for ($i = 0; $i -lt 15; $i++)`, 2 function definitions (`Assert-ExcelNotRunning` + `Remove-LegacyVbaAddin`), zero PowerShell parse errors.

---
*Phase: quick-260712-hzj*
*Completed: 2026-07-12*
