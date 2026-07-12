---
phase: quick-260712-hbu
plan: 01
subsystem: installer
tags: [powershell, com-interop, migration, legacy-cleanup]

dependency-graph:
  requires: []
  provides:
    - "Remove-LegacyVbaAddin function in scripts/install.ps1"
    - "Automatic legacy VBA (.xlam) detection/removal before C# install"
  affects:
    - scripts/install.ps1

tech-stack:
  added: []
  patterns:
    - "Nested try/catch/finally with unconditional COM object release + GC.Collect/WaitForPendingFinalizers, mirroring the archived VBA-era installer's COM cleanup pattern"
    - "Never-blocking auxiliary automation: inner catch swallows all COM automation failures with Write-Warn2, letting the outer C# install flow proceed regardless"

key-files:
  created: []
  modified:
    - scripts/install.ps1

decisions:
  - "VbaAddinTitle hardcoded to 'Finance Fmt Tools' (same literal as the C# add-in's FriendlyName) — must match the legacy .xlam's document Title property exactly, per plan's canonical reference to the archived installer"
  - "$script:VbaRemoved is set to $true unconditionally after Remove-Item (with -ErrorAction SilentlyContinue), not gated on removal success — matches plan's explicit action spec and accepted threat disposition T-quick-hbu-04"

metrics:
  duration: "~15 minutes"
  completed: 2026-07-12
---

# Quick Task 260712-hbu: Auto-detect and remove legacy VBA installation Summary

Added automatic detection, COM-based unregistration, and disk removal of a legacy VBA `.xlam` add-in installation to `scripts/install.ps1`, running before any C# binary resolution or HKCU registration, and never blocking the C# install if Excel COM automation fails.

## What Was Built

**Task 1 — Constants, state flag, and `Remove-LegacyVbaAddin` function** (commit `cd364be`):
- Added `$VbaAddinTitle` (`'Finance Fmt Tools'`), `$VbaAddinDir` (`%APPDATA%\Microsoft\AddIns`), `$VbaXlamPath` (`FinanceFmtTools.xlam` under that dir) — new "Legado VBA (.xlam)" block right after the existing "Identidade fixa" block.
- Added `$script:VbaRemoved = $false`, initialized alongside `$script:TempExtractDir`.
- Added `Remove-LegacyVbaAddin`, inserted between `Assert-ExcelNotRunning` and `Test-PeMachine`:
  - Early-return (no Excel COM touched) when `$VbaXlamPath` doesn't exist.
  - Otherwise opens `Excel.Application` (invisible, `DisplayAlerts=$false`), calls `Workbooks.Add()` to access the `AddIns` collection, iterates with a classic indexed `for` loop, releases every non-matching `AddIns.Item()` immediately, and matches by `.Title -eq $VbaAddinTitle`.
  - On match: `$foundAddin.Installed = $false` + `Write-Ok`. On no match: `Write-Info`, no error.
  - All COM automation wrapped in an inner `try/catch` that only logs `Write-Warn2` on failure — never re-throws, never exits.
  - Outer `finally` always releases `$wb`/`$foundAddin`/`$excel` via `Marshal.ReleaseComObject` (each close/quit wrapped in its own silent try/catch) and always calls `[GC]::Collect()` + `[GC]::WaitForPendingFinalizers()`.
  - After the try/finally, unconditionally removes the file (`Remove-Item -Force -ErrorAction SilentlyContinue`), logs success, and sets `$script:VbaRemoved = $true`.

**Task 2 — Wiring into main flow, final report, and comment-based help** (commit `e0db2a9`):
- Added `Write-Step 'Detectando instalação VBA legada'` + `Remove-LegacyVbaAddin` call inside PASSO 0, immediately after `Assert-ExcelNotRunning` and before the Office-bitness informational checks — runs before PASSO 1 (binary resolution) and PASSO 2 (HKCU registration).
- Added a conditional block in PASSO 4's final report: `if ($script:VbaRemoved) { ... }` prints a "Migração automática:" section with the removed path, only when a legacy install was actually found and removed. Placed after "O que foi instalado" and before "Próximos passos".
- Updated `.SYNOPSIS` with a sentence documenting the new automatic migration capability.
- Updated `.DESCRIPTION`'s "FLUXO PRINCIPAL" list: inserted a new item 1 describing the legacy-VBA detection/removal step, renumbering the previous 5 items to 2-6. `.PARAMETER`/`.EXAMPLE`/`.NOTES` left untouched (no new parameters introduced).

## Verification

- Structural `grep` checks (function definition count, constants, `ReleaseComObject` count ≥3, `GC]::Collect`, `WaitForPendingFinalizers`, `$script:VbaRemoved` state, call-site count, `.SYNOPSIS`/`.DESCRIPTION` mentions) — all passed for both tasks.
- Full-file PowerShell syntax check via `[System.Management.Automation.PSParser]::Tokenize` (invoked through `powershell.exe -ExecutionPolicy Bypass -File` on the WSL2→Windows interop path) reported `SYNTAX OK` with zero parse errors after each task.
- Manual review confirmed: `Remove-LegacyVbaAddin` is defined exactly once and called exactly once, before PASSO 1/2; every `$excel`/`$wb`/`$foundAddin` obtained has a matching `ReleaseComObject`; `$foundAddin` is always initialized to `$null` before the try block; `git diff --stat` shows only `scripts/install.ps1` changed — `scripts/uninstall.ps1` and `scripts/verify-environment.ps1` are untouched.
- Live-Excel end-to-end validation (creating a genuine legacy `.xlam` registered via COM and confirming removal) is explicitly out of scope for this plan, deferred to the orchestrator per the plan's `<verification>` item 4.

## Deviations from Plan

None — plan executed exactly as written. Both tasks' `<action>` specifications were followed literally (variable names, function placement, call-site placement, report-block wording, help renumbering).

## Threat Flags

None — all new surface (Excel COM session, file removal at a fixed well-known path) was already anticipated and dispositioned in the plan's `<threat_model>` (T-quick-hbu-01 through 04), and the implementation matches the mitigations described there (unconditional COM cleanup, never-blocking inner try/catch, fixed non-attacker-controlled path).

## Self-Check: PASSED

- FOUND: scripts/install.ps1 (modified, contains `Remove-LegacyVbaAddin` function + call site)
- FOUND: commit cd364be (Task 1)
- FOUND: commit e0db2a9 (Task 2)
