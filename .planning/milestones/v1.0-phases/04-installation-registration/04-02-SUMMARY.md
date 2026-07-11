---
phase: 04-installation-registration
plan: 02
subsystem: installer
tags: [powershell, hkcu, com-registration, uninstaller, diagnostics]

# Dependency graph
requires:
  - phase: 04-installation-registration (plan 01)
    provides: "The exact identity constants and 3 HKCU registry tree shapes install.ps1 writes, which uninstall.ps1 must remove byte-identically"
provides:
  - "scripts/uninstall.ps1 — idempotent removal of all 3 HKCU registry trees install.ps1 writes (COM class, Excel discovery, Resiliency value-only) plus the 4 installed files (INST-02)"
  - "scripts/verify-environment.ps1 — read-only diagnostic (Windows/Excel/Office-bitness/.NET-Framework-4.8/.NET-SDK/PowerShell/Git/VS-Code), discretionary per 04-RESEARCH.md"
affects: [phase-4-plan-03-live-verification]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Resiliency removal never deletes the shared DoNotDisableAddinList parent key — only Remove-ItemProperty on this add-in's own named value, since other add-ins may share the same key"
    - "File removal iterates the exact named file list and only removes the install directory itself once confirmed empty afterward — never a blanket recursive delete"

key-files:
  created:
    - scripts/uninstall.ps1
    - scripts/verify-environment.ps1
  modified: []

key-decisions:
  - "Wrote both files directly in the main conversation (same recovery approach as Plan 01), using the sibling project's install.ps1/verify-environment.ps1 as structural templates adapted from Outlook to Excel nouns and this project's fixed COM identity."
  - "uninstall.ps1 has no -RemoveLogs concept (unlike the sibling project) — this add-in's TraceLog writes via System.Diagnostics.Trace, not to a logs\\ folder on disk, so there is nothing file-based to preserve/purge."

patterns-established: []

requirements-completed: [INST-02]

# Metrics
duration: unknown (written directly in main context, no standalone timer)
completed: 2026-07-11
---

# Phase 4 Plan 2: uninstall.ps1 + verify-environment.ps1 Summary

**`scripts/uninstall.ps1` (176 lines) idempotently removes every HKCU registry tree and file `install.ps1` writes, never touching the shared `DoNotDisableAddinList` key wholesale; `scripts/verify-environment.ps1` (279 lines) is a genuinely read-only diagnostic reporting Windows/Excel/.NET Framework 4.8/.NET SDK/PowerShell/Git/VS Code status — both verified via every grep-based structural check embedded in the plan.**

## Performance

- **Duration:** unknown — written directly in main context (no standalone executor timer)
- **Completed:** 2026-07-11
- **Tasks:** 2 (both `type="auto"`)
- **Files modified:** 2 (both created)

## Accomplishments
- `scripts/uninstall.ps1`: declares the same fixed identity constants as `install.ps1` (independently, no shared module), removes the CLSID subtree and ProgId mapping key via `Remove-KeyIfExists`, removes the non-versioned Excel discovery key, and removes *only* the named Resiliency value (`Remove-ItemProperty`, never `Remove-Item` on the parent `DoNotDisableAddinList` key). Removes the 4 installed files by exact name, then removes the install directory only if it is empty afterward. Exits 0 unconditionally except when Excel is running and `-Force` was not passed.
- `scripts/verify-environment.ps1`: read-only diagnostic with zero registry/file-write cmdlets. Checks Windows version, Excel presence + PE bitness (`Test-PeMachine`), Office Click-to-Run bitness, `.NET SDK`, `.NET Framework 4.8` (`NDP\v4\Full` Release ≥ 528040), MSBuild, PowerShell version, Git, VS Code + C# extension, and winget. `-RuntimeOnly` switch narrows BUILD-time items (SDK/Git/VS Code) to informational-only, for checking a target (non-developer) machine before running `install.ps1`. No `#Requires -Version 5.1` gate (deliberately, so it can report on older PowerShell versions too).
- All plan-embedded automated `<verify>` grep checks pass for both files (brace-balance 45/45 and 88/88, GUID count 1 in uninstall.ps1, `Remove-KeyIfExists` ×4+, `Remove-ItemProperty` present, Resiliency value-only removal confirmed, non-versioned discovery key correct, `EXCEL.EXE`/no `OUTLOOK.EXE` in verify-environment.ps1, `NDP\v4\Full` present, `-RuntimeOnly` present, both `exit 0`/`exit 1` paths present, zero write cmdlets).

## Task Commits

1. **Task 1 + Task 2 (uninstall.ps1 and verify-environment.ps1, written together):** `90876f4` (feat)

## Files Created/Modified
- `scripts/uninstall.ps1` - new file, 176 lines
- `scripts/verify-environment.ps1` - new file, 279 lines

## Decisions Made
- See `key-decisions` above.

## Deviations from Plan

None.

## Issues Encountered
- Same subagent-session-limit recovery approach as Plan 01: written directly in the main conversation rather than dispatching a fresh executor, since the plan's task instructions (exact constants, exact registry shapes, exact check list) were explicit enough to implement directly and verify with the plan's own grep-based checks.

## User Setup Required
None for this plan's own scope. Live uninstall/idempotency behavior on a real Windows+Excel machine remains `human_needed`, tracked in Plan 03.

## Next Phase Readiness
- Plan 03's live-verification checkpoint can now exercise `install.ps1` → `uninstall.ps1` → `install.ps1` (idempotency) → `uninstall.ps1` on a real machine.
- No blockers. Phase 4's three plans (01, 02, 03) are all either complete or, for 03, a `human_needed` checkpoint — ready for phase-level code review and `VERIFICATION.md`.

---
*Phase: 04-installation-registration*
*Completed: 2026-07-11*
