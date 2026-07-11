---
phase: 04-installation-registration
plan: 03
subsystem: installer
tags: [powershell, hkcu, com-registration, human-verify]

# Dependency graph
requires:
  - phase: 04-installation-registration (plan 01, plan 02)
    provides: "scripts/install.ps1 and scripts/uninstall.ps1, structurally verified but never executed"
provides:
  - "Automated precondition (dotnet build) confirmed green; itemized live-Excel install/uninstall/idempotency/Resiliency checklist recorded as an explicit human_needed item, not faked or assumed"
affects: [phase-4-verification, milestone-completion]

# Tech tracking
tech-stack:
  added: []
  patterns: []

key-files:
  created: []
  modified: []

key-decisions:
  - "This plan's Task 1 is a checkpoint:human-verify gate=\"blocking\" task requiring a real Windows+Excel machine — genuinely not executable in this Linux/WSL environment (no Windows, no Excel, no registry, no PowerShell). Per this session's full-autonomy directive and the identical precedent set by Phase 3's 03-02-SUMMARY.md, this is resolved by running every automatable precondition here, then recording the itemized checklist verbatim as an open human_needed item rather than pausing to block the milestone or fabricating a pass result."

patterns-established: []

requirements-completed: []

# Metrics
duration: unknown
completed: 2026-07-11
---

# Phase 4 Plan 3: Live Install/Uninstall/Idempotency/Resiliency Verification Summary

**The one automatable precondition (`dotnet build src/FinanceFmtTools.sln -c Release`) is confirmed green with all 4 files `install.ps1`/`uninstall.ps1` expect present in the build output. The actual live-Excel checklist (fresh install, Ribbon tab appearance, idempotent re-install, the Resiliency behavioral test, uninstall, idempotent uninstall) cannot run in this Linux/WSL environment and is recorded here as an explicit, unresolved `human_needed` item — not approved, faked, or assumed.**

## Performance

- **Duration:** N/A — automated precondition check only; the task's own substance is a human checkpoint
- **Completed:** 2026-07-11 (precondition verified; live checklist still open)
- **Tasks:** 1 (`checkpoint:human-verify`, `gate="blocking"`)
- **Files modified:** 0

## Accomplishments
- Ran `dotnet build src/FinanceFmtTools.sln -c Release`: **0 Warning(s), 0 Error(s)**.
- Confirmed `src/FinanceFmtTools.ComAddin/bin/Release/net48/` contains all 4 files both scripts require by exact name: `FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, `office.dll` — validating the `$AllFiles` assumption baked into both `install.ps1` and `uninstall.ps1` against the actual build output, in this environment, before deferring the rest.

## Task Commits

None — this task modifies no files (verification-only), per its own `<files>` declaration (`N/A`).

## Files Created/Modified
None.

## Decisions Made
- See `key-decisions` above.

## Deviations from Plan

None — this plan's own design anticipates and explicitly permits a `human_needed` outcome for its Task 1; recording the checklist as open is the expected, by-design result in this environment, not a deviation.

## Issues Encountered
None beyond the expected environment limitation the plan itself documents (04-CONTEXT.md's non-discretionary no-Windows/PowerShell constraint).

## User Setup Required

**None for the precondition** — `dotnet build` requires no external service configuration and is already confirmed green above.

**Live install/uninstall verification is explicitly deferred to the user's own Windows+Excel machine.** This environment (Linux/WSL) has no Windows, no Excel, no PowerShell, no registry — there is no way to actually run `install.ps1`/`uninstall.ps1` or observe Excel's Ribbon/Resiliency behavior here. Per 04-CONTEXT.md's non-discretionary environment constraint and this plan's `checkpoint:human-verify` Task 1, the following itemized checklist is recorded **verbatim from the plan** for the user to run manually. It is **not** approved, faked, or assumed — it is an open item.

### Itemized live-Excel install/uninstall checklist (human_needed)

1. Build the add-in locally: run `dotnet build src\FinanceFmtTools.sln -c Release` from the repo root on the Windows machine. Confirm `src\FinanceFmtTools.ComAddin\bin\Release\net48\` contains `FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, and `office.dll`.
2. (Optional but recommended) Run `powershell -ExecutionPolicy Bypass -File .\scripts\verify-environment.ps1 -RuntimeOnly` and confirm it reports `.NET Framework 4.8` and `Excel` as `[OK]`, exiting 0.
3. **FRESH INSTALL (INST-01, local-testing escape hatch):** close Excel if open, then run `powershell -ExecutionPolicy Bypass -File .\scripts\install.ps1 -Source src\FinanceFmtTools.ComAddin\bin\Release\net48`. Confirm the script reports `[OK]` for every registry key/file check in its post-install validation section and exits 0. Confirm no admin/UAC prompt ever appeared (every write targets `HKCU`).
4. **RIBBON TAB APPEARS (INST-01):** open Excel. Confirm the "Finance Fmt" tab appears with the same groups/buttons/tooltips as Phase 3's smoke test already confirmed (Numérico/Percentual/Data/Texto/Info). Close Excel.
5. **IDEMPOTENT RE-INSTALL (roadmap criterion #4):** with Excel closed, run the exact same `install.ps1 -Source ...` command again. Confirm it completes without error (exit 0), does not prompt about "already exists," and the Ribbon tab still appears correctly on the next Excel launch.
6. **RESILIENCY BEHAVIOR TEST (INST-03, the one item static review cannot prove — 04-RESEARCH.md Pitfall 2 / Open Question 1):** with the add-in installed, deliberately force a slow or crashing load once (e.g. temporarily rename `FinanceFmtTools.Engine.dll` so `FinanceFmtTools.ComAddin.dll` fails to load, or introduce a temporary unhandled exception in `Connect.OnConnection`, then restore it after this check). Launch Excel with the broken DLL in place, let it fail/crash once, then restore the correct DLL and relaunch Excel again. Confirm Excel's next launch does NOT show its native "this add-in has been disabled" notification, and the Ribbon tab still loads normally — this is the actual behavioral proof that `DoNotDisableAddinList` is honored for Excel, not just a documentation-corroborated guess.
7. **UNINSTALL (INST-02):** close Excel, run `powershell -ExecutionPolicy Bypass -File .\scripts\uninstall.ps1`. Confirm it reports removal of all 3 registry trees and the installed files, and exits 0. Open Excel and confirm the "Finance Fmt" tab no longer appears.
8. **IDEMPOTENT UNINSTALL (roadmap criterion #4):** with the add-in already uninstalled, run `uninstall.ps1` again. Confirm it reports "already absent" for every key/file and still exits 0 (no error).

**Resolution recorded when the user responds:** "approved" (every item passed) or an itemized list of which item(s) failed and how. This result should be captured in this phase's `VERIFICATION.md` as `human_needed`.

## Next Phase Readiness
- Phase 4 (Installation & Registration) is now code-complete — 3/3 plans (install.ps1, uninstall.ps1 + verify-environment.ps1, and this recorded human-verify checkpoint) — mirroring Phase 3's exact pattern: fully buildable/statically-verified here, with its own goal statement's live-machine confirmation left as an explicit open item.
- Phase 4 code review and `VERIFICATION.md` can now proceed against the 3 PowerShell scripts (no C# files were touched in this phase).
- Phase 5 (CI/CD Pipeline & Release Runbook) has no hard dependency on this checklist having passed — it can proceed in parallel with the user eventually running this checklist — but Phase 5's CI must still publish the `FinanceFmtTools.zip` asset name this phase's `install.ps1` already assumes.

---
*Phase: 04-installation-registration*
*Completed: 2026-07-11*
