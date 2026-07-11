---
phase: 03-com-entry-point-real-excel-integration
plan: 02
subsystem: com-interop
tags: [dotnet, csharp, com-interop, excel, net48, ribbon]

# Dependency graph
requires:
  - phase: 03-com-entry-point-real-excel-integration (plan 01)
    provides: "FinanceFmtTools.ComAddin project (net48-only) with real RealExcelGateway/RealRangeHandle/TraceLog implementations of Phase 2's unmodified IExcelGateway/IRangeHandle/ILog interfaces, plus the hand-rolled Extensibility.IDTExtensibility2 shim"
provides:
  - "AddInHost — composition root wiring RealExcelGateway/TraceLog/RibbonController, ApplyFormat (FMT-06 live friendly MessageBox), ShowAbout, OpenDocs, SetForceAlign/SetZeroDash with defensive IRibbonUI.InvalidateControl"
  - "Connect — the actual COM entry point Excel instantiates: IDTExtensibility2 + IRibbonExtensibility, ClassInterfaceType.AutoDispatch, fixed Guid/ProgId, all 17 Ribbon callback methods matching src/customUI14.xml exactly"
  - "A documented, itemized live-Excel smoke-test checklist (RIB-01..04) ready for the user to run on their own Windows+Excel machine — explicitly NOT executed in this environment"
affects: [phase-4-installation-registration]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Thin COM entry point delegating to a composition-root host (Connect -> AddInHost), every method a 1-3 line try/catch-wrapped delegation, matching src/modRibbon.bas's '1 line of logic per callback' convention"
    - "FormatEngine.Apply (not ApplyToSelection) called directly from AddInHost after its own gateway guard, so Phase 2's dotnet-test-proven ApplyToSelection contract is never touched to add the live MessageBox"
    - "Defensive IRibbonUI caching + InvalidateControl on checkbox toggle, even though VBA never called it — cheap insurance against an unverifiable-without-live-Excel behavior"

key-files:
  created:
    - src/FinanceFmtTools.ComAddin/AddInHost.cs
    - src/FinanceFmtTools.ComAddin/Connect.cs
  modified: []

key-decisions:
  - "Self-caught acceptance-criteria failure: Connect.cs's explanatory comment above [ClassInterface(ClassInterfaceType.AutoDispatch)] originally spelled out 'ClassInterfaceType.None' by name, making grep -c \"ClassInterfaceType.None\" return 1 instead of the required 0. Reworded to 'the \"no default dispinterface\" alternative' — no functional change, re-verified, re-built 0 Warnings/0 Errors."
  - "Task 3 (checkpoint:human-verify — full live-Excel smoke test) is explicitly NOT marked complete and NOT claimed as verified in this environment. Per 03-CONTEXT.md's non-discretionary environment constraint (Linux/WSL, no Windows/Excel/COM runtime available) and explicit orchestrator instruction, the itemized checklist is documented below for the user to run on their own Windows+Excel machine. This is expected, by-design deferral, not a plan failure."

patterns-established:
  - "Connect.cs carries a fixed-value doc-comment block (GUID/ProgId/AssemblyName/Version/discovery-key) for Phase 4's installer to consume verbatim — no registry/regasm code written in Phase 3."

requirements-completed: [RIB-01, RIB-02, RIB-03, RIB-04]

# Metrics
duration: ~15min
completed: 2026-07-11
---

# Phase 3 Plan 2: COM Entry Point & AddInHost Composition Root Summary

**`Connect` (the real COM entry point Excel instantiates: `IDTExtensibility2` + `IRibbonExtensibility`, `ClassInterfaceType.AutoDispatch`, fixed GUID/ProgId, all 17 Ribbon callbacks) now delegates to `AddInHost` (composition root wiring the real `RealExcelGateway`/`TraceLog`/`RibbonController`), completing the phase's actual deliverable — proven by `dotnet build`/`dotnet test` as far as this Linux/WSL environment allows; live-Excel behavior is explicitly deferred to human verification.**

## Performance

- **Duration:** ~15 min
- **Started:** 2026-07-11T09:55Z (immediately following 03-01's completion)
- **Completed:** 2026-07-11T09:59:55-03:00
- **Tasks:** 3 (2 auto + 1 checkpoint:human-verify, deferred per checkpoint_preapproval instructions)
- **Files modified:** 2 (both created)

## Accomplishments
- Added `src/FinanceFmtTools.ComAddin/AddInHost.cs` — the composition root. Wires the real `RealExcelGateway`/`TraceLog` (Plan 01 outputs) and a `RibbonController` (Phase 2, unmodified, constructed in the field initializer so `GetCustomUI` works even before `OnConnection` completes). `ApplyFormat(formatKey)` guards the selection itself and calls `FormatEngine.Apply` directly (never re-wraps `FormatEngine.ApplyToSelection`) — showing the FMT-06 friendly `MessageBox` ("Selecione um intervalo de células antes de aplicar a formatação.") on an invalid selection, exact VBA parity with `modUtils.bas`'s `SafeSelection`. `ShowAbout()` matches `modUtils.bas`'s `ShowAbout` text/title exactly. `OpenDocs()` uses the explicit `ProcessStartInfo(DocsUrl){UseShellExecute=true}` form. `SetForceAlign`/`SetZeroDash` mutate `RibbonController.Config` and defensively call `InvalidateControl` on the cached `IRibbonUI`.
- Added `src/FinanceFmtTools.ComAddin/Connect.cs` — the actual `[ComVisible(true)][Guid("881EFDF3-424C-4240-BCA0-714DAC2B9CD7")][ProgId("FinanceFmtTools.Connect")][ClassInterface(ClassInterfaceType.AutoDispatch)]` class Excel instantiates via COM, implementing `IDTExtensibility2` + `Office.IRibbonExtensibility`. All 17 `src/customUI14.xml` `onAction`/`getPressed` names (`RibbonInteger`, `RibbonFin2D`, `RibbonFin4D`, `RibbonFin8D`, `RibbonPct4D`, `RibbonPct2D`, `RibbonSpreadBps`, `RibbonDateISO`, `RibbonDateBR`, `RibbonDateBRLong`, `RibbonText`, `RibbonChkForceAlign`, `RibbonGetForceAlign`, `RibbonChkZeroDash`, `RibbonGetZeroDash`, `RibbonFinInfo`, `RibbonAbout`) have a matching public method, each a thin (1-3 line) try/catch-wrapped delegation to `AddInHost`. A doc-comment block above the class records the fixed GUID/ProgId/AssemblyName/Version/discovery-key values Phase 4's installer must reuse verbatim.
- Full solution (`FinanceFmtTools.Engine` net48+net8.0, `FinanceFmtTools.Engine.Tests` net8.0, `FinanceFmtTools.ComAddin` net48) builds via `dotnet build src/FinanceFmtTools.sln -c Release`: **0 Warning(s), 0 Error(s)**.
- `dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release`: **40/40 passing**, zero changes to any Engine/Engine.Tests source file.
- Task 3 (the live-Excel smoke-test checklist) is documented below, explicitly NOT executed or claimed as verified — see "Next Phase Readiness".

## Task Commits

1. **Task 1: AddInHost composition root** - `6c943ef` (feat)
2. **Task 2: Connect.cs — COM entry point + all 17 Ribbon callbacks** - `5c563b1` (feat)
3. **Task 3: Live Excel smoke test checklist (human_needed)** — not executed in this environment (no commit; deferred per checkpoint_preapproval, itemized checklist recorded below)

**Plan metadata:** (pending — committed alongside this SUMMARY)

## Files Created/Modified
- `src/FinanceFmtTools.ComAddin/AddInHost.cs` - composition root: real gateway/log/ribbon wiring, FMT-06 live MessageBox, About/Docs, checkbox setters with InvalidateControl
- `src/FinanceFmtTools.ComAddin/Connect.cs` - COM entry point: `IDTExtensibility2` + `IRibbonExtensibility`, `ClassInterfaceType.AutoDispatch`, fixed Guid/ProgId, all 17 Ribbon callbacks

## Decisions Made
- `AddInHost.ApplyFormat` calls `FormatEngine.Apply` (the lower-level Phase 2 method) directly after its own `TryGetSelectedRange` guard, rather than adding a callback/delegate parameter to `FormatEngine.ApplyToSelection` — keeps Phase 2's `dotnet test`-proven contract (log warning, never throw, never show UI) completely untouched, per 03-RESEARCH.md Pattern 3.
- `RibbonController` is constructed in `AddInHost`'s field initializer (not deferred to `Wire`), so `Connect.GetCustomUI` can return real Ribbon XML even if Excel calls it before `OnConnection` finishes — avoids a null-Ribbon race at Excel startup.
- Every Ribbon callback in `Connect` wraps its `AddInHost` call in try/catch, logging via a private `TryLog` helper that itself never throws — matching the sibling `outlook-classic-delay-send` project's defensive pattern (an unhandled exception escaping a COM entry-point method risks Excel silently disabling the add-in via Resiliency).

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 1 - Bug] `Connect.cs`'s explanatory comment caused the `ClassInterfaceType.None` grep acceptance criterion to fail**
- **Found during:** Task 2 acceptance-criteria verification
- **Issue:** The comment above `[ClassInterface(ClassInterfaceType.AutoDispatch)]` explained the choice by naming the literal string `ClassInterfaceType.None` as the rejected alternative, so `grep -c "ClassInterfaceType.None" src/FinanceFmtTools.ComAddin/Connect.cs` returned 1, not the plan's required 0.
- **Fix:** Reworded the comment to say "the 'no default dispinterface' alternative" instead of the literal type name — same explanation, no functional change.
- **Files modified:** `src/FinanceFmtTools.ComAddin/Connect.cs`
- **Verification:** `grep -c "ClassInterfaceType.None"` returns 0; `grep -c "ClassInterfaceType.AutoDispatch"` still returns 1; `dotnet build` still 0 Warnings/0 Errors after the edit.
- **Committed in:** `5c563b1` (Task 2 commit — fixed before commit, not a separate commit)

---

**Total deviations:** 1 auto-fixed (1 bug — self-inflicted acceptance-criteria failure, fixed before the task's commit).
**Impact on plan:** No scope creep; purely a comment-wording fix caught by the plan's own hard verification gate. Same pattern as Plan 01's own self-caught GUID-comment issue.

## Issues Encountered
None beyond the auto-fixed comment-wording issue above.

## User Setup Required

**None for building this code** — `dotnet build`/`dotnet test` require no external service configuration.

**Live-Excel verification is explicitly deferred to the user's own Windows+Excel machine.** This environment (Linux/WSL) has no Windows, no Excel, no COM runtime — there is no way to instantiate `Connect`, load the Ribbon, or click a button here. Per 03-CONTEXT.md's non-discretionary environment constraint and this plan's `checkpoint:human-verify` Task 3, the following itemized checklist is recorded **verbatim from the plan** for the user to run manually. It is **not** approved, faked, or assumed — it is an open item.

### Itemized live-Excel smoke-test checklist (human_needed)

1. On a Windows machine with Excel 2016+ installed, run `dotnet build src\FinanceFmtTools.sln -c Release` from the repo root — confirm 0 Warnings/0 Errors and that `src\FinanceFmtTools.ComAddin\bin\Release\net48\FinanceFmtTools.ComAddin.dll` exists alongside `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, and `office.dll` (Microsoft.Office.Core).
2. Temporary manual registration for this smoke test only (Phase 4 automates this properly later) — in PowerShell, create these HKCU keys (no admin required), pointing at the built DLL's full path:
   - `HKCU:\Software\Classes\FinanceFmtTools.Connect` (default) = `FinanceFmtTools.Connect`
   - `HKCU:\Software\Classes\FinanceFmtTools.Connect\CLSID` (default) = `{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}`
   - `HKCU:\Software\Classes\CLSID\{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}` (default) = `FinanceFmtTools.ComAddin.Connect`
   - `HKCU:\Software\Classes\CLSID\{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}\InprocServer32` (default) = `C:\Windows\System32\mscoree.dll`, `ThreadingModel` = `Both`, `Class` = `FinanceFmtTools.ComAddin.Connect`, `Assembly` = `FinanceFmtTools.ComAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null`, `RuntimeVersion` = `v4.0.30319`, `CodeBase` = `file:///` + the full path to `FinanceFmtTools.ComAddin.dll` with backslashes replaced by forward slashes
   - `HKCU:\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect`, `LoadBehavior` (DWORD) = `3`
3. Open Excel. Verify the "Finance Fmt" tab appears with groups Numérico/Percentual/Data/Texto/Info, matching `src\customUI14.xml`'s labels/tooltips exactly (RIB-01).
4. Type a number in a cell, select it, click each of the 11 format buttons in turn (Fin 8D, Fin 4D, Fin 2D, Fin 0D, % 4D, % 2D, Spread bps, ISO, BR, BR Extenso, Texto) — confirm each applies the expected `NumberFormat`.
5. Select a Chart or a Shape (not a cell range) and click any format button — confirm a MessageBox appears ("Selecione um intervalo de células antes de aplicar a formatação.") instead of a crash or silent no-op.
6. Toggle "Forçar à direita" — confirm it starts unchecked, and that toggling it changes the alignment of a subsequently-applied Fin format. Toggle "Zero contábil" — confirm it starts checked, and that toggling it changes whether a zero value in a Fin-formatted cell displays as "-" or as `0,00`/etc. Confirm both checkboxes visually reflect their own click state correctly (this specific detail was flagged in 03-RESEARCH.md as unresolvable without a live session).
7. Close and reopen Excel — confirm both checkboxes reset to their default state (unchecked / checked respectively) with no persistence.
8. Click "Sobre" — confirm a MessageBox shows "Finance Fmt Tools v1.0.0" plus the description and author line. Click "Documentação" — confirm the default browser opens `https://github.com/tpougy/finance-fmt-tools`.

**Resolution recorded when the user responds:** "approved" (every item passed) or an itemized list of which item(s) failed and how. This result should be captured in this phase's eventual VERIFICATION.md as `human_needed`.

## Next Phase Readiness
- `FinanceFmtTools.ComAddin` now contains the complete, real COM add-in: `Connect` (entry point) + `AddInHost` (composition root) + `RealExcelGateway`/`RealRangeHandle` (Plan 01) + `TraceLog` (Plan 01), all wired to Phase 1/2's unmodified `FormatEngine`/`FormatRegistry`/`RibbonController`. `dotnet build`/`dotnet test` is the ceiling of what has been verified in this Linux/WSL environment — no Windows, no Excel, no COM runtime available here.
- **Phase 3 (COM Entry Point & Real Excel Integration) is code-complete — 2/2 plans — but its own goal statement ("verified by manual smoke test, not unit tests alone") is only partially satisfiable in this environment.** The live-Excel smoke test above (RIB-01 through RIB-04) remains an open `human_needed` item, to be run and reported back by the user on their own Windows+Excel machine before this phase can be considered fully verified end-to-end.
- Phase 4 (Installation & Registration, not yet planned) can proceed once Phase 3 code is in place — it does not strictly require the live-Excel smoke test to have passed first, but should treat the smoke-test result as a leading indicator worth checking before shipping a real installer. Phase 4's installer must reuse, verbatim, the fixed values documented in `Connect.cs`'s doc-comment block: GUID `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`, ProgId `FinanceFmtTools.Connect`, AssemblyName `FinanceFmtTools.ComAddin`, Version `1.0.0.0`, discovery key `HKCU\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect`.

---
*Phase: 03-com-entry-point-real-excel-integration*
*Completed: 2026-07-11*
