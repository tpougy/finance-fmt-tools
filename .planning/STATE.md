---
gsd_state_version: 1.0
milestone: v1.0
milestone_name: milestone
status: Awaiting next milestone
stopped_at: Completed 05-04-PLAN.md — Phase 5 fully code-complete (4/4 plans). Proceeding to Phase 5 code review + verification, then milestone lifecycle.
last_updated: "2026-07-11T19:18:37.256Z"
last_activity: 2026-07-12 - Released v2.1.0 (VBA auto-migration in installer, PadHyphen removal, race-condition fix)
progress:
  total_phases: 5
  completed_phases: 5
  total_plans: 14
  completed_plans: 14
  percent: 100
---

# Project State

## Project Reference

See: .planning/PROJECT.md (updated 2026-07-10)

**Core value:** Aplicar formatos financeiros/contábeis padronizados a células do Excel com um clique — agora sobre uma base de código C# testável, com dev/build/release 100% via terminal.
**Current focus:** Awaiting next milestone. v1.0 shipped (code-complete) with 3 explicit `human_needed` open items — see Deferred Items below.

## Current Position

Phase: Milestone v1.0 complete
Plan: —
Status: Awaiting next milestone
Last activity: 2026-07-11 — Milestone v1.0 completed and archived (see `.planning/MILESTONES.md`, `.planning/milestones/v1.0-*`, `.planning/RETROSPECTIVE.md`)

## Performance Metrics

**v1.0 final:** 5 phases, 14 plans, 28 tasks, 1 continuous autonomous session (with several mid-session compactions and 3 recovered subagent session-limit interruptions). Full per-plan timing detail archived in `.planning/RETROSPECTIVE.md`.

## Accumulated Context

### Decisions

Full decision log for v1.0 is preserved in `.planning/RETROSPECTIVE.md` and `.planning/milestones/v1.0-ROADMAP.md`/`v1.0-REQUIREMENTS.md`. Condensed outcomes in PROJECT.md's Key Decisions table. Carried-forward decisions relevant to any future milestone touching this codebase:

- Fixed COM identity (GUID `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`, ProgId `FinanceFmtTools.Connect`, AssemblyName `FinanceFmtTools.ComAddin`) lives in `Connect.cs`'s header comment and must be reused verbatim by any future installer/CI changes — never re-invented.
- Real `git push`/`gh release create` against this project's real public remote requires explicit user authorization, never autonomous execution — a durable safety boundary, not a one-time v1.0 decision.
- `src/customUI14.xml` is NOT legacy — it's an active `EmbeddedResource` in `FinanceFmtTools.Engine.csproj`, unlike the `.bas` files it shipped alongside in the VBA era.

### Pending Todos

None yet.

### Blockers/Concerns

None currently open (all v1.0 blockers resolved or reclassified as Deferred Items below).

### Quick Tasks Completed

| # | Description | Date | Commit | Directory |
|---|-------------|------|--------|-----------|
| 260712-glm | Refatorar AccountingFormatBuilder para usar tabela de tokens combináveis em C# | 2026-07-12 | afe34c7 | [260712-glm-refatorar-accountingformatbuilder-para-u](./quick/260712-glm-refatorar-accountingformatbuilder-para-u/) |
| 260712-h44 | Remover token PadHyphen do formato contábil final (sobra apenas Open/CloseParenPad) | 2026-07-12 | caa55bf | [260712-h44-remover-padhyphen-do-formato-contabil-fi](./quick/260712-h44-remover-padhyphen-do-formato-contabil-fi/) |
| 260712-hbu | Adicionar detecção e remoção automática de instalação VBA legada (.xlam) a scripts/install.ps1 | 2026-07-12 | e0db2a9 | [260712-hbu-adicionar-deteccao-e-remocao-de-instalac](./quick/260712-hbu-adicionar-deteccao-e-remocao-de-instalac/) |
| 260712-hzj | Corrigir race condition (EXCEL.EXE ainda de pé) em Remove-LegacyVbaAddin, achada em teste ao vivo | 2026-07-12 | 528082f | [260712-hzj-corrigir-race-condition-no-remove-legacy](./quick/260712-hzj-corrigir-race-condition-no-remove-legacy/) |

## Deferred Items

Items acknowledged and deferred at v1.0 milestone close on 2026-07-11:

| Category | Item | Status |
|----------|------|--------|
| ~~verification_gap~~ | ~~Phase 03 — live-Excel smoke test~~ — **resolved 2026-07-12**: WSL2 has interop access to a real Windows host with Excel 16.0 (Click-to-Run x64) installed (`powershell.exe`/`cscript.exe` at `/mnt/c/Windows/System32/...`); the "no Windows+Excel available" assumption from milestone close was **wrong**. Verified live via `Excel.Application` COM automation: add-in connects (`COMAddIns(...).Connect = True`), `LoadBehavior` stable at `3` across sessions. | done |
| ~~verification_gap~~ | ~~Phase 04 — live install/uninstall test~~ — **resolved 2026-07-12**: ran the real documented one-liner install (downloads from GitHub Releases) and `scripts/uninstall.ps1` repeatedly against the real Windows host — idempotent, clean install/uninstall/reinstall cycles confirmed, registry left clean after uninstall. | done |
| ~~verification_gap~~ | ~~Phase 05 — real git push + tag + live CI run~~ — **resolved 2026-07-12**: user explicitly authorized the push; `main` + `archive/vba-legacy` pushed to `origin`, tag `v2.0.0` pushed, `.github/workflows/release.yml` ran green on `windows-latest`, release published and asset verified (`gh release download` + zip integrity check, 7/7 files) | done |

**All three original Deferred Items are now closed.** However, live testing (Phase 03/04) surfaced a **critical bug that the entire v1.0 test suite, code review, and verification process had missed**: the add-in never actually connected in real Excel. `LoadBehavior` silently downgraded `3 -> 2` on first load — no managed exception, no Windows Event Log entry, `dotnet test` 40/40 green throughout. Root cause: the hand-rolled `Extensibility.IDTExtensibility2` COM shim (`src/FinanceFmtTools.ComAddin/Extensibility.cs`) was missing `[DispId]`/`[In]`/`[MarshalAs]` attributes the real COM interface requires (verified by reflecting on a genuine `Extensibility.dll` from the sibling project and diffing exact method signatures). `QueryInterface` for the interface succeeded either way, which is why this looked like nothing was wrong — the vtable ABI itself was incompatible, not the interface identity. Fixed and released as **v2.0.1** (2026-07-12), confirmed working end-to-end via the real documented one-liner installer against real Excel.

**Process lesson**: none of Phases 1-5's automated tests (unit tests, code review, plan-checker, verifier) could have caught this — it only surfaces when a real native host (Excel) tries to call through the vtable. Pure managed-code testing and even direct COM activation/QueryInterface checks are insufficient to validate a hand-rolled classic COM interop shim; only a real native caller exercising the actual method calls proves it. See `RELEASE_NOTES.md`'s v2.0.1 entry for the full writeup.

## Session Continuity

Last session: 2026-07-11T19:30:00-03:00
Stopped at: v1.0 milestone completed and archived. Awaiting next milestone.
Resume file: None

## Operator Next Steps

- **v2.1.0 released 2026-07-12** (https://github.com/tpougy/finance-fmt-tools/releases/tag/v2.1.0) — current recommended release. Three changes, all via `/gsd-quick` tasks (260712-glm, h44, hbu, hzj — see Quick Tasks Completed above):
  1. `FormatTokens.cs` extracted (refactor, no behavior change) — named constants for every combinable piece of the accounting number-format string.
  2. `PadHyphen` token ("_-") removed from the final accounting format — only `OpenParenPad`/`CloseParenPad` ("_(" / "_)") remain as padding. Real behavior change (visual: one less trailing blank char), all 40 tests updated and passing.
  3. `scripts/install.ps1` now detects a legacy VBA install (`FinanceFmtTools.xlam` in `%APPDATA%\Microsoft\AddIns`), unregisters it from Excel via COM (`AddIns.Installed = $false`) and deletes the file, automatically, before installing the C# version — no manual step needed anymore (supersedes the manual instruction published in v2.0.0's release notes).
  - A real bug (race condition: `Remove-LegacyVbaAddin` didn't wait for the `EXCEL.EXE` process it spawned to actually exit before returning, causing an intermittent "Excel is open" failure in the next step) was found via live-Excel testing and fixed before release (260712-hzj).
  - Verified end-to-end twice against real Excel via WSL2 interop: once with local `-Source` binaries, once with the real documented one-liner installer downloading the actual published v2.1.0 asset from GitHub Releases. Both times: legacy VBA fixture created via genuine Excel COM automation (Title="Finance Fmt Tools", registered `Installed=True`) → detected → unregistered → file deleted → C# add-in installed → `Connect=True`, `LoadBehavior=3` → no VBA residue left in the `AddIns` collection or on disk.
- Start the next milestone with /gsd-new-milestone
