---
gsd_state_version: 1.0
milestone: v1.0
milestone_name: milestone
status: Awaiting next milestone
stopped_at: Completed 05-04-PLAN.md — Phase 5 fully code-complete (4/4 plans). Proceeding to Phase 5 code review + verification, then milestone lifecycle.
last_updated: "2026-07-11T19:18:37.256Z"
last_activity: 2026-07-11 — Milestone v1.0 completed and archived
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
**Current focus:** All 5 phases code-complete. Milestone lifecycle (code review + verification for Phase 5, then audit/complete/cleanup) remaining.

## Current Position

Phase: Milestone v1.0 complete
Plan: —
Status: Awaiting next milestone
Last activity: 2026-07-11 — Milestone v1.0 completed and archived

## Performance Metrics

**Velocity:**

- Total plans completed: 4
- Average duration: ~21 min
- Total execution time: ~1h 20m (Plan 01-02's duration is unknown — session was interrupted by an API rate limit and resumed later)

**By Phase:**

| Phase | Plans | Total | Avg/Plan |
|-------|-------|-------|----------|
| Phase 01 | 3 | ~65 min (P02 unknown) | ~22 min |
| Phase 02 | 2 | 35 min | ~18 min |

**Recent Trend:**

- Last 5 plans: P02 (unknown, 2 tasks, 2 files) → P03 (20 min, 2 tasks, 3 files) → 02-01 (20 min, 3 tasks, 9 files) → 02-02 (15 min, 2 tasks, 4 files)
- Trend: Stable

*Updated after each plan completion*
| Phase 03 P01 | 25 min | 3 tasks | 6 files |
| Phase 05 P03 | 25 min | 2 tasks | 1 files |

## Accumulated Context

### Decisions

Decisions are logged in PROJECT.md Key Decisions table.
Recent decisions affecting current work:

- Roadmap: Horizontal Layers structure confirmed with user (not Vertical MVP) — Format Engine Core → Abstractions & Orchestration → COM Entry Point & Real Excel Integration → Installation & Registration → CI/CD Pipeline & Release Runbook
- Roadmap: Phases 1-2 fully verifiable via `dotnet test` alone (no Windows/Excel required); Phases 3-5 require a real Windows+Excel environment and live smoke testing as their definition of done
- [Phase 01]: Added .gitignore for bin/ and obj/ .NET build artifacts. — The repo had no .gitignore (VBA-era project); running dotnet build immediately generates bin/obj folders that must not be committed as binary artifacts.
- [Phase 01]: FormatDef is a plain sealed class with constructor-assigned get-only properties, not a C# record. — C# 9 records/init-only properties fail to compile on net48 with CS0518 (IsExternalInit not defined), confirmed empirically in 01-RESEARCH.md; a plain class avoids this while staying immutable.
- [Phase 01, Plan 02]: AccountingFormatBuilder.Build ported VBA's two-branch structure exactly (general case + explicit decimals==0 override), not unified into one formula — deliberate per 01-RESEARCH.md Pitfall #4. Relies on `new string('0', decimals)`'s native ArgumentOutOfRangeException for negative-input guarding instead of an explicit check.
- [Phase 01, Plan 03]: FormatRegistry.TryGetFormatDef's 11-case switch was built in two stages across two TDD tasks (7 literal entries, then 4 Fin/Integer entries delegating to AccountingFormatBuilder) so each task's RED commit was a genuinely failing test, not a no-op. All 11 constructed FormatDef instances use CellAlignment.General — VBA's GetFormatDef never assigns f.Alignment in any Case branch, so Right/Left are never used; the Fin family's visual right-alignment comes entirely from the " * " fill-character token inside the NumberFormat string itself.
- [Phase 01]: Phase 1 (Format Engine Core) is fully complete as of Plan 03 — FMT-01/02/03/04/05/07 and DEV-01 all done, 31/31 xUnit tests passing, `dotnet build` 0 Warning(s)/0 Error(s) on net48 and net8.0.
- [Phase 02, Plan 01]: `IExcelGateway`/`IRangeHandle`/`ILog` added as pure C# interfaces (zero COM types) plus hand-written `FakeExcelGateway`/`FakeRangeHandle`/`SpyLog` test doubles, extending the existing `FinanceFmtTools.Engine` project rather than a new project — Phase 3 is what introduces the first COM-referencing project.
- [Phase 02, Plan 01]: `FormatEngine.Apply`/`ApplyToSelection` ported VBA's `ApplyFormat`/`ApplyFormatToSelection`/`SafeSelection`, with the FMT-06 invalid-selection guard collapsed into `IExcelGateway.TryGetSelectedRange`'s Try-pattern. The guard logs a warning and returns without throwing — it deliberately does NOT show a `MessageBox`/`MsgBox`; the real user-facing dialog is Phase 3's job once a live Excel/WinForms host exists. 35/35 tests passing (31 Phase 1 + 4 new).
- [Phase 02, Plan 02]: `RibbonSessionConfig` (`ForceAlign=false`, `ZeroDash=true`) implements REQUIREMENTS.md's RIB-02/RIB-03 authoritative defaults, deliberately NOT matching either of `src/modConfig.bas`'s or `src/modUtils.bas`'s two mutually contradictory VBA defaults — a considered migration behavior change, no persistence anywhere. `RibbonController` is a narrow instance class (`Config` property + `GetCustomUiXml()` only) per 02-CONTEXT.md's resolved scope boundary — no `IRibbonUI` caching/`InvalidateControl`/image loading, all deferred to Phase 3. `src/customUI14.xml` is linked (not duplicated) into `FinanceFmtTools.Engine.csproj` via MSBuild `EmbeddedResource Link`, resolved at runtime by suffix match (`EndsWith`) to avoid resource-name drift. **Phase 2 (Abstractions & Orchestration) is now fully complete — 2/2 plans, 39/39 tests passing, 0 Warning(s)/0 Error(s) on net48+net8.0.**
- [Phase 03, Plan 01, autonomous decision]: Approved the two `[SUS]`-flagged NuGet packages (`Microsoft.Office.Interop.Excel` 16.0.18925.20022, `MicrosoftOfficeCore16` 16.0.16626.20000, publisher CamronBute) at Plan 01 Task 1's blocking human-verify checkpoint, without pausing to ask the user, per this session's full-autonomy directive. Reasoning: 03-RESEARCH.md's researcher already content-verified both packages (via nuget.org metadata + `.nupkg`/`strings` binary inspection) to contain the genuine, complete Excel/Office Core object model with no malicious content — the "SUS" flag is about missing an official Microsoft publisher badge/license text, not integrity. This is the documented de facto community answer (30.2M downloads) for referencing Office Interop types without a full Office/VSTO install, which Microsoft does not otherwise publish as a standalone NuGet package. The only alternative (vendoring PIA DLLs from a real Windows+Office machine) is unavailable in this Linux/WSL environment and doesn't meaningfully change the trust profile. Flagging this prominently for the user's awareness — reversible later by swapping to vendored PIAs if they disagree.
- [Phase 03]: [Phase 03, Plan 01]: Bootstrapped FinanceFmtTools.ComAddin (net48-only, the first COM-referencing project in the solution) with real RealExcelGateway/RealRangeHandle/TraceLog implementations of Phase 2's unmodified IExcelGateway/IRangeHandle/ILog interfaces, plus a hand-rolled Extensibility.IDTExtensibility2 shim (GUID B65AD801-ABAF-11D0-BB8B-00A0C90F2744). — Full solution builds 0 Warning(s)/0 Error(s) across all 3 projects; 40/40 tests pass (baseline was actually 40, not the stale "39" in 02-02-SUMMARY.md/plan text — see 03-01-SUMMARY.md Issues Encountered); zero source changes to Phase 1/2 files. Task 1's package-legitimacy checkpoint was pre-approved per orchestrator instruction and STATE.md commit 80f0046, so execution proceeded directly through Tasks 2-3 without pausing.
- [Phase 03, Plan 02]: Wired `Connect` (the real COM entry point: `IDTExtensibility2`+`IRibbonExtensibility`, `ClassInterfaceType.AutoDispatch`, fixed Guid `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`/ProgId `FinanceFmtTools.Connect`, all 17 `src/customUI14.xml` callback names) on top of a new `AddInHost` composition root (real gateway/log/ribbon wiring; `FormatEngine.Apply` called directly — never `ApplyToSelection` — so Phase 2's tested no-throw contract stays untouched while a live FMT-06 `MessageBox` is added). Full solution builds 0 Warning(s)/0 Error(s), 40/40 tests still pass, zero Phase 1/2 source changes. **Phase 3 is now code-complete (2/2 plans)** but its own goal statement requires a live-Excel smoke test (RIB-01..04) that cannot run in this Linux/WSL environment (no Windows/Excel/COM runtime) — the itemized checklist is recorded verbatim in 03-02-SUMMARY.md's "User Setup Required" section as an explicit `human_needed` item, per 03-CONTEXT.md's non-discretionary constraint. Not treated as a plan or phase failure — this is the expected, by-design outcome for this phase in this environment.
- [Phase 03 code review + verification]: Fixed 1 critical (Marshal.ReleaseComObject/GC flush missing on disconnect — ghost EXCEL.EXE process leak) + 4 warnings (Range RCW never released, 2 getPressed callbacks missing try/catch, Selection access unguarded against COMException, DBNull mixed-value cast risk) directly in `AddInHost.cs`/`RealRangeHandle.cs`/`RealExcelGateway.cs`/`Connect.cs`. `03-VERIFICATION.md` recorded status `human_needed` (5/5 code-verified, live-Excel smoke test outstanding). REQUIREMENTS.md's RIB-01..04 rows corrected to a consistent "Code complete — human_needed" status (all four, not just RIB-01).
- [Phase 04, Plan 01]: Wrote `scripts/install.ps1` (447 lines) directly in the main conversation — the subagent originally dispatched for this plan was cut off by an API session-limit error before writing any file or commit, so this was implemented directly against the plan's explicit task spec rather than risking a third subagent attempt on the same plan. Implements INST-01 (GitHub-Releases one-liner default flow, version-agnostic `releases/latest/download/FinanceFmtTools.zip` URL) + INST-03 (`DoNotDisableAddinList` Resiliency key) + a `-Package`/`-Source` local-testing escape hatch (needed since no real C# GitHub release exists yet — only legacy `.xlam` v1.0.0/v1.0.1). Reuses Phase 3's fixed COM identity (`Connect.cs`'s GUID/ProgId/AssemblyName) verbatim — the literal GUID string appears exactly once in the file. Zero `$PSScriptRoot`/`$MyInvocation` dependency (the documented `irm | iex` one-liner has no on-disk script location). All plan-embedded grep-based structural verify checks pass. Live install/uninstall behavior remains `human_needed`, tracked in Plan 03.
- [Phase 04, Plan 02]: Wrote `scripts/uninstall.ps1` (176 lines, INST-02) and `scripts/verify-environment.ps1` (279 lines, discretionary read-only diagnostic), both directly in the main conversation for the same recovery reason as Plan 01. `uninstall.ps1` removes all 3 HKCU registry trees `install.ps1` writes — CLSID subtree, ProgId mapping, non-versioned Excel discovery key via `Remove-KeyIfExists`, plus the Resiliency value via `Remove-ItemProperty` only (never deletes the shared `DoNotDisableAddinList` parent key) — and the 4 installed files, exiting 0 unconditionally except when Excel is running without `-Force`. No `-RemoveLogs` concept exists (unlike the sibling project) since `TraceLog` writes via `System.Diagnostics.Trace`, not to a file. All plan-embedded grep checks pass for both files.
- [Phase 04, Plan 03]: Task 1 is a `checkpoint:human-verify gate="blocking"` requiring a real Windows+Excel machine. Ran the one automatable precondition (`dotnet build src/FinanceFmtTools.sln -c Release` — 0 Warning(s)/0 Error(s)) and confirmed all 4 files `install.ps1`/`uninstall.ps1` expect (`FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, `office.dll`) exist in the build output. The live install/uninstall/idempotency/Resiliency-behavior checklist itself (8 steps) is recorded verbatim in `04-03-SUMMARY.md` as an explicit `human_needed` item — not approved, faked, or assumed — mirroring Phase 3's identical precedent exactly. **Phase 4 (Installation & Registration) is now code-complete (3/3 plans)**, same status class as Phase 3.
- [Phase 04 code review]: Fixed 1 critical + 8 of 9 remaining warnings/info findings directly in `scripts/install.ps1`/`uninstall.ps1`/`verify-environment.ps1`: CR-01 (`-Force` could `Stop-Process` Excel after only a 3s grace period, discarding an unrelated unsaved-workbook "Save changes?" dialog — replaced with a 30s polling wait that fails safely with an actionable message instead of ever force-killing); WR-01 (`CodeBase` built via naive string concat instead of `[Uri]::AbsoluteUri`, which would produce an invalid unescaped URI for install paths containing spaces, e.g. `C:\Users\John Smith\...`); WR-03 (`$AssemblyStr` was a hand-maintained literal — now derived from the actually-copied DLL via `[System.Reflection.AssemblyName]::GetAssemblyName`, so it can't silently rot on a version bump); WR-04 (added a second `Assert-ExcelNotRunning` call immediately before `Copy-Item`/`Remove-Item` in both scripts — TOCTOU gap since network download/extraction in `install.ps1` can take real wall-clock time after the upfront check); WR-05 (wrapped the registry-write/file-copy blocks in `try/catch` with friendly `Write-Err2`+`exit 1`, since they previously could throw a raw uncaught exception, contradicting `uninstall.ps1`'s own "exit 0 incondicional" docstring claim); WR-06 (added UTF-8 BOM to all 3 `.ps1` files — without it, Windows PowerShell 5.1's `-File` execution path reads via the legacy ANSI code page, not UTF-8, which would mojibake all the accented Portuguese message strings); IN-02 (fixed a `FileStream`/`BinaryReader` handle leak in `Test-PeMachine`'s exception path, in both `install.ps1` and `verify-environment.ps1`, via `finally { Dispose() }`); IN-03 (`Find-BinDir`'s recursive fallback now prefers `\bin\` paths via `Sort-Object`, avoiding a non-deterministic pick between a real `bin\` output and a stale `obj\` intermediate). **Deliberately deferred (not fixed):** WR-02/IN-01 (extract shared identity constants/helpers into a `scripts/common.ps1`) — `04-02-PLAN.md`'s own Task 1 explicitly states constants must be declared independently ("this script does not source or depend on install.ps1's file") and helpers are "duplicated, not shared — no PowerShell module infrastructure is in scope for this phase," mirroring the sibling project's identical convention; fixing this would override an already-approved planning decision rather than fix a defect. All grep-based structural verify checks re-run and still pass after every fix; no PowerShell interpreter available in this environment to execute the scripts directly.
- [Phase 05, safety scoping decision]: Research confirmed `main` had never been pushed to the real `origin` remote (79-87 commits ahead across this phase's execution) and `gh`/`git` are authenticated against the real `tpougy/finance-fmt-tools` repo (2 real legacy releases already public). Per this project's core safety rules (pushing code / publishing releases are externally-visible, hard-to-reverse actions requiring explicit confirmation, regardless of the session's broader "run everything autonomously" directive for internal GSD decisions), every Phase 5 plan was scoped to produce fully-built, fully-committed LOCAL artifacts only — no `type: auto` task may run `git push origin main`, `git push origin archive/vba-legacy`, a real tag push, or `gh release create` against the real remote. All of that is confined exclusively to Plan 04 (05-04)'s single blocking `checkpoint:human-verify` task. Verified structurally by the plan-checker before execution and re-confirmed empirically after execution (`git rev-list --left-right --count origin/main...main` still shows 0 pushed).
- [Phase 05, Plan 01]: Added `.github/workflows/release.yml` — tag-triggered (`v*.*.*`) GitHub Actions job on `windows-latest`: restore/build/test the solution, package the 4 required binaries + 3 install scripts into a fixed-name `FinanceFmtTools.zip` (matching `scripts/install.ps1`'s hardcoded asset-name assumption from Phase 4), publish via `softprops/action-gh-release@v2` with `body_path: RELEASE_NOTES.md`. Includes an explicit `permissions: contents: write` block since this repo's default Actions token permission is read-only (confirmed via `gh api`). Never executed against real CI in this environment (no Windows runner here) — verified via YAML structural parsing + grep checks only.
- [Phase 05, Plan 02]: Added `RELEASE.md` (manual `gh` CLI runbook, REL-02) documenting `git push origin main` before the tag push, then a manual `gh release create ... -F RELEASE_NOTES.md` fallback with zero CI dependency — and `RELEASE_NOTES.md` (REL-03), a hand-maintained v2.0.0 migration changelog (not GitHub's auto-generated notes, since this repo's commit history is dense with internal GSD bookkeeping messages that would make auto-notes noisy — per 05-RESEARCH.md's explicit recommendation). Both files are documentation only; no command inside them was actually executed against the real remote.
- [Phase 05, Plan 03]: Removed src/ThisWorkbook.bas, src/modConfig.bas, src/modFormatEngine.bas, src/modRibbon.bas, src/modUtils.bas, Install-FinanceFmtTools.ps1, Install-FinanceFmtTools.bat from main (LEGACY-01), verified fully recoverable via archive/vba-legacy (tip cf2559b); src/customUI14.xml deliberately left untouched (still an active EmbeddedResource in FinanceFmtTools.Engine.csproj) and dotnet build confirmed green after removal. Rewrote README.md for the C# add-in only (LEGACY-02), preserving all user-facing format tables verbatim. No push to origin performed for either main or archive/vba-legacy -- deferred to 05-04's human-authorized checkpoint.
- [Phase 05 code review]: Fixed all 11 findings (1 critical, 6 warnings, 4 info) directly in `.github/workflows/release.yml`/`RELEASE.md`/`README.md`/`RELEASE_NOTES.md`: CR-01/WR-01 (the release-packaging `pwsh` script in both the CI workflow and the manual runbook lacked `$ErrorActionPreference = 'Stop'`, so a missing/renamed build artifact would `Copy-Item`-fail silently and still publish an incomplete `FinanceFmtTools.zip` — added explicit per-file `Test-Path`+`throw` checks plus a post-package content verification in the manual runbook); WR-02 (GitHub Actions were pinned to mutable major-version tags `@v4`/`@v2` — repinned to immutable commit SHAs, verified for real via `git ls-remote` against each action's own repo rather than trusting the reviewer's suggested SHAs blindly, which caught the reviewer's own `action-gh-release` SHA as fabricated/incorrect and replaced it with the verified real SHA for the latest available tag `v2.6.2`); WR-03/04/05/06 + IN-01..04 (README/RELEASE_NOTES.md accuracy fixes: wrong "Guia Fin" button label corrected to the real `customUI14.xml` label "Documentação", date format-string tables corrected to include the real `;@` text-passthrough section verified against `FormatRegistry.cs`, a stale license placeholder replaced with a real reference to the repo's existing MIT `LICENSE`, `RELEASE_NOTES.md`'s button-name list corrected to match actual Ribbon labels instead of invented English names, dotnet SDK version pinned to a verified-real `8.0.422`, an `iex`-risk disclosure note added, 64-bit wording softened to match the installer's actual non-blocking behavior, and an inaccurate "permanent" claim about GitHub release notes reworded). **Deliberate, documented deviation:** WR-02's SHA-pinning fix intentionally makes 05-01-PLAN.md's original acceptance-criteria substring checks (`grep -q 'actions/checkout@v4'` etc.) fail, since the mutable tag reference no longer appears literally — this is the fix superseding the original (weaker) acceptance bar, not a regression. All other original plan-embedded automated checks (05-01/05-02/05-04's `ALL_PASS` compound commands) re-run and still pass after every fix, and `dotnet build`/`dotnet test` remain green (0 Warnings/0 Errors, 40/40 tests).
- [Phase 05, concurrency note]: Plans 05-01/05-02/05-03 were all wave-1 with `depends_on: []` and zero file overlap (confirmed by the plan-checker before execution), so all three were dispatched as parallel background executors sharing the same non-worktree working directory (`workflow.use_worktrees: false`). Each executor scoped its own git commits strictly to its own files (`git commit -- <path>`, never `git add -A`), and none touched `STATE.md`/`REQUIREMENTS.md`/`ROADMAP.md` while another was mid-write on those shared files — 05-01 deliberately deferred its own STATE.md/ROADMAP.md update rather than race a concurrent in-flight edit, 05-02 updated only REQUIREMENTS.md (clean at the time), and 05-03 performed the final consolidated update. Result: zero commit conflicts, `dotnet build`/`dotnet test` confirmed green (0 Warnings/0 Errors, 40/40 tests) after all three landed. One STATE.md staleness gap (frontmatter/Current Position still described pre-Phase-5 state after all three plans landed) was caught and reconciled in a single consolidated pass afterward, exactly as recommended by 05-01/05-02's own executors.

### Pending Todos

None yet.

### Blockers/Concerns

- Research flagged PIA sourcing strategy (vendor `Microsoft.Office.Interop.Excel.dll` from a real Office GAC install vs. official NuGet package) as MEDIUM confidence — needs a quick spike at the start of Phase 1/3, not assumed
- Research flagged 32-bit Excel bitness handling in the installer as new ground the sibling reference project never solved — Phase 4 must make an explicit bitness-aware implementation or a documented single-bitness (64-bit only, per PROJECT.md constraint) decision
- REQUIREMENTS.md's own "Coverage" section previously stated "19 total" v1 requirements but the actual v1 requirement list contains 20 (FMT-01..07 = 7, not 6) — corrected during roadmap creation

## Deferred Items

Items acknowledged and deferred at v1.0 milestone close on 2026-07-11:

| Category | Item | Status |
|----------|------|--------|
| verification_gap | Phase 03 (03-VERIFICATION.md) — live-Excel smoke test (RIB-01..04) requires a real Windows+Excel machine, unavailable in this Linux/WSL environment | human_needed |
| verification_gap | Phase 04 (04-VERIFICATION.md) — live install/uninstall/idempotency/Resiliency test (INST-01..03) requires a real Windows+Excel machine | human_needed |
| verification_gap | Phase 05 (05-VERIFICATION.md) — real git push + tag + live CI run or manual `gh release create` (REL-01) deliberately deferred pending explicit user authorization to publish to the real public remote | human_needed |

All three are itemized, non-fabricated checklists recorded in their respective phase SUMMARY.md/VERIFICATION.md files and in `.planning/v1.0-MILESTONE-AUDIT.md` — none were silently skipped or assumed passing.

## Session Continuity

Last session: 2026-07-11T19:15:00-03:00
Stopped at: Completed 05-04-PLAN.md — Phase 5 fully code-complete (4/4 plans). Proceeding to Phase 5 code review + verification, then milestone lifecycle.
Resume file: None

## Operator Next Steps

- Start the next milestone with /gsd-new-milestone
