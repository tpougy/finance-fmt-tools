---
phase: 05-cicd-pipeline-release-runbook
plan: 03
subsystem: legacy-cleanup
tags: [git, documentation, vba-archival, readme]

# Dependency graph
requires: []
provides:
  - "main's working tree no longer contains src/*.bas files or the two legacy root installer scripts (LEGACY-01), while remaining fully recoverable from archive/vba-legacy (tip cf2559b)"
  - "README.md rewritten to document only the C# add-in's install/uninstall/architecture flow (LEGACY-02)"
affects: [phase-5-verification, milestone-completion, 05-04-real-release]

# Tech tracking
tech-stack:
  added: []
  patterns: []

key-files:
  created: []
  modified:
    - README.md
  deleted:
    - src/ThisWorkbook.bas
    - src/modConfig.bas
    - src/modFormatEngine.bas
    - src/modRibbon.bas
    - src/modUtils.bas
    - Install-FinanceFmtTools.ps1
    - Install-FinanceFmtTools.bat

key-decisions:
  - "Verified (read-only) that archive/vba-legacy already exists locally, tip commit cf2559b, containing exactly the 6 expected src/ paths, before touching main — no branch creation/force-push performed, per plan constraint."
  - "git rm targeted exactly the 7 VBA-only files; src/customUI14.xml was deliberately excluded since it is still an active EmbeddedResource in FinanceFmtTools.Engine.csproj, read at runtime by RibbonController.GetCustomUiXml() — confirmed via grep before removal and via a green dotnet build after."
  - "README.md's 'Persistência de configurações' section was replaced with a session-only note that avoids the literal string 'CustomXMLPart' entirely (per Task 2's acceptance criteria), rephrasing the historical VBA persistence mechanism descriptively instead."
  - "No git push performed for either main or archive/vba-legacy — deferred to 05-04's human-authorized checkpoint, per explicit plan/task instruction."

patterns-established: []

requirements-completed: [LEGACY-01, LEGACY-02]

# Metrics
duration: ~25 min
completed: 2026-07-11
---

# Phase 5 Plan 3: Archive VBA Source Off Main + Rewrite README.md Summary

**Removed the 5 retired `.bas` modules and the 2 legacy root PowerShell/.bat installer files from `main`'s working tree in one atomic commit (LEGACY-01), confirmed `archive/vba-legacy` (tip `cf2559b`) still holds a complete pre-migration snapshot including `src/customUI14.xml`, confirmed `dotnet build` stays green (0 Warning(s)/0 Error(s)) proving `customUI14.xml` survived untouched, then rewrote `README.md` end-to-end for the C# add-in only (LEGACY-02) — new install/uninstall one-liners, an "upgrading from VBA" callout, a session-only preferences note replacing the old `CustomXMLPart` persistence section, and a rewritten architecture/development section — while preserving every user-facing format table and the ribbon-tree reference verbatim.**

## Performance

- **Duration:** ~25 min
- **Completed:** 2026-07-11
- **Tasks:** 2/2
- **Files modified:** 1 (`README.md`) + 7 deleted

## Accomplishments

- Confirmed `archive/vba-legacy`'s tip commit is `cf2559b` and that it contains exactly the 6 expected `src/` paths (5 `.bas` files + `customUI14.xml`) — the pre-migration snapshot is intact and fully recoverable.
- Removed `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`, `Install-FinanceFmtTools.ps1`, and `Install-FinanceFmtTools.bat` from `main` via `git rm`, committed in one commit (`555d06d`).
- Left `src/customUI14.xml` completely untouched — verified it is still present and that `dotnet build src/FinanceFmtTools.sln -c Release` completes with `0 Warning(s)`/`0 Error(s)` after the removal (proves `FinanceFmtTools.Engine.csproj`'s `EmbeddedResource Include="../customUI14.xml"` reference is intact).
- Rewrote `README.md` (commit `990969e`): removed the hardcoded `Versão: 1.0.0` line and updated the platform line to mention the C# (.NET Framework 4.8) COM add-in; rewrote "Instalação" with the `scripts/install.ps1`/`scripts/uninstall.ps1` one-liners and `scripts/verify-environment.ps1 -RuntimeOnly` as an optional diagnostic, plus an "Atualizando da versão VBA" callout; replaced the `CustomXMLPart`-based "Persistência de configurações" section with a "Preferências de sessão" note (session-only, no persistence, without using the literal string "CustomXMLPart"); rewrote "Arquitetura do projeto" to describe the real C# solution layout (`FinanceFmtTools.Engine`, `FinanceFmtTools.ComAddin`, `FinanceFmtTools.Engine.Tests`, `src/customUI14.xml`) with a reference to `archive/vba-legacy` for historical VBA source; added a "Desenvolvimento" section documenting `dotnet build`/`dotnet test` commands. All user-facing format tables ("Família Fin xD", "Outros formatos", "Referência rápida do ribbon") were preserved verbatim.
- Re-ran every acceptance-criteria grep check for both tasks after the edits — all pass (see Self-Check below).

## Task Commits

| Task | Name | Commit | Files |
|------|------|--------|-------|
| 1 | Archive VBA source off main (LEGACY-01) | `555d06d` | 7 files deleted (5 `.bas`, 2 legacy installer scripts) |
| 2 | Rewrite README.md for the C# add-in only (LEGACY-02) | `990969e` | `README.md` |

## Files Created/Modified

**Deleted:**
- `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`
- `Install-FinanceFmtTools.ps1`, `Install-FinanceFmtTools.bat`

**Modified:**
- `README.md` — full rewrite of install/persistence/architecture sections; format tables and ribbon-tree reference untouched.

**Untouched (verified, not part of this plan):**
- `src/customUI14.xml` — still present, still embedded by `FinanceFmtTools.Engine.csproj`.

## Decisions Made

See `key-decisions` above.

## Deviations from Plan

None — plan executed exactly as written. Both tasks' automated verify commands and acceptance criteria pass as specified.

## Issues Encountered

**Concurrent wave-1 execution on a shared working tree (not a defect in this plan):** Plans `05-01` and `05-02` are also `wave: 1`/`depends_on: []` and, per this project's `workflow.use_worktrees: false` config, execute against the same working directory as this plan, concurrently, rather than in isolated git worktrees. This was observed directly: `.github/workflows/release.yml` and `RELEASE_NOTES.md` appeared/disappeared as untracked files between consecutive `git status` calls during this plan's execution, and `git log` picked up an interleaved `05-01` commit (`bdf38ec chore(05-01): gitignore local release packaging artifacts`) between this plan's own two commits. This plan's own git operations were scoped exclusively to the 8 files explicitly listed in its `<files>` frontmatter (never `git add -A`/`git add .`), so no cross-contamination occurred in either of this plan's two commits — confirmed via `git diff --stat` on both commits showing only the intended files. `.planning/STATE.md` was also observed mid-edit (uncommitted) by a concurrent agent's `gsd-sdk query state.*` calls at the time this plan reached its own state-update step; see "State Update Note" below.

**`gsd-sdk query commit` wrapper fails on already-`git rm`'d deletions:** Task 1's commit was made via a direct `git commit` (not the `gsd-sdk query commit` wrapper) because the wrapper's underlying `git add -- <path>` step fails with `fatal: pathspec '<path>' did not match any files` for a file that was already removed from the working tree via `git rm` (a standard, reproducible git behavior for deletions — not a race condition; confirmed by re-running `git add -- src/ThisWorkbook.bas` directly and getting the identical error in isolation). Task 2's commit used the same direct `git commit` path for consistency. Both commits followed the same "stage only the task's own files, commit with a descriptive message" protocol the wrapper itself implements; no `git add -A`/`-f` was used at any point.

## User Setup Required

None. This plan performed only local git operations (no push) and a local documentation rewrite — no external service configuration needed.

## State Update Note

Per this plan's execution instructions, STATE.md/ROADMAP.md/REQUIREMENTS.md updates were attempted via the standard `gsd-sdk query state.*`/`roadmap.*`/`requirements.*` handlers. Because `05-01`/`05-02` execute concurrently against the same shared working tree (see "Issues Encountered" above) and were independently mutating `.planning/STATE.md` at the same time, this plan's own state-tracking calls may interleave with theirs. `requirements.mark-complete` was called with `LEGACY-01 LEGACY-02` (this plan's frontmatter `requirements` field) so both rows get checked off in `REQUIREMENTS.md`'s traceability table regardless of which agent's write lands last. If the final `.planning/STATE.md`/`ROADMAP.md` state looks inconsistent with this plan's own contributions after all three wave-1 plans finish, reconcile by cross-referencing this SUMMARY's Task Commits table (`555d06d`, `990969e`) and the two other wave-1 plans' own SUMMARY files.

## Next Phase Readiness

- LEGACY-01 and LEGACY-02 are both complete — `main`'s active working tree has zero `.bas` files and no root-level legacy installer scripts, `src/customUI14.xml` is unchanged, `dotnet build` remains green, and `README.md` documents only the C# add-in.
- Nothing has been pushed to `origin` for either `main` or `archive/vba-legacy` — both are deferred to `05-04-PLAN.md`'s human-authorized checkpoint, exactly as this plan's constraints required.
- Phase 5's remaining plan (`05-04`) can proceed once `05-01` (CI workflow) and `05-02` (manual runbook + changelog) also complete — this plan has no blocking dependency on either of them and vice versa.

## Self-Check: PASSED

- FOUND: `.planning/phases/05-cicd-pipeline-release-runbook/05-03-SUMMARY.md`
- FOUND: `README.md`
- FOUND: `src/customUI14.xml`
- CONFIRMED ABSENT: `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`, `Install-FinanceFmtTools.ps1`, `Install-FinanceFmtTools.bat`
- FOUND commit: `555d06d`
- FOUND commit: `990969e`

---
*Phase: 05-cicd-pipeline-release-runbook*
*Completed: 2026-07-11*
