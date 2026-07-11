---
phase: 05-cicd-pipeline-release-runbook
plan: 02
subsystem: release-docs
tags: [release, gh-cli, changelog, runbook, documentation]

# Dependency graph
requires:
  - phase: 05-cicd-pipeline-release-runbook (plan 01, in progress concurrently)
    provides: ".github/workflows/release.yml (tag-triggered CI pipeline) — RELEASE.md's automatic-path subsection references it by name; not a hard build/runtime dependency of this plan's own files"
provides:
  - "RELEASE.md — gh CLI manual release runbook (REL-02), zero dependency on GitHub Actions"
  - "RELEASE_NOTES.md — hand-maintained v2.0.0 changelog (REL-03), read by both CI's body_path and this runbook's -F flag"
affects: [phase-5-plan-03, phase-5-plan-04-live-release-checkpoint]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Hand-maintained RELEASE_NOTES.md (not --generate-notes) as the single changelog source shared by both the automated CI workflow (body_path) and the manual gh runbook (-F flag) — avoids noisy auto-generated notes from this repo's GSD-bookkeeping-heavy commit history"
    - "RELEASE.md documents git push origin main strictly before git push origin vX.Y.Z, with an explicit inline warning explaining why order matters"

key-files:
  created:
    - RELEASE.md
    - RELEASE_NOTES.md
  modified: []

key-decisions:
  - "Used v2.0.0 as the changelog heading per 05-RESEARCH.md Assumption A1 (SemVer major bump signaling the full VBA -> C# replacement) — adjustable later at the 05-04-PLAN.md human checkpoint if disagreed with."
  - "Committed each task with git commit -- <pathspec> (not git add -A / plain git commit) after discovering pre-existing staged deletions (Install-FinanceFmtTools.bat/.ps1, src/*.bas) and an untracked .github/ directory in the working tree at task start, belonging to a concurrently-running Plan 05-01 executor sharing the same working tree (workflow.use_worktrees: false). This kept both of this plan's commits scoped exactly to RELEASE_NOTES.md and RELEASE.md respectively, leaving Plan 05-01's in-flight work untouched throughout."
  - "Neither task executed any git push, git tag, or gh release create against the real remote — both files document these commands as text/examples only, per the plan's explicit constraint and this milestone's deferral of live release execution to Plan 05-04's human-authorization checkpoint."

patterns-established:
  - "Any future release-notes entry follows RELEASE_NOTES.md's established section order: opening summary, 'O que mudou', a clearly-labeled behavior-change subsection when applicable, upgrade instructions, installation one-liner, closing 'Notas técnicas'."

requirements-completed: [REL-02, REL-03]

# Metrics
duration: ~3 min
completed: 2026-07-11
---

# Phase 5 Plan 2: RELEASE.md + RELEASE_NOTES.md Summary

**Authored the CI-independent `gh` CLI manual release runbook (`RELEASE.md`, REL-02) and the hand-maintained v2.0.0 changelog (`RELEASE_NOTES.md`, REL-03) that both the CI workflow and a human/AI-agent-run manual release depend on — pure local documentation authoring, nothing pushed to `origin`.**

## Performance

- **Duration:** ~3 min
- **Started:** 2026-07-11T18:29:59Z
- **Completed:** 2026-07-11T18:32:31Z
- **Tasks:** 2 (both `type="auto"`)
- **Files modified:** 2 (both created)

## Accomplishments
- `RELEASE_NOTES.md` (66 lines): opening migration summary, "O que mudou" (format engine port with automated test coverage, new HKCU installer/uninstaller replacing the old Excel-COM-automation installer, new automated tag-triggered release pipeline), a clearly-labeled "Mudança de comportamento" subsection explicitly stating the "Alinhar à direita"/"Zero contábil" checkboxes no longer persist between Excel sessions (framed as a deliberate scope decision, not a regression), an "Atualizando a partir da versão VBA" upgrade note, the exact install one-liner, and a closing "Notas técnicas" section referencing `archive/vba-legacy` and `scripts/uninstall.ps1`.
- `RELEASE.md` (166 lines): "Ferramentas necessárias" table, a 5-step numbered "Fluxo de release" (test → update changelog → update the two hardcoded version fields → build/package locally → tag+push, with `git push origin main` explicitly before the tag push and an inline explanation of the Pitfall-3 ordering risk), an automatic-path subsection (CI workflow trigger), a manual-fallback subsection (`gh release create ... -F RELEASE_NOTES.md`, explicitly noting zero CI dependency), a verification section (`gh release view`), a "Comandos úteis" section, and a final release checklist. A prominent callout near the top states `tpougy/finance-fmt-tools` is public and that `git push origin main` will publish all local migration history there for the first time.
- Both tasks' automated verify commands pass exactly as specified in the plan (confirmed via direct re-run of each `<verify><automated>` command, `ALL_PASS` both times).

## Task Commits

Each task was committed atomically:

1. **Task 1: Write RELEASE_NOTES.md — first C#-migration changelog entry (REL-03)** - `fedfe93` (docs)
2. **Task 2: Write RELEASE.md — gh CLI manual release runbook (REL-02)** - `5272740` (docs)

_Note: commits used `git commit -- <pathspec>` (not `git add -A`) to stay scoped to only this plan's files — see Deviations below for why this mattered in this execution._

## Files Created/Modified
- `RELEASE_NOTES.md` - new file, 66 lines — hand-maintained v2.0.0 changelog (REL-03)
- `RELEASE.md` - new file, 166 lines — gh CLI manual release runbook (REL-02)

## Decisions Made
See `key-decisions` in frontmatter above. Most notable: this plan's execution overlapped in time with a separate, concurrently-running executor for Plan 05-01 in the same (non-worktree) working directory — handled by staging and committing only this plan's own files via `git commit -- <pathspec>`, never `git add -A`/plain `git commit`.

## Deviations from Plan

None affecting the plan's own deliverables — both files match every `<action>`/`<acceptance_criteria>` item in `05-02-PLAN.md` verbatim, and both automated verify commands pass.

### Auto-fixed Issues
None — no bugs, missing functionality, or blocking issues encountered in this plan's own scope.

### Notable environmental observation (not a deviation, not fixed)
At task start, `git status --short` showed pre-existing staged file deletions (`Install-FinanceFmtTools.bat`, `Install-FinanceFmtTools.ps1`, `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`) and an untracked `.github/workflows/release.yml`, belonging to Plan 05-01 (LEGACY-01 VBA removal + REL-01 CI workflow), executed by a separate, concurrently-running agent in the same shared working tree (`workflow.use_worktrees: false` for this project, confirmed via `.planning/config.json`). Mid-execution, that other process committed `defc25f feat(05-01): add tag-triggered release workflow (REL-01)` between this plan's two commits — visible in `git log`/`git reflog`. This is out of scope for Plan 05-02 (files_modified: `RELEASE.md`, `RELEASE_NOTES.md` only) and was left entirely untouched: neither staged nor committed by this execution. Confirmed via `git show --stat` on both of this plan's own commits (`fedfe93`, `5272740`) that each contains exactly one file, matching the plan's `files_modified` list.

## Issues Encountered
None blocking. The concurrent-execution observation above required using `git commit -- <pathspec>` scoping instead of the more naive `git add -A && git commit`, to avoid accidentally bundling Plan 05-01's in-flight, uncommitted work into this plan's commits.

## User Setup Required
None — this plan is pure local documentation authoring. Actual release execution (running the commands `RELEASE.md` documents against the real `origin` remote — `git push origin main`, tagging, `gh release create`) is explicitly out of this plan's scope and deferred to Plan 05-04's human-authorization checkpoint, per the orchestrator's instructions and this plan's own `<verification>` block.

## Next Phase Readiness
- `RELEASE_NOTES.md` is ready to be read by Plan 05-01's CI workflow (`body_path: RELEASE_NOTES.md`) once that plan's `.github/workflows/release.yml` is committed.
- `RELEASE.md`'s manual fallback command (`gh release create vX.Y.Z FinanceFmtTools.zip --title "Finance Fmt Tools vX.Y.Z" -F RELEASE_NOTES.md`) is ready for Plan 05-04's live checkpoint to execute for real, once the human authorizes it.
- No blockers for Plan 05-03 or 05-04.

## Self-Check: PASSED
- `FOUND: RELEASE_NOTES.md` — file exists on disk.
- `FOUND: RELEASE.md` — file exists on disk.
- `FOUND: fedfe93` — commit exists in `git log --oneline --all`.
- `FOUND: 5272740` — commit exists in `git log --oneline --all`.

---
*Phase: 05-cicd-pipeline-release-runbook*
*Completed: 2026-07-11*
