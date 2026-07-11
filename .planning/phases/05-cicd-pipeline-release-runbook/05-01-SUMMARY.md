---
phase: 05-cicd-pipeline-release-runbook
plan: 01
subsystem: ci-cd
tags: [github-actions, ci, release-automation, windows-latest, gitignore]

# Dependency graph
requires:
  - phase: 04-installation-registration (plan 01)
    provides: "scripts/install.ps1's fixed $AssetName = 'FinanceFmtTools.zip' constant and its releases/latest/download/FinanceFmtTools.zip URL — this plan's packaging step must reproduce that exact filename"
provides:
  - ".github/workflows/release.yml — tag-triggered (v*.*.*) build/test/package/publish pipeline (REL-01), publishing the fixed-name FinanceFmtTools.zip asset via softprops/action-gh-release@v2"
  - ".gitignore entries (FinanceFmtTools.zip, staging/) — prevents local manual-runbook test artifacts from being accidentally committed"
affects: [phase-5-plan-02-release-runbook, phase-5-plan-04-live-release-checkpoint]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Single tag-triggered workflow (on.push.tags: v*.*.*) with an explicit top-level permissions: contents: write block — this repo's default Actions token permission is read-only, confirmed via gh api repos/tpougy/finance-fmt-tools/actions/permissions/workflow"
    - "Fixed, non-interpolated release asset filename (FinanceFmtTools.zip) — never $tag/${{ github.ref_name }}-suffixed — so scripts/install.ps1's stable latest/download/ URL never breaks between releases"

key-files:
  created:
    - .github/workflows/release.yml
  modified:
    - .gitignore

key-decisions:
  - "Reused outlook-classic-delay-send's proven-working release.yml pattern verbatim, adapted only for the fixed (non-versioned) zip filename this repo's install.ps1 already hard-depends on — no other structural deviation from the sibling's pattern."
  - "Committed each task individually via 'gsd-sdk query commit ... --files <path>' (not git add -A), after observing a concurrently-running Plan 05-02 executor sharing the same working tree (workflow.use_worktrees: false) — kept both commits scoped exactly to this plan's own two files."

patterns-established:
  - "Any future additional CI workflow in this repo should keep the same single-job, tag-only-trigger shape unless a new requirement explicitly calls for a matrix or push/PR-triggered CI workflow."

requirements-completed: [REL-01]

# Metrics
duration: ~10 min (approximate — start time not formally recorded before context-loading began; the two task commits themselves are 47s apart)
completed: 2026-07-11
---

# Phase 5 Plan 1: Tag-Triggered Release Workflow (release.yml + .gitignore) Summary

**Authored `.github/workflows/release.yml` (REL-01) — a single tag-triggered (`v*.*.*`) GitHub Actions job on `windows-latest` that restores/builds/tests/packages the C# COM add-in and publishes it as a fixed-name `FinanceFmtTools.zip` GitHub Release asset via `softprops/action-gh-release@v2` — plus two new `.gitignore` entries for local packaging artifacts. Nothing pushed to `origin`; both files committed locally only.**

## Performance

- **Duration:** ~10 min (approximate)
- **Completed:** 2026-07-11
- **Tasks:** 2 (both `type="auto"`)
- **Files modified:** 2 (1 created, 1 modified)

## Accomplishments
- `.github/workflows/release.yml` (56 lines): `name: Release`, trigger `on.push.tags: ['v*.*.*']` only (no branch/PR triggers, no matrix), explicit top-level `permissions: contents: write` (mandatory since this repo's default Actions token permission is `read`), single job `build-and-release` on `runs-on: windows-latest` with steps in order: `actions/checkout@v4` → `actions/setup-dotnet@v4` (`dotnet-version: '8.x'`) → `Restore` (`dotnet restore src\FinanceFmtTools.sln`) → `Build` (`dotnet build src\FinanceFmtTools.sln -c Release --no-restore`) → `Test` (`dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release --no-build`) → `Package` (`shell: pwsh`, stages the 4 binaries from `src\FinanceFmtTools.ComAddin\bin\Release\net48` plus the 3 `scripts\*.ps1` files into `staging\`, then `Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force` — literal filename, zero `${{`/`$tag` interpolation) → `Create GitHub Release` (`softprops/action-gh-release@v2`, `tag_name: ${{ github.ref_name }}`, `name: Finance Fmt Tools ${{ github.ref_name }}`, `body_path: RELEASE_NOTES.md`, `files: FinanceFmtTools.zip`).
- `.gitignore` gained a new `# Release packaging artifacts (local manual-runbook test runs)` section with two entries: `FinanceFmtTools.zip` and `staging/` — all 5 pre-existing lines (`bin/`, `obj/`, `.vscode/`, `.DS_Store`, `Thumbs.db`) left untouched.
- Both tasks' automated verify commands pass exactly as specified in the plan (Python `yaml.safe_load` structural assertions, `FinanceFmtTools.zip` literal count ≥2, non-interpolated `Compress-Archive` line, all 3 action version tags present, `body_path: RELEASE_NOTES.md` present; `.gitignore` grep checks for all 7 expected lines) — confirmed `ALL_PASS`/`YAML OK` on direct re-run.
- Confirmed via `git rev-list --left-right --count origin/main...main` (`0  84`) that nothing was pushed to `origin` — `origin/main` remains at the pre-migration commit `8804778`.

## Task Commits

1. **Task 1: Create the tag-triggered release workflow (.github/workflows/release.yml)** - `defc25f` (feat)
2. **Task 2: Gitignore local release packaging artifacts (.gitignore)** - `bdf38ec` (chore)

_Note: both commits used `gsd-sdk query commit "..." --files <path>` scoped to exactly one file each — confirmed via `git show --stat` on both hashes that each touches only its own single file._

## Files Created/Modified
- `.github/workflows/release.yml` - new file, 56 lines — tag-triggered build/test/package/publish pipeline (REL-01)
- `.gitignore` - modified, +4 lines — ignores `FinanceFmtTools.zip` and `staging/`

## Decisions Made
See `key-decisions` in frontmatter above. Most notable: this plan's execution overlapped in time with a separate, concurrently-running executor for Plan 05-02 in the same (non-worktree) working directory — handled by committing only this plan's own files via the SDK's `--files`-scoped commit, never a blanket `git add -A`/plain `git commit`.

## Deviations from Plan

None affecting the plan's own deliverables — `.github/workflows/release.yml` matches every structural requirement in `05-01-PLAN.md`'s Task 1 `<action>`/`<acceptance_criteria>` verbatim (exact step order, exact commands, literal fixed zip filename, exact action version tags), and `.gitignore` matches Task 2's `<action>`/`<acceptance_criteria>` verbatim.

### Auto-fixed Issues
None — no bugs, missing functionality, or blocking issues encountered in this plan's own scope.

### Notable environmental observation (not a deviation, not fixed)
Partway through execution, `git status --short` showed staged file deletions (`Install-FinanceFmtTools.bat`, `Install-FinanceFmtTools.ps1`, `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`) that do not belong to this plan (`files_modified: [.github/workflows/release.yml, .gitignore]` only) — these are LEGACY-01 VBA-archival work from a separate, concurrently-running agent sharing the same working tree (`workflow.use_worktrees: false`, confirmed via `.planning/config.json`). Left entirely untouched: neither staged further, unstaged, nor committed by this execution. Confirmed via `git show --stat` on both of this plan's own commits (`defc25f`, `bdf38ec`) that each contains exactly one file, matching the plan's `files_modified` list. Separately, a concurrent Plan 05-02 executor's two commits (`fedfe93` RELEASE_NOTES.md, `5272740` RELEASE.md) landed interleaved with this plan's own commits in `git log` — also out of scope and left untouched.

## Issues Encountered
None blocking. The concurrent-execution observation above required using the SDK's `--files`-scoped commit instead of a naive `git add -A && git commit`, to avoid accidentally bundling other in-flight plans' uncommitted/staged work into this plan's commits.

## User Setup Required
None — this plan is pure local file authoring (workflow YAML + gitignore entries). Actual execution of the workflow (pushing a real `v*.*.*` tag against the real `origin` remote and observing a live `windows-latest` Actions run) is explicitly out of this plan's scope and deferred to Plan 05-04's human-authorization checkpoint, per the orchestrator's instructions and this plan's own `<verification>` block.

## Next Phase Readiness
- `.github/workflows/release.yml`'s `body_path: RELEASE_NOTES.md` reference is already satisfied — Plan 05-02 (running concurrently) authored `RELEASE_NOTES.md` in this same session.
- `.gitignore`'s new `staging/`/`FinanceFmtTools.zip` entries are ready for Plan 05-02's `RELEASE.md` manual runbook and Plan 05-04's live release checkpoint to use without risk of accidental commits.
- No blockers for Plan 05-03 or 05-04.

## Self-Check: PASSED
- `FOUND: .github/workflows/release.yml` — file exists on disk.
- `FOUND: .gitignore` (contains `FinanceFmtTools.zip` and `staging/`) — confirmed via grep.
- `FOUND: defc25f` — commit exists in `git log --oneline --all`.
- `FOUND: bdf38ec` — commit exists in `git log --oneline --all`.

---
*Phase: 05-cicd-pipeline-release-runbook*
*Completed: 2026-07-11*
