---
phase: 05-cicd-pipeline-release-runbook
plan: 04
subsystem: release-authorization
tags: [git, gh-cli, human-verify, release]

# Dependency graph
requires:
  - phase: 05-cicd-pipeline-release-runbook (plan 01, plan 02, plan 03)
    provides: ".github/workflows/release.yml, RELEASE.md, RELEASE_NOTES.md, VBA removed from main, README.md rewritten — all local, all unpushed"
provides:
  - "Confirmed-unregressed local Phase 5 deliverables (Task 1, ALL_PASS) and a recorded, non-fabricated human-authorization checkpoint for the actual real-remote release (Task 2)"
affects: [milestone-completion]

# Tech tracking
tech-stack:
  added: []
  patterns: []

key-files:
  created: []
  modified: []

key-decisions:
  - "Task 2's checkpoint was NOT self-authorized. Per this project's core operating rules (pushing code, publishing releases, and other actions visible to others/affecting shared state always require explicit user confirmation, regardless of how autonomous the rest of this milestone's workflow has been) this is categorically different from Phase 3/4's human_needed checkpoints, which were about physical inability to test in this sandbox, not about authorization. Real git/gh access exists here — the reason to stop is that pushing 79+ commits and cutting a real GitHub Release against the user's real public repo (tpougy/finance-fmt-tools) for the first time is a consequential, hard-to-reverse, externally-visible action that only the user may authorize. Resolved by recording this as an explicit open item (resume-signal: \"deferred — not releasing yet\") rather than fabricating a push/release that never happened."

patterns-established: []

requirements-completed: []

# Metrics
duration: unknown
completed: 2026-07-11
---

# Phase 5 Plan 4: Pre-flight Integration Check + Human-Authorized Real Release Summary

**Task 1's compound integration re-check passed (`ALL_PASS`) — every Phase 5 local deliverable (workflow YAML, RELEASE.md ordering, RELEASE_NOTES.md content, VBA removal + `customUI14.xml` survival + green build, README.md cleanliness) is confirmed unregressed. Task 2, the human-authorization checkpoint for the actual `git push`/`gh release create` against the real `tpougy/finance-fmt-tools` remote, is recorded here as an explicit, non-fabricated open item — no push, tag, or release was executed.**

## Performance

- **Duration:** N/A — Task 1 automated; Task 2 is a human-authorization gate, not executed work
- **Completed:** 2026-07-11 (Task 1 verified; Task 2 open)
- **Tasks:** 2 (1 automated integration check, 1 `checkpoint:human-verify gate="blocking"`)
- **Files modified:** 0

## Accomplishments
- Ran Task 1's exact compound verification command: confirmed `.github/workflows/release.yml` parses as valid YAML with `permissions.contents: write` and `body_path: RELEASE_NOTES.md`; confirmed `RELEASE.md` contains `gh release create` and orders `git push origin main` strictly before any `git push origin v<tag>` line; confirmed `RELEASE_NOTES.md` documents the checkbox-persistence removal; confirmed zero `src/*.bas` files remain, `src/customUI14.xml` is present, and `dotnet build src/FinanceFmtTools.sln -c Release` succeeds with 0 Warning(s)/0 Error(s); confirmed `README.md` has no `CustomXMLPart` or `Install-FinanceFmtTools.ps1` references. Result: `ALL_PASS`.

## Task Commits

1. **Task 1 (automated integration check):** no commit — read-only verification, modifies no files, per its own `<files>` declaration.
2. **Task 2 (human-authorization checkpoint):** no commit for the checkpoint itself; this SUMMARY.md is the record.

## Files Created/Modified
- `.planning/phases/05-cicd-pipeline-release-runbook/05-04-SUMMARY.md` (this file)

## Decisions Made
- See `key-decisions` above.

## Deviations from Plan

None. This plan's own design anticipates and explicitly permits a "deferred — not releasing yet" resume-signal outcome for Task 2 — recording that outcome is the expected, by-design result, not a deviation.

## Issues Encountered

None. Task 1 passed cleanly on the first run.

## User Setup Required

**Task 2's entire content is the open item.** Nothing in this task was executed against the real GitHub remote — per this phase's hard scoping constraint (no `type: auto` task may run `git push origin main`, `git push origin archive/vba-legacy`, a real tag push, or `gh release create` against the real remote), this is confined exclusively to explicit human action.

**Resume-signal recorded:** `"deferred — not releasing yet"` — the assistant executing this milestone autonomously reached this checkpoint but did not self-authorize crossing the local→real-remote trust boundary, since doing so is an irreversible-ish, publicly-visible action (this repo is public; pushing would publish 79+ commits of the VBA→C# migration to `tpougy/finance-fmt-tools` for the first time, and would sit alongside 2 pre-existing real legacy releases). This decision is presented to the user as part of the final milestone report, with the exact command sequence below ready to run whenever the user chooses.

### Itemized human-authorization checklist (deferred, human_needed)

1. Review `RELEASE_NOTES.md` — this is the permanent public release body text. Edit now if wording needs changing.
2. Decide the release tag — `v2.0.0` is 05-RESEARCH.md's recommendation (SemVer major bump for the full VBA→C# replacement), not a technical requirement.
3. Update the two hardcoded version fields to match the chosen tag: `<Version>` in `src/FinanceFmtTools.ComAddin/FinanceFmtTools.ComAddin.csproj` (currently `1.0.0.0`) and the `AddinVersion` const in `src/FinanceFmtTools.ComAddin/AddInHost.cs` (currently `"1.0.0"`). Commit locally on `main`.
4. `git push origin main` — publishes all 79+ local migration commits to the public repo for the first time. **This step requires explicit user confirmation before running.**
5. `git push origin archive/vba-legacy` — publishes the preserved VBA snapshot branch (safe, read-only historical reference).
6. `git tag -a vX.Y.Z -m "vX.Y.Z — Migração completa VBA -> C# COM add-in"` then `git push origin vX.Y.Z`.
7. Either watch the triggered `windows-latest` GitHub Actions workflow (`gh run list -R tpougy/finance-fmt-tools`) until it succeeds, or run `RELEASE.md`'s manual fallback: build/package locally, then `gh release create vX.Y.Z FinanceFmtTools.zip --title "Finance Fmt Tools vX.Y.Z" -F RELEASE_NOTES.md`.
8. Confirm: `gh release view vX.Y.Z -R tpougy/finance-fmt-tools` shows the release with exactly one asset named `FinanceFmtTools.zip`.
9. Optional: run `scripts/install.ps1`'s one-liner on a real Windows+Excel machine to confirm the Ribbon tab appears — re-exercises Phase 3/4's own already-recorded `human_needed` checklists, not new scope.

REL-01/REL-02 remain "code complete" in REQUIREMENTS.md, not "live-verified" — that upgrade only happens once the user confirms a real, successful release via one of the resume-signal outcomes above.

## Next Phase Readiness
- Phase 5 is code-complete (4/4 plans) with all local artifacts verified consistent. The milestone's remaining work is the audit/complete/cleanup lifecycle, which can proceed independently of whether/when the user chooses to authorize the real release.
- No blockers to closing out Phase 5's own code review and VERIFICATION.md.

---
*Phase: 05-cicd-pipeline-release-runbook*
*Completed: 2026-07-11*
