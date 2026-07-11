# Project Retrospective

*A living document updated after each milestone. Lessons feed forward into future planning.*

## Milestone: v1.0 — VBA to C# Migration

**Shipped:** 2026-07-11
**Phases:** 5 | **Plans:** 14 | **Sessions:** 1 (single long-running autonomous session, `/gsd-autonomous`)

### What Was Built
- A pure C# (.NET Framework 4.8, buildable with the .NET 8 SDK) port of the entire VBA format engine — 11 format keys, 16 accounting-format combinations — proven byte-for-byte identical to the original via 40 xUnit tests, runnable with `dotnet test` alone before any Excel/COM code existed.
- A COM-free `IExcelGateway`/`IRangeHandle`/`ILog` abstraction seam plus `FormatEngine`/`RibbonController` orchestration, unit-tested against fakes.
- A real COM entry point (`Connect.cs`+`AddInHost.cs`) implementing `IDTExtensibility2`/`IRibbonExtensibility` against real `Microsoft.Office.Interop.Excel` types, with all 17 Ribbon callbacks wired 1:1 to `customUI14.xml`.
- A fully HKCU-only, no-admin PowerShell installer/uninstaller (`scripts/install.ps1`/`uninstall.ps1`/`verify-environment.ps1`) with `DoNotDisableAddinList` Resiliency protection.
- A tag-triggered GitHub Actions release pipeline (`.github/workflows/release.yml`) plus a manual `gh` CLI runbook (`RELEASE.md`) and changelog (`RELEASE_NOTES.md`), with the legacy VBA source archived to `archive/vba-legacy` and `README.md` fully rewritten for the C# add-in.

### What Worked
- **Layer-by-layer horizontal roadmap** (format engine → abstractions → real COM → installer → CI/CD) meant Phases 1-2 were 100% `dotnet test`-verifiable with zero Windows/Excel dependency, isolating the "cannot verify here" problem to only Phases 3-5 instead of spreading it across the whole project.
- **Explicit `human_needed` recording pattern**, established in Phase 3 and reused verbatim through Phases 4 and 5: rather than faking a live-Excel/live-install/live-release result, every phase that hit this environment's ceiling recorded an itemized, non-fabricated checklist in its own SUMMARY.md/VERIFICATION.md. This kept the milestone honest (audit status `tech_debt`, never `gaps_found`) while still allowing forward progress on everything that *was* verifiable.
- **Code review as a real gate, not a formality**: every phase's review caught genuine, fixable issues (Phase 3's missing `Marshal.ReleaseComObject`/GC-flush ghost-process leak; Phase 4's `-Force` flag that could silently kill Excel and discard unsaved work after only 3 seconds; Phase 5's missing fail-fast in the release packaging step that could have silently shipped a broken public release). All were fixed in the same session, not deferred.
- **Verifying agent-suggested fixes empirically rather than trusting them**: the Phase 5 code reviewer suggested a commit SHA for pinning `softprops/action-gh-release@v2` that turned out to be fabricated/incorrect on `git ls-remote` verification — caught before it shipped, by treating "the reviewer said so" as a claim to verify, not a fact to copy-paste.
- **Sibling project (`outlook-classic-delay-send`) as a structural template**: reused near-verbatim for `Connect.cs`, `install.ps1`/`uninstall.ps1`, and `release.yml`/`RELEASE.md`, saving significant research/design time across three separate phases.

### What Was Inefficient
- **Recurring subagent API session-limit interruptions** (at least 3 times across the session) required a "verify what actually survived" recovery step each time rather than blind retries. Two of three interruptions left real partial progress (commits already made, or a fully-written research file); the third left nothing. Each required manually re-checking `git log`/`git status`/`ls` before deciding whether to resume or restart — a repeated tax on session time.
- **Concurrent parallel-agent execution on a shared non-worktree working directory** (`workflow.use_worktrees: false`) caused `STATE.md`/`ROADMAP.md` write races during Phase 5's Wave 1 (3 plans dispatched in parallel with `depends_on: []`). Individual plan executors handled this correctly by scoping their own commits to their own files and skipping shared-state updates mid-race, but the shared files still ended up transiently inconsistent (stale frontmatter, a regressed "Current Position" section) and needed a manual consolidated reconciliation pass afterward. `workflow.use_worktrees: true` would have avoided this entirely at the cost of merge overhead.
- **STATE.md progress-counter semantics drifted** over the session (`completed_phases` was inconsistently incremented for "code complete but human_needed" phases at different points) — never affected actual deliverables, but required a couple of manual reconciliation passes to keep the bookkeeping internally consistent.

### Patterns Established
- **Fixed COM identity carried forward verbatim across phases**: Phase 3's `Connect.cs` header comment (GUID/ProgId/AssemblyName/Version) became the single source of truth that Phase 4's `install.ps1`/`uninstall.ps1` and Phase 5's `release.yml` packaging step all had to match byte-for-byte — verified explicitly at each phase boundary rather than assumed.
- **Push/publish-to-real-remote as its own explicit authorization boundary**, distinct from the "cannot physically test here" class of `human_needed` item. Phase 5 established that even under a full-autonomy operating directive, actions that publish to a real, public, shared remote (first-ever push of ~95 commits, cutting a real GitHub Release) require the user's own explicit go-ahead — recorded as a deferred, non-executed checklist rather than either blocking the whole milestone or silently pushing.
- **Deep code review fix passes, not superficial**: across all 5 phases, code review findings (critical + warning, and often info-level too) were fixed directly in the same session rather than deferred to a follow-up cycle, with a documented rationale whenever a finding was deliberately *not* fixed (e.g., Phase 4's WR-02/IN-01 shared-module extraction, explicitly out of scope per the plan's own stated design).

### Key Lessons
1. When a phase's own definition-of-done requires infrastructure the execution environment doesn't have (live Excel, a real Windows registry, authorization to push to a real remote), the fix is not to skip verification — it's to verify everything that *is* locally checkable, then record the gap as an itemized, non-fabricated checklist. This keeps a milestone audit honest (`tech_debt`, not `gaps_found`) without blocking all forward progress.
2. Treat subagent-suggested fixes (especially security-relevant ones like commit-SHA pinning) as claims requiring independent verification, not facts to apply blindly — a reviewer's own suggested fix contained a fabricated SHA that direct `git ls-remote` verification caught before it shipped.
3. Parallel background-agent execution without git worktrees is workable if each agent scrupulously scopes its own commits to its own files, but shared bookkeeping files (`STATE.md`/`ROADMAP.md`) will still need a manual consolidated reconciliation pass afterward — budget for that step rather than assuming the parallel agents will leave shared state clean.

### Cost Observations
- Model mix: primarily Sonnet 5 for planning/execution/review agents, per this project's `model_profile: balanced` config.
- Sessions: 1 (single continuous autonomous session spanning all 5 phases plus milestone close, with several mid-session compactions).
- Notable: three separate subagent API session-limit interruptions were absorbed without losing net progress, by treating each recovery as "verify what's real, don't redo correct work" rather than a blind restart.

---

## Cross-Milestone Trends

### Process Evolution

| Milestone | Sessions | Phases | Key Change |
|-----------|----------|--------|------------|
| v1.0 | 1 | 5 | First milestone — established the `human_needed` explicit-deferral pattern and the real-remote-push authorization boundary as durable conventions for future milestones on this project. |

### Cumulative Quality

| Milestone | Tests | Coverage | Zero-Dep Additions |
|-----------|-------|----------|-------------------|
| v1.0 | 40 (xUnit) | Format engine + orchestration layers fully covered; COM/installer/CI layers covered by structural/static checks only (no live Excel/Windows in dev environment) | 2 NuGet packages (`Microsoft.Office.Interop.Excel`, `MicrosoftOfficeCore16` — unofficial repackages, content-verified genuine) + 3 GitHub Actions (`checkout`, `setup-dotnet`, `action-gh-release`, all SHA-pinned) |

### Top Lessons (Verified Across Milestones)

1. Explicit, itemized, non-fabricated deferral of what genuinely cannot be verified in the current environment is better than either blocking all progress or silently assuming success — this milestone's `tech_debt`-not-`gaps_found` audit outcome is the proof point.
