# Phase 5: CI/CD Pipeline & Release Runbook - Context

**Gathered:** 2026-07-11
**Status:** Ready for planning
**Mode:** Auto-generated (discuss skipped via workflow.skip_discuss)

<domain>
## Phase Boundary

Releases are fully automated from a `v*.*.*` tag push (build/test/package/publish via GitHub Actions on `windows-latest`), a documented manual `gh` CLI fallback exists for a person or an AI agent to cut a release without CI, every release has changelog notes, and the VBA legacy source/docs are fully out of `main`'s active flow (archived to `archive/vba-legacy`).

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions to guide decisions.

### Carried-over fixed identity and file layout from Phases 3-4 (not discretionary — must be reused verbatim)
The CI workflow's build/package step MUST produce a release asset that matches what `scripts/install.ps1` already assumes: a zip literally named `FinanceFmtTools.zip` (the fixed convention `install.ps1` documents in its own comments), containing at minimum the 4 files `install.ps1`/`uninstall.ps1` require: `FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, `office.dll` (from `src/FinanceFmtTools.ComAddin/bin/Release/net48/`), plus `scripts/install.ps1`, `scripts/uninstall.ps1`, and `scripts/verify-environment.ps1` so a downloaded release is self-contained. Do not invent a different asset name or layout — read `scripts/install.ps1`'s `$AssetName`/`$AllFiles` constants and its own comment ("todo release do Phase 5 (CI) deve publicar seu zip sob este nome literal fixo") before writing the workflow.

### Environment constraint (same class as Phases 3-4, but narrower)
This dev environment is Linux/WSL with no Windows, no Excel. Unlike Phases 3/4, this phase's core deliverable (a GitHub Actions YAML workflow, a release runbook markdown file, and a git branch operation for VBA archival) IS fully authorable and largely verifiable here: workflow YAML syntax/structure can be validated, the `gh` CLI is likely available in this environment for both authoring the runbook and potentially exercising a real manual release, and the `archive/vba-legacy` branch operation is a plain git action executable in this sandbox. The one thing that cannot be verified here is an actual `windows-latest` GitHub Actions run (no way to trigger/observe a real CI run without pushing to a real remote and consuming Actions minutes) — that piece may need to be deferred as `human_needed` or done via a real (user-authorized) tag push, not silently assumed to work.

</decisions>

<code_context>
## Existing Code Insights

`Install-FinanceFmtTools.ps1` (legacy VBA installer, repo root) downloads `FinanceFmtTools.xlam` from GitHub Releases — this is the artifact being fully retired by this phase (LEGACY-01/02). The sibling project `outlook-classic-delay-send` (CLAUDE.md's stated workflow inspiration) is the closest analog for a GitHub Actions release pipeline building a .NET Framework COM add-in on `windows-latest` — check for a `.github/workflows/*.yml` and a `RELEASE.md`/runbook there to reuse patterns from, the same way Phases 3/4 reused its `Connect.cs`/`install.ps1`/`uninstall.ps1` structure. `README.md` currently documents the VBA `.xlam` install flow exclusively and needs a full rewrite (LEGACY-02) pointing only at the new C# add-in's `scripts/install.ps1` one-liner.

</code_context>

<specifics>
## Specific Ideas

No specific requirements — discuss phase skipped. Refer to ROADMAP phase description and success criteria (REL-01, REL-02, REL-03, LEGACY-01, LEGACY-02 in REQUIREMENTS.md).

</specifics>

<deferred>
## Deferred Ideas

None — discuss phase skipped.

</deferred>
