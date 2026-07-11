---
phase: 05-cicd-pipeline-release-runbook
verified: 2026-07-11T19:03:29Z
status: human_needed
score: 5/5 must-haves code-verified; 0/5 live-remote-verified (authorization-boundary constraint, by design)
overrides_applied: 0
human_verification:
  - test: "Push `main` and `archive/vba-legacy` to the real `origin` remote (`tpougy/finance-fmt-tools`, public), then push a `vX.Y.Z` tag and either watch `.github/workflows/release.yml` run to success on `windows-latest`, or run `RELEASE.md`'s manual `gh release create` fallback."
    expected: "The real GitHub Actions run (or the manual `gh release create` command) completes successfully, publishing a GitHub Release with exactly one asset named `FinanceFmtTools.zip`, and `gh release view vX.Y.Z` confirms it."
    why_human: "This project's explicit, documented hard safety constraint prohibits any autonomous execution agent from pushing to the real `origin` remote or cutting a real GitHub Release — this repo is real, public, and has never had this migration's ~95 commits pushed. Only a human (or an agent the human separately and explicitly authorizes outside this task) may cross that boundary. This is an authorization-policy constraint, not a technical/environment limitation — `git`/`gh` are both authenticated and functional in this sandbox."
  - test: "Once released, run `scripts/install.ps1`'s documented one-liner on a real Windows+Excel machine and confirm the 'Finance Fmt' ribbon tab appears, matching Phase 3/4's own still-open live-Excel/live-install checklists."
    expected: "Add-in installs and the Ribbon tab renders exactly as documented."
    why_human: "Requires a live Windows+Excel environment; re-exercises Phase 3/4's own already-recorded `human_needed` items, not new scope introduced by this phase."
---

# Phase 5: CI/CD Pipeline & Release Runbook Verification Report

**Phase Goal:** Releases are fully automated from a `v*.*.*` tag push, a documented manual fallback exists for a person or AI agent to cut a release without CI, and the VBA legacy code/docs are completely out of the active `main` flow.
**Verified:** 2026-07-11
**Status:** human_needed
**Re-verification:** No — initial verification

## Environment / Authorization Constraint (documented, non-discretionary)

This phase's hard safety constraint (05-CONTEXT.md, every Phase 5 plan's threat model, `.planning/STATE.md`'s "[Phase 05, safety scoping decision]" entry) is that **no autonomous execution agent may push to the real `origin` remote or cut a real GitHub Release** — `tpougy/finance-fmt-tools` is a real, public repository that has never had this milestone's C#-migration commits pushed. Unlike Phase 3/4's `human_needed` verdicts (which stemmed from a genuine technical inability — no Windows/Excel/registry in this Linux/WSL sandbox), this phase's live-remote step is fully *executable* here (`git`/`gh` are authenticated and working) but is deliberately *not authorized* for an agent to run autonomously. Every Phase 5 deliverable was therefore built and independently re-verified in this session as a **local artifact**: workflow YAML structural correctness (re-parsed with `yaml.safe_load`), the three pinned GitHub Actions SHAs (cross-checked against the real upstream repos via `git ls-remote`, not just trusted from source comments), `dotnet build`/`dotnet test` (re-run fresh, not taken from SUMMARY.md), VBA file removal + `customUI14.xml` survival (confirmed via `ls`/`git ls-tree`), and README/RELEASE.md/RELEASE_NOTES.md content (read directly, cross-checked against `FormatRegistry.cs`/`customUI14.xml`/`scripts/install.ps1`). The actual live trigger of a real `windows-latest` Actions run, and the actual `git push`/`gh release create` against the real remote, are correctly deferred as an explicit, non-fabricated `human_needed` item — this mirrors Phase 3/4's own precedent and is the correct, by-design outcome, not a phase failure.

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | REL-01: Pushing a `v*.*.*` tag triggers a GitHub Actions workflow on `windows-latest` that builds, tests, packages, and publishes a release with the fixed-name `FinanceFmtTools.zip` asset, no manual steps | ✓ VERIFIED (code/structure) — real trigger `human_needed` | `.github/workflows/release.yml` re-parsed fresh in this session: `on.push.tags: ['v*.*.*']` only, `permissions.contents: write`, single job `build-and-release` on `runs-on: windows-latest`. Steps in order: checkout → setup-dotnet → restore → build → test → package (fail-fast `$ErrorActionPreference='Stop'` + explicit `Test-Path`/`throw` per required file, confirmed present — this is the CR-01 review fix, verified in current file bytes, not just trusted from the review's disposition table) → `Compress-Archive -DestinationPath FinanceFmtTools.zip` (literal, non-interpolated, confirmed no `${{` on that line) → `softprops/action-gh-release` with `tag_name`, `body_path: RELEASE_NOTES.md`, `files: FinanceFmtTools.zip`. Real live execution against the real remote is the one deliberately-deferred item — see Human Verification. |
| 2 | The packaged zip is named exactly `FinanceFmtTools.zip` in every release, never per-tag-versioned, matching `scripts/install.ps1`'s hardcoded `$AssetName` | ✓ VERIFIED | `grep -c 'FinanceFmtTools.zip' .github/workflows/release.yml` → 4 occurrences; `scripts/install.ps1:90` → `$AssetName = 'FinanceFmtTools.zip'` (exact literal match); `RELEASE.md` uses the identical fixed name in both its packaging step and its manual `gh release create` command. |
| 3 | REL-02: A person or an AI agent can read `RELEASE.md` and run its exact `gh` CLI commands to cut a release manually, zero dependency on GitHub Actions | ✓ VERIFIED | `RELEASE.md` (184 lines) contains a complete, ordered "Fluxo de release" (test → update changelog → bump 2 hardcoded version fields → build/package locally with the same fail-fast fix as CI (`$ErrorActionPreference='Stop'` + `$required`/`throw` loop + a post-package `Compare-Object` content check — WR-01 review fix, confirmed in current bytes) → tag/push → `gh release create vX.Y.Z FinanceFmtTools.zip --title "..." -F RELEASE_NOTES.md`, explicitly labeled "REL-02, zero dependência de CI"). |
| 4 | `RELEASE.md` documents pushing `main` to `origin` strictly before pushing the release tag | ✓ VERIFIED | `git push origin main` appears at line 100 (inside the "Fluxo de release" step 5 code block); the first `git push origin v` line appears at line 101 — main-push line number is lower, confirmed both programmatically (`grep -n`) and by direct reading. An inline explanatory note (lines 104-108) states why the order matters (orphaned-commit risk), matching 05-RESEARCH.md's Pitfall 3. |
| 5 | REL-03: Every release's changelog lives in `RELEASE_NOTES.md`, read by both CI's `body_path` and the manual runbook's `-F` flag | ✓ VERIFIED | `RELEASE_NOTES.md` (68 lines) exists, headed `## Finance Fmt Tools v2.0.0`; `.github/workflows/release.yml`'s `body_path: RELEASE_NOTES.md` and `RELEASE.md`'s `-F RELEASE_NOTES.md` both reference the identical file — single shared source of truth, no drift risk. |
| 6 | The first C#-migration release's changelog explicitly calls out that checkbox preferences no longer persist across Excel restarts, framed as intentional, not a regression | ✓ VERIFIED | `RELEASE_NOTES.md`'s "Mudança de comportamento" section states verbatim that "Alinhar à direita"/"Zero contábil" no longer persist between sessions, explains the old VBA `CustomXMLPart` mechanism being removed, and explicitly labels it "uma decisão deliberada de simplificação de escopo... não uma regressão ou bug" — cross-checked against `REQUIREMENTS.md`'s "Out of Scope" table row ("Persistência das preferências... removida deliberadamente nesta migração"), which matches exactly. |
| 7 | LEGACY-01: The VBA source (`.bas` files) and legacy root PowerShell/`.bat` installer no longer exist on `main`'s working tree, while remaining fully recoverable from `archive/vba-legacy` | ✓ VERIFIED | `ls src/*.bas` → "no matches found"; `ls Install-FinanceFmtTools.*` → "no matches found"; `git ls-tree -r archive/vba-legacy --name-only -- src` → exactly the 6 expected paths (5 `.bas` files + `customUI14.xml`), confirming full recoverability with zero history rewriting. |
| 8 | `src/customUI14.xml` still exists on `main` unmodified, since `FinanceFmtTools.Engine.csproj` actively embeds it at build time | ✓ VERIFIED | `src/customUI14.xml` present (7001 bytes); a fresh `dotnet build src/FinanceFmtTools.sln -c Release` (re-run in this session, not taken from SUMMARY.md) succeeds with **0 Warning(s), 0 Error(s)** across all 3 projects — proves the `EmbeddedResource` reference to `customUI14.xml` is intact. |
| 9 | `dotnet build`/`dotnet test` still succeed after VBA removal | ✓ VERIFIED | Independently re-run in this session: `dotnet build src/FinanceFmtTools.sln -c Release` → 0 Warning(s)/0 Error(s); `dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release` → **40/40 tests passed** (includes 16 `[InlineData]` cases in `AccountingFormatBuilderTests.cs`, confirmed via `grep -c`, matching `RELEASE_NOTES.md`'s "16 combinações" claim). |
| 10 | LEGACY-02: `README.md` documents only the new C# add-in's install/uninstall flow, with the old VBA `.xlam` flow fully absent from active instructions | ✓ VERIFIED | `README.md` (312 lines): contains `scripts/install.ps1`/`scripts/uninstall.ps1` one-liners (byte-identical to `scripts/install.ps1`'s own `.EXAMPLE`); zero occurrences of `CustomXMLPart` or `Install-FinanceFmtTools.ps1`; zero references to any removed `.bas` file names. The only remaining VBA/`.xlam` mentions (4 occurrences) are all inside the intentional "Atualizando da versão VBA" migration-upgrade callout and the historical-reference note pointing at `archive/vba-legacy` — exactly the scoped exception the plan's own acceptance criteria required, not a leftover. |

**Score:** 10/10 code-level truths verified. The one remaining item — an actual live tag push / `windows-latest` Actions run / real `gh release create` against `origin` — is correctly routed to `human_needed`, not counted as a failure (see Human Verification section and Environment/Authorization Constraint above).

### Deferred Items

None — no gaps were deferred to a later milestone phase; this is the final phase of the milestone.

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `.github/workflows/release.yml` | Tag-triggered build/test/package/publish pipeline (REL-01) | VERIFIED | 71 lines. Re-parsed with `yaml.safe_load` in this session: `permissions.contents=='write'` ✓, `runs-on=='windows-latest'` ✓. Actions pinned to real, verified commit SHAs (see Key Link Verification below) — post-review-fix state, not the original mutable `@v4`/`@v2` tags. |
| `.gitignore` | Ignore rules for local release packaging artifacts | VERIFIED | Contains `FinanceFmtTools.zip` and `staging/` as standalone lines, alongside all pre-existing entries (`bin/`, `obj/`, `.vscode/`, `.DS_Store`, `Thumbs.db`) — all 7 confirmed present via direct read. |
| `RELEASE.md` | `gh` CLI manual release runbook (REL-02) | VERIFIED | 184 lines. Contains `gh release create`, `gh release view`, `AddinVersion` version-bump reference, fail-fast packaging script matching the CI fix, and correct push-order (`git push origin main` before any `git push origin v` line). |
| `RELEASE_NOTES.md` | Hand-maintained changelog for the current release (REL-03) | VERIFIED | 68 lines (≥20 required). Headed `v2.0.0`, documents checkbox-persistence removal, references `archive/vba-legacy` and `scripts/install.ps1`. |
| `README.md` | C#-only user-facing documentation (LEGACY-02) | VERIFIED | 312 lines. `scripts/install.ps1` referenced; format tables (`Fin 2D`, `0.00%`, etc.) preserved verbatim; `archive/vba-legacy` referenced as historical pointer. |
| `src/customUI14.xml` | Ribbon XML still embedded by the C# build — must survive LEGACY-01's removal untouched | VERIFIED | Present, unmodified (7001 bytes); `dotnet build` green, proving the `EmbeddedResource` link is intact. |
| `src/*.bas`, `Install-FinanceFmtTools.ps1`/`.bat` | Must be absent from `main`'s working tree | VERIFIED ABSENT | `ls src/*.bas` / `ls Install-FinanceFmtTools.*` both report no matches; fully recoverable on `archive/vba-legacy` (tip `cf2559b`, confirmed via `git ls-tree`). |

### Key Link Verification

| From | To | Via | Status | Details |
|------|-----|-----|--------|---------|
| `.github/workflows/release.yml` | `scripts/install.ps1` | Fixed asset filename `FinanceFmtTools.zip` shared by both | WIRED | Identical literal string in both files (`$AssetName = 'FinanceFmtTools.zip'` in `install.ps1:90`; 4 occurrences in `release.yml`). |
| `.github/workflows/release.yml` | `RELEASE_NOTES.md` | `body_path: RELEASE_NOTES.md` | WIRED | Confirmed present at `release.yml:69`. |
| `.github/workflows/release.yml` | `actions/checkout` (upstream) | Pinned commit SHA `34e114876b0b11c390a56381ad16ebd13914f8d5` | WIRED, VERIFIED REAL | `git ls-remote https://github.com/actions/checkout` (run live in this session, not trusted from the source comment) confirms this exact SHA is tagged both `refs/tags/v4` and `refs/tags/v4.3.1` — the SHA is genuine, not fabricated, and the trailing `# v4.3.1` comment is accurate. |
| `.github/workflows/release.yml` | `actions/setup-dotnet` (upstream) | Pinned commit SHA `67a3573c9a986a3f9c594539f4ab511d57bb3ce9` | WIRED, VERIFIED REAL | `git ls-remote https://github.com/actions/setup-dotnet` confirms this SHA is tagged `refs/tags/v4` and `refs/tags/v4.3.1` — genuine. |
| `.github/workflows/release.yml` | `softprops/action-gh-release` (upstream) | Pinned commit SHA `3bb12739c298aeb8a4eeaf626c5b8d85266b0e65` | WIRED, VERIFIED REAL | `git ls-remote https://github.com/softprops/action-gh-release` confirms this SHA is tagged `refs/tags/v2` and `refs/tags/v2.6.2` — genuine. (05-REVIEW.md itself notes the *reviewer's own* first-suggested SHA for this action was fabricated/incorrect and was replaced during the fix pass — this verification independently re-confirms the *replacement* SHA that actually landed in the file is the real one.) |
| `RELEASE.md` | `RELEASE_NOTES.md` | `-F RELEASE_NOTES.md` flag on `gh release create` | WIRED | Confirmed at `RELEASE.md:125`. |
| `src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj` | `src/customUI14.xml` | `EmbeddedResource Include="../customUI14.xml"` (pre-existing, untouched) | WIRED | Confirmed via successful `dotnet build` (0 Warning(s)/0 Error(s)) after VBA removal — if this link were broken, the build would fail immediately with a missing-file MSBuild error. |
| `README.md` | `scripts/install.ps1` | Documented one-liner install command | WIRED | `README.md:179` reproduces `scripts/install.ps1`'s own `.EXAMPLE` doc-comment content exactly. |

### Data-Flow Trace (Level 4)

Not a UI/data-rendering phase in the traditional sense; the equivalent trace here is "do the pinned Action SHAs actually resolve to real, correct upstream commits, or are they fabricated/copy-paste errors?" — this was independently verified via **live `git ls-remote` calls against the real `actions/checkout`, `actions/setup-dotnet`, and `softprops/action-gh-release` GitHub repositories** (not merely trusted from the workflow file's own trailing comments or from 05-REVIEW.md's fix-disposition claims). All three SHAs resolved to the exact version tags claimed in their trailing comments. This closes the exact failure mode 05-REVIEW.md's own fix-pass note flagged (the reviewer's *first* suggested SHA for `action-gh-release` was itself fabricated and had to be corrected) — this verification did not repeat that mistake, and confirms the *final, shipped* SHA is genuine.

### Behavioral Spot-Checks

| Behavior | Command | Result | Status |
|----------|---------|--------|--------|
| Solution builds cleanly after VBA removal | `dotnet build src/FinanceFmtTools.sln -c Release` (re-run fresh in this session) | `Build succeeded. 0 Warning(s) 0 Error(s)` | ✓ PASS |
| Full test suite passes | `dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release` (re-run fresh) | `Passed! - Failed: 0, Passed: 40, Skipped: 0, Total: 40` | ✓ PASS |
| Release workflow YAML is structurally valid | `python3 -c "import yaml; yaml.safe_load(open('.github/workflows/release.yml'))"` + key assertions | `permissions: {'contents': 'write'}`, `runs-on: windows-latest` | ✓ PASS |
| Pinned Action SHAs are real (not fabricated) | `git ls-remote` against 3 real upstream Action repos | All 3 SHAs found, tagged with the exact versions claimed in comments | ✓ PASS |
| Required packaging binaries actually exist in a real local build output | `ls src/FinanceFmtTools.ComAddin/bin/Release/net48/` | All 4 required DLLs present (`FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, `office.dll`) | ✓ PASS |
| A real `windows-latest` GitHub Actions run against `finance-fmt-tools` | N/A — cannot run without pushing to the real remote | Not executed (by design) | ? SKIP — routed to Human Verification |

### Probe Execution

No `scripts/*/tests/probe-*.sh` files exist in this repository (`find scripts -path '*/tests/probe-*.sh'` returns nothing), and no PLAN/SUMMARY file for this phase references any probe script. Step 7c: SKIPPED (no probes declared or found).

### Independent Checks (run in this session, not taken from SUMMARY.md/REVIEW.md claims)

| Check | Result |
|-------|--------|
| `git status --short` | Clean working tree |
| `git rev-list --left-right --count origin/main...main` | `0  95` — confirms zero commits pushed to `origin`, local `main` is 95 commits ahead |
| `git ls-remote --heads origin` | Only `refs/heads/main` exists remotely, pointing at `8804778` (the pre-migration VBA-era commit — `git show origin/main --stat` confirms commit message "Update README.md com instruções de instalação") — `archive/vba-legacy` has genuinely never been pushed |
| `git branch -a` | Confirms `archive/vba-legacy` exists locally but has no `remotes/origin/archive/vba-legacy` counterpart |
| `git show --stat` on the review-fix commits (`b6871dc`, `51ce0e4`) | `b6871dc` touches exactly the 4 claimed files (`release.yml`, `README.md`, `RELEASE.md`, `RELEASE_NOTES.md`); `51ce0e4` touches only `STATE.md` — matches SUMMARY/REVIEW claims exactly |
| `git ls-tree -r archive/vba-legacy --name-only -- src \| sort` | Exactly the 6 expected paths (5 `.bas` + `customUI14.xml`) |
| `grep -c InlineData AccountingFormatBuilderTests.cs` | 16 — matches `RELEASE_NOTES.md`'s "16 combinações" claim |
| `grep -n "AssetName\s*=" scripts/install.ps1` | `$AssetName = 'FinanceFmtTools.zip'` — matches every reference in `release.yml`/`RELEASE.md` |
| Anti-pattern scan (`TBD\|FIXME\|XXX\|TODO\|HACK\|PLACEHOLDER`, case-insensitive) across all 5 phase deliverables | Zero genuine matches — only Portuguese "todo/todos" ("all"/"every") false positives, same false-positive class Phase 4's verification already documented |
| `LICENSE` file existence (referenced by README's fixed "Licença" section) | Present, 1069 bytes, predates this milestone |

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|-------------|--------------|--------|----------|
| REL-01 | 05-01, 05-04 | Tag push triggers automated build/test/package/publish workflow | CODE VERIFIED — NEEDS HUMAN for live trigger | Workflow YAML structurally correct and re-verified in this session; the actual live `windows-latest` run against the real repo has never happened (by design — authorization boundary, not a defect). |
| REL-02 | 05-02, 05-04 | Documented `gh` CLI manual runbook, zero CI dependency | SATISFIED | `RELEASE.md` fully documents the manual flow; independently re-verified. |
| REL-03 | 05-02, 05-04 | Every release includes changelog notes | SATISFIED | `RELEASE_NOTES.md` exists, is the shared source for both CI and manual paths. |
| LEGACY-01 | 05-03, 05-04 | VBA source preserved on `archive/vba-legacy`, removed from `main`'s active flow | SATISFIED | Confirmed via `git ls-tree`/`ls` in this session. |
| LEGACY-02 | 05-03, 05-04 | README/install instructions reference only the C# add-in | SATISFIED | Confirmed via direct `README.md` read + grep checks in this session. |

No orphaned requirements: `.planning/REQUIREMENTS.md` maps exactly REL-01/02/03 and LEGACY-01/02 to Phase 5, and all five are declared across this phase's 4 plans' `requirements` frontmatter (05-01: REL-01; 05-02: REL-02, REL-03; 05-03: LEGACY-01, LEGACY-02; 05-04: all five, as the closing checkpoint).

**Minor documentation-consistency note (not a code defect):** `REQUIREMENTS.md`'s Traceability table (line 91) marks REL-01's row status as bare `"Pending"`, while the analogous still-open, human-needed rows for Phase 3/4 (RIB-01..04, INST-01..03) all use the more precise `"Code complete — human_needed (...)"` phrasing. Both conventions correctly leave the requirement unchecked in the checklist above it (`- [ ] REL-01`), so there is no functional inconsistency (nothing is prematurely marked done) — only a minor wording/style drift worth tidying up during the milestone's closing housekeeping pass. Does not block phase completion or affect this verification's status.

### Anti-Patterns Found

None. Scanned `.github/workflows/release.yml`, `RELEASE.md`, `RELEASE_NOTES.md`, `README.md`, and `.gitignore` for `TBD`/`FIXME`/`XXX`/`TODO`/`HACK`/`PLACEHOLDER`, "not yet implemented"/"coming soon", empty-body implementations, and hardcoded-empty stub patterns — zero genuine matches (only Portuguese "todo/todos" substring false positives, same class already documented in Phase 4's verification). No debt markers requiring a blocker classification.

**05-REVIEW.md cross-check (fix-pass claims independently re-verified against current file bytes, not trusted from the review's own "fixed" disposition table):**
- CR-01 (packaging step missing fail-fast): confirmed fixed — `$ErrorActionPreference = 'Stop'` plus an explicit `foreach`/`Test-Path`/`throw` loop over all 7 required files now precedes `Compress-Archive` in `release.yml:38-57`.
- WR-01 (manual runbook same gap): confirmed fixed — identical fail-fast pattern plus a post-package `Compare-Object` content-verification block added to `RELEASE.md:59-93`.
- WR-02 (mutable Action tags): confirmed fixed and independently re-verified as *genuinely real* SHAs (not just present in the file) via live `git ls-remote` against all 3 upstream Action repositories — see Key Link Verification / Data-Flow Trace above.
- WR-03 (wrong ribbon button label "Guia Fin"): confirmed fixed — `README.md:243` now reads `Documentação`, matching `src/customUI14.xml`'s actual `label="Documentação"`.
- WR-04 (missing `;@` in date format table): confirmed fixed — `README.md:162-164` now shows `yyyy-mm-dd;@`, `[$-pt-BR]dd/mm/yyyy;@`, `[$-pt-BR]dd/mmm/yyyy;@` with an explanatory sentence, matching `FormatRegistry.cs`'s actual literals.
- WR-05 (stale "Licença" placeholder): confirmed fixed — `README.md:309-311` now reads `MIT — ver [LICENSE](./LICENSE)`; the referenced `LICENSE` file is confirmed present.
- WR-06 (RELEASE_NOTES.md button-name mismatches): confirmed fixed — `RELEASE_NOTES.md:22-23` now uses the real Ribbon labels (`Fin 0D/2D/4D/8D`, `% 2D`/`% 4D`, `Spread bps`, `ISO`/`BR`/`BR Extenso`, `Texto`).
- IN-01 (floating `8.x` SDK version): confirmed fixed — `release.yml:24` now pins `dotnet-version: '8.0.422'`.
- IN-02 (no `iex` risk disclosure): confirmed fixed — `README.md:182-183` adds the disclosure note.
- IN-03 (64-bit worded as hard requirement): confirmed fixed — `README.md:190` softened to match `install.ps1`'s actual non-blocking bitness check.
- IN-04 ("permanent" overstatement): confirmed fixed — `RELEASE_NOTES.md:3-6` reworded to the accurate "overwritten before each new tag" phrasing.

## Human Verification Required

See YAML frontmatter `human_verification` section for the full itemized checklist (2 items). Summary:

### 1. Real Release Execution (REL-01, REL-02 live confirmation)

**Test:** Push `main` and `archive/vba-legacy` to `origin`, push a `vX.Y.Z` tag, and either watch the real `windows-latest` Actions run succeed or execute `RELEASE.md`'s manual `gh release create` fallback.
**Expected:** A real GitHub Release is published with exactly one asset named `FinanceFmtTools.zip`; `gh release view vX.Y.Z` confirms it.
**Why human:** This project's explicit, non-discretionary safety constraint (documented in 05-CONTEXT.md and every Phase 5 plan's threat model) prohibits an autonomous execution agent from pushing to the real, public `origin` remote or cutting a real GitHub Release. This is an authorization-policy boundary, not a technical limitation — the itemized command sequence is fully ready and pre-verified (05-04-SUMMARY.md's resume-signal: `"deferred — not releasing yet"`).

### 2. Live Install Confirmation (post-release)

**Test:** Once a real release exists, run `scripts/install.ps1`'s one-liner on a real Windows+Excel machine.
**Expected:** Add-in installs, "Finance Fmt" Ribbon tab appears.
**Why human:** Requires a live Windows+Excel environment; this re-exercises Phase 3/4's own still-open `human_needed` checklists, not new scope introduced by Phase 5.

## Gaps Summary

No code gaps found. All five Phase 5 local deliverables (`.github/workflows/release.yml`, `.gitignore`, `RELEASE.md`, `RELEASE_NOTES.md`, `README.md`) exist, are substantive (not stubs), are correctly cross-wired to each other and to Phase 4's `scripts/install.ps1` fixed-asset-name contract, and every one of 05-REVIEW.md's 11 findings (1 critical, 6 warnings, 4 info) was independently re-verified as genuinely fixed in the current file bytes — including going one step further than the review itself by live-checking the replacement Action SHAs against their real upstream repositories (`git ls-remote`), since the review's own fix-pass note flagged that its *first* suggested SHA had itself been fabricated.

The phase's own goal statement — "releases are fully automated from a tag push... a documented manual fallback exists... the VBA legacy code/docs are completely out of the active `main` flow" — is **fully achieved at the code/artifact level**. The one remaining piece, an actual live trigger of the real `windows-latest` workflow (or an actual manual `gh release create` against the real remote), is a **deliberate, explicit, non-fabricated `human_needed` item**, not a gap or failure: this project's hard safety constraint requires a human (or a human-authorized agent, in a separate session) to cross the local→real-remote boundary. This mirrors Phase 3/4's own `human_needed` precedent (live-Excel smoke test, live install/uninstall test) and is the correct, by-design outcome for this phase and this environment.

One minor, non-blocking documentation-consistency note was found (REQUIREMENTS.md's REL-01 traceability-row wording differs stylistically from the Phase 3/4 convention for the same class of open item) — flagged above for optional tidy-up, not a phase-blocking issue.

---

_Verified: 2026-07-11T19:03:29Z_
_Verifier: Claude (gsd-verifier)_
