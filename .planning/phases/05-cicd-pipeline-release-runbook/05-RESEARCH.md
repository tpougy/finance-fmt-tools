# Phase 5: CI/CD Pipeline & Release Runbook - Research

**Researched:** 2026-07-11
**Domain:** GitHub Actions CI/CD for a .NET Framework 4.8 COM add-in, `gh` CLI release automation, git branch archival
**Confidence:** HIGH

## Summary

This phase has no unknown technology — it reuses the exact pattern already built and **verified working in production** by the sibling project `outlook-classic-delay-send` (`.github/workflows/release.yml`, `RELEASE.md`), which shipped a real GitHub Actions release 6 days ago on `windows-latest` (`gh run view 28721027038` → success). The build/test/package mechanics for *this* repo were independently re-verified live in this session: `dotnet restore src/FinanceFmtTools.sln`, `dotnet build src/FinanceFmtTools.sln -c Release --no-restore` (0 Warnings/0 Errors), and `dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release --no-build` (40/40 passed) all ran successfully in this Linux/WSL sandbox via `~/.dotnet/dotnet` (SDK 8.0.422, not on default `$PATH`). The 4 files `scripts/install.ps1` requires were confirmed present at `src/FinanceFmtTools.ComAddin/bin/Release/net48/` after that build.

The single most important deviation from the sibling's pattern: the sibling's release zip is **versioned** (`outlook-undo-send-$tag.zip`), but this project's `scripts/install.ps1` (already written in Phase 4) hard-depends on a **fixed, literal, non-versioned** asset name — `FinanceFmtTools.zip` — so that its `releases/latest/download/FinanceFmtTools.zip` URL never has to change between releases. The CI packaging step **must not** copy the sibling's `$tag`-suffixed naming convention.

A second load-bearing discovery: `src/customUI14.xml` is **not** purely-VBA legacy — `FinanceFmtTools.Engine.csproj` embeds it directly (`<EmbeddedResource Include="../customUI14.xml" .../>`) and the running C# add-in reads it at runtime via `RibbonController.GetCustomUiXml()`. Deleting it as part of "remove VBA source" (LEGACY-01) would break the Phase 1-4 build. Only the `.bas` files and the two legacy root-level installer scripts (`Install-FinanceFmtTools.ps1`, `.bat`) are true VBA-only legacy; `customUI14.xml` must stay on `main`.

A branch named `archive/vba-legacy` **already exists locally** (created 2026-07-10, forked from commit `cf2559b`, the last commit before the C# migration started) but has **never been pushed to `origin`** (`gh api .../branches` shows only `main` remotely). This is exactly the safest archival mechanism (a plain branch-and-freeze, zero history rewriting) — it just needs `git push origin archive/vba-legacy`.

**Primary recommendation:** Reuse the sibling's `release.yml`/`RELEASE.md`/`softprops/action-gh-release@v2` pattern near-verbatim, with three deltas: (1) fixed-name zip (`FinanceFmtTools.zip`, no tag suffix), (2) a hand-maintained `RELEASE_NOTES.md` (not `--generate-notes`) as the changelog source given this repo's commit history is dominated by internal GSD phase/plan bookkeeping messages that would make auto-generated notes noisy for end users, and (3) push the already-existing local `archive/vba-legacy` branch, then `git rm` only the `.bas` files + root VBA installer scripts from `main` (never `customUI14.xml`).

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Build/test/package automation | CI/Build Runner (`windows-latest`) | — | Only environment that can compile+package the net48 COM add-in exactly like a release; must mirror the dev-machine `dotnet` commands 1:1 |
| Release publishing (asset upload, Release object) | GitHub Releases (distribution) | CI/Build Runner | GitHub Releases hosts the artifact; CI's last step (`softprops/action-gh-release`) or `gh release create` pushes to it |
| Manual release fallback (REL-02) | Developer / AI-agent local machine | GitHub Releases | Same distribution target, different trigger — a human/agent runs the same `gh`/`dotnet`/`Compress-Archive` commands the CI job runs |
| Changelog / release notes (REL-03) | Git repository (`RELEASE_NOTES.md`) | GitHub Releases (permanent body text) | Authored in-repo before tagging; both CI and manual paths read the same file; once a release is created, its body text is permanent regardless of later file edits |
| VBA legacy archival (LEGACY-01) | Git repository (`archive/vba-legacy` branch) | — | Pure version-control operation, no runtime component, no code changes needed to the archived snapshot |
| End-user installation | End-user Windows/Excel client machine | GitHub Releases | `scripts/install.ps1` (built in Phase 4) fetches the fixed-name asset this phase publishes |
| Documentation (README rewrite, LEGACY-02) | Git repository (`README.md`) | End-user client machine | Read by humans/agents before running the installer one-liner; must contain zero remaining VBA/`.xlam` references |

## Standard Stack

### Core

| Tool | Version | Purpose | Why Standard |
|------|---------|---------|---------------|
| `actions/checkout` | `v4` | Checkout repo in the workflow | `[VERIFIED]` — this exact tag ran successfully on `windows-latest` in the sibling repo's real production release 6 days ago (`gh run view 28721027038`, job green) |
| `actions/setup-dotnet` | `v4`, `dotnet-version: '8.x'` | Install .NET 8 SDK on the runner (builds both net48 via reference-assemblies package and net8.0 targets) | `[VERIFIED]` — same real run confirmed this installs a working SDK that builds a net48 COM add-in project on `windows-latest` |
| `softprops/action-gh-release` | `v2` | Create the GitHub Release + upload the zip asset | `[VERIFIED]` — same real run; note: `v3` exists upstream requiring the Node 24 Actions runtime, `v2.6.2` is the last Node 20 line — the run's annotation shows `v2` still executes correctly today (`gh run view` showed a deprecation *warning*, not a failure) |
| GitHub CLI (`gh`) | 2.4.0 present locally, `2.0+` sufficient per sibling's own documented minimum | Manual release runbook (REL-02) | `[VERIFIED]` — `gh --version` / `gh auth status` run in this session; confirmed authenticated as `tpougy`; `gh release create --help` confirms `--generate-notes`, `-F/--notes-file`, and positional `[<tag>] [<files>...]` all work on this exact installed version |
| `.NET SDK` | 8.0.422 (local), `8.x` (CI) | `dotnet restore/build/test` | `[VERIFIED]` — ran directly in this session via `~/.dotnet/dotnet` (not on default `$PATH` in this sandbox — see Environment Availability) |

### Supporting

| Tool | Version | Purpose | When to Use |
|------|---------|---------|-------------|
| `git` | 2.34.1 (local) | Tagging, branch archival, pushing | Always — no alternative |
| PowerShell `Compress-Archive` | built into PowerShell 5.1+ (present on `windows-latest` by default) | Zip packaging, both in CI (`shell: pwsh`) and in the manual runbook | Standard, zero extra install needed on Windows |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Hand-maintained `RELEASE_NOTES.md` (`body_path`) | `--generate-notes` / `generate_release_notes: true` (auto-generated from merged PRs/commits) | Auto-notes are zero-maintenance but this repo's git history is dominated by internal GSD phase/plan bookkeeping commit messages (`docs(03): auto-generated context`, `chore: enable full autonomous mode`, etc.) — would produce a noisy, low-quality changelog for end users. Sibling project deliberately chose the hand-file approach for its real production release for the same reason. |
| `softprops/action-gh-release` (community Action) | `gh release create` invoked directly via a `run:` step inside the workflow (using the auto-provided `GITHUB_TOKEN`) | Fully equivalent capability, marginally fewer third-party-Action supply-chain dependencies; not chosen because the sibling's proven, already-working YAML uses the Action form and the phase should reuse working patterns rather than re-derive an equivalent — flagged as a legitimate lower-dependency alternative for the planner if minimizing third-party Actions is a priority |
| Fixed-name zip (`FinanceFmtTools.zip`) | Per-tag-versioned zip name (`FinanceFmtTools-vX.Y.Z.zip`), sibling's pattern | **Not available as a real choice for the primary asset** — `scripts/install.ps1`'s `$AssetName`/`$DownloadUrl` (Phase 4, already shipped) hard-requires the fixed literal name. A **second, additional** versioned-name asset can optionally be uploaded alongside it in the same release (for manual/archival download) without conflict — `gh release create`/`softprops` both accept multiple files per release. |
| Single `release.yml` (tag-triggered only) | An additional `ci.yml` running build+test on every push/PR to `main` | Not required by REL-01/REQUIREMENTS.md (only tag-triggered pipeline is a v1 requirement); sibling project also ships with exactly one workflow file. Noted as a defensible future addition, not in scope here. |

**Installation:** No new package-manager installs for this phase — no `npm install`/`pip install`/`dotnet add package`. The phase adds one new file, `.github/workflows/release.yml`, referencing three GitHub Actions Marketplace actions by tag (no local install step).

**Version verification:** All three GitHub Actions above were verified not via `npm view`/registry lookup (not applicable — GitHub Actions Marketplace, not a package registry) but via direct evidence of a real, successful, recent execution:
```bash
gh run list -R tpougy/outlook-classic-delay-send
gh run view 28721027038 -R tpougy/outlook-classic-delay-send
```
Output confirmed: job `build-and-release` succeeded in 51s, 6 days before this research date, using exactly `actions/checkout@v4`, `actions/setup-dotnet@v4`, `softprops/action-gh-release@v2`. The only issue surfaced was an *informational* annotation (not a failure):
> "Node.js 20 is deprecated. The following actions target Node.js 20 but are being forced to run on Node.js 24: actions/checkout@v4, actions/setup-dotnet@v4, softprops/action-gh-release@v2."

This is `[VERIFIED: real GitHub Actions run]` — the strongest possible evidence available without triggering a live run against *this* repo (which is explicitly deferred to the user, see Open Questions).

## Package Legitimacy Audit

This phase introduces **no new npm/PyPI/crates packages** — nothing for `slopcheck`/`npm view`/`pip index versions` to check. The only "external" additions are 3 GitHub Actions Marketplace references inside the new workflow YAML, which are not package-registry artifacts, so the standard slopcheck/registry-verification protocol does not directly apply. A supply-chain-equivalent check was done manually instead:

| Action | Publisher | Track record | Disposition |
|--------|-----------|---------------|-------------|
| `actions/checkout@v4` | `actions` org — GitHub's own official org | Maintained directly by GitHub; used in effectively all GitHub Actions workflows | Approved |
| `actions/setup-dotnet@v4` | `actions` org — GitHub's own official org | Maintained directly by GitHub | Approved |
| `softprops/action-gh-release@v2` | Third-party (`softprops`), not a GitHub org | Long-lived (multi-year), extremely widely used community action; **empirically confirmed working** in a real production release of the sibling project 6 days before this research | Approved — third-party, but proven in this exact organization's own infrastructure days ago |

**Packages removed due to slopcheck [SLOP] verdict:** none (not applicable — no registry packages introduced).
**Packages flagged as suspicious [SUS]:** none.

**Supply-chain hardening note (Security Domain, defense-in-depth):** all three are currently pinned to a mutable major-version tag (`@v4`, `@v2`), matching the sibling's proven pattern and GitHub's own documented convention. Pinning to a full commit SHA (e.g. `actions/checkout@<sha> # v4.x.x`) is a stricter but non-default posture; not required by CLAUDE.md or REQUIREMENTS.md, and not used by the sibling project either — noted as an optional discretionary hardening, not a blocking recommendation.

## Architecture Patterns

### System Architecture Diagram

```
Developer / AI-agent                                    GitHub (origin)
┌─────────────────────────┐                              ┌────────────────────────────┐
│ 1. finish work on main   │──git push origin main───────▶│ main (currently 78 commits │
│ 2. update RELEASE_NOTES  │                              │ behind local — must push   │
│    .md + AddinVersion    │                              │ first, see Open Questions) │
│ 3. git tag -a vX.Y.Z     │──git push origin vX.Y.Z─────▶│ new tag ref                │
└─────────────────────────┘                              └──────────┬─────────────────┘
                                                                      │ tag matches
                                                                      │ 'v*.*.*'
                                                                      ▼
                                                       ┌───────────────────────────────┐
                                                       │ GitHub Actions                │
                                                       │ .github/workflows/release.yml │
                                                       │ runs-on: windows-latest       │
                                                       │                               │
                                                       │  checkout (tag ref)           │
                                                       │  → setup-dotnet 8.x           │
                                                       │  → dotnet restore sln         │
                                                       │  → dotnet build -c Release    │
                                                       │  → dotnet test (Engine.Tests) │
                                                       │  → package step (pwsh):       │
                                                       │      copy 4 DLLs from         │
                                                       │      ComAddin/bin/Release/    │
                                                       │      net48/ + 3 scripts/*.ps1 │
                                                       │      Compress-Archive →       │
                                                       │      FinanceFmtTools.zip      │
                                                       │      (FIXED name, no $tag)    │
                                                       │  → softprops/action-gh-release│
                                                       │      tag_name, body_path:     │
                                                       │      RELEASE_NOTES.md,        │
                                                       │      files: FinanceFmtTools   │
                                                       │      .zip                    │
                                                       └──────────┬────────────────────┘
                                                                  ▼
                                                       ┌───────────────────────────────┐
                                                       │ GitHub Releases                │
                                                       │ /releases/latest/download/     │
                                                       │   FinanceFmtTools.zip          │
                                                       └──────────┬────────────────────┘
                                                                  │ irm ... | iex
                                                                  ▼
                                                       ┌───────────────────────────────┐
                                                       │ End-user Windows + Excel       │
                                                       │ scripts/install.ps1 (Phase 4)  │
                                                       │ downloads latest/download/     │
                                                       │ FinanceFmtTools.zip, extracts, │
                                                       │ registers HKCU                 │
                                                       └───────────────────────────────┘

MANUAL FALLBACK (REL-02, no CI dependency):
Developer/agent runs the identical restore/build/test/package commands locally on
Windows, then `gh release create vX.Y.Z FinanceFmtTools.zip --title ... -F RELEASE_NOTES.md`
directly against GitHub Releases — same destination, bypasses the Actions runner entirely.
```

### Recommended Project Structure

```
.github/
└── workflows/
    └── release.yml         # NEW — tag-triggered build/test/package/publish (REL-01)
RELEASE.md                  # NEW — gh CLI manual runbook (REL-02)
RELEASE_NOTES.md            # NEW — current-release changelog, read by CI's body_path (REL-03)
README.md                   # REWRITTEN — C#-only install/usage docs (LEGACY-02)
scripts/
├── install.ps1             # existing (Phase 4) — unchanged, defines the fixed asset name
├── uninstall.ps1           # existing (Phase 4) — unchanged
└── verify-environment.ps1  # existing (Phase 4) — unchanged
src/
├── customUI14.xml          # STAYS on main — actively embedded by FinanceFmtTools.Engine.csproj
├── FinanceFmtTools.Engine/
├── FinanceFmtTools.Engine.Tests/
├── FinanceFmtTools.ComAddin/
└── FinanceFmtTools.sln
# REMOVED from main (git rm, preserved via archive/vba-legacy branch history):
#   src/ThisWorkbook.bas, src/modConfig.bas, src/modFormatEngine.bas,
#   src/modRibbon.bas, src/modUtils.bas
#   Install-FinanceFmtTools.ps1, Install-FinanceFmtTools.bat  (repo root)
```

### Pattern 1: Tag-triggered release workflow (single job, `windows-latest`)
**What:** One workflow file, triggered only on `push: tags: v*.*.*`, that does restore → build → test → package → publish in a single job — no matrix, no separate CI-on-push workflow.
**When to use:** Exactly this phase's scope (REL-01). Matches the sibling project's proven, minimal pattern; do not add complexity (multi-job, matrix builds) not required by any REQUIREMENTS.md line.
**Example (adapted from sibling's real, working `.github/workflows/release.yml`):**
```yaml
# Source: outlook-classic-delay-send/.github/workflows/release.yml (verified working,
# gh run view 28721027038, 2026-07-04/05), adapted for FinanceFmtTools's fixed-name zip.
name: Release

on:
  push:
    tags:
      - 'v*.*.*'

permissions:
  contents: write

jobs:
  build-and-release:
    runs-on: windows-latest

    steps:
      - name: Checkout
        uses: actions/checkout@v4

      - name: Setup .NET
        uses: actions/setup-dotnet@v4
        with:
          dotnet-version: '8.x'

      - name: Restore
        run: dotnet restore src\FinanceFmtTools.sln

      - name: Build
        run: dotnet build src\FinanceFmtTools.sln -c Release --no-restore

      - name: Test
        run: dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release --no-build

      - name: Package
        shell: pwsh
        run: |
          $binSrc = "src\FinanceFmtTools.ComAddin\bin\Release\net48"
          New-Item -ItemType Directory -Path staging -Force | Out-Null

          Copy-Item "$binSrc\FinanceFmtTools.ComAddin.dll"       staging\
          Copy-Item "$binSrc\FinanceFmtTools.Engine.dll"         staging\
          Copy-Item "$binSrc\Microsoft.Office.Interop.Excel.dll" staging\
          Copy-Item "$binSrc\office.dll"                         staging\
          Copy-Item scripts\install.ps1            staging\
          Copy-Item scripts\uninstall.ps1          staging\
          Copy-Item scripts\verify-environment.ps1 staging\

          # FIXED literal name -- scripts/install.ps1's $AssetName constant and its
          # ".../releases/latest/download/FinanceFmtTools.zip" URL depend on this exact
          # filename never changing between releases (unlike a per-tag-versioned name).
          Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force

      - name: Create GitHub Release
        uses: softprops/action-gh-release@v2
        with:
          tag_name: ${{ github.ref_name }}
          name: Finance Fmt Tools ${{ github.ref_name }}
          body_path: RELEASE_NOTES.md
          files: FinanceFmtTools.zip
```

### Pattern 2: Manual `gh` CLI release fallback (REL-02)
**What:** The exact same restore/build/test/package sequence, run locally by a human or an AI agent on a Windows machine, followed by `gh release create` — no dependency on GitHub Actions at all.
**When to use:** CI is down, or a release needs to be cut without waiting for/triggering a workflow run.
**Example (adapted from sibling's real `RELEASE.md`):**
```powershell
# Source: outlook-classic-delay-send/RELEASE.md section "5. (Opcional) Criar a release
# manualmente via gh" -- adapted for this repo's fixed asset name and repo path.
dotnet restore src\FinanceFmtTools.sln
dotnet build src\FinanceFmtTools.sln -c Release --no-restore
dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release --no-build

$binSrc = "src\FinanceFmtTools.ComAddin\bin\Release\net48"
New-Item -ItemType Directory -Path staging -Force | Out-Null
Copy-Item "$binSrc\FinanceFmtTools.ComAddin.dll", "$binSrc\FinanceFmtTools.Engine.dll", `
          "$binSrc\Microsoft.Office.Interop.Excel.dll", "$binSrc\office.dll" staging\
Copy-Item scripts\install.ps1, scripts\uninstall.ps1, scripts\verify-environment.ps1 staging\
Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force

git tag -a v2.0.0 -m "v2.0.0 — Migração completa VBA -> C# COM add-in"
git push origin main
git push origin v2.0.0

gh release create v2.0.0 FinanceFmtTools.zip `
  --title "Finance Fmt Tools v2.0.0" `
  -F RELEASE_NOTES.md

gh release view v2.0.0
```

### Anti-Patterns to Avoid
- **Versioning the primary release zip's filename (`FinanceFmtTools-v2.0.0.zip`):** breaks `scripts/install.ps1`'s fixed `.../latest/download/FinanceFmtTools.zip` URL, which was written in Phase 4 assuming a literal, unchanging asset name. If a versioned copy is wanted too, upload it as a *second, additional* asset in the same release — never as a replacement for the fixed-name one.
- **Deleting `src/customUI14.xml` as part of "remove VBA source":** it is an active build input for `FinanceFmtTools.Engine.csproj` (`<EmbeddedResource Include="../customUI14.xml" .../>`) — removing it breaks `dotnet build` for the currently-shipping C# add-in. Only the `.bas` files and the two root-level legacy installer scripts are true VBA-only legacy.
- **Relying on `--generate-notes`/auto-generated release notes as the sole REL-03 mechanism here:** this repo's commit history is dominated by internal GSD-workflow bookkeeping commits (`docs(0X): auto-generated context`, `chore: enable full autonomous mode`, etc.) that would surface directly in an auto-generated changelog, producing noise with no value to an end user installing an Excel add-in.
- **Force-pushing or rewriting history to remove the VBA files' past commits:** unnecessary and destructive — a plain `git rm` on `main` (with the `.bas` files' full history still reachable via the already-existing `archive/vba-legacy` branch) fully satisfies LEGACY-01 with zero history loss.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Creating a GitHub Release + uploading an asset from CI | A custom `curl`/`Invoke-RestMethod` call against the GitHub REST API with a hand-rolled multipart upload | `softprops/action-gh-release@v2` (already proven working in this exact organization's sibling repo) or `gh release create` inside a `run:` step | Release-asset upload has real edge cases (large file chunking, correct `Content-Type`, retry-on-5xx) already solved by both `gh` and this Action; hand-rolling risks subtly broken uploads that only surface on a real tag push |
| Auto-generating a semantically meaningful changelog from raw git history | A custom commit-parser/changelog generator | Hand-maintained `RELEASE_NOTES.md`, one entry per release, overwritten before each tag (matches the sibling project's real, working convention) | This repo's commit messages are optimized for GSD-workflow bookkeeping, not end-user communication — a custom parser would need non-trivial filtering logic to be useful; a 5-minute manual write is simpler and higher quality |
| Detecting the .NET Framework 4.8 reference assemblies on the CI runner | Installing the full .NET Framework Developer Pack via a custom runner step | `Microsoft.NETFramework.ReferenceAssemblies` NuGet package (already referenced in both `.csproj` files since Phase 1/3) | Already solved, already verified in this session (`dotnet build` succeeded for net48 with zero extra runner setup) — the package supplies compile-time reference assemblies independent of any OS-level Framework install |

**Key insight:** every piece of this phase's core mechanism (build, test, package, publish, archive, document) already has a proven, working reference implementation one directory away (`outlook-classic-delay-send`). The only genuinely new work is: (1) the two fixed-name/asset-layout deltas described above, (2) the VBA archival git operations, and (3) the README rewrite content itself.

## Runtime State Inventory

This phase is an archival/removal operation (VBA source leaving `main`'s active flow) with a real "still-referenced-at-build-time" nuance, so this inventory is included even though it is not a rename phase per se.

| Category | Items Found | Action Required |
|----------|-------------|------------------|
| Stored data | None — the old VBA add-in's persisted preferences (`CustomXMLPart` inside the `.xlam`) live only inside a compiled `.xlam` binary that was never committed to this git repo (confirmed: `README.md`/`CLAUDE.md` both state the `.xlam` is a GitHub Release asset only, never tracked in git). Nothing in this repo's history needs a data migration. | None |
| Live service config | Two pre-existing GitHub Releases (`v1.0.0`, `v1.0.1`) already exist remotely, each carrying a `FinanceFmtTools.xlam` asset (`gh release list` confirmed both). These are not deleted by this phase (out of scope) but are worth an explicit decision: leave as historical/legacy releases, or add a note to their descriptions marking them superseded. Not required by any REQUIREMENTS.md line — flagged as an Open Question below. | Decision needed (optional) |
| OS-registered state | Out of this repo's control, but real: any end-user machine that installed the **old** VBA `.xlam` add-in has it registered via `%APPDATA%\Microsoft\AddIns` + Excel's own `AddIns` collection (a completely different mechanism than the new HKCU/COM registration). Running both simultaneously risks a duplicate/confusing "Finance Fmt" ribbon tab. | README (LEGACY-02) should include an explicit "upgrading from the VBA version" note instructing users to remove the old `.xlam` (Excel > File > Options > Add-ins > manage > uncheck/remove, or delete the file from `%APPDATA%\Microsoft\AddIns`) before running the new installer. |
| Secrets/env vars | None added. `GITHUB_TOKEN` is auto-provided by the Actions runtime (no repo secret to create); the manual `gh` runbook uses the developer/agent's own already-authenticated `gh auth login` session (confirmed present and authenticated as `tpougy` in this sandbox). Repo's default Actions workflow permission is `read` (`gh api repos/tpougy/finance-fmt-tools/actions/permissions/workflow` → `"default_workflow_permissions":"read"`) — this is exactly why the workflow YAML must explicitly declare `permissions: contents: write` (an explicit workflow-level block overrides the repo default). | Workflow YAML must include the explicit `permissions:` block (already in the Code Example above) — not optional here, this repo's default is read-only. |
| Build artifacts | `bin/`/`obj/` already gitignored. No `releases/`/`staging/` ignore entries exist yet in this repo's `.gitignore` (sibling project ignores `releases/*.zip` and `staging/`) — if the manual runbook or local testing produces a `FinanceFmtTools.zip`/`staging/` folder at the repo root, it could be accidentally `git add`-ed. | Add `FinanceFmtTools.zip` (or a `releases/`/`staging/` pattern) to `.gitignore` as part of this phase's work. |

**Additional build-dependency finding (not a rename-phase category but same class of risk):** `src/customUI14.xml` is simultaneously (a) historically VBA-era Ribbon XML and (b) a currently-active embedded resource for the C# build (`FinanceFmtTools.Engine.csproj`, confirmed via `grep` and via `RibbonController.cs`'s suffix-matching resource loader). It must **not** be deleted as part of LEGACY-01's ".bas removal" — see Common Pitfalls below.

## Common Pitfalls

### Pitfall 1: Deleting `src/customUI14.xml` along with the `.bas` files
**What goes wrong:** `dotnet build` fails immediately (`FinanceFmtTools.Engine.csproj`'s `<EmbeddedResource Include="../customUI14.xml" .../>` line references a now-missing file), breaking every phase's tested code (Phases 1-4, 40/40 tests, 0 Warnings/0 Errors baseline).
**Why it happens:** `customUI14.xml` sits in the same `src/` folder as the `.bas` files and looks like it belongs to the "VBA legacy" bucket being removed by LEGACY-01, but it is Ribbon XML (Office Fluent UI schema), not VBA code, and Phase 2 deliberately linked it into the C# build rather than duplicating it (`02-CONTEXT.md`/STATE.md: "linked (not duplicated) into FinanceFmtTools.Engine.csproj via MSBuild EmbeddedResource Link").
**How to avoid:** LEGACY-01's `git rm` must target only: `src/ThisWorkbook.bas`, `src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`, `Install-FinanceFmtTools.ps1`, `Install-FinanceFmtTools.bat`. Leave `src/customUI14.xml` untouched on `main` (it already exists, unmodified, on `archive/vba-legacy` too, since that branch forked from a later commit).
**Warning signs:** a `dotnet build` failure referencing a missing embedded resource file, or `RibbonController.GetCustomUiXml()` throwing `InvalidOperationException("Embedded resource 'customUI14.xml' not found...")` at runtime if the deletion somehow passed a stale/cached build.

### Pitfall 2: Publishing the release zip under a versioned filename
**What goes wrong:** `scripts/install.ps1`'s one-liner flow (`.../releases/latest/download/FinanceFmtTools.zip`) 404s — INST-01 (Phase 4, already shipped and code-reviewed) silently stops working for every future release.
**Why it happens:** the sibling project's own working pattern versions its zip filename (`outlook-undo-send-$tag.zip`); copying that convention verbatim is the natural (wrong) move.
**How to avoid:** the packaging step's `Compress-Archive -DestinationPath` must be the literal string `FinanceFmtTools.zip`, every release, no `$tag`/`${{ github.ref_name }}` interpolation in that specific filename. Confirmed via `scripts/install.ps1`'s own constants: `$AssetName = 'FinanceFmtTools.zip'` and its inline comment: *"todo release do Phase 5 (CI) deve publicar seu zip sob este nome literal fixo."*
**Warning signs:** `install.ps1`'s `Invoke-WebRequest -Uri $DownloadUrl` throwing a 404, or `gh release view` showing an asset name that doesn't match `FinanceFmtTools.zip` exactly.

### Pitfall 3: Tagging/pushing before pushing `main`
**What goes wrong:** `git push origin main` has not happened in this session — local `main` is **78 commits ahead of `origin/main`** (verified via `git status`/`git log origin/main..main`), meaning `origin/main` on GitHub is still sitting at the pre-Phase-1 commit (`8804778`, VBA era). If a release tag is pushed without first pushing `main`, GitHub's UI/API will show the release's target commit as unreachable from any pushed branch (confusing, though the tag push itself will still succeed and still trigger the workflow, since `git push` transfers whatever commit objects the tag needs regardless of branch state).
**Why it happens:** all of Phases 1-4's work happened as local commits only; nothing has been pushed since the milestone started.
**How to avoid:** both the CI-triggered flow and the manual runbook must document `git push origin main` **before** (or together with) `git push origin vX.Y.Z` — exactly as the sibling's own `RELEASE.md` step 4 already does (`git push origin main` immediately before `git push origin v1.4.0`).
**Warning signs:** `gh release view` or the GitHub web UI showing "This tag has no matching branch" or the release's commit history looking truncated/disconnected.

### Pitfall 4: First C#-release changelog omitting the preference-persistence behavior change
**What goes wrong:** users upgrading from the VBA add-in silently lose their "Alinhar à direita"/"Zero contábil" checkbox preferences across Excel restarts (VBA persisted these via `CustomXMLPart`; the C# rewrite explicitly does not — `REQUIREMENTS.md`'s "Out of Scope" table: *"Persistência das preferências ... removida deliberadamente nesta migração"*). Without a call-out, this reads as a bug report, not an intentional design decision.
**Why it happens:** it is a genuine, deliberate scope-narrowing decision made early in the milestone, easy to forget when writing release notes focused on "what's new" rather than "what's different."
**How to avoid:** the first C#-migration release's `RELEASE_NOTES.md` entry should explicitly call out (a) the new install/uninstall mechanism (HKCU + PowerShell, replacing the old `.xlam`-in-`%APPDATA%\Microsoft\AddIns` flow) and (b) that the two checkboxes no longer persist between sessions, and always start at their documented defaults (off / on respectively).
**Warning signs:** none automatic — this is a documentation-completeness pitfall, not a code defect; catch it in the plan-checker/review step by cross-referencing `RELEASE_NOTES.md` content against `REQUIREMENTS.md`'s "Out of Scope" table.

### Pitfall 5: Assuming `dotnet` is on `$PATH` in every execution environment
**What goes wrong:** `command -v dotnet` / `which dotnet` returns nothing in this sandbox even though a working 8.0.422 SDK is installed at `~/.dotnet/dotnet` — any verification step written as a bare `dotnet ...` command will silently appear to "not exist" rather than fail loudly, if run in a raw shell without the user's normal profile sourced.
**Why it happens:** the SDK was installed to `~/.dotnet` (the typical dotnet-install.sh location) but that directory is not guaranteed to be on `$PATH` in every non-interactive shell invocation.
**How to avoid:** any locally-run verification command in this environment should either source the user's shell profile first or use the absolute path `~/.dotnet/dotnet`. This does not affect the CI workflow itself (GitHub-hosted `windows-latest` runners have `dotnet` on `PATH` natively via `actions/setup-dotnet`).
**Warning signs:** `command not found: dotnet` immediately followed by a `dotnet --version` success once `export PATH="$HOME/.dotnet:$PATH"` is added — exactly what happened in this research session.

## Code Examples

### `RELEASE.md` runbook skeleton (REL-02)
```markdown
# Runbook de Release — Finance Fmt Tools

## Ferramentas necessárias
| Ferramenta | Versão mínima |
|---|---|
| GitHub CLI (gh) | 2.0+ |
| .NET 8 SDK | 8.0+ |
| PowerShell | 5.1+ |
| Windows | necessário para compilar o net48 COM add-in |

## Fluxo de release
1. Testes passam: `dotnet test src\FinanceFmtTools.Engine.Tests\FinanceFmtTools.Engine.Tests.csproj -c Release`
2. Atualize `RELEASE_NOTES.md` com o changelog desta versão
3. Atualize a versão em `src\FinanceFmtTools.ComAddin\FinanceFmtTools.ComAddin.csproj`
   (`<Version>`) e a constante `AddinVersion` em `AddInHost.cs` (usada pelo diálogo "Sobre")
4. Compile e empacote (ver Pattern 2 do 05-RESEARCH.md) -> `FinanceFmtTools.zip`
5. `git tag -a vX.Y.Z -m "..."`; `git push origin main`; `git push origin vX.Y.Z`
6. (Automático) o workflow dispara e publica -- OU (manual) `gh release create vX.Y.Z FinanceFmtTools.zip -F RELEASE_NOTES.md`
7. `gh release view vX.Y.Z`
```

### Auditing this repo's real Actions permissions (verified in this session)
```bash
# Confirms the repo's default token permission is read-only, which is exactly why
# the workflow YAML's explicit `permissions: contents: write` block is required.
gh api repos/tpougy/finance-fmt-tools/actions/permissions/workflow
# => {"default_workflow_permissions":"read","can_approve_pull_request_reviews":false}
```

### Archiving the VBA branch (git operations verifiable in this sandbox)
```bash
# archive/vba-legacy already exists locally (forked from cf2559b, pre-migration commit)
# but has never been pushed -- confirmed via:
git ls-remote --heads origin              # shows only 'main' remotely today
gh api repos/tpougy/finance-fmt-tools/branches --jq '.[].name'   # => "main"

# Safest archival step: just push the already-existing branch, no rewrite needed.
git push origin archive/vba-legacy

# Then on main, remove only the true VBA-only files (never customUI14.xml):
git rm src/ThisWorkbook.bas src/modConfig.bas src/modFormatEngine.bas \
       src/modRibbon.bas src/modUtils.bas \
       Install-FinanceFmtTools.ps1 Install-FinanceFmtTools.bat
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|---------------|--------|
| VBA `.xlam` manually assembled, uploaded as a GitHub Release asset by hand | `dotnet build`-produced COM add-in, packaged and published by an automated `windows-latest` GitHub Actions workflow | This phase (Phase 5) | Removes the manual "reopen VBA project, re-import modules, re-inject Ribbon XML via Custom UI Editor, Save As .xlam, upload by hand" process entirely — `DEV-01`'s "100% via terminal" goal is only fully realized once release, not just build/test, is terminal-only |
| `Install-FinanceFmtTools.ps1` root script installing via Excel COM automation (`New-Object -ComObject Excel.Application`, `$excel.AddIns.Add(...)`) into `%APPDATA%\Microsoft\AddIns` | `scripts/install.ps1` (Phase 4) registering directly into `HKCU` registry keys, no Excel COM automation, no admin | Phase 4 (already shipped) | This phase's README rewrite (LEGACY-02) must describe *this* new mechanism only |

**Deprecated/outdated:**
- `Install-FinanceFmtTools.ps1`/`.bat` (repo root): fully superseded by `scripts/install.ps1`; also contains a pre-existing, now-irrelevant bug (`.bat` references a stale filename `Install-RBRFinanceTools.ps1` that doesn't match `Install-FinanceFmtTools.ps1` — moot once both files are removed under LEGACY-01/02).
- `README.md`'s entire current content: 100% VBA/`.xlam`-oriented (format tables are still accurate/reusable for the *user-facing* format descriptions, but every install/architecture/persistence section must be rewritten).

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | Recommending `v2.0.0` as the first C#-migration release tag (vs. continuing `v1.x`) | Code Examples / Pattern 2 | Low — purely a SemVer-labeling choice with no functional effect; if the user prefers continuing `v1.x.y`, only the example tag string changes, nothing structural |
| A2 | Recommending a hand-maintained `RELEASE_NOTES.md` file name (vs. `CHANGELOG.md` or another name) | Standard Stack / Architecture Patterns | Low — arbitrary naming choice, matches the sibling project's convention for pattern-reuse consistency; any filename works as long as the workflow's `body_path` matches it |
| A3 | Recommending the old GitHub Releases `v1.0.0`/`v1.0.1` (VBA `.xlam` assets) be left as-is rather than annotated/deleted | Runtime State Inventory | Low — no REQUIREMENTS.md line mandates touching them; if the user wants them marked "superseded," it's a one-line `gh release edit` addition |

**None of the above are load-bearing for REL-01/REL-02/REL-03/LEGACY-01/LEGACY-02** — all core mechanics (workflow YAML structure, asset naming, branch archival, permissions requirement) are `[VERIFIED]` via direct execution or via the sibling project's real, recent, successful production run.

## Open Questions

1. **Should `main` be pushed to `origin` as part of this phase's own work, or is that a separate/manual step the user performs later?**
   - What we know: local `main` is 78 commits ahead of `origin/main`; the tag-push workflow will trigger correctly regardless (tags carry their own commit objects), but a clean release requires `main` to be pushed too (matches sibling's `RELEASE.md` step 4 ordering).
   - What's unclear: whether pushing 78 commits of internal migration history to a public repo (`visibility: public`, confirmed via `gh api`) during this phase is within scope, or whether the user wants to review/curate history first.
   - Recommendation: the plan should include pushing `main` as an explicit task (it is a prerequisite for any real release, automated or manual), but flag the actual tag-push/release-creation as the `human_needed`/`checkpoint:human-verify` step this phase's own environment note already anticipates — consistent with Phases 3/4's precedent of deferring the one genuinely unverifiable live step to the user.

2. **Should the two pre-existing legacy GitHub Releases (`v1.0.0`, `v1.0.1`, VBA `.xlam` assets) be edited/labeled or left untouched?**
   - What we know: both exist and are unaffected by anything in this phase's REQUIREMENTS.md.
   - What's unclear: whether leaving a "latest" C# release next to two older VBA releases in the same release list could confuse a future user browsing `github.com/tpougy/finance-fmt-tools/releases`.
   - Recommendation: out of scope for REL-01..03/LEGACY-01..02; leave untouched unless the user asks otherwise during plan review.

3. **Should the new C# add-in's version numbering (`FinanceFmtTools.ComAddin.csproj`'s `<Version>`, `AddInHost.cs`'s hardcoded `AddinVersion` const) be wired to the git tag automatically (e.g. `dotnet build -p:Version=X.Y.Z`), or updated manually per the runbook checklist?**
   - What we know: today both are hardcoded (`1.0.0.0` / `"1.0.0"`), completely disconnected from git tags; the sibling project's own convention is also manual (`RELEASE.md`: "AssemblyInfo.cs tem a versão correta").
   - What's unclear: whether automating this now is worth touching already-tested Phase 3 code (`AddInHost.cs`) vs. adding one manual checklist line to `RELEASE.md`.
   - Recommendation: manual checklist line (lowest risk, matches proven sibling convention) — automating this is a reasonable v2 improvement, not required by any REQUIREMENTS.md line.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|--------------|-----------|---------|----------|
| `.NET SDK` (`dotnet`) | REL-01 build/test steps, manual runbook | ✓ (verified — ran in this session) | 8.0.422, installed at `~/.dotnet`, not on default `$PATH` | `export PATH="$HOME/.dotnet:$PATH"` or invoke `~/.dotnet/dotnet` directly |
| `git` | LEGACY-01 branch push, tagging | ✓ | 2.34.1 | — |
| GitHub CLI (`gh`) | REL-02 manual runbook, verification | ✓ authenticated as `tpougy` | 2.4.0 (old, but all needed flags confirmed present via `--help`) | — |
| GitHub Actions `windows-latest` runner | REL-01 automated pipeline | Cannot execute a *real* run against **this** repo from this sandbox (no way to push a real tag and observe Actions without user authorization/cost) | — | Sibling project's analogous workflow verified working 6 days ago (proof-of-pattern); an actual live tag push against `finance-fmt-tools` is the one item this research cannot close — see Open Question 1 |
| PowerShell / `Compress-Archive` | Packaging step (CI + manual) | N/A in this Linux sandbox (not needed locally — CI runs on `windows-latest`, which ships PowerShell 5.1+ natively) | — | — |

**Missing dependencies with no fallback:**
- None. The only unverifiable item (a real `windows-latest` Actions run against this specific repo) has a documented, non-blocking path forward: push `main`, push a real version tag, and let the workflow run for real — an explicit human-authorized action, not something this research can or should fake.

**Missing dependencies with fallback:**
- `dotnet` not on default `$PATH` in this sandbox — resolved via `export PATH="$HOME/.dotnet:$PATH"` (harmless local-environment quirk, does not affect CI).

## Security Domain

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-------------------|
| V2 Authentication | No | No new authentication surface — `gh`/Actions use existing GitHub token-based auth, not custom |
| V3 Session Management | No | N/A |
| V4 Access Control | Yes | `permissions: contents: write` scoped narrowly to exactly what release creation needs (not `write-all`); repo default is already `read` (verified), so the workflow must explicitly elevate only this one permission |
| V5 Input Validation | No (minimal) | The workflow has no user-controlled input surface beyond the tag name itself (`github.ref_name`), which GitHub already validates against the `v*.*.*` trigger pattern before the job even starts |
| V6 Cryptography | No | N/A — no crypto operations in this phase |

### Known Threat Patterns for this stack

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|----------------------|
| Overly broad `GITHUB_TOKEN` permissions (e.g. `permissions: write-all`) letting a compromised/malicious dependency in the build graph push arbitrary changes | Elevation of Privilege | Explicit least-privilege `permissions: contents: write` block (already the plan) — nothing broader |
| Supply-chain compromise of a third-party Action via a mutable tag being repointed (`@v2` retargeted to a malicious commit) | Tampering | Currently mitigated only by using a well-known, long-lived, empirically-proven action (`softprops/action-gh-release`); optional stricter mitigation is pinning to a commit SHA instead of a tag (not done by the sibling project either — documented as optional hardening, not a blocking requirement here) |
| Secrets exposure via workflow logs (e.g., accidentally echoing `GITHUB_TOKEN`) | Information Disclosure | Not applicable here — no step in this pattern ever echoes the token; both `gh release create` and `softprops/action-gh-release` consume it internally without exposing it in logs |
| Zip-slip / path traversal when a downloaded release zip is extracted | Tampering | Already handled by Phase 4's `scripts/install.ps1` (extracts to a fresh temp directory under `%TEMP%`, never directly into the install target) — this phase only needs to keep producing a well-formed, flat zip; no new risk introduced by the packaging step itself |

## Sources

### Primary (HIGH confidence)
- `outlook-classic-delay-send/.github/workflows/release.yml` — read directly, full file
- `outlook-classic-delay-send/RELEASE.md` — read directly, full file
- `outlook-classic-delay-send/RELEASE_NOTES.md` — read directly, full file
- `gh run view 28721027038 -R tpougy/outlook-classic-delay-send` — real, successful workflow run, 6 days before this research
- Direct execution in this session: `dotnet restore`/`dotnet build`/`dotnet test` against `src/FinanceFmtTools.sln` (0 Warnings/0 Errors, 40/40 tests)
- `gh api repos/tpougy/finance-fmt-tools/actions/permissions/workflow` — real repo configuration (`default_workflow_permissions: read`)
- `scripts/install.ps1`, `scripts/uninstall.ps1` — read directly, full files (fixed asset-name constants, file-layout requirements)
- `src/FinanceFmtTools.ComAddin/Connect.cs`, `FinanceFmtTools.ComAddin.csproj`, `FinanceFmtTools.Engine.csproj`, `RibbonController.cs`, `AddInHost.cs` — read directly
- `.planning/REQUIREMENTS.md`, `.planning/STATE.md`, `.planning/ROADMAP.md`, `05-CONTEXT.md` — read directly

### Secondary (MEDIUM confidence)
- WebSearch results on `actions/checkout`/`actions/setup-dotnet`/`softprops/action-gh-release` latest major versions — internally inconsistent dates (likely fetch-model fabrication) on exact release dates; version *numbers* (v4/v5/v7, v3 vs v2) are directionally consistent across two independent searches, but superseded in this document's actual recommendation by the stronger, dated, empirical evidence from the sibling repo's real run (see Primary sources) — not used as the basis for any final recommendation, included only as background context on the wider ecosystem's version trajectory

### Tertiary (LOW confidence)
- None retained — every claim that could be verified against the sibling repo, this repo's own files, or a live command in this session was verified that way instead of left as a raw web search claim.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — every tool/action version was either run directly in this session or confirmed via a real, recent, successful GitHub Actions run in a sibling repo under the same GitHub account
- Architecture: HIGH — the workflow/runbook pattern is a near-verbatim adaptation of a proven-working sibling implementation, with the two genuinely project-specific deltas (fixed zip name, `customUI14.xml` build dependency) independently verified against this repo's own source
- Pitfalls: HIGH — all 5 pitfalls are grounded in direct evidence from this repo's own files/git state (not speculative), except Pitfall 4 (changelog completeness) which is a documentation-quality judgment call, not a verifiable fact

**Research date:** 2026-07-11
**Valid until:** 30 days (GitHub Actions Marketplace action versions and `gh` CLI syntax are relatively stable; re-verify action versions if this research is used more than ~30 days from now, especially given the observed Node 20→24 Actions runtime transition in progress)
