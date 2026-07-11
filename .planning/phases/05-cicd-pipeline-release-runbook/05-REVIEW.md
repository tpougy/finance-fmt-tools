---
phase: 05-cicd-pipeline-release-runbook
reviewed: 2026-07-11T18:53:28Z
depth: standard
files_reviewed: 4
files_reviewed_list:
  - .github/workflows/release.yml
  - RELEASE.md
  - RELEASE_NOTES.md
  - README.md
findings:
  critical: 1
  warning: 6
  info: 4
  total: 11
status: fixed
fix_disposition:
  CR-01: fixed
  WR-01: fixed
  WR-02: fixed
  WR-03: fixed
  WR-04: fixed
  WR-05: fixed
  WR-06: fixed
  IN-01: fixed
  IN-02: fixed
  IN-03: fixed
  IN-04: fixed
---

**Fix pass completed 2026-07-11. All 11 findings fixed directly.** Notable: WR-02's fix (pinning `actions/checkout`/`actions/setup-dotnet`/`softprops/action-gh-release` to immutable commit SHAs, verified for real via `git ls-remote` against each action's actual repo — the reviewer's own suggested SHA for `action-gh-release` was in fact incorrect/fabricated and was replaced with the verified real SHA for the latest available tag, `v2.6.2`) supersedes 05-01-PLAN.md's original acceptance-criteria substring checks (`grep -q 'actions/checkout@v4'` etc.), which now fail by design since the mutable tag reference no longer appears literally in the file — the SHA + trailing `# v4.3.1`-style comment is the intentionally stronger replacement. Documented in STATE.md.

# Phase 5: Code Review Report

**Reviewed:** 2026-07-11T18:53:28Z
**Depth:** standard
**Files Reviewed:** 4
**Status:** issues_found

## Summary

Reviewed the tag-triggered GitHub Actions release workflow, the manual `gh` CLI release runbook,
the hand-maintained changelog, and the rewritten README for the C# migration. Cross-referenced
claims in these documents against the actual source they describe (`scripts/install.ps1`,
`src/FinanceFmtTools.Engine/FormatRegistry.cs`, `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs`,
`src/customUI14.xml`, the `.csproj` files, and a stale local `bin/Release/net48` build) to verify
factual accuracy, not just internal consistency.

Positives confirmed by cross-reference: the `FinanceFmtTools.zip` fixed-name convention is
consistent across `release.yml`, `RELEASE.md`, and `scripts/install.ps1`'s `$AssetName`; the
`Compress-Archive`/`Copy-Item` file list in `release.yml` and in `RELEASE.md` step 4 are
byte-for-byte identical and match the 4 files `install.ps1` actually expects to find
(`$AllFiles`); the workflow's `permissions: contents: write` is correctly minimal-scoped; the
accounting number-format strings documented in the README's four combination tables were checked
against `AccountingFormatBuilder.Build` and are accurate; the "16 combinations" test-coverage claim
in `RELEASE_NOTES.md` matches the actual `[InlineData]` count in `AccountingFormatBuilderTests.cs`.

The most significant defect is a missing fail-fast in the release workflow's packaging step: a
failed `Copy-Item` (e.g., a renamed/missing build output) would not fail the CI job, so a broken,
incomplete release zip could be silently published with no CI signal. Several documentation-accuracy
gaps were also found between the README/RELEASE_NOTES.md prose and the actual Ribbon XML/format
registry they describe.

## Critical Issues

### CR-01: Release packaging step has no fail-fast — a missing/renamed build artifact silently ships an incomplete release

**File:** `.github/workflows/release.yml:35-52`
**Issue:** The `Package` step is a multi-line `pwsh` script block that never sets
`$ErrorActionPreference = 'Stop'` (or checks `$LASTEXITCODE`/`$?` after each `Copy-Item`).
PowerShell's default `$ErrorActionPreference` is `Continue`, and cmdlets like `Copy-Item` write a
**non-terminating** error to the error stream when the source path doesn't exist — they do not throw,
and they do not set a non-zero exit code. Concretely: if the build output path changes (e.g. a
TargetFramework rename, an assembly rename, or `MicrosoftOfficeCore16`/`Microsoft.Office.Interop.Excel`
package upgrade that ships a differently-cased/renamed DLL), one or more of the four
`Copy-Item "$binSrc\...\" staging\` lines (lines 41-44) would fail silently, the script would continue
to `Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force` (line 52) with
whatever subset of files actually got copied, the step would report success, and
`softprops/action-gh-release@v2` (line 55-60) would publish and upload an **incomplete**
`FinanceFmtTools.zip` to a public GitHub Release with zero indication anything went wrong. Every
user who runs the documented `scripts/install.ps1` one-liner against that release would get a COM
add-in that fails to load in Excel (missing dependency DLL) with no way to know the release itself was
broken.
**Fix:**
```yaml
      - name: Package
        shell: pwsh
        run: |
          $ErrorActionPreference = 'Stop'
          $binSrc = "src\FinanceFmtTools.ComAddin\bin\Release\net48"
          New-Item -ItemType Directory -Path staging -Force | Out-Null

          $required = @(
            "$binSrc\FinanceFmtTools.ComAddin.dll",
            "$binSrc\FinanceFmtTools.Engine.dll",
            "$binSrc\Microsoft.Office.Interop.Excel.dll",
            "$binSrc\office.dll",
            "scripts\install.ps1",
            "scripts\uninstall.ps1",
            "scripts\verify-environment.ps1"
          )
          foreach ($f in $required) {
            if (-not (Test-Path -LiteralPath $f)) { throw "Missing required packaging input: $f" }
            Copy-Item -LiteralPath $f -Destination staging\ -Force
          }

          Compress-Archive -Path staging\* -DestinationPath FinanceFmtTools.zip -Force
```
Setting `$ErrorActionPreference = 'Stop'` at the top of the script block (or wrapping the whole
block in `try { ... } catch { Write-Error $_; exit 1 }`) is the minimum fix — it converts the
existing non-terminating `Copy-Item` errors into job failures instead of silent omissions.

## Warnings

### WR-01: Manual release runbook script has the same missing-fail-fast gap as CR-01

**File:** `RELEASE.md:58-77`
**Issue:** Step 4 ("Compilar e empacotar localmente") reproduces the exact same `Copy-Item` /
`Compress-Archive` sequence as `release.yml`, for a human (or an AI agent, per the file's own stated
audience) to run interactively. It has the same default-`Continue` behavior: a failed `Copy-Item`
prints red text but the script keeps going and still produces a `FinanceFmtTools.zip`, which the
runbook's own step 5 then instructs the operator to tag and push. Because this is run interactively a
human is somewhat more likely to notice the red error text than in unattended CI, but the runbook
gives no explicit instruction to verify the resulting zip's contents before publishing it (manual or
automatic path), and an AI agent running this non-interactively (this file explicitly says it's
"destinado a uma pessoa ou a um agente de IA") would not "notice" a printed error unless it greps
output for it.
**Fix:** Add `$ErrorActionPreference = 'Stop'` as the first line of the step-4 code block (mirroring
the CR-01 fix), and add a post-step verification line to the checklist, e.g.:
```powershell
$expected = @('FinanceFmtTools.ComAddin.dll','FinanceFmtTools.Engine.dll','Microsoft.Office.Interop.Excel.dll','office.dll','install.ps1','uninstall.ps1','verify-environment.ps1')
(Get-ChildItem staging -Name) | Sort-Object | Should be $expected # or an explicit foreach/Test-Path check
```

### WR-02: GitHub Actions pinned to mutable major-version tags, not immutable commit SHAs

**File:** `.github/workflows/release.yml:19,22,55`
**Issue:** `actions/checkout@v4`, `actions/setup-dotnet@v4`, and `softprops/action-gh-release@v2`
are all referenced by a mutable major-version tag. Tags can be moved (by the action maintainer, or
by an attacker who compromises the maintainer's account/repo), silently changing what code runs in
a workflow that holds `contents: write` and executes on every `v*.*.*` tag push — i.e. exactly the
kind of workflow supply-chain attacks target (this is the same class of risk CVE advisories about
compromised GitHub Actions describe). `softprops/action-gh-release` in particular is a third-party,
non-Microsoft/non-GitHub action with broad `contents: write` usage in this workflow.
**Fix:** Pin to a full-length commit SHA with the version as a trailing comment, e.g.:
```yaml
      - uses: actions/checkout@11bd71901bbe5b1630ceea73d27597364c9af683 # v4.2.2
      - uses: actions/setup-dotnet@67a3573c9a986a3f9c594539f4ab511d57bb3ce9 # v4.3.1
      - uses: softprops/action-gh-release@da05d552573ad5aba039eaac05058a918a7bf631 # v2.2.1
```
(Use Dependabot/Renovate to keep SHAs current, which also gives an auditable diff on every bump.)

### WR-03: README ribbon-reference tree names a button that does not exist ("Guia Fin" vs. actual label "Documentação")

**File:** `README.md:240` (cross-referenced against `src/customUI14.xml:135-141`)
**Issue:** The "Referência rápida do ribbon" tree lists the Info group as:
```
└── Info
    ├── Guia Fin        → abre esta documentação
    └── Sobre           → versão do add-in
```
but `src/customUI14.xml`'s `btnFinInfo` button (`onAction="RibbonFinInfo"`) declares
`label="Documentação"`, not "Guia Fin". A user reading the README and then looking at the actual
Excel ribbon for a button labeled "Guia Fin" will not find one — this is the button surface actually
labeled "Documentação".
**Fix:**
```
└── Info
    ├── Documentação    → abre esta documentação
    └── Sobre           → versão do add-in
```

### WR-04: README's date format-string table omits the actual `;@` text-passthrough section

**File:** `README.md:160-166` (cross-referenced against `src/FinanceFmtTools.Engine/FormatRegistry.cs:45-55`)
**Issue:** The "Datas" table documents the format strings as `yyyy-mm-dd`, `[$-pt-BR]dd/mm/yyyy`, and
`[$-pt-BR]dd/mmm/yyyy`. The actual literals registered in `FormatRegistry.cs` are two-section formats:
`"yyyy-mm-dd;@"`, `"[$-pt-BR]dd/mm/yyyy;@"`, and `"[$-pt-BR]dd/mmm/yyyy;@"` — each with a trailing
`;@` section that tells Excel to render text values verbatim instead of coercing/erroring on them.
This is a real, user-visible behavioral difference (e.g. pasting a text label into a date-formatted
cell) that the documented format strings don't reflect, and it means a user who copies the README's
string verbatim into a custom Excel number format will get slightly different behavior than the
add-in's own buttons produce.
**Fix:**
```
| ISO | `yyyy-mm-dd;@` | `2025-03-15` |
| BR | `[$-pt-BR]dd/mm/yyyy;@` | `15/03/2025` |
| BR Extenso | `[$-pt-BR]dd/mmm/yyyy;@` | `15/mar/2025` |
```
(and add a short note explaining the `;@` text-passthrough section, mirroring the "Decomposição"
explanation already given for the Zero Dash token in the Fin xD section).

### WR-05: README "Licença" section is still an unresolved placeholder despite a complete LICENSE file already existing

**File:** `README.md:306-308`
**Issue:**
```markdown
## Licença

<!-- Adicionar licença aqui -->
```
An MIT `LICENSE` file (21 lines, copyright 2026 Thomaz Pougy) has existed in the repository since
the very first commit (`f6ec3e4 Initial commit`) — it predates this phase entirely. The Phase 5
README rewrite (commit `990969e`, which is one of the four files under review) had the opportunity
to fix this and left the stale placeholder comment in place instead.
**Fix:**
```markdown
## Licença

MIT — ver [`LICENSE`](./LICENSE).
```

### WR-06: RELEASE_NOTES.md's changelog bullet misnames/duplicates format buttons, inconsistent with the actual Ribbon and FormatRegistry

**File:** `RELEASE_NOTES.md:22-23` (cross-referenced against `src/customUI14.xml` and `src/FinanceFmtTools.Engine/FormatRegistry.cs:33-58`)
**Issue:** This is the changelog body actually published to the public GitHub Release (via
`body_path: RELEASE_NOTES.md` in the workflow and `-F RELEASE_NOTES.md` in the manual runbook), so
its accuracy matters for end users, not just internal readers. The bullet reads:
> "...contábil `Fin 0D/2D/4D/8D`, percentual `Pct 0,00%`/`Pct 0,0000%`, `Spread (bps)`, datas
> `Date ISO`/`Date BR`/`Date BR Longa`, `Integer` e `Text`) foram portados 1:1..."

Problems:
1. `Pct 0,00%`/`Pct 0,0000%` matches neither the actual Ribbon button labels (`% 2D`/`% 4D`) nor the
   `FormatRegistry` DisplayNames (`"% 2 casas"`/`"% 4 casas"`).
2. `Date ISO`/`Date BR`/`Date BR Longa` switches to English "Date" for no apparent reason — every
   other label in this same sentence, the actual Ribbon buttons (`ISO`/`BR`/`BR Extenso`), and the
   `FormatRegistry` DisplayNames (`"Data ISO"`/`"Data BR"`/`"Data BR Longa"`) all use Portuguese
   "Data". "BR Longa" also doesn't match either the button label ("BR Extenso") or the DisplayName
   ("Data BR Longa") consistently.
3. `Integer` and `Text` are listed as if they were two additional, distinct buttons on top of
   `Fin 0D/2D/4D/8D` and the rest — but `Integer` is the internal `FormatKeys` constant for the
   already-listed `Fin 0D` button, and `Text` is the internal `FormatKeys` constant for the Ribbon's
   `Texto` button. As written, the sentence enumerates what reads like 12 distinct format buttons
   when the product only has 11 (`Fin 8D/4D/2D/0D`, `% 4D/2D`, `Spread bps`, `ISO`/`BR`/`BR Extenso`,
   `Texto`).
**Fix:** Use the actual Ribbon button labels consistently:
```
todos os botões de formatação (contábil Fin 0D/2D/4D/8D, percentual % 2D/% 4D, Spread bps,
datas ISO/BR/BR Extenso e Texto) foram portados 1:1 a partir do VBA original...
```

## Info

### IN-01: `.NET SDK` version floated to `8.x` rather than pinned in the release workflow

**File:** `.github/workflows/release.yml:24`
**Issue:** `dotnet-version: '8.x'` lets `actions/setup-dotnet` install whatever the newest published
8.x SDK patch is at the time the workflow runs. For a release pipeline this is a minor reproducibility
concern — the exact toolchain used to produce a given release's binaries can silently drift between
tag pushes months apart, without any change to the workflow file itself.
**Fix:** Pin to a specific SDK version tested against this repo, e.g. `dotnet-version: '8.0.404'`,
and bump it deliberately (with a commit) when a newer SDK is adopted.

### IN-02: One-liner installer executes remote script directly via `iex` with no integrity check

**File:** `README.md:179` (also duplicated in `RELEASE_NOTES.md:58` and `RELEASE.md`'s "Testado o
instalador" checklist item)
**Issue:** `Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/.../scripts/install.ps1 | iex`
downloads and executes arbitrary PowerShell over HTTPS with no checksum, signature, or pinned commit
verification — a compromised GitHub account/repo (or a MITM if the user is on a hostile network
despite TLS) could serve modified code that the user runs with their own privileges. This mirrors a
documented, deliberate convention from the sibling `outlook-classic-delay-send` project (per
CLAUDE.md), so it is not a novel choice introduced by this phase, but it also is not called out
anywhere in these docs as a known/accepted tradeoff for the reader.
**Fix:** Not necessarily requiring a change (matches project convention), but consider adding a
one-line note near the install command, e.g. "Este comando baixa e executa `install.ps1` diretamente
do branch `main`; revise o script antes de rodar se preferir uma verificação manual."

### IN-03: README documents 64-bit as a hard requirement; installer only warns

**File:** `README.md:187` (cross-referenced against `scripts/install.ps1:237-253`)
**Issue:** The "Requisitos" section lists "Windows + Excel 2016 ou superior, **64-bit**" as a
requirement (bold, unqualified). `scripts/install.ps1`'s own bitness checks (`Test-PeMachine`,
Click-to-Run `Platform` check) are explicitly non-blocking — they only `Write-Warn2` and continue
installing regardless of detected architecture (per the script's own comments, this is deliberate:
32-bit is merely untested, not technically incompatible, per "FUT-01"). The README's phrasing reads
as a hard gate the installer doesn't actually enforce.
**Fix:** Soften the wording slightly, e.g. "Windows + Excel 2016 ou superior (baseline validado:
64-bit — 32-bit não é bloqueado pelo instalador, mas não é testado)."

### IN-04: RELEASE_NOTES.md's "permanent once published" claim about the changelog is not strictly accurate

**File:** `RELEASE_NOTES.md:3-7`
**Issue:** "Depois de uma release ser criada no GitHub, o texto publicado é permanente" overstates
GitHub's actual behavior — a published release's notes/body *can* be edited after the fact (via
`gh release edit --notes-file ...` or the web UI). The intended point (this file itself gets
overwritten in-repo before each new tag, so old entries aren't preserved as running changelog
history in git) is reasonable, but "permanente" is the wrong word for what GitHub actually
guarantees.
**Fix:** Reword to something like: "este arquivo é sobrescrito com a entrada da próxima versão antes
de cada nova tag — o histórico de versões anteriores só fica preservado na página de Releases do
GitHub (editável via `gh release edit`, não neste arquivo)."

---

_Reviewed: 2026-07-11T18:53:28Z_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
