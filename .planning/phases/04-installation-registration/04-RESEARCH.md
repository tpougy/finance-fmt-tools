# Phase 4: Installation & Registration - Research

**Researched:** 2026-07-11
**Domain:** Per-user (HKCU) COM Shared Add-in registration for a classic (non-VSTO) .NET Framework 4.8 Excel add-in, via self-contained PowerShell install/uninstall scripts
**Confidence:** HIGH

## Summary

This phase does not need to invent a registration mechanism — it needs to adapt two mechanisms that already exist and are already proven in this exact codebase family. First, the legacy `Install-FinanceFmtTools.ps1` (still in this repo's root) already proves the GitHub-Releases-download + TLS1.2 + Excel-process-detection flow works in production (two real releases, v1.0.0/v1.0.1, exist today) — but it registers a `.xlam` via `Excel.AddIns.Add()`, which is architecturally irrelevant to a COM Shared Add-in (no `.xlam`, no `Excel.AddIns` collection involved at all). Second, and far more directly relevant, the sibling project `outlook-classic-delay-send` (`/home/thomaz/pessoal/outlook-classic-delay-send`) already solved the *exact* problem this phase faces — a plain `[ComVisible]`/`[Guid]`/`[ProgId]` .NET Framework 4.8 class registered as a per-user Office COM Shared Add-in, no admin, no VSTO, no `regasm` — for Outlook instead of Excel. Its `scripts/install.ps1`/`scripts/uninstall.ps1` are a near-complete template: swap `Outlook` for `Excel` in three registry paths, swap the fixed GUID/ProgId/AssemblyName values for this project's own (already declared in `src/FinanceFmtTools.ComAddin/Connect.cs`'s header comment and confirmed by direct inspection of this repo's own compiled `FinanceFmtTools.ComAddin.dll`), and swap the 4-file binary set for this project's actual 4 files (verified by an actual local build already present in `src/FinanceFmtTools.ComAddin/bin/Release/net48/`).

Three official-Microsoft-Learn-confirmed facts anchor the registry design: (1) the non-versioned `Root\Software\Microsoft\Office\<application>\Addins\<ProgId>` discovery key is real and documented (not a Outlook-only convention — it is the generic "all Office applications" pattern, confirmed via the *Registry entries for VSTO Add-ins* Microsoft Learn page, which literally has an "All Other" row covering every Office app by that exact path shape); (2) `HKEY_CURRENT_USER\Software` is **not** subject to WOW6432Node redirection, so a single HKCU registration is visible to both 32-bit and 64-bit Excel without any bitness-specific duplication (confirmed in the same Microsoft Learn page: "If the Installer is targeting the current user, it doesn't need to install to the WOW6432Node because the HKEY_CURRENT_USER\Software path is shared"); (3) `regasm.exe` is not required — it normally needs admin and only writes HKLM/HKCR, so this phase (like the sibling project) writes the five/six registry values directly via `New-Item`/`Set-ItemProperty`, which is fully idempotent and admin-free by construction.

One area needs an explicit, honest caveat rather than a confident claim: Microsoft's canonical "DoNotDisableAddinList" documentation page lives under the **Outlook** VBA docs namespace, and a Microsoft Q&A thread contains one answer claiming the key does *not* work for Word/Excel, contradicted by another citing KB 2758876's own applies-to list (which explicitly includes Excel 2016/2013). The registry path shape generalizes trivially (`HKCU\...\Office\<ver>\<app>\Resiliency\DoNotDisableAddinList`), and this is corroborated by community documentation (blog posts describing the same pattern for Excel/Word), but this phase's own live-Excel verification step (like Phase 3's) should explicitly confirm the key is honored by Excel, not just written successfully — write success and behavioral effectiveness are two different things.

**Primary recommendation:** Write one self-contained `install.ps1` (supporting both the documented `irm .../install.ps1 | iex` one-liner flow AND a `-Package <local-zip>` / `-Source <local-dir>` convenience flow for local testing before a real GitHub Release exists) plus one `uninstall.ps1`, both modeled directly on `outlook-classic-delay-send/scripts/{install,uninstall}.ps1`, targeting Excel's non-versioned Addins key and the versioned Resiliency key, using the exact fixed identity values already declared in `Connect.cs`.

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| COM class registration (CLSID -> InprocServer32 -> mscoree.dll) | OS Registry (HKCU\Software\Classes) | .NET Framework CLR (activated by Excel via COM) | Registry entries are pure data; the CLR is what actually loads `FinanceFmtTools.ComAddin.dll` when Excel calls `CoCreateInstance` |
| Excel add-in discovery (LoadBehavior=3) | OS Registry (HKCU\...\Office\Excel\Addins) | Excel.exe (reads this key at startup) | Non-versioned, all-Excel-versions key per Microsoft Learn; Excel enumerates this key before instantiating any CLSID |
| Resiliency / auto-disable protection (INST-03) | OS Registry (HKCU\...\Office\<ver>\Excel\Resiliency\DoNotDisableAddinList) | Excel.exe's Resiliency subsystem | Versioned (per Office release), unlike the Addins discovery key — different key shape, easy to get wrong |
| Binary deployment (DLL + interop assemblies) | Local Filesystem (%LocalAppData%\FinanceFmtTools\) | — | Xcopy-deploy, no GAC, no shared location — per-user install mirrors HKCU-only registration |
| Release artifact acquisition | GitHub Releases (external service) | PowerShell installer script (client) | Phase 5 (not this phase) builds the CI that produces the release zip; this phase only consumes and documents the expected zip shape |
| Bitness detection/guard (64-bit-only scope) | Installer script (client-side logic) | OS Registry (Office ClickToRun\Configuration) + EXCEL.EXE PE header | Read-only checks against the target machine; no separate 32-bit code path needs registering since HKCU is bitness-transparent |

## User Constraints (from CONTEXT.md)

### Locked Decisions

**Carried-over fixed identity from Phase 3 (not discretionary — must be reused verbatim), confirmed by direct inspection of both source and the already-compiled binary in this repo:**

| Value | Content | Verified via |
|---|---|---|
| GUID (CLSID) | `881EFDF3-424C-4240-BCA0-714DAC2B9CD7` | `Connect.cs` `[Guid(...)]` attribute AND `strings` on the compiled `FinanceFmtTools.ComAddin.dll` in this repo (`$881EFDF3-424C-4240-BCA0-714DAC2B9CD7` found embedded) |
| ProgId | `FinanceFmtTools.Connect` | `Connect.cs` `[ProgId(...)]` attribute AND found embedded in the compiled DLL |
| Full class name | `FinanceFmtTools.ComAddin.Connect` | namespace `FinanceFmtTools.ComAddin` + class `Connect` in `Connect.cs` |
| AssemblyName | `FinanceFmtTools.ComAddin` | `FinanceFmtTools.ComAddin.csproj` `<AssemblyName>` |
| Version | `1.0.0.0` | `FinanceFmtTools.ComAddin.csproj` `<Version>` |
| Excel discovery key | `HKCU\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect` | `Connect.cs` header comment (non-versioned form, confirmed generically correct by Microsoft Learn — see Sources) |
| FriendlyName (recommend) | `Finance Fmt Tools` | matches `AddInHost.cs`'s `AddinName` constant used in the About dialog |

### Environment constraint (carries forward from Phase 3)

This dev environment (Linux/WSL) has no Windows, no Excel, no PowerShell (confirmed: `pwsh`/`powershell` both absent from PATH in this session), no registry. The install/uninstall PowerShell scripts can be **written and reviewed for correctness** here but **cannot be executed** here. Expect this phase's verification to again come back `human_needed` for the actual install/uninstall run and the "Ribbon tab appears" check, exactly as Phase 3 did. There is also a second-order environment gap unique to this phase: **no real C# release exists yet** on GitHub (only the legacy VBA `v1.0.0`/`v1.0.1` `.xlam` releases exist — confirmed via `gh release list`), so the documented one-liner's "download the latest GitHub release" step cannot be end-to-end verified until Phase 5 ships a real release asset. Design the installer to be testable against a **locally-built** zip in the meantime (see Pattern 2 below).

### Claude's Discretion

All implementation choices are at Claude's discretion (discuss phase skipped, full-autonomy run). Areas where this research surfaces a real choice rather than a single correct answer:
- Exact install directory name/shape under `%LocalAppData%` (recommend mirroring the sibling project: `%LocalAppData%\FinanceFmtTools\` with an optional `logs\` subfolder)
- Whether to hard-block or soft-warn on detecting 32-bit Excel (recommend soft-warn, matching the sibling project's `Write-Warn2`-and-continue pattern, since the compiled add-in is AnyCPU/PIA-based and would very likely still load on 32-bit Excel — FUT-01 defers *support/testing*, not necessarily technical compatibility)
- Whether to also ship a `verify-environment.ps1`-style read-only diagnostic script (sibling project has one; not required by INST-01/02/03 but low-cost and matches the explicitly-named inspiration project's convention)
- Exact file names/locations of the two scripts in this repo (recommend `scripts/install.ps1` + `scripts/uninstall.ps1`, new files — do NOT overwrite or repurpose the existing root `Install-FinanceFmtTools.ps1`, which is legacy-VBA-specific and explicitly Phase 5's concern to retire per LEGACY-01/LEGACY-02)

### Deferred Ideas (OUT OF SCOPE)

None — discuss phase skipped, no deferred ideas recorded in CONTEXT.md. Note from REQUIREMENTS.md's own Out of Scope table: **32-bit Excel support (FUT-01)** is explicitly deferred to a future release; this phase targets 64-bit Excel only, per explicit user decision, not a technical limitation of the registration mechanism itself.

## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| INST-01 | Instalador PowerShell one-liner (`irm ... \| iex`) baixa a última release do GitHub e registra o add-in via HKCU para Excel 64-bit, sem exigir admin | Pattern 1 (registry keys) + Pattern 2 (self-contained download flow, reusing legacy `Install-FinanceFmtTools.ps1`'s GitHub Releases API + TLS1.2 pattern) + Code Examples below |
| INST-02 | Script de desinstalação remove o registro HKCU e os arquivos instalados | Pattern 1 (mirrors `outlook-classic-delay-send/scripts/uninstall.ps1` verbatim structure, swapped to Excel keys) |
| INST-03 | O instalador grava a chave `DoNotDisableAddinList` para evitar que o Excel desative o add-in silenciosamente após um erro transiente (Resiliency) | Common Pitfall "Resiliency key generalization is corroborated, not 100% Microsoft-guaranteed for non-Outlook apps" + Code Examples below |

## Standard Stack

### Core

| Tool | Version | Purpose | Why Standard |
|------|---------|---------|---------------|
| Windows PowerShell | 5.1+ (script declares `#Requires -Version 5.1`) | Install/uninstall script host | Ships with every supported Windows target (Win10/11); matches both the legacy VBA installer and the sibling project's `install.ps1`/`uninstall.ps1`, which are both PS 5.1-compatible |
| `New-Item` / `Set-ItemProperty` / `Remove-Item` / `Remove-ItemProperty` (built-in PowerShell registry provider) | built-in | HKCU registry read/write | Directly idempotent (`-Force` overwrites, `Test-Path` guards deletes); no admin required for HKCU; no external module needed |
| `Invoke-RestMethod` / `[System.Net.HttpWebRequest]` | built-in | GitHub Releases API query + asset download | Already proven working in this exact repo's `Install-FinanceFmtTools.ps1` (two real releases downloaded successfully by real users) |
| `Expand-Archive` (Microsoft.PowerShell.Archive, built into PS 5.1+) | built-in | Extract the release `.zip` asset | Standard, no third-party zip library needed; used verbatim by the sibling project's `install.ps1` |

### Supporting

| Tool | Version | Purpose | When to Use |
|------|---------|---------|-------------|
| `RegAsm.exe` | ships with .NET Framework 4.8 (`%WINDIR%\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe`) | Optional dev-time bootstrap / `.reg` file generation only | NOT used in the production install script (needs admin for its default HKLM/HKCR write mode); only useful as a one-time dev-machine sanity check (`regasm Asm.dll /codebase /regfile:out.reg` — generates a `.reg` file without touching the registry, useful to cross-check the exact `Assembly`/`RuntimeVersion` string format if ever in doubt) |
| `gh` CLI | any recent | Verify current release state during planning/testing | Already available in this environment; confirmed `v1.0.0`/`v1.0.1` are the only releases that exist today (both legacy `.xlam`) |

### Alternatives Considered

| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Direct HKCU registry writes (`New-Item`/`Set-ItemProperty`) | `regasm.exe /codebase` at install time | regasm needs admin for its default write targets and does not support writing directly to HKCU; would require an install-time `/regfile` + text-substitution dance for no benefit over writing the keys directly |
| Self-contained `install.ps1` (downloads zip itself) | Two-step flow like the sibling project (`build.ps1` produces a zip; user manually downloads and runs `install.ps1 -Package <zip>`) | INST-01 explicitly requires the one-liner `irm ... \| iex` UX (matching the legacy VBA installer's exact UX); the sibling project's two-step flow does not by itself satisfy INST-01 and must be adapted, not copied as-is |
| `.zip` release asset (bin\ folder + scripts) | A single self-extracting `.exe` or an MSI | CLAUDE.md explicitly forbids VSTO/ClickOnce/MSI-based installers; a plain `.zip` + PowerShell matches the project's stated "no admin, no Visual Studio" installer philosophy and both the legacy and sibling-project precedents |

**Installation:** Not applicable — this phase produces PowerShell scripts, not a package to install via a package manager.

**Version verification:** N/A — no NuGet/npm/pip packages are introduced by this phase. All dependencies are Windows-builtin (PowerShell, registry provider, `Expand-Archive`) or come from Phase 3's already-approved dependencies (see Package Legitimacy Audit below).

## Package Legitimacy Audit

**Not applicable to this phase.** Phase 4 introduces zero new NuGet/npm/pip packages — it only writes PowerShell scripts and Windows Registry key definitions. The only external binaries this phase's installer *deploys* (it does not add new dependencies at build time) are the two NuGet-sourced interop assemblies already vetted in `.planning/phases/03-com-entry-point-real-excel-integration/03-RESEARCH.md` and approved via the autonomous decision recorded in `.planning/STATE.md` at commit `80f0046`:
- `Microsoft.Office.Interop.Excel` 16.0.18925.20022 (unofficial CamronBute repackage, content-verified genuine — already `[SUS]`-flagged-then-approved in Phase 3, not re-litigated here)
- `MicrosoftOfficeCore16` 16.0.16626.20000 (same publisher family, same prior approval)

No new slopcheck/registry-verification run was performed for this phase since no new package names are introduced.

## Architecture Patterns

### System Architecture Diagram

```
                     ┌─────────────────────────────┐
  User runs:         │  GitHub raw.githubusercontent│
  irm .../install.ps1│  .com/.../scripts/install.ps1│
  | iex              └──────────────┬──────────────┘
        │                            │ (1) script text downloaded, executed in caller's session
        ▼                            ▼
┌───────────────────────────────────────────────────────────┐
│ install.ps1 (runs entirely in user's PowerShell session)   │
│                                                             │
│  (2) Assert-ExcelNotRunning ──► if running: warn/close      │
│                                                             │
│  (3) Get-LatestReleaseTag ──► GitHub Releases API           │
│         │                     (api.github.com/.../latest)  │
│         ▼                                                  │
│  (4) Download release .zip asset (TLS1.2 forced)            │
│         │                                                  │
│         ▼                                                  │
│  (5) Expand-Archive to %TEMP%\financefmt-install-<guid>\    │
│         │                                                  │
│         ▼                                                  │
│  (6) Find bin\ (FinanceFmtTools.ComAddin.dll + 3 files)      │
│         │                                                  │
│         ▼                                                  │
│  (7) Copy 4 files → %LocalAppData%\FinanceFmtTools\          │
│         │                                                  │
│         ▼                                                  │
│  (8) Write HKCU registry keys:                               │
│        (a) Software\Classes\{CLSID}\InprocServer32           │
│        (b) Software\Microsoft\Office\Excel\Addins\<ProgId>    │
│        (c) Software\Microsoft\Office\16.0\Excel\Resiliency\   │
│            DoNotDisableAddinList\<ProgId> = 1                 │
│         │                                                  │
│         ▼                                                  │
│  (9) Post-install validation (files exist, LoadBehavior==3,  │
│      CodeBase matches installed DLL path)                    │
│         │                                                  │
│         ▼                                                  │
│  (10) Remove %TEMP% extraction folder                        │
└───────────────────────────────────────────────────────────┘
        │
        ▼
┌───────────────────────────────────────────────────────────┐
│  Next Excel.exe launch:                                    │
│   Excel enumerates HKCU\...\Office\Excel\Addins\*            │
│   → finds FinanceFmtTools.Connect, LoadBehavior=3            │
│   → CoCreateInstance({CLSID}) → mscoree.dll → CLR loads       │
│     FinanceFmtTools.ComAddin.dll → Connect.OnConnection       │
│   → GetCustomUI() → "Finance Fmt" Ribbon tab renders          │
└───────────────────────────────────────────────────────────┘

uninstall.ps1 runs the mirror image: Test-Path-guarded Remove-Item
on each of the 3 registry trees (idempotent — safe if never installed),
then removes the 4 files from %LocalAppData%\FinanceFmtTools\.
```

### Recommended Project Structure

```
scripts/
├── install.ps1           # Self-contained: GitHub Releases download + HKCU registration
│                          #   Supports -Package <zip> / -Source <dir> for local testing
│                          #   before a real GitHub Release exists (see Pitfall below)
├── uninstall.ps1          # Mirror of install.ps1 — Test-Path-guarded key/file removal
└── verify-environment.ps1 # OPTIONAL (Claude's discretion) — read-only .NET FW 4.8 /
                           #   Excel bitness diagnostic, mirroring the sibling project's
                           #   script of the same name
```

Do NOT modify the existing root `Install-FinanceFmtTools.ps1` / `Install-FinanceFmtTools.bat` — those belong to the legacy VBA flow and are explicitly Phase 5's responsibility to retire (LEGACY-01/LEGACY-02), not this phase's.

### Pattern 1: HKCU-only, no-admin COM Shared Add-in registration (three registry trees)

**What:** Three independent registry trees must all be written for Excel to discover, load, and keep loading the add-in. Missing any one breaks a different symptom (missing = doesn't load at all; wrong CodeBase = loads but throws; missing Resiliency = loads fine but a later transient crash silently disables it).

**When to use:** Every install; every uninstall must remove exactly these three trees (and nothing beyond the add-in's own values within any tree shared with other add-ins).

**Example (adapted directly from `outlook-classic-delay-send/scripts/install.ps1`, swapped to Excel + this project's fixed identity):**
```powershell
# Source pattern: outlook-classic-delay-send/scripts/install.ps1 lines 349-391
# (verbatim structural adaptation — Outlook -> Excel, GUID/ProgId/Assembly swapped)

$Guid         = '{881EFDF3-424C-4240-BCA0-714DAC2B9CD7}'   # from Connect.cs, braces added for registry form
$ProgId       = 'FinanceFmtTools.Connect'
$ClassName    = 'FinanceFmtTools.ComAddin.Connect'          # full namespace.class
$AssemblyStr  = 'FinanceFmtTools.ComAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'
$RuntimeVer   = 'v4.0.30319'
$Shim         = 'C:\Windows\System32\mscoree.dll'            # bitness-transparent — see Pitfall below
$ThreadingMdl = 'Both'
$FriendlyName = 'Finance Fmt Tools'
$Description  = 'Formatação financeira padronizada para mercado de capitais.'
$OfficeVerKey = '16.0'   # Resiliency key IS versioned, unlike the Addins discovery key

$InstallDir = Join-Path $env:LOCALAPPDATA 'FinanceFmtTools'
$dllPath    = Join-Path $InstallDir 'FinanceFmtTools.ComAddin.dll'
$codeBase   = 'file:///' + ($dllPath -replace '\\', '/')

# (a) COM class registration — HKCU\Software\Classes (NOT redirected by WOW6432Node)
$kProg      = "HKCU:\Software\Classes\$ProgId"
$kProgClsid = "HKCU:\Software\Classes\$ProgId\CLSID"
$kClsid     = "HKCU:\Software\Classes\CLSID\$Guid"
$kClsidProg = "HKCU:\Software\Classes\CLSID\$Guid\ProgId"
$kInproc    = "HKCU:\Software\Classes\CLSID\$Guid\InprocServer32"

foreach ($p in @($kProg,$kProgClsid,$kClsid,$kClsidProg,$kInproc)) { New-Item -Path $p -Force | Out-Null }

Set-ItemProperty -Path $kProg      -Name '(default)' -Value $ProgId
Set-ItemProperty -Path $kProgClsid -Name '(default)' -Value $Guid
Set-ItemProperty -Path $kClsid     -Name '(default)' -Value $ClassName
Set-ItemProperty -Path $kClsidProg -Name '(default)' -Value $ProgId
Set-ItemProperty -Path $kInproc -Name '(default)'      -Value $Shim
Set-ItemProperty -Path $kInproc -Name 'ThreadingModel' -Value $ThreadingMdl
Set-ItemProperty -Path $kInproc -Name 'Class'          -Value $ClassName
Set-ItemProperty -Path $kInproc -Name 'Assembly'       -Value $AssemblyStr
Set-ItemProperty -Path $kInproc -Name 'RuntimeVersion' -Value $RuntimeVer
Set-ItemProperty -Path $kInproc -Name 'CodeBase'       -Value $codeBase

# (b) Excel discovery key — NON-VERSIONED ("all versions of Excel"), NOT "16.0\Excel\Addins"
$kAddin = "HKCU:\Software\Microsoft\Office\Excel\Addins\$ProgId"
New-Item -Path $kAddin -Force | Out-Null
Set-ItemProperty -Path $kAddin -Name 'FriendlyName' -Value $FriendlyName
Set-ItemProperty -Path $kAddin -Name 'Description'  -Value $Description
Set-ItemProperty -Path $kAddin -Name 'LoadBehavior' -Value 3 -Type DWord

# (c) Resiliency anti-soft-disable (INST-03) — VERSIONED, unlike (b)
$kResil = "HKCU:\Software\Microsoft\Office\$OfficeVerKey\Excel\Resiliency\DoNotDisableAddinList"
New-Item -Path $kResil -Force | Out-Null
Set-ItemProperty -Path $kResil -Name $ProgId -Value 1 -Type DWord
```

### Pattern 2: Self-contained one-liner installer with a local-testing escape hatch

**What:** INST-01 requires `irm .../install.ps1 | iex` to work end-to-end against a real GitHub Release. But no C# release exists yet (only legacy `.xlam` `v1.0.0`/`v1.0.1` — confirmed via `gh release list` in this session), and Phase 5 (not this phase) builds the CI that will produce one. Design `install.ps1` so it can be tested *now*, before a real release exists, without contradicting INST-01's documented default flow.

**When to use:** Always for this phase — it resolves a real chicken-and-egg dependency between Phase 4 and Phase 5 without either phase blocking on the other.

**Example (parameter shape adapted from `outlook-classic-delay-send/scripts/install.ps1`'s `-Package`/`-Source` convenience params):**
```powershell
[CmdletBinding()]
param(
    [string]$Package,   # optional: local .zip path, bypasses GitHub download (for testing)
    [string]$Source,    # optional: already-extracted folder containing bin\
    [switch]$Force
)
# Default (no params): the documented one-liner path — query GitHub Releases API,
# download the latest .zip asset, extract to %TEMP%, proceed exactly as Pattern 1.
# -Package/-Source: skip the GitHub download entirely, useful for testing this
# installer against a locally `dotnet build`-produced bin\ before Phase 5 ships CI.
```

**Important gotcha:** when a script is executed via `irm ... | iex`, there is no on-disk script file — `$PSScriptRoot` and `$MyInvocation.MyCommand.Path` are both empty/null in that execution context. Do NOT design any part of the default (no-params) flow to depend on the script's own file location (the legacy `Install-FinanceFmtTools.ps1` already gets this right by using only `$env:TEMP`/`$env:APPDATA`, never `$PSScriptRoot`). The `-Source`/`-Package` local-testing params are fine to use `$MyInvocation` internally since they're only invoked via direct local file execution (`.\install.ps1 -Package ...`), never via the piped one-liner.

### Pattern 3: Idempotent registry writes / idempotent removal

**What:** `New-Item -Force` + `Set-ItemProperty` never errors on "already exists" (unlike `New-ItemProperty` without `-Force`, which does). `Remove-Item`/`Remove-ItemProperty` guarded by `Test-Path`/`Get-ItemProperty -ErrorAction SilentlyContinue` never errors on "already absent."

**When to use:** Every registry write/removal in both scripts — this is what makes the roadmap's idempotency criterion (#4: "Running the installer twice... or the uninstaller when never installed... completes without error") true by construction, not by special-casing.

**Example (from `outlook-classic-delay-send/scripts/uninstall.ps1`):**
```powershell
function Remove-KeyIfExists {
    param([string]$Path)
    if (Test-Path $Path) {
        Remove-Item -Path $Path -Recurse -Force -ErrorAction Stop
    }
    # else: silently no-op — this IS the idempotency guarantee
}

# Resiliency key needs special care: remove only THIS add-in's named value,
# never the whole key (other add-ins' entries may share it).
$kResil = "HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList"
if (Test-Path $kResil) {
    $prop = Get-ItemProperty -Path $kResil -Name $ProgId -ErrorAction SilentlyContinue
    if ($null -ne $prop -and $null -ne $prop.$ProgId) {
        Remove-ItemProperty -Path $kResil -Name $ProgId -Force
    }
}
```

### Anti-Patterns to Avoid

- **Using `regasm.exe` in the production install script:** requires admin for its default HKLM/HKCR write mode and has no HKCU switch — directly contradicts CLAUDE.md's "no admin" constraint. Direct registry cmdlets are strictly simpler and already proven by the sibling project.
- **Removing the entire `DoNotDisableAddinList` key on uninstall:** other add-ins may have their own value under the same key. Only `Remove-ItemProperty -Name <ProgId>`, never `Remove-Item` the whole key.
- **Depending on `$PSScriptRoot` in the default (piped one-liner) code path:** breaks silently when invoked via `irm | iex` (see Pattern 2).
- **Writing the Excel Addins discovery key under a versioned path (`...\16.0\Excel\Addins\...`):** this is the *Resiliency* key's shape, not the discovery key's. The discovery key is non-versioned (confirmed by Microsoft Learn — see Sources); using the wrong shape means Excel never finds the add-in at all, silently.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Zip extraction | Custom byte-level zip parser | `Expand-Archive` (built into PS 5.1+) | Built-in, handles path validation, zero dependencies |
| GitHub Releases API response parsing | Custom JSON string parsing | `Invoke-RestMethod` (auto-deserializes JSON to PSCustomObject) | Already proven in this repo's own legacy installer |
| TLS negotiation | Custom socket/handshake code | `[System.Net.ServicePointManager]::SecurityProtocol = Tls12` (one line, set once before any network call) | PS 5.1 on Windows defaults to TLS 1.0, which GitHub has rejected since 2018 — already a solved, documented one-liner in the legacy installer |
| COM registration | Custom native `RegisterTypeLib`/`ITypeLib` calls, or invoking `regasm.exe` | Direct `New-Item`/`Set-ItemProperty` against the well-documented `InprocServer32` shape | `mscoree.dll`'s job is to read exactly these five values (`Assembly`,`Class`,`RuntimeVersion`,`CodeBase`,`ThreadingModel`) — no COM-level API call is needed, it's pure registry data |
| PE header bitness detection | A NuGet dependency just to check x86 vs x64 | ~15-line hand-rolled `BinaryReader` reading the PE header's Machine field at offset `e_lfanew + 4` | Already written, tested, and battle-proven in the sibling project's `Test-PeMachine` function — small enough that a dependency would be worse than the code it replaces |

**Key insight:** every piece of this phase's actual mechanism (registry writes, zip extraction, HTTPS download, JSON parsing) is a solved problem with a one-line built-in PowerShell answer. The only genuinely custom code needed is ~15 lines of PE-header bitness detection, and even that already exists verbatim in the sibling project.

## Common Pitfalls

### Pitfall 1: Confusing the versioned Resiliency key with the non-versioned Addins discovery key
**What goes wrong:** Writing `HKCU\...\Office\16.0\Excel\Addins\<ProgId>` (versioned) instead of `HKCU\...\Office\Excel\Addins\<ProgId>` (non-versioned) for the *discovery* key — or vice-versa for the Resiliency key.
**Why it happens:** Both keys look superficially similar and both live under `HKCU\Software\Microsoft\Office`; it's easy to copy-paste one key's shape into the other's slot.
**How to avoid:** Discovery key = non-versioned (works across all installed Excel versions, confirmed official). Resiliency key = versioned (per Office release branch, e.g. `16.0`). Keep both as separate named constants in the script, never derive one from the other.
**Warning signs:** Add-in never appears in Excel even though registry keys look "present" on inspection — check the exact path shape, not just presence.

### Pitfall 2: Assuming the DoNotDisableAddinList key's effectiveness for Excel is 100% Microsoft-guaranteed
**What goes wrong:** Treating INST-03 as "done" the moment the registry value is written, without a live-Excel behavioral check.
**Why it happens:** Microsoft's canonical documentation page for this key lives under the *Outlook* VBA docs namespace, and one Microsoft Q&A answer explicitly disputes it applies to Excel/Word, even though another citation (KB 2758876's applies-to list, which is the same underlying KB as the "add-ins are user re-enabled" article) lists Excel 2016/2013 as covered products.
**How to avoid:** Write the key (it's cheap, well-formed, and corroborated by multiple independent sources including a working sibling-project implementation for the same Office resiliency subsystem) but do not claim INST-03 is behaviorally verified until the phase's live-Excel checklist (same `human_needed` pattern as Phase 3) includes an explicit test: force a transient error in the add-in and confirm Excel does NOT silently disable it on the next launch.
**Warning signs:** None observable at install time — this only manifests after a real crash/slow-load event in a live Excel session, which is exactly why it needs its own checklist item, not just a registry-write assertion.

### Pitfall 3: Depending on `$PSScriptRoot` in the piped one-liner code path
**What goes wrong:** `$PSScriptRoot`/`$MyInvocation.MyCommand.Path` are empty when the script runs via `irm ... | iex` (no on-disk file exists in that execution context).
**Why it happens:** The sibling project's `install.ps1` (and this repo's own legacy installer) both support a "no-args, sibling bin\ folder" fallback mode that DOES use `$MyInvocation.MyCommand.Path` — but that mode is only reachable via direct local file execution (`.\install.ps1`), never via the documented one-liner.
**How to avoid:** The DEFAULT (no-args) execution path — the one INST-01's one-liner actually exercises — must do 100% of its file-location work via `$env:TEMP`/`$env:LOCALAPPDATA`, never `$PSScriptRoot`. Reserve any `$PSScriptRoot` usage for the `-Source`/local-testing convenience path only.
**Warning signs:** Script works when run locally (`.\install.ps1`) but fails mysteriously when run via `irm | iex` — a classic PowerShell gotcha unrelated to this project specifically.

### Pitfall 4: Excel file lock on the DLL during re-install/update
**What goes wrong:** `Copy-Item -Force` to overwrite `FinanceFmtTools.ComAddin.dll` fails with a file-in-use error if Excel is currently running with the add-in loaded.
**Why it happens:** COM in-process servers keep their DLL file handle open for the lifetime of the hosting process.
**How to avoid:** Reuse the exact `Assert-ExcelNotRunning`-style guard already proven in this repo's legacy `Install-FinanceFmtTools.ps1` (and the sibling project's `install.ps1`/`uninstall.ps1`) — check for the `EXCEL` process before any file copy, prompt or `-Force`-close as appropriate.
**Warning signs:** Install "succeeds" (no error) but the Ribbon still shows old behavior after Excel restart — a stale/partially-overwritten DLL is a classic symptom.

### Pitfall 5: `TraceLog`'s `System.Diagnostics.Trace` output currently has no configured listener
**What goes wrong:** Not a registration bug, but a troubleshooting gap: `AddInHost`'s real `ILog` implementation (`TraceLog`, Phase 3) calls `Trace.TraceWarning`/`TraceInformation`/`TraceError`, but no `App.config`/`.dll.config` exists anywhere in `src/FinanceFmtTools.ComAddin/` to attach a `TextWriterTraceListener`. Confirmed by direct inspection: no `.config` file exists in that project directory.
**Why it happens:** `System.Diagnostics.Trace` without an explicit listener only reaches the default trace listener (visible via a debugger or DebugView, not written to any file).
**How to avoid:** This is outside INST-01/02/03's literal scope, but if the installer/uninstaller checklist wants a troubleshooting story (matching the sibling project's `logs\` folder convention), consider (a) deploying a `FinanceFmtTools.ComAddin.dll.config` alongside the DLL that configures a `TextWriterTraceListener` writing to `%LocalAppData%\FinanceFmtTools\logs\addin.log`, and (b) preserving that log folder on uninstall by default (mirroring the sibling project's `-RemoveLogs` opt-in pattern). Flag this as a discretionary nice-to-have for the planner, not a hard requirement — it is not named in INST-01/02/03.
**Warning signs:** A user reports "the Ribbon doesn't work" with zero diagnosable output anywhere on disk.

## Code Examples

### Verified: existing 4-file binary set (from an actual local build in this repo)
```
$ ls src/FinanceFmtTools.ComAddin/bin/Release/net48/
FinanceFmtTools.ComAddin.dll     # the add-in itself
FinanceFmtTools.ComAddin.pdb     # (optional — debug symbols, can be omitted from release asset)
FinanceFmtTools.Engine.dll       # ProjectReference dependency (Phase 1/2's format engine)
FinanceFmtTools.Engine.pdb       # (optional)
Microsoft.Office.Interop.Excel.dll
office.dll                      # this IS the Microsoft.Office.Core interop assembly's actual filename
```
No `Extensibility.dll`/`stdole.dll` are needed (unlike the sibling Outlook project) — Phase 3 hand-rolled the `IDTExtensibility2` interface directly in `Extensibility.cs` rather than referencing an external interop assembly (03-01-SUMMARY.md), so there is one fewer moving part to deploy. **This 4-file (or 6, if `.pdb`s are kept) list is what the Phase 5 CI-built release zip's `bin\` folder must contain, and what this phase's installer must expect.**

### GitHub Releases download flow (already proven in this repo's own `Install-FinanceFmtTools.ps1`)
```powershell
# Source: Install-FinanceFmtTools.ps1 lines 46-49, 97-109 (this repo, already in production)
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12

function Get-LatestReleaseTag {
    $apiUrl = "https://api.github.com/repos/tpougy/finance-fmt-tools/releases/latest"
    $release = Invoke-RestMethod -Uri $apiUrl -Headers @{ 'User-Agent' = 'FinanceFmtTools-Install' } -ErrorAction Stop
    return $release.tag_name
}
```

### Post-install validation (adapted from `outlook-classic-delay-send/scripts/install.ps1` lines 396-423)
```powershell
$lb = (Get-ItemProperty -Path $kAddin -Name 'LoadBehavior' -ErrorAction SilentlyContinue).LoadBehavior
if ($lb -eq 3) { Write-Ok "LoadBehavior = 3 (carrega no início)." }
else { Write-Err2 ("LoadBehavior inesperado: {0} (esperado 3)." -f $lb) }

$cb = (Get-ItemProperty -Path $kInproc -Name 'CodeBase' -ErrorAction SilentlyContinue).CodeBase
if ($cb -eq $codeBase) { Write-Ok ("CodeBase correto: {0}" -f $cb) }
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|---------------|--------|
| `.xlam` registered via `Excel.AddIns.Add()` + `Installed = $true` (legacy VBA installer, this repo) | HKCU registry keys directly registering a COM class (CLSID/ProgId/InprocServer32) + Excel Addins discovery key | This phase (Phase 4), following the C# migration started in Phase 1 | Completely different registration mechanism — the legacy installer's `Excel.AddIns` collection is irrelevant to a COM Shared Add-in; do not reuse that code path, only its download/TLS/process-detection scaffolding |
| VSTO ClickOnce (`|vstolocal` manifest registration) | Plain COM Shared Add-in registry keys (this phase) | N/A — VSTO was never used in this project; explicitly excluded per CLAUDE.md/REQUIREMENTS.md Out of Scope | Simpler: no manifest, no ClickOnce cache, no trust-prompt/signature concerns that VSTO's registry-only deploy mode has |

**Deprecated/outdated:** Nothing in this phase's own mechanism is deprecated — HKCU-based classic COM Shared Add-in registration via `mscoree.dll` has been the stable, documented mechanism for roughly two decades and is not tied to any Office version-specific behavior beyond the versioned Resiliency key.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `HKCU\Software\Microsoft\Office\<ver>\Excel\Resiliency\DoNotDisableAddinList\<ProgId>=1` is honored by Excel the same way it is documented (and community-corroborated) for Outlook — Microsoft's own canonical doc page is written under the Outlook VBA docs namespace, and one Microsoft Q&A answer disputes cross-application applicability | Common Pitfall 2, Phase Requirements (INST-03) | If wrong, INST-03's registry write would be cosmetic only — the add-in could still get silently disabled by a transient error. Mitigated by this phase's own live-Excel verification checklist explicitly testing the behavior, not just the registry write |
| A2 | 32-bit Excel would very likely still load this AnyCPU add-in via the same HKCU keys (the project defers 32-bit *support/testing*, not necessarily technical incompatibility) | User Constraints / Claude's Discretion | Low risk — this is an informational note for the planner's bitness-guard design choice, not a functional claim this phase depends on (project scope is 64-bit only regardless) |
| A3 | `FriendlyName` value of `Finance Fmt Tools` (matching `AddInHost.cs`'s `AddinName` constant) is the right value to register, as opposed to `Finance Fmt` (the Ribbon tab's own label) | Standard Stack / User Constraints | Low risk — cosmetic only, affects what's shown in Excel's "COM Add-ins" dialog, not functional behavior; easy to change later without any migration concern since it's a REG_SZ value, not an identity key |

## Open Questions (RESOLVED)

1. **RESOLVED — Does the Resiliency `DoNotDisableAddinList` key structurally and behaviorally apply to Excel exactly as documented for Outlook?**
   - RESOLVED: Write the key regardless (low cost, well-formed, best current understanding per KB 2758876). Add an explicit item to this phase's human-verification checklist: deliberately trigger a slow/crashing load once, confirm Excel's next launch does NOT show the "add-in was disabled" notification.

2. **RESOLVED — Can this phase's installer be meaningfully tested before Phase 5 produces a real release asset?**
   - RESOLVED: Use the local-testing escape hatch (`-Package`/`-Source` flags, Pattern 2) against a manually zipped, locally `dotnet build`-produced `bin\` folder. This phase's plans must include this escape hatch explicitly, not silently defer everything to `human_needed`.

3. **RESOLVED — Should the installer detect and warn about a still-installed legacy `.xlam` add-in (potential dual-Ribbon-tab conflict)?**
   - RESOLVED: Out of scope for this phase's INST-01/02/03 requirements — do not add active detection/warning logic. This is documentation/UX territory better owned by Phase 5's LEGACY-01/02 (formal VBA retirement), not a Phase 4 installer behavior.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| Windows | Running/testing the installer at all | ✗ (this dev environment is Linux/WSL) | — | None — human_needed, same as Phase 3 |
| PowerShell (`pwsh`/`powershell`) | Executing/testing install.ps1/uninstall.ps1 | ✗ (confirmed absent from PATH in this session) | — | Scripts can be written/reviewed for syntax and logic correctness here; cannot be executed. Human runs on real Windows+Excel machine |
| A real GitHub Release with a C# add-in zip asset | The documented `irm .../install.ps1 \| iex` one-liner's download step | ✗ (only legacy `.xlam` v1.0.0/v1.0.1 releases exist today, confirmed via `gh release list`) | — | Use the `-Package`/`-Source` local-testing escape hatch (Pattern 2) against a manually zipped, locally `dotnet build`-produced `bin\` folder, until Phase 5 ships a real release |
| `gh` CLI | Confirming current release state during planning | ✓ | (available in this session) | — |

**Missing dependencies with no fallback:**
- Windows/Excel/PowerShell itself — this phase's actual install/uninstall execution and the "Ribbon tab appears" / "Resiliency key is honored" behavioral checks are inherently `human_needed`, exactly as Phase 3's live-Excel smoke test was.

**Missing dependencies with fallback:**
- Real GitHub Release asset — covered by the local-testing escape hatch (Pattern 2) described above.

## Security Domain

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-------------------|
| V2 Authentication | No | No authentication surface — this is a local, per-user install script, not a network service |
| V3 Session Management | No | Not applicable |
| V4 Access Control | Partially | HKCU-only registration is itself an access-control decision (per-user, not machine-wide) — this is the entire point of CLAUDE.md's "no admin" constraint, already correctly scoped |
| V5 Input Validation | Yes | Validate the `-Package`/`-Source` local-testing parameters (path exists, is a `.zip`, extracted contents actually contain the expected files) before acting on them — mirrors the sibling project's `Find-BinDir`/file-existence checks |
| V6 Cryptography | Yes | Force TLS 1.2 (`ServicePointManager.SecurityProtocol`) before any network call — PowerShell 5.1's default (TLS 1.0) is rejected by GitHub; already a solved, proven pattern in this repo's own legacy installer |

### Known Threat Patterns for this stack

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|-----------------------|
| Unauthenticated HTTP download of executable content (MITM tampering of the downloaded zip/DLL) | Tampering | GitHub Releases are served over HTTPS with TLS 1.2 forced; consider (optional, discretionary) a SHA-256 checksum published alongside the release asset and verified post-download — not currently done by either the legacy installer or the sibling project, so treat as a nice-to-have, not a blocking requirement |
| Zip-slip (malicious zip entries with `../` path traversal escaping the extraction directory) | Tampering / Elevation of Privilege | `Expand-Archive`'s underlying `System.IO.Compression.ZipFile.ExtractToDirectory` has validated entry paths against directory traversal since a documented .NET Framework security update; still, extract only to a fresh, uniquely-named `%TEMP%` subfolder (never directly to `%LocalAppData%\FinanceFmtTools\`) so a hypothetical traversal bug can't overwrite arbitrary user files in the final install location |
| Running `irm ... \| iex` from an untrusted or spoofed URL | Spoofing | The one-liner in README/docs must reference the exact, canonical `raw.githubusercontent.com/tpougy/finance-fmt-tools/main/...` path — no shortlink, no third-party mirror |
| Registry key injection via a maliciously-crafted ProgId/GUID | Tampering | Not applicable here — all identity values (GUID/ProgId/AssemblyName) are fixed constants embedded in the script, never derived from external/user input |

## Sources

### Primary (HIGH confidence)
- Microsoft Learn — *Registry entries for VSTO Add-ins* (`learn.microsoft.com/en-us/visualstudio/vsto/registry-entries-for-vsto-add-ins`) — confirmed the non-versioned `Root\Software\Microsoft\Office\<application>\Addins\<add-in ID>` discovery key shape (generic, not Outlook-specific), the `LoadBehavior`/`FriendlyName`/`Description` value names/types, and the HKCU-vs-WOW6432Node non-redirection fact
- Microsoft Learn — *Registering Assemblies with COM* (`learn.microsoft.com/en-us/dotnet/framework/interop/registering-assemblies-with-com`) — confirmed the `InprocServer32` (`Assembly`/`Class`/`RuntimeVersion`, `mscoree.dll` shim) structure that `regasm.exe` normally writes, and which this phase's scripts write directly instead
- Microsoft Learn — *Support for keeping add-ins enabled* (`learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled`) — confirmed the exact `DoNotDisableAddinList` key shape and hex reason-code values (Outlook-namespaced page, applicability caveat noted in Pitfall 2)
- Microsoft Learn / KB 2758876 — *Add-ins are user re-enabled after being disabled* (`learn.microsoft.com/en-us/troubleshoot/outlook/performance/add-ins-are-user-re-enabled-after-being-disabled`) — same key structure, corroborating source
- This repo's own compiled binary: `src/FinanceFmtTools.ComAddin/bin/Release/net48/FinanceFmtTools.ComAddin.dll` — directly inspected via `strings` in this session, confirming the GUID and ProgId are genuinely embedded exactly as `Connect.cs` declares
- This repo's own `Install-FinanceFmtTools.ps1` — proven-in-production (two real GitHub releases exist) download/TLS1.2/Excel-process-detection pattern
- `/home/thomaz/pessoal/outlook-classic-delay-send/scripts/install.ps1` and `uninstall.ps1` — the sibling project explicitly named in CLAUDE.md as this project's dev/build/release inspiration; solves the structurally identical HKCU COM Shared Add-in registration problem for Outlook

### Secondary (MEDIUM confidence)
- `/home/thomaz/pessoal/outlook-classic-delay-send/.planning/RESEARCH-CSHARP.md` — that project's own Phase-1 research, independently reaching the same "HKCU-only, no regasm, no admin, direct registry writes" conclusion with its own cited sources
- Microsoft Q&A — *Does the DoNotDisableAddInList registry key apply to Word and Excel?* (`learn.microsoft.com/en-us/answers/questions/4903695/`) — contradictory community answers, used only to flag the honest uncertainty in Pitfall 2/Open Question 1, not as a definitive source

### Tertiary (LOW confidence)
- `blog.blue929.com` — *Enable or Disable Office Add-in Resiliency* — asserts (without its own citation) that the same `DoNotDisableAddinList` pattern must be added per-application (Word/Excel/PowerPoint) if the add-in targets those apps; directionally consistent with the primary sources but not independently authoritative

## Metadata

**Confidence breakdown:**
- Standard stack (registry mechanism, download flow, zip extraction): HIGH — every piece is either official-Microsoft-Learn-confirmed or already proven working in this exact codebase/sibling codebase
- Architecture (three-registry-tree pattern, self-contained one-liner + local-testing escape hatch): HIGH — directly adapted from a working, structurally-identical sibling-project implementation, cross-checked against official docs
- Pitfalls: HIGH for registry-shape/file-lock/PSScriptRoot pitfalls (all directly observed/documented); MEDIUM for the Resiliency key's Excel-specific behavioral guarantee (genuine, disclosed ambiguity — see Pitfall 2/Open Question 1)

**Research date:** 2026-07-11
**Valid until:** 2026-08-10 (30 days — this domain, per-user Office COM add-in registration via HKCU, has been stable for ~two decades and is not expected to change; the one genuinely time-sensitive fact — whether a real GitHub Release exists yet — should be re-checked at plan time via `gh release list` regardless of this expiry window)
