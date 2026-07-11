---
phase: 04-installation-registration
verified: 2026-07-11T18:30:00Z
status: human_needed
score: 4/4 must-haves code-verified; 0/4 live-behavior-verified (environment constraint)
overrides_applied: 0
human_verification:
  - test: "Run `powershell -ExecutionPolicy Bypass -File .\\scripts\\install.ps1 -Source src\\FinanceFmtTools.ComAddin\\bin\\Release\\net48` on a real Windows+Excel machine (local-testing escape hatch, since no real GitHub C# release exists yet)."
    expected: "Script reports [OK] for every registry key/file check in its post-install validation section, exits 0, and no admin/UAC prompt ever appears (every write targets HKCU)."
    why_human: "No Windows/PowerShell/registry available in this Linux/WSL environment; registry-write execution cannot be observed here."
  - test: "Open Excel after install.ps1 completes and confirm the 'Finance Fmt' Ribbon tab appears."
    expected: "Tab renders (parity already confirmed by Phase 3's own pending Ribbon-render check)."
    why_human: "Requires a live Excel session — no COM runtime here."
  - test: "Run install.ps1 a second time immediately after a successful install (idempotency), then run uninstall.ps1, confirm the Ribbon tab disappears, then run uninstall.ps1 again (idempotency)."
    expected: "Both re-runs complete without error (exit 0); the Ribbon tab disappears after uninstall and stays gone; re-running uninstall reports 'already absent' for every key/file."
    why_human: "Requires actually executing PowerShell against a live Windows registry and observing Excel's Ribbon state across restarts."
  - test: "Deliberately force a slow/crashing add-in load once (e.g. temporarily rename FinanceFmtTools.Engine.dll or throw in Connect.OnConnection), relaunch Excel, restore the DLL, relaunch again, and confirm Excel's native 'this add-in has been disabled' notification never appears."
    expected: "DoNotDisableAddinList is behaviorally honored — Excel does not silently disable the add-in after the one transient failure (INST-03's actual behavioral proof, not just registry-write success)."
    why_human: "This is the one item 04-RESEARCH.md itself flags as not 100% Microsoft-guaranteed for Excel (Pitfall 2 / Open Question 1, documentation lives under Outlook's namespace) — only observable by triggering the real Excel Resiliency subsystem."
---

# Phase 4: Installation & Registration Verification Report

**Phase Goal:** A non-admin user can install and uninstall the add-in on 64-bit Excel with a single PowerShell command, and the installed add-in survives transient errors without being silently disabled by Excel.
**Verified:** 2026-07-11
**Status:** human_needed
**Re-verification:** No — initial verification

## Environment Constraint (documented, non-discretionary)

This verification runs in a Linux/WSL environment with **no Windows, no Excel, no PowerShell interpreter, and no registry**. 04-CONTEXT.md explicitly documents this as a non-discretionary constraint carried forward from Phase 3: this phase's own goal statement requires a single PowerShell command to actually install/uninstall the add-in on a real Windows+Excel machine, which cannot be executed here. This is the expected, by-design outcome for this phase in this environment — not a missing or incomplete implementation, and mirrors Phase 3's own `human_needed` verdict (03-VERIFICATION.md: 5/5 code-verified, 0/5 live-behavior-verified). What follows verifies everything that IS provable by reading the actual script text in this environment (constants, registry-key shapes, control flow, idempotency guards, error handling, encoding) and produces an itemized human-verification checklist for the remaining live-Excel behavior.

## Goal Achievement

### Observable Truths

| # | Truth | Status | Evidence |
|---|-------|--------|----------|
| 1 | INST-01: One-liner install downloads the latest GitHub release, registers entirely under HKCU with no admin prompt, and the "Finance Fmt" tab appears next time Excel opens | ? UNCERTAIN (code verified, live behavior human_needed) | `scripts/install.ps1` default (no-args) branch calls `Get-LatestReleaseTag` + `Invoke-WebRequest -Uri $DownloadUrl` against `https://github.com/tpougy/finance-fmt-tools/releases/latest/download/FinanceFmtTools.zip` (line 91, 304-335), forces TLS1.2 before any network call (line 103), extracts only to a fresh `%TEMP%` GUID-named subfolder (zip-slip mitigation, never into `$InstallDir` directly), then writes exactly 3 HKCU registry trees (`Software\Classes\...`, `Software\Microsoft\Office\Excel\Addins\...`, `Software\Microsoft\Office\16.0\Excel\Resiliency\...` — lines 386-424). Zero `HKLM:` **writes** anywhere in the file (`HKLM:` appears only in 2 **read-only** `Test-Path`/`Get-ItemProperty` checks for informational bitness detection, lines 231-253). Actual download/registration execution and Ribbon-tab rendering not observable in this environment. |
| 2 | INST-02: Uninstall script removes the HKCU registration keys and installed files, and the Ribbon tab no longer appears after Excel restarts | ? UNCERTAIN (code verified, live behavior human_needed) | `scripts/uninstall.ps1` removes the CLSID subtree, the ProgId→CLSID mapping key, and the non-versioned Excel discovery key via `Remove-KeyIfExists` (lines 124-128), removes only the named Resiliency **value** via `Remove-ItemProperty` (never the parent key — lines 130-143), and removes each of the 4 named files individually before removing the (now-empty) install directory (lines 148-174). Actual registry removal and "tab disappears" behavior not observable here. |
| 3 | INST-03: `DoNotDisableAddinList` registry key is present for the add-in's ProgID after install, so a transient runtime error does not cause Excel to silently disable the add-in | ? UNCERTAIN (registry-write code verified; behavioral honoring by Excel is explicitly flagged as unverified even in research) | `install.ps1` writes `HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList` with `Set-ItemProperty -Name $ProgId -Value 1 -Type DWord` (lines 421-424) — versioned key, correctly distinguished from the non-versioned discovery key (confirmed: zero occurrences of the incorrectly-versioned `$OfficeVerKey\Excel\Addins` form in either script). 04-RESEARCH.md's own Pitfall 2 / Open Question 1 explicitly states this key's behavioral effectiveness for Excel (as opposed to Outlook, where Microsoft's canonical doc page lives) is corroborated but not 100% Microsoft-guaranteed — the phase's own design correctly routes this to a live-Excel behavioral test (04-03-PLAN.md step 6) rather than claiming it proven by the registry write alone. |
| 4 | Idempotency: running the installer twice, or the uninstaller when never installed, completes without error | VERIFIED (code-level) — live confirmation still human_needed | Every registry write uses `New-Item -Force` + `Set-ItemProperty` (never errors on "already exists"). Every removal in `uninstall.ps1` is `Test-Path`-guarded (`Remove-KeyIfExists`, the Resiliency-value removal, and the per-file removal loop) and the script's only unconditional `exit 1` path is inside `Assert-ExcelNotRunning`'s Excel-is-running guard — the main flow always reaches `exit 0` otherwise (confirmed by reading lines 175-193). This is a code-level guarantee provable by structural reading alone, not requiring a live environment to reason about, though the phase itself still schedules a live re-run as a checklist item for full confidence. |

**Score:** 4/4 must-haves are code-complete and structurally verified; 0/4 have live-Excel behavioral confirmation (expected, per documented environment constraint — routed to human_verification, not treated as failures).

### Required Artifacts

| Artifact | Expected | Status | Details |
|----------|----------|--------|---------|
| `scripts/install.ps1` | Self-contained, admin-free installer: one-liner GitHub-Releases flow (INST-01) + `-Package`/`-Source` local-testing escape hatch + 3-registry-tree HKCU registration incl. Resiliency (INST-03) | VERIFIED | 479 lines. Fixed identity constants (`$Guid`, `$ProgId`, `$ClassName`, `$RuntimeVer`, `$Shim`, `$ThreadingMdl`, `$FriendlyName`, `$OfficeVerKey`) match `Connect.cs`'s header doc-comment verbatim (GUID `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`, ProgId `FinanceFmtTools.Connect`) — literal GUID string appears exactly once. Zero `$PSScriptRoot`/`$MyInvocation` dependency. TLS1.2 forced before any network call. Zip-slip mitigation confirmed (extraction always to a fresh `%TEMP%` subfolder). Brace-balanced (125/125). |
| `scripts/uninstall.ps1` | Idempotent removal of the 3 HKCU trees + 4 installed files (INST-02) | VERIFIED | 193 lines. Mirrors `install.ps1`'s identity constants exactly (independently declared, values match). `Remove-KeyIfExists` used 3x; Resiliency value removed via `Remove-ItemProperty` only (never `Remove-Item` on the parent key — confirmed zero matches). Brace-balanced (50/50). |
| `scripts/verify-environment.ps1` | Read-only diagnostic (discretionary, per 04-RESEARCH.md) | VERIFIED | 284 lines. Zero `Set-ItemProperty`/`New-Item`/`Remove-Item` write cmdlets found — genuinely read-only. References `EXCEL.EXE`, never `OUTLOOK.EXE`. Checks `.NET Framework 4.8` via `NDP\v4\Full` `Release>=528040`. `-RuntimeOnly` switch present and functional. Brace-balanced (91/91). |

### Key Link Verification

| From | To | Via | Status | Details |
|------|-----|-----|--------|---------|
| `scripts/install.ps1` | `HKCU:\Software\Classes\CLSID\{881EFDF3-...}\InprocServer32` | `New-Item -Force` + `Set-ItemProperty` (Assembly/Class/CodeBase/RuntimeVersion/ThreadingModel) | WIRED | All 6 values written (lines 404-409); `CodeBase` built via `([Uri]$dllPath).AbsoluteUri` (percent-encodes spaces in paths); `Assembly` value read live from the copied DLL via `[System.Reflection.AssemblyName]::GetAssemblyName` rather than hardcoded, eliminating version-drift risk. |
| `scripts/install.ps1` | `HKCU:\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect` | `Set-ItemProperty -Name LoadBehavior -Value 3 -Type DWord` | WIRED | Non-versioned discovery key confirmed correct shape (line 413); `LoadBehavior=3` set and re-validated post-write (lines 445-447). |
| `scripts/install.ps1` | `HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList` | `Set-ItemProperty -Name $ProgId -Value 1 -Type DWord` | WIRED | Versioned key (correctly distinguished from the discovery key), value written (lines 421-424). |
| `scripts/uninstall.ps1` | `HKCU:\Software\Classes\CLSID\{881EFDF3-...}` | `Remove-KeyIfExists` (Test-Path guarded `Remove-Item -Recurse -Force`) | WIRED | Confirmed idempotent — logs "already absent" rather than erroring when not present. |
| `scripts/uninstall.ps1` | `HKCU:\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList` | `Remove-ItemProperty -Name $ProgId` (never `Remove-Item` on the parent key) | WIRED | Confirmed: `Remove-Item` is never called against `$kResil` itself anywhere in the file; only the named value is targeted, preserving other add-ins' entries under the same shared key. |

### Data-Flow Trace (Level 4)

Not applicable in the traditional sense (no UI framework rendering dynamic data from a store). Traced instead as "does the script's registry-write path actually derive from real, freshly-resolved values rather than stale/hardcoded ones": `$codeBase` and `$AssemblyStr` are both computed from the just-copied `$dllPath` at runtime (not pre-computed constants), so a future assembly version bump cannot silently desynchronize the registered `Assembly=` value from the shipped DLL — this closes the exact drift risk the code review (WR-03) flagged, confirmed fixed by direct reading of `install.ps1:375-382`.

### Behavioral Spot-Checks

Not run — this phase's runnable artifacts are Windows PowerShell scripts that write to the Windows registry and interact with a live Excel process; no PowerShell interpreter (`pwsh`/`powershell`) is available in this Linux/WSL environment (confirmed: both commands absent from PATH). This is the same constraint driving the phase's own `human_needed` design; see Human Verification section instead.

### Probe Execution

No `scripts/*/tests/probe-*.sh` files exist in this repository (checked via `find scripts -path '*/tests/probe-*.sh'` and a repo-wide equivalent search), and neither PLAN nor SUMMARY files reference any probe scripts. Step 7c: SKIPPED (no probes declared or found).

### Independent Checks (run in this session, not taken from SUMMARY.md/REVIEW.md claims)

| Check | Result |
|-------|--------|
| `git status --short` | Clean working tree |
| `git log --oneline -- scripts/` | 3 commits: `4430019` (install.ps1), `90876f4` (uninstall.ps1 + verify-environment.ps1), `a38ce24` (fix pass) — all exist in history, `git show --stat` confirms only `scripts/*.ps1` files touched, zero C# files modified in any Phase 4 commit |
| `Install-FinanceFmtTools.ps1` (legacy, repo root) modification history | Untouched by Phase 4 — last touched by pre-milestone commits (`a187b8a`, `a5c8fbb`), confirming the plan's explicit "do not modify" instruction was honored |
| Brace balance (`{` vs `}`) | install.ps1: 125/125, uninstall.ps1: 50/50, verify-environment.ps1: 91/91 — all balanced |
| `grep -c 'PSScriptRoot\|MyInvocation'` on install.ps1/uninstall.ps1 | 0/0 — confirmed the documented one-liner path has zero dependency on the script's own on-disk location |
| Literal GUID string occurrence count | Exactly 1 in each of install.ps1/uninstall.ps1 (only in the `$Guid` declaration) |
| Non-versioned vs. versioned key-shape cross-check | `$OfficeVerKey\Excel\Addins` (the incorrect, versioned discovery-key shape) — 0 matches in either script. `Office\Excel\Addins` (correct, non-versioned) present in both. `$OfficeVerKey\Excel\Resiliency` (correct, versioned) present in both. |
| `Remove-Item.*kResil` (whole-key deletion of the shared Resiliency key) | 0 matches — only `Remove-ItemProperty -Path $kResil -Name $ProgId` found |
| File encoding (`xxd`/`file`) | All 3 `.ps1` files confirmed UTF-8 **with BOM** (`efbb bf` header) — the WR-06 code-review fix is genuinely present in the current file bytes, not just claimed |
| Anti-pattern scan (`TBD\|FIXME\|XXX\|TODO\|HACK\|PLACEHOLDER`, case-insensitive) | Zero genuine matches; the only substring hits were Portuguese words containing "todo"/"todos" ("all"/"every"), which is a false-positive of the case-insensitive regex, not an actual debt marker |
| `dotnet build src/FinanceFmtTools.sln -c Release` (re-run in this session) | 0 Warning(s), 0 Error(s); `src/FinanceFmtTools.ComAddin/bin/Release/net48/` confirmed to contain all 4 files (`FinanceFmtTools.ComAddin.dll`, `FinanceFmtTools.Engine.dll`, `Microsoft.Office.Interop.Excel.dll`, `office.dll`) both scripts' `$AllFiles` list expects |

### Requirements Coverage

| Requirement | Source Plan | Description | Status | Evidence |
|-------------|-------------|--------------|--------|----------|
| INST-01 | 04-01, 04-03 | One-liner PowerShell installer downloads latest GitHub release, registers via HKCU for Excel 64-bit, no admin | NEEDS HUMAN | Script logic fully verified structurally (download flow, TLS1.2, zip-slip mitigation, HKCU-only writes, zero HKLM writes); live download/registration/Ribbon-render behavior unobservable here. |
| INST-02 | 04-02, 04-03 | Uninstall script removes HKCU registration and installed files | NEEDS HUMAN | Removal logic fully verified structurally (idempotent, correctly scoped Resiliency-value-only removal); live "tab disappears" behavior unobservable here. |
| INST-03 | 04-01, 04-03 | Installer writes `DoNotDisableAddinList` to prevent silent disable after transient error | NEEDS HUMAN | Registry-write logic verified structurally and key-shape-correct; behavioral honoring by Excel is the one item this phase's own research (04-RESEARCH.md Pitfall 2) flags as needing a live test, not just a registry-write assertion. |

No orphaned requirements: REQUIREMENTS.md maps exactly INST-01/02/03 to Phase 4, and all three plans collectively declare them (`04-01-PLAN.md requirements: [INST-01, INST-03]`, `04-02-PLAN.md requirements: [INST-02]`, `04-03-PLAN.md requirements: [INST-01, INST-02, INST-03]`). REQUIREMENTS.md itself currently marks all three `[ ]` (unchecked) — consistent with the `human_needed` status; no premature "complete" marking to flag, unlike the minor inconsistency noted in Phase 3's verification.

### Anti-Patterns Found

None. Scanned all 3 `scripts/*.ps1` files for `TBD`/`FIXME`/`XXX`/`TODO`/`HACK`/`PLACEHOLDER`, "not yet implemented"/"coming soon", empty-body implementations, and hardcoded-empty stub patterns — zero genuine matches (only Portuguese-language false positives from the "todo/todos" substring, addressed above). No debt markers requiring a blocker classification.

**04-REVIEW.md cross-check (fix-pass claims independently re-verified against current file bytes, not trusted from the review's own "fixed" disposition table):**
- CR-01 (Excel force-kill could discard unsaved work after a 3s grace period): confirmed fixed — `Stop-Process` no longer appears anywhere in either script; both now poll for up to 30s and fail safely with an actionable message (`install.ps1:126-153`, `uninstall.ps1:66-93`).
- WR-01 (unescaped `CodeBase` URI): confirmed fixed — `$codeBase = ([Uri]$dllPath).AbsoluteUri` (`install.ps1:379`).
- WR-03 (hardcoded `$AssemblyStr` version-drift risk): confirmed fixed — derived from `[System.Reflection.AssemblyName]::GetAssemblyName($dllPath).FullName` (`install.ps1:382`).
- WR-04 (TOCTOU gap on Excel-running check): confirmed fixed — `Assert-ExcelNotRunning` called a second time immediately before `Copy-Item`/`Remove-Item` in both scripts (`install.ps1:363`, `uninstall.ps1:153`).
- WR-05 (missing try/catch around registry/file operations): confirmed fixed — both scripts now wrap their entire registration/removal block in `try { ... } catch { Write-Err2 ...; exit 1 } finally { ... }` (`install.ps1:359-430`, `uninstall.ps1:117-178`).
- WR-06 (missing UTF-8 BOM, mojibake risk for accented Portuguese text): confirmed fixed — all 3 files verified via `xxd` to start with the `EF BB BF` BOM byte sequence.
- IN-02 (`Test-PeMachine` handle leak): confirmed fixed — `finally { if ($br) {$br.Dispose()}; if ($fs) {$fs.Dispose()} }` present in both `install.ps1` and `verify-environment.ps1`.
- IN-03 (`Find-BinDir` non-deterministic stale-artifact pick): confirmed fixed — `Sort-Object { $_.FullName -notmatch '\\bin\\' }` preference added (`install.ps1:199-201`).
- WR-02/IN-01 (shared `scripts/common.ps1` module): confirmed genuinely deferred, not silently dropped — no `scripts/common.ps1` file exists; the deferral rationale (04-02-PLAN.md's explicit "declared independently... no PowerShell module infrastructure is in scope for this phase") is a real, documented planning decision, not an unaddressed defect. This is a reasonable, low-risk deferral (duplication risk, not a functional bug) and does not block phase completion.

## Human Verification Required

See YAML frontmatter `human_verification` section for the full itemized checklist (4 items, consolidating 04-03-SUMMARY.md's 8-step "User Setup Required" checklist). Summary:

### 1. Fresh Install (INST-01)

**Test:** Run `install.ps1 -Source <locally-built bin folder>` on a real Windows+Excel machine (local-testing escape hatch, since no real GitHub C# release exists yet — confirmed via `gh release list`, only legacy `.xlam` v1.0.0/v1.0.1 releases exist).
**Expected:** Script reports `[OK]` for every post-install validation check, exits 0, no admin/UAC prompt.
**Why human:** No Windows/PowerShell/registry in this environment.

### 2. Ribbon Tab Appears (INST-01)

**Test:** Open Excel after install completes.
**Expected:** "Finance Fmt" tab renders.
**Why human:** Requires a live Excel session.

### 3. Idempotency + Uninstall (INST-02, roadmap criterion #4)

**Test:** Re-run install.ps1 immediately after success; run uninstall.ps1 and confirm the tab disappears; run uninstall.ps1 again with nothing installed.
**Expected:** All re-runs complete without error; tab disappears after uninstall; re-running uninstall reports "already absent" for everything.
**Why human:** Requires executing PowerShell against a live registry and observing Excel across restarts.

### 4. Resiliency Behavioral Test (INST-03)

**Test:** Deliberately break the add-in's load once (rename a dependency DLL or throw in `OnConnection`), let Excel fail/crash once, restore, relaunch.
**Expected:** Excel's native "this add-in has been disabled" notification never appears; the add-in loads normally on the next launch.
**Why human:** This is the one behavioral claim 04-RESEARCH.md itself says is not 100%-Microsoft-guaranteed for Excel (the canonical `DoNotDisableAddinList` doc page lives under Outlook's namespace) — only a live Resiliency-subsystem trigger can prove it, not a registry-write assertion.

## Gaps Summary

No code gaps found. All three scripts exist, are substantive (not stubs — 479/193/284 lines respectively, each implementing its full documented scope), are correctly wired to the fixed COM identity Phase 3 established (`Connect.cs`'s GUID/ProgId/AssemblyName, reused verbatim, not re-invented), and independently re-verified against `04-REVIEW.md`'s "fixed" claims by reading the current file bytes rather than trusting the review's own disposition table — every claimed fix (CR-01, WR-01, WR-03, WR-04, WR-05, WR-06, IN-02, IN-03) is genuinely present in the code, and the one deliberately-deferred item (WR-02/IN-01, shared `common.ps1`) is genuinely absent (not silently dropped), with a documented, defensible rationale.

The phase's own goal statement — "a non-admin user can install and uninstall the add-in... with a single PowerShell command" and "survives transient errors without being silently disabled" — cannot be satisfied by this Linux/WSL verification environment, which has no Windows, PowerShell interpreter, Excel, or registry. This is a **documented, non-discretionary environment constraint** (04-CONTEXT.md, carried forward verbatim from Phase 3's identical constraint), not a missing or incomplete implementation, and correctly routes this phase's status to `human_needed` rather than `gaps_found` — mirroring Phase 3's own verified precedent (03-VERIFICATION.md: 5/5 code-verified, 0/5 live-behavior-verified, `human_needed`). The 4-item live-Excel checklist above (consolidated from 04-03-SUMMARY.md's 8-step itemized checklist) must be run and reported back by the user on their own Windows+Excel machine before INST-01 through INST-03 can be considered fully verified end-to-end.

No documentation inconsistency to flag this time: REQUIREMENTS.md correctly leaves INST-01/02/03 unchecked pending human verification (unlike Phase 3's minor RIB-01 premature-checkmark issue, which does not recur here).

---

_Verified: 2026-07-11_
_Verifier: Claude (gsd-verifier)_
