---
phase: 04-installation-registration
reviewed: 2026-07-11T17:50:10Z
depth: standard
files_reviewed: 3
files_reviewed_list:
  - scripts/install.ps1
  - scripts/uninstall.ps1
  - scripts/verify-environment.ps1
findings:
  critical: 1
  warning: 6
  info: 3
  total: 10
status: fixed
fix_disposition:
  CR-01: fixed
  WR-01: fixed
  WR-02: deferred (out of phase scope — see note below)
  WR-03: fixed
  WR-04: fixed
  WR-05: fixed
  WR-06: fixed
  IN-01: deferred (same rationale as WR-02)
  IN-02: fixed
  IN-03: fixed
---

**Fix pass completed 2026-07-11.** 9 of 10 findings fixed directly (CR-01, WR-01, WR-03, WR-04, WR-05, WR-06, IN-02, IN-03). WR-02/IN-01 (extracting a shared `scripts/common.ps1` module) deliberately deferred: `04-02-PLAN.md`'s own Task 1 instructions state constants must be "declared independently — this script does not source or depend on install.ps1's file" and that helpers are "duplicated, not shared — no PowerShell module infrastructure is in scope for this phase," mirroring the sibling `outlook-classic-delay-send` project's identical convention. Fixing this would mean overriding an explicit, already-approved planning decision rather than a code defect — left as a documented, intentional non-fix. See STATE.md for the full rationale.

# Phase 04: Code Review Report

**Reviewed:** 2026-07-11T17:50:10Z
**Depth:** standard
**Files Reviewed:** 3
**Status:** issues_found

## Summary

Reviewed `scripts/install.ps1`, `scripts/uninstall.ps1`, and `scripts/verify-environment.ps1` — the HKCU-only, no-admin, no-regasm installer/uninstaller/diagnostic trio for the C# COM add-in.

Two specific patterns called out in the review brief were checked and found **correctly implemented**:
- The zip-slip mitigation is real: both `-Package` and the default GitHub-download flow always `Expand-Archive` into a fresh per-run temp directory (`$script:TempExtractDir`, GUID-named) and only ever copy a fixed whitelist of files (`$AllFiles`) discovered by exact name into the final install directory — the zip is never extracted directly into `$InstallDir`.
- The Resiliency `DoNotDisableAddinList` key is handled correctly in `uninstall.ps1`: only the named value (`$ProgId`) is removed via `Remove-ItemProperty`; the shared key itself is never deleted wholesale.

However, the review surfaced one blocker-level data-loss risk (a very short, non-verified force-close of Excel), plus several correctness/maintainability warnings, chiefly: unescaped characters in the generated `CodeBase` file URI, no single source of truth for the identity constants (GUID/ProgId/file list) that are duplicated verbatim between `install.ps1` and `uninstall.ps1`, a TOCTOU gap between the "Excel is closed" check and the actual file copy/delete, inconsistent error handling between the network/extract steps (which have friendly try/catch) and the registry/file steps (which do not), and a missing UTF-8 BOM that risks mojibake of all the (extensively used) accented Portuguese message strings when the scripts are run locally via Windows PowerShell 5.1's `-File` path.

## Critical Issues

### CR-01: `-Force` can kill Excel and discard unsaved user work after only a 3-second grace period

**File:** `scripts/install.ps1:125-140`, `scripts/uninstall.ps1:66-81` (identical logic in both files)

**Issue:** When Excel is running and `-Force` is passed (explicitly documented and recommended as a normal usage mode — see `install.ps1`'s own `.EXAMPLE` block, `powershell ... -File .\scripts\install.ps1 -Force`), `Assert-ExcelNotRunning` does:

```powershell
$excelProcs | ForEach-Object { $_.CloseMainWindow() | Out-Null }
Start-Sleep -Seconds 3
$excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
if ($excelProcs) {
    Write-Warn2 'Excel não fechou sozinho; encerrando o processo (-Force)...'
    $excelProcs | Stop-Process -Force -ErrorAction Stop
    Start-Sleep -Seconds 2
}
```

`CloseMainWindow()` posts `WM_CLOSE`, which for Excel triggers the standard "Save changes to <workbook>?" modal dialog whenever there are unsaved changes in *any* open workbook — not just a workbook related to this install. The script gives the user a single, unconditional 3-second window to notice that dialog (which may not even have focus, since the user's attention is on the terminal) and respond to it, after which it force-kills the process via `Stop-Process -Force`, unconditionally discarding whatever the pending dialog would have saved. There is no check for `Workbook.Saved` state via COM automation, no increased/adaptive grace period, and no re-confirmation before the kill. This is a realistic, easily triggered data-loss path — any user who happens to have unrelated unsaved work open in Excel and reinstalls/upgrades with `-Force` (the documented convenience flag) risks silently losing it.

**Fix:** Either drop the automatic force-kill entirely (only ever `CloseMainWindow()` and then fail with an actionable message asking the user to close Excel and re-run, without a timeout-based `Stop-Process`), or, at minimum, use COM automation to check for unsaved workbooks before offering `-Force`, and give a materially longer/interactive grace period (e.g. poll every second for 30-60s with a running countdown message) before ever calling `Stop-Process -Force`:

```powershell
if ($Force) {
    Write-Warn2 'Excel está aberto. -Force informado: tentando fechar com segurança...'
    $excelProcs | ForEach-Object { $_.CloseMainWindow() | Out-Null }
    for ($i = 0; $i -lt 30; $i++) {
        Start-Sleep -Seconds 1
        $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
        if (-not $excelProcs) { break }
    }
    if ($excelProcs) {
        Write-Err2 'Excel ainda está aberto (possível diálogo de "salvar alterações?" pendente). Salve seu trabalho e feche o Excel manualmente, depois rode novamente.'
        exit 1
    }
}
```

## Warnings

### WR-01: `CodeBase` file URI does not percent-encode special characters (e.g. spaces) in the install path

**File:** `scripts/install.ps1:352-353`

**Issue:**
```powershell
$dllPath  = Join-Path $InstallDir $DllName
$codeBase = 'file:///' + ($dllPath -replace '\\', '/')
```
`$InstallDir` is `Join-Path $env:LOCALAPPDATA 'FinanceFmtTools'`. On machines where the Windows user profile folder contains a space (not uncommon in corporate/domain environments, e.g. `C:\Users\John Smith\AppData\Local\...`), this produces an unescaped space in the `CodeBase` registry value written under `InprocServer32`. This is a plausible, if narrow, environment-dependent activation failure for the exact users most likely to need a no-admin/HKCU-only installer (corporate machines).

**Fix:** Build the URI via `System.Uri` so reserved/unsafe characters are properly percent-encoded:
```powershell
$codeBase = ([Uri]$dllPath).AbsoluteUri
```

### WR-02: Identity/config constants duplicated verbatim between `install.ps1` and `uninstall.ps1` with no shared source of truth

**File:** `scripts/install.ps1:70-98`, `scripts/uninstall.ps1:45-50`

**Issue:** `$Guid`, `$ProgId`, `$OfficeVerKey`, and the 4-file `$AllFiles` list are declared independently, as literal strings, in both scripts (the comment in `uninstall.ps1:41-44` explicitly acknowledges this is by design: "declarada de forma independente aqui; este script não depende do arquivo install.ps1"). There is no shared module, no test, and no runtime cross-check that these stay in sync with each other or with the actual values baked into `Connect.cs`/the `.csproj` (currently consistent, but nothing prevents drift on the next release/refactor — e.g. if the GUID or file list changes in one script but not the other, `uninstall.ps1` would silently fail to fully clean up a newer install, or `install.ps1` would register a GUID that no longer matches the shipped assembly).

**Fix:** Extract the shared identity/config constants (and the duplicated `Write-Step`/`Write-Ok`/`Write-Info`/`Write-Warn2`/`Write-Err2`/`Assert-ExcelNotRunning`/`Test-PeMachine` helpers — see IN-01) into a single dot-sourced `scripts/common.ps1`, imported by both `install.ps1` and `uninstall.ps1`.

### WR-03: Hardcoded assembly version string has no automated verification against the actual build output

**File:** `scripts/install.ps1:73`

**Issue:** `$AssemblyStr = 'FinanceFmtTools.ComAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null'` is a hand-maintained literal that must always match `FinanceFmtTools.ComAddin.csproj`'s `<Version>` (currently `1.0.0.0`, so today it is correct). Nothing in this script (or elsewhere in the repo) verifies this at install time or at build/release time; the `Assembly=` value registered under `InprocServer32` will silently go stale the first time the assembly version is bumped for a real release, which can break COM/CLR activation for end users with no signal other than "the add-in doesn't load."

**Fix:** Derive the value from the actually-copied DLL at install time instead of hardcoding it:
```powershell
$asmName  = [System.Reflection.AssemblyName]::GetAssemblyName($dllPath)
$AssemblyStr = $asmName.FullName
```

### WR-04: TOCTOU gap between the "Excel is closed" check and the file copy/delete operations

**File:** `scripts/install.ps1:118-146` (check) vs. `345-350` (copy); `scripts/uninstall.ps1:59-86` (check) vs. `140-161` (delete)

**Issue:** `Assert-ExcelNotRunning` runs once, at the very start of each script (Passo 0 / Passo 1). In `install.ps1`, an arbitrary amount of time can elapse between that check and the actual `Copy-Item` calls (network download + zip extraction + file-existence validation), during which the user could relaunch Excel — which is exactly the "classic file lock" scenario the check comment (`install.ps1:116-117`) says it exists to prevent. The check is never repeated immediately before the risky file operations, so a locked-file failure at that point degrades ungracefully (see WR-05).

**Fix:** Re-run `Assert-ExcelNotRunning` (or at least a cheap `Get-Process -Name EXCEL` check) immediately before `Copy-Item`/`Remove-Item` on the target binaries, in addition to the upfront check.

### WR-05: Registration/removal steps lack the friendly try/catch used by earlier steps

**File:** `scripts/install.ps1:341-398` (Passo 2: file copy + 3 registry trees); `scripts/uninstall.ps1:89-161` (Passo 2/3: registry + file removal); docstring `scripts/uninstall.ps1:17-19`

**Issue:** The download/extract steps in `install.ps1` (lines 259-266, 298-309) are wrapped in `try/catch` with `Write-Err2`/`Remove-TempExtract`/`exit 1` on failure. In contrast, the entire registration block — file copy plus all three registry trees — is wrapped only in `try { ... } finally { Remove-TempExtract }`, with no `catch`. Any failure here (permission issue, locked DLL per WR-04, corrupted registry hive, etc.) propagates as a raw, unhandled PowerShell exception/stack trace instead of the script's own friendly `[ERRO]` reporting convention used everywhere else. The same applies to `uninstall.ps1`'s registry/file removal steps, which directly contradicts its own docstring's claim ("IDEMPOTENTE: ... exit 0 incondicional ao final, exceto quando o Excel está aberto") — an unexpected removal failure (e.g. `AccessDenied` on a locked DLL) will not reach the unconditional `exit 0` at the bottom; it will throw uncaught.

**Fix:** Wrap the registry/file operations in both scripts with `try/catch` blocks that log a friendly `Write-Err2` message and `exit 1`, consistent with the rest of the script's error-handling convention.

### WR-06: Scripts saved as UTF-8 without a BOM — accented Portuguese text will likely be mojibake under Windows PowerShell 5.1's `-File` execution path

**File:** `scripts/install.ps1`, `scripts/uninstall.ps1`, `scripts/verify-environment.ps1` (all three; confirmed via `file`/`xxd`: UTF-8 text, no BOM)

**Issue:** All three scripts are UTF-8-encoded without a byte-order mark, and all are saturated with accented Portuguese literals passed to `Write-Host` (`não`, `está`, `instalação`, `é`, `Descrição`, etc.), including in the value written to the `Description` registry property (`install.ps1:78`). Windows PowerShell 5.1 (explicitly declared as the supported baseline — `#Requires -Version 5.1`, and `-File` execution is a documented `.EXAMPLE` for all three scripts) reads script files without a BOM using the system's legacy ANSI code page, not UTF-8. On a typical pt-BR Windows install this will very likely render every accented character incorrectly (mojibake) whenever a user runs these scripts locally via `-File` rather than via the `irm | iex` one-liner (which sidesteps the issue because `Invoke-RestMethod` decodes the HTTP response using the server's declared charset, not the file's on-disk BOM).

**Fix:** Save all three `.ps1` files as UTF-8 **with BOM** (`Set-Content -Encoding UTF8` in PowerShell 5.1 writes a BOM by default; most editors have an explicit "UTF-8 with BOM" save option), which both Windows PowerShell 5.1 and PowerShell 7+ interpret correctly.

## Info

### IN-01: `Test-PeMachine`, `Assert-ExcelNotRunning`, and the `Write-Step`/`Write-Ok`/`Write-Info`/`Write-Warn2`/`Write-Err2` helpers are duplicated verbatim across files

**File:** `scripts/install.ps1:110-114,118-146,149-165`; `scripts/uninstall.ps1:52-56,59-86`; `scripts/verify-environment.ps1:76-95`

**Issue:** `Test-PeMachine` is byte-for-byte identical in `install.ps1` and `verify-environment.ps1`. `Assert-ExcelNotRunning` and the five `Write-*` output helpers are byte-for-byte identical between `install.ps1` and `uninstall.ps1`. This is pure copy-paste with no shared module, compounding the drift risk already described in WR-02.

**Fix:** Move these into a single `scripts/common.ps1`, dot-sourced (`. "$PSScriptRoot\common.ps1"`) from all three scripts.

### IN-02: `Test-PeMachine` leaks the `FileStream`/`BinaryReader` if reading the PE header throws

**File:** `scripts/install.ps1:151-164`, `scripts/verify-environment.ps1:79-94`

**Issue:**
```powershell
$fs = [System.IO.File]::OpenRead($Path)
$br = New-Object System.IO.BinaryReader($fs)
$fs.Seek(0x3C, 'Begin') | Out-Null
$peOff = $br.ReadInt32()
$fs.Seek($peOff + 4, 'Begin') | Out-Null
$machine = $br.ReadUInt16()
$br.Close(); $fs.Close()
```
`$br.Close(); $fs.Close()` only run on the success path; the surrounding `try/catch` catches any exception (e.g. a truncated or non-PE file causing `Seek`/`ReadInt32` to throw) and returns `'desconhecido'`, but the handles opened on `$Path` (which is `EXCEL.EXE` in both call sites) are never released in that case, leaking a file handle for the remainder of the script's process lifetime.

**Fix:** Use `finally` to guarantee disposal:
```powershell
$fs = $null
try {
    $fs = [System.IO.File]::OpenRead($Path)
    $br = New-Object System.IO.BinaryReader($fs)
    ...
} catch {
    'desconhecido'
} finally {
    if ($br) { $br.Dispose() }
    if ($fs) { $fs.Dispose() }
}
```

### IN-03: `Find-BinDir`'s recursive fallback can non-deterministically pick a stale/duplicate DLL

**File:** `scripts/install.ps1:177-186`

**Issue:** When `-Source` points at a project root rather than directly at a `bin\` output folder, `Find-BinDir` falls back to `Get-ChildItem -Path $Root -Recurse -Filter $DllName -File | Select-Object -First 1`. If both a `bin\...\FinanceFmtTools.ComAddin.dll` and an `obj\...\FinanceFmtTools.ComAddin.dll` exist (common for SDK-style .NET Framework builds), the file actually selected depends on filesystem enumeration order rather than an explicit `bin`-first preference, which could install a stale intermediate build artifact instead of the real output.

**Fix:** Prefer paths containing `\bin\` when multiple matches are found, e.g. `... | Sort-Object { $_.FullName -notmatch '\\bin\\' } | Select-Object -First 1`.

---

_Reviewed: 2026-07-11T17:50:10Z_
_Reviewer: Claude (gsd-code-reviewer)_
_Depth: standard_
