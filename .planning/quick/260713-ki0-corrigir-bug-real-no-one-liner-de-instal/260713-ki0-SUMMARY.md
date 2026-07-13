---
status: complete
phase: quick-260713-ki0
plan: 01
subsystem: infra
tags: [powershell, installer, encoding, bom, ascii, one-liner, irm-iex]

# Dependency graph
requires: []
provides:
  - "scripts/install.ps1, scripts/uninstall.ps1 and scripts/verify-environment.ps1 are 100% ASCII with no UTF-8 BOM, so Invoke-RestMethod | Invoke-Expression no longer hits a parse error on the residual BOM"
affects: [installer, uninstaller, scripts/install.ps1, scripts/uninstall.ps1, scripts/verify-environment.ps1]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "ASCII-only PowerShell source files (no BOM, no byte > 127) as the fix for irm | iex parse failures — ASCII is a valid subset of every code page and of UTF-8, eliminating encoding ambiguity at the root instead of patching around it with TrimStart/BOM handling in the one-liner"

key-files:
  created: []
  modified:
    - "scripts/install.ps1 — BOM removed, 78 lines transliterated (comments, comment-based help, Write-* messages)"
    - "scripts/uninstall.ps1 — BOM removed, 43 lines transliterated"
    - "scripts/verify-environment.ps1 — BOM removed, 48 lines transliterated (consistency fix; not on the irm|iex path)"

key-decisions:
  - "Followed the plan's exact line-by-line before→after transliteration table verbatim (no re-derivation) to avoid introducing new mistakes across the 169 listed lines"
  - "Task 4's verification command as literally written in PLAN.md invokes powershell.exe via -File with three comma-joined -Targets arguments; from WSL, Win32 argv passing collapses this into a single joined argument (confirmed via a debug harness — even space-separated -File args only bind the first array element). Re-ran the same ParseFile/ParseInput logic via -Command \"& 'script' -Targets 'a','b','c'\" instead, which correctly binds all 3 files as a real string[] array. This is an invocation-only fix (Rule 3) — no target file content changed, and all three files had already independently passed Tasks 1-3's per-file ParseFile gates."

requirements-completed: []

# Metrics
duration: ~15min
completed: 2026-07-13
---

# Quick Task 260713-ki0: Corrigir bug real no one-liner de instalação Summary

**Remoção do BOM UTF-8 e transliteração para ASCII puro de scripts/install.ps1, scripts/uninstall.ps1 e scripts/verify-environment.ps1, corrigindo o parse error real do one-liner documentado `irm <url> | iex`**

## Performance

- **Duration:** ~15 min
- **Completed:** 2026-07-13
- **Tasks:** 4 (3 file-modification tasks + 1 verification-only task)
- **Files modified:** 3

## Accomplishments
- `scripts/install.ps1`, `scripts/uninstall.ps1` e `scripts/verify-environment.ps1` não têm mais o BOM UTF-8 (`EF BB BF`) inicial e são 100% ASCII (zero bytes com código > 127).
- As 169 linhas de prosa humana (comentários, blocos de comment-based help, mensagens `Write-Step`/`Write-Ok`/`Write-Info`/`Write-Warn2`/`Write-Err2`/`Write-Host`/`Write-Status`) foram transliteradas para ASCII exatamente conforme a tabela before→after do plano — nenhum nome de variável, GUID, caminho de registro, URL ou valor de lógica de código foi tocado.
- Contagem de linhas dos 3 arquivos permanece idêntica à original (583 / 193 / 283).
- Confirmado via `[System.Management.Automation.Language.Parser]::ParseInput` sobre `Get-Content -Raw` (simulando exatamente o que `Invoke-Expression` recebe depois de `Invoke-RestMethod` no fluxo `irm | iex`) que `install.ps1` e `uninstall.ps1` parseiam sem nenhum erro — a causa raiz do bug reportado pelo usuário está eliminada.

## Task Commits

Each task was committed atomically:

1. **Task 1: Remover BOM e transliterar scripts/install.ps1 para ASCII puro** - `eeec7ea` (fix)
2. **Task 2: Remover BOM e transliterar scripts/uninstall.ps1 para ASCII puro** - `60bb960` (fix)
3. **Task 3: Remover BOM e transliterar scripts/verify-environment.ps1 para ASCII puro** - `bf7a384` (fix)
4. **Task 4: Verificação cruzada final** - somente leitura/verificação, nenhum arquivo modificado, nenhum commit próprio (todas as checagens passaram, ver seção Verification Gates abaixo)

## Files Created/Modified
- `scripts/install.ps1` - BOM removido; 78 linhas transliteradas (78 listadas na tabela do plano + 1 linha 1 alterada apenas pela remoção do BOM em si, sem mudança de texto visível)
- `scripts/uninstall.ps1` - BOM removido; 43 linhas transliteradas (+ linha 1 idem)
- `scripts/verify-environment.ps1` - BOM removido; 48 linhas transliteradas (+ linha 1 idem)

## Verification Gates (all passed)

- **Tasks 1-3 (per file):** ausência de BOM (`head -c3 | xxd -p` ≠ `efbbbf`), ausência de bytes não-ASCII (`LC_ALL=C grep -P '[^\x00-\x7F]'` vazio), contagem de linhas inalterada (583/193/283), e `[System.Management.Automation.Language.Parser]::ParseFile` via `powershell.exe` (WSL2↔Windows interop) reportando `SYNTAX OK` — todos passaram individualmente.
- **Task 4 (agregado):** `git status --porcelain` (fora de `.planning/`) confirmado como mostrando somente os 3 arquivos deste plano; BOM/ASCII/contagem de linhas re-confirmados para os 3 juntos; `ParseFile` sem erros para os 3; `ParseInput` sobre `Get-Content -Raw` de `install.ps1`/`uninstall.ps1` (simulando o caminho real `irm | iex`) sem erros de parse.

## Decisions Made
- Seguida a tabela de transliteração linha-a-linha do plano exatamente como escrita, sem re-derivar traduções — evita introduzir novos erros.
- Task 4's verification command as written in PLAN.md invokes `powershell.exe -File` with `-Targets "a","b","c"` (no space between the quoted, comma-separated tokens). From WSL2, this collapses into a single comma-joined argv token before reaching PowerShell's `-File` argument binder, which only bound one array element (confirmed with a minimal debug harness, testing both the exact plan syntax and a space-separated variant — neither correctly produced a 3-element `string[]`). Re-ran the identical ParseFile/ParseInput logic via `powershell.exe -Command "& 'script' -Targets 'a','b','c'"`, which correctly binds all 3 paths as a `string[]` of length 3 and confirmed all gates pass. This is a Rule 3 (auto-fix blocking issue) fix to the verification *invocation* only — no target script content was touched, and all 3 files had already independently passed Tasks 1-3's per-file `ParseFile` gates before this was even discovered.

## Deviations from Plan

### Auto-fixed Issues

**1. [Rule 3 - Blocking] Task 4's `-File`-based multi-target PowerShell invocation only bound 1 of 3 `-Targets` array elements from WSL2**
- **Found during:** Task 4 (verificação cruzada final)
- **Issue:** The plan's Task 4 automated verification command invokes `powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$ALL_WIN" -Targets "$I1_WIN","$I2_WIN","$I3_WIN"` (and similarly for the `ParseInput` live check). Because bash concatenates the three adjacent quoted, comma-separated substrings into a single argv token (no space between them), and because `powershell.exe -File` binds `[string[]]` parameters from raw Win32 argv without re-parsing embedded commas the way the native PowerShell command-line parser does, the script's `foreach ($t in $Targets)` loop ran exactly once with `$t` equal to all three Windows paths joined by literal commas — producing `PARSEFILE ERRORS in ...` for a single bogus combined "path" rather than validating each of the 3 real files.
- **Fix:** Verified the root cause with a minimal debug harness printing `$Targets.Count` and each element, testing both the plan's exact syntax and a space-separated variant of `-File ... -Targets "$a" "$b" "$c"` (still only bound 1 element), then confirmed `-Command "& 'script.ps1' -Targets '$a','$b','$c'"` correctly binds a 3-element array. Re-ran Task 4's full gate set (git status scope, BOM/ASCII/line-count re-check, `ParseFile`, `ParseInput`) using the `-Command` invocation form with identical inline script logic to the plan's — all passed for all 3 files.
- **Files modified:** None (verification-invocation-only fix; no target `.ps1` files changed by this deviation).
- **Verification:** `PARSEFILE ALL OK` and `PARSEINPUT OK` printed for each of `install.ps1`, `uninstall.ps1`, `verify-environment.ps1` individually; `git status --porcelain` scope check passed; line counts and BOM/ASCII re-confirmed for all 3.
- **Committed in:** N/A (Task 4 is verification-only per the plan; no commit corresponds to this task).

---

**Total deviations:** 1 auto-fixed (1 blocking - verification invocation only, zero impact on shipped files)
**Impact on plan:** No scope creep. The deviation is confined to how the Task 4 verification gate was invoked from this WSL2 environment; the underlying `.ps1` files were unaffected and had already passed Tasks 1-3's independent per-file syntax gates before Task 4 ran.

## Issues Encountered
None beyond the Task 4 invocation quirk documented above.

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- The root cause of the reported bug (residual UTF-8 BOM surviving `Invoke-RestMethod` and breaking `Invoke-Expression`'s parse of `#Requires`/`[CmdletBinding()]`/`param()`) is eliminated: all 3 scripts are now BOM-free and 100% ASCII, verified via both `ParseFile` (local `-File` flow) and `ParseInput` over raw string content (simulating the `irm | iex` flow exactly).
- README.md, RELEASE.md, and the `.EXAMPLE` one-liner text inside the scripts themselves were left byte-for-byte untouched, per the plan's explicit scope boundary — `git status --porcelain` (excluding `.planning/`) confirms only the 3 target scripts were modified.
- Live end-to-end verification of the actual documented one-liner (`Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex`) against the real GitHub-hosted raw content requires this fix to be pushed/merged first — out of scope for this local plan, called out explicitly in the plan's own `<verification>` section as a follow-up for the orchestrator/user.

## Self-Check: PASSED

- FOUND: scripts/install.ps1 (modified; no BOM, no non-ASCII bytes, 583 lines, `ParseFile` → SYNTAX OK)
- FOUND: scripts/uninstall.ps1 (modified; no BOM, no non-ASCII bytes, 193 lines, `ParseFile` → SYNTAX OK)
- FOUND: scripts/verify-environment.ps1 (modified; no BOM, no non-ASCII bytes, 283 lines, `ParseFile` → SYNTAX OK)
- FOUND: commit eeec7ea (`git log --oneline` confirms it exists in history)
- FOUND: commit 60bb960 (`git log --oneline` confirms it exists in history)
- FOUND: commit bf7a384 (`git log --oneline` confirms it exists in history)
- Confirmed: `git status --porcelain -- .` (excluding `.planning/`) is empty — no other tracked file was modified.
- Confirmed: `ParseInput` over `Get-Content -Raw` of `install.ps1` and `uninstall.ps1` reports zero parse errors, simulating the exact `irm | iex` code path.

---
*Phase: quick-260713-ki0*
*Completed: 2026-07-13*
