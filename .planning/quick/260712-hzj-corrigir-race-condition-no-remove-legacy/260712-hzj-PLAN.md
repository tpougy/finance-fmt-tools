---
phase: quick-260712-hzj
plan: 01
type: execute
wave: 1
depends_on: []
files_modified:
  - scripts/install.ps1
autonomous: true
requirements: []

must_haves:
  truths:
    - "Depois que Remove-LegacyVbaAddin retorna, o processo EXCEL.EXE que ela mesma abriu já terminou (ou o loop de espera de ~15s expirou), eliminando a race condition em que o PASSO 2 (Assert-ExcelNotRunning) ainda encontra o processo em fase de encerramento"
    - "A correção é puramente interna ao bloco finally de Remove-LegacyVbaAddin — nenhuma mensagem Write-* nova é exibida ao usuário para esse loop de espera"
    - "Se o processo EXCEL.EXE não desaparecer dentro do timeout de ~15s, a função simplesmente sai do loop e retorna normalmente — sem logar, sem lançar erro, sem chamar exit — deixando o Assert-ExcelNotRunning do PASSO 2 reportar o problema de forma apropriada caso ainda persista"
    - "Assert-ExcelNotRunning (função separada, mais acima no arquivo) permanece byte-a-byte inalterada"
  artifacts:
    - path: "scripts/install.ps1"
      provides: "Bloco finally de Remove-LegacyVbaAddin com uma segunda chamada a [GC]::Collect() e um loop de espera (polling de Get-Process -Name 'EXCEL', até ~15s) pelo término do processo EXCEL.EXE antes da função retornar"
  key_links:
    - from: "Remove-LegacyVbaAddin — loop de espera pelo término do EXCEL.EXE dentro do finally"
      to: "PASSO 2 do fluxo principal — segunda chamada a Assert-ExcelNotRunning, logo em seguida"
      via: "sequência de execução dentro do mesmo script: a espera garante que o processo aberto por Remove-LegacyVbaAddin já tenha terminado (ou o timeout já tenha expirado) antes que Assert-ExcelNotRunning rode de novo"
      pattern: "Get-Process -Name 'EXCEL'"
---

<objective>
Corrigir uma race condition real (observada em teste ao vivo contra Excel real) em
`Remove-LegacyVbaAddin` (`scripts/install.ps1`): a função abre o Excel via COM para desregistrar um
add-in VBA legado, chama `$excel.Quit()` no `finally`, mas retorna sem esperar o processo
`EXCEL.EXE` correspondente realmente terminar — fazendo o PASSO 2 (`Assert-ExcelNotRunning`, chamado
logo em seguida) às vezes encontrar o processo ainda em fase de encerramento e falhar a instalação
inteira com "Excel ainda está aberto".

Purpose: eliminar essa falha intermitente sem alterar nenhum outro comportamento do instalador —
a correção é estritamente aditiva dentro do `finally` já existente de `Remove-LegacyVbaAddin`.

Output: `scripts/install.ps1` com uma segunda chamada a `[GC]::Collect()` logo após
`[GC]::WaitForPendingFinalizers()`, seguida de um loop de espera (mesmo estilo de
`Assert-ExcelNotRunning`) que faz polling de `Get-Process -Name 'EXCEL'` a cada 1 segundo, por até
~15 segundos, antes de `Remove-LegacyVbaAddin` retornar.
</objective>

<execution_context>
@$HOME/.claude/get-shit-done/workflows/execute-plan.md
@$HOME/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@.planning/STATE.md
@scripts/install.ps1

# Escopo estritamente restrito à função Remove-LegacyVbaAddin (scripts/install.ps1, atualmente
# linhas ~179-237). NÃO tocar em Assert-ExcelNotRunning (linhas ~139-174, logo acima), no fluxo
# principal (PASSO 0/1/2/4), nem em nenhum outro script (uninstall.ps1, verify-environment.ps1).
#
# Assert-ExcelNotRunning já usa exatamente o estilo de wait-loop a replicar aqui:
#   for ($i = 0; $i -lt 30; $i++) {
#       Start-Sleep -Seconds 1
#       $excelProcs = Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue
#       if (-not $excelProcs) { break }
#   }
# (sleep primeiro, depois checa o processo, break se sumiu) — reaproduzir esse MESMO estilo dentro
# de Remove-LegacyVbaAddin, apenas com 15 iterações em vez de 30 e sem nenhuma mensagem Write-*.
#
# Bloco atual de Remove-LegacyVbaAddin a modificar (finally, final da função, logo antes do
# Remove-Item que apaga o .xlam do disco):
#     if ($excel) {
#         try { $excel.Quit() } catch { }
#         [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
#     }
#     [GC]::Collect()
#     [GC]::WaitForPendingFinalizers()
# }
#
# Assert-ExcelNotRunning já rodou (no PASSO 0, ANTES de Remove-LegacyVbaAddin ser chamada) e
# confirmou que não havia Excel do usuário aberto — logo, qualquer EXCEL.EXE visto dentro do loop
# de espera desta função é necessariamente o que ela mesma abriu e mandou fechar via $excel.Quit().
</context>

<tasks>

<task type="auto">
  <name>Task 1: Esperar o processo EXCEL.EXE terminar antes de Remove-LegacyVbaAddin retornar</name>
  <files>scripts/install.ps1</files>
  <action>
    Dentro do bloco `finally` de `Remove-LegacyVbaAddin` (o único `finally` da função — começa
    logo após o `catch` que trata falha da automação COM), localizar a sequência final já
    existente: `[GC]::Collect()` seguido de `[GC]::WaitForPendingFinalizers()` (essas são as duas
    últimas linhas dentro do `finally`, logo antes do `}` que fecha o bloco `finally`). Não alterar
    nada antes dessas duas linhas (a liberação de `$wb`/`$foundAddin`/`$excel` via
    `ReleaseComObject` permanece exatamente como está).

    Imediatamente depois de `[GC]::WaitForPendingFinalizers()`, ainda dentro do mesmo `finally` e
    ainda antes do `}` que o fecha, adicionar duas coisas, nesta ordem:

    1. Uma segunda chamada a `[GC]::Collect()` (uma única linha, sem argumentos — mesma sintaxe da
       chamada já existente). Motivo: uma única passagem `Collect()`+`WaitForPendingFinalizers()`
       nem sempre libera todos os RCWs do Excel; uma segunda `Collect()` finaliza o que a primeira
       passagem deixou enfileirado para finalização.

    2. Um loop de espera pelo término do processo `EXCEL.EXE`, no MESMO estilo de código já usado
       em `Assert-ExcelNotRunning` (função separada, logo acima no arquivo — não modificar essa
       função, apenas replicar seu estilo aqui): um `for ($i = 0; $i -lt 15; $i++)` (15 iterações,
       não 30 — timeout de ~15s, não 30s, já que aqui é apenas para cobrir o encerramento do
       processo que a própria função abriu, não uma espera pelo usuário fechar suas planilhas).
       Dentro do loop, na mesma ordem do padrão existente: primeiro `Start-Sleep -Seconds 1`, depois
       `Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue` atribuído a uma variável local
       nova (não reaproveitar nomes de variável já usados na função, como `$excel`/`$wb`/
       `$foundAddin` — usar um nome novo, ex. `$legacyExcelProc`), e então `if (-not
       $legacyExcelProc) { break }`. Se o processo não sumir depois das 15 iterações, o loop
       simplesmente termina e a função segue seu fluxo normal — NÃO adicionar nenhum `Write-Warn2`,
       `Write-Err2`, `Write-Info` ou qualquer outra mensagem nova para esse loop, e NÃO chamar
       `exit`: isso é um detalhe interno de robustez, não uma etapa visível ao usuário, e o
       `Assert-ExcelNotRunning` do PASSO 2 (mais à frente no fluxo principal) já cobre o caso de o
       Excel realmente continuar aberto.

    Adicionar um comentário curto (1-3 linhas, no mesmo estilo de comentários em português já usado
    no restante do arquivo) logo acima do novo loop, explicando por que esse polling é seguro aqui:
    `Assert-ExcelNotRunning` já confirmou, antes desta função ser chamada, que não havia Excel do
    usuário aberto — logo, qualquer `EXCEL.EXE` visto neste ponto é necessariamente o processo que
    esta função abriu e mandou fechar via `$excel.Quit()`; sem essa espera, o PASSO 2 pode chamar
    `Assert-ExcelNotRunning` de novo logo em seguida e ainda encontrar esse processo em fase de
    encerramento (a race condition que este fix elimina).

    NÃO tocar em nenhuma outra parte do arquivo: nem em `Assert-ExcelNotRunning`, nem no restante de
    `Remove-LegacyVbaAddin` (guard de saída antecipada, automação COM no `try`/`catch` interno,
    `Remove-Item` final, `$script:VbaRemoved`), nem no fluxo principal (PASSO 0/1/2/4), nem em
    `scripts/uninstall.ps1` ou `scripts/verify-environment.ps1`.
  </action>
  <verify>
    <automated>test "$(grep -c 'GC\]::Collect()' scripts/install.ps1)" -eq 2 && test "$(grep -c "Get-Process -Name 'EXCEL'" scripts/install.ps1)" -eq 3 && grep -q 'for (\$i = 0; \$i -lt 15; \$i++)' scripts/install.ps1 && test "$(grep -c 'function Assert-ExcelNotRunning\|function Remove-LegacyVbaAddin' scripts/install.ps1)" -eq 2 && printf '%s\n' 'param([string]$TargetPath)' '$content = Get-Content -Raw -LiteralPath $TargetPath' '$errs = $null' '$null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errs)' 'if ($errs -and $errs.Count -gt 0) { Write-Output ("PARSE ERRORS: {0}" -f $errs.Count); $errs | ForEach-Object { Write-Output $_.Message }; exit 1 }' 'Write-Output "SYNTAX OK"' 'exit 0' > /tmp/gsd-hzj-check-syntax.ps1 && SCRIPT_WIN=$(wslpath -w /tmp/gsd-hzj-check-syntax.ps1) && TARGET_WIN=$(wslpath -w scripts/install.ps1) && /mnt/c/windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_WIN" -TargetPath "$TARGET_WIN"</automated>
  </verify>
  <done>
    O `finally` de `Remove-LegacyVbaAddin` chama `[GC]::Collect()` duas vezes (a original mais a
    nova, logo após `[GC]::WaitForPendingFinalizers()`), seguido de um `for ($i = 0; $i -lt 15;
    $i++)` que faz `Start-Sleep -Seconds 1` e então checa `Get-Process -Name 'EXCEL'`, saindo
    (`break`) assim que o processo desaparecer e sem logar/falhar caso o timeout expire.
    `Assert-ExcelNotRunning` permanece inalterada (mesmo texto de antes). Nenhum outro arquivo foi
    tocado. A checagem de sintaxe completa (`PSParser::Tokenize` via `powershell.exe`) reporta
    `SYNTAX OK` sem erros de parse.
  </done>
</task>

</tasks>

<threat_model>
## Trust Boundaries

| Boundary | Description |
|----------|--------------|
| `Remove-LegacyVbaAddin` (finally block) ↔ `EXCEL.EXE` process it opened | Same-machine, same-user process lifecycle boundary — the process being polled is one this function itself spawned via `New-Object -ComObject Excel.Application` earlier in the same call; not attacker-controlled input. |
| `Remove-LegacyVbaAddin` ↔ PASSO 2's `Assert-ExcelNotRunning` call | Internal sequencing boundary within the same script run — the fix's entire purpose is to make this handoff race-free. |

## STRIDE Threat Register

| Threat ID | Category | Component | Disposition | Mitigation Plan |
|-----------|----------|-----------|-------------|-----------------|
| T-quick-hzj-01 | Denial of Service | Remove-LegacyVbaAddin's finally block (process-exit wait loop) | mitigate | Second `[GC]::Collect()` after `WaitForPendingFinalizers()` plus a `for` loop (up to 15s, same polling style as `Assert-ExcelNotRunning`) on `Get-Process -Name 'EXCEL'` ensures the EXCEL.EXE process this function opened has actually exited before the function returns, eliminating the observed race with PASSO 2's `Assert-ExcelNotRunning` call. |
| T-quick-hzj-02 | Denial of Service | Wait loop silently exceeding its ~15s timeout | accept | If the process genuinely fails to exit within 15s, the loop exits without logging or failing — this is intentional per the task spec; PASSO 2's own `Assert-ExcelNotRunning` (with its own 30s wait + `CloseMainWindow()` + actionable error message) is the real safety net for a truly stuck Excel process, not this internal optimization. |
</threat_model>

<verification>
1. Structural grep checks confirm: exactly two `[GC]::Collect()` calls remain in the file (one original, one new); exactly three `Get-Process -Name 'EXCEL'` occurrences (two pre-existing in `Assert-ExcelNotRunning`, one new in `Remove-LegacyVbaAddin`); a `for ($i = 0; $i -lt 15; $i++)` loop is present; both `Assert-ExcelNotRunning` and `Remove-LegacyVbaAddin` function definitions still exist exactly once each.
2. `[System.Management.Automation.PSParser]::Tokenize` (invoked via `powershell.exe -ExecutionPolicy Bypass -File`) reports `SYNTAX OK` with zero parse errors for the fully-edited `scripts/install.ps1`.
3. Manual review confirms: (a) `Assert-ExcelNotRunning` is byte-identical to before this change (`git diff` shows no hunk touching its line range); (b) the new wait loop uses a fresh local variable name, not reusing `$excel`/`$wb`/`$foundAddin`; (c) no new `Write-*` call was introduced for the wait loop; (d) `git diff --stat` shows only `scripts/install.ps1` changed.
4. Live-Excel end-to-end validation of the actual race-condition fix (rerunning the real VBA→C# migration install scenario against a real Excel install) is explicitly OUT of scope for this plan — the orchestrator will perform that separately right after this plan completes.
</verification>

<success_criteria>
- `Remove-LegacyVbaAddin`'s `finally` block calls `[GC]::Collect()` a second time immediately after `[GC]::WaitForPendingFinalizers()`.
- A `for` loop (up to 15 iterations, `Start-Sleep -Seconds 1` then `Get-Process -Name 'EXCEL' -ErrorAction SilentlyContinue`, `break` when the process is gone) runs before the function returns, in the same style as `Assert-ExcelNotRunning`'s existing wait loop.
- No new `Write-*` message, logging, or `exit` call was added for this wait loop — it fails open (silently exits after timeout) since PASSO 2's own `Assert-ExcelNotRunning` is the actual safety net.
- `Assert-ExcelNotRunning` and every other part of `scripts/install.ps1` outside `Remove-LegacyVbaAddin`'s `finally` block remain untouched.
- `scripts/uninstall.ps1` and `scripts/verify-environment.ps1` remain unmodified.
- Full-file PowerShell syntax check (`PSParser::Tokenize`) passes with zero errors.
</success_criteria>

<output>
Create `.planning/quick/260712-hzj-corrigir-race-condition-no-remove-legacy/260712-hzj-SUMMARY.md` when done
</output>
