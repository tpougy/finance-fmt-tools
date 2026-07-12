---
phase: quick-260712-hbu
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
    - "Ao rodar install.ps1 numa máquina com %APPDATA%\\Microsoft\\AddIns\\FinanceFmtTools.xlam existente, o script detecta e remove essa instalação legada ANTES de localizar/copiar os binários C# (PASSO 1) e ANTES do registro HKCU (PASSO 2)"
    - "Se a automação COM do Excel encontrar um add-in registrado com Title 'Finance Fmt Tools', ele é desregistrado (.Installed = $false) antes do arquivo .xlam ser apagado do disco"
    - "O arquivo .xlam legado é removido do disco independentemente do resultado da automação COM (desregistro bem-sucedido, add-in não encontrado, ou falha total da automação COM)"
    - "Se a automação COM falhar completamente (Excel não instalado, erro COM), o script apenas avisa via Write-Warn2 e continua — a instalação da versão C# NUNCA é bloqueada por essa falha"
    - "Se não existir instalação legada (.xlam ausente), o script NUNCA abre o Excel — retorna cedo sem automação COM"
    - "O relatório final (PASSO 4) menciona a remoção da instalação legada apenas quando ela de fato ocorreu ($script:VbaRemoved -eq $true)"
    - "Nenhum objeto COM (Excel.Application, Workbook, itens enumerados da coleção AddIns) fica sem ReleaseComObject em nenhum caminho de execução, incluindo caminhos de erro"
  artifacts:
    - path: "scripts/install.ps1"
      provides: "Constantes de identidade do add-in VBA legado, função Remove-LegacyVbaAddin, call site no PASSO 0, bloco condicional no relatório do PASSO 4, e comment-based help atualizado"
  key_links:
    - from: "PASSO 0 (Pré-instalação), logo após Assert-ExcelNotRunning"
      to: "function Remove-LegacyVbaAddin"
      via: "chamada direta 'Remove-LegacyVbaAddin' antes da resolução dos binários C# (PASSO 1)"
      pattern: "Remove-LegacyVbaAddin"
    - from: "Remove-LegacyVbaAddin (seta $script:VbaRemoved = $true ao remover o arquivo)"
      to: "PASSO 4 — Relatório final"
      via: "variável script-scoped compartilhada lida em 'if ($script:VbaRemoved)'"
      pattern: "\\$script:VbaRemoved"
---

<objective>
Adicionar a `scripts/install.ps1` a capacidade de detectar uma instalação legada da versão VBA
(`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`), desregistrá-la do Excel via automação COM
e removê-la do disco, executando essa migração automaticamente ANTES do fluxo normal de instalação
da versão C# (PASSO 1/2 já existentes) — sem nunca bloquear a instalação C# caso a automação COM
falhe.

Purpose: usuários que ainda têm o add-in VBA antigo instalado (distribuído nas releases `v1.0.0`/
`v1.0.1`, hoje arquivado em `archive/vba-legacy`) devem poder rodar o instalador C# uma única vez e
ter a versão antiga limpa automaticamente, sem passo manual extra e sem risco de duas versões do
add-in "Finance Fmt" convivendo na Ribbon ao mesmo tempo.

Output: `scripts/install.ps1` com (1) três novas constantes de identidade do add-in VBA legado,
(2) a função `Remove-LegacyVbaAddin` com automação COM robusta e nunca-bloqueante, (3) uma chamada a
essa função no início do PASSO 0, (4) um bloco condicional no relatório final do PASSO 4, e (5)
comment-based help (`.SYNOPSIS`/`.DESCRIPTION`) atualizado para documentar a nova capacidade.
</objective>

<execution_context>
@$HOME/.claude/get-shit-done/workflows/execute-plan.md
@$HOME/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@.planning/STATE.md
@scripts/install.ps1
@scripts/uninstall.ps1

# NÃO modificar scripts/uninstall.ps1 nem scripts/verify-environment.ps1 — escopo restrito a
# scripts/install.ps1. uninstall.ps1 é referência de estilo (Write-Step/Write-Ok/Write-Info/
# Write-Warn2/Write-Err2, Assert-ExcelNotRunning) já usada por install.ps1 — reaproveitar as
# mesmas convenções de output/logging, não inventar novas.

# Instalador VBA legado real (arquivado, fora da árvore de trabalho atual). Referência canônica de
# como a versão VBA foi instalada — ler sob demanda com:
#   git show archive/vba-legacy:Install-FinanceFmtTools.ps1
# Pontos relevantes já extraídos e traduzidos em prosa nas <action> abaixo: usa
# `New-Object -ComObject Excel.Application` + `.Workbooks.Add()` (necessário para acessar
# `$excel.AddIns`), itera com `for ($i = 1; $i -le $excel.AddIns.Count; $i++)` +
# `$excel.AddIns.Item($i)`, casa por `.Title -eq 'Finance Fmt Tools'`, desativa com
# `.Installed = $false`, e faz cleanup em `finally` com `$wb.Close($false)`, `$excel.Quit()`,
# `[System.Runtime.InteropServices.Marshal]::ReleaseComObject(...)` em cada objeto COM (incluindo
# os itens não-casados iterados da coleção AddIns) e `[GC]::Collect(); [GC]::WaitForPendingFinalizers()`.
</context>

<tasks>

<task type="auto">
  <name>Task 1: Adicionar constantes do add-in VBA legado e a função Remove-LegacyVbaAddin</name>
  <files>scripts/install.ps1</files>
  <action>
    Constantes (canonical reference): logo após o bloco "Identidade fixa" existente (termina na
    linha com `$OfficeVerKey = '16.0'`, antes do bloco "GitHub Releases (INST-01)"), adicionar um
    novo bloco de comentário delimitado por `# ====...====` intitulado algo como "Legado VBA (.xlam)
    — deteccao/remocao automatica antes de instalar a versao C#", seguido de três novas variáveis:
    `$VbaAddinTitle` com valor literal `'Finance Fmt Tools'` (deve bater exatamente com o document
    property Title do .xlam legado, não inventar outro valor), `$VbaAddinDir` como
    `Join-Path $env:APPDATA 'Microsoft\AddIns'`, e `$VbaXlamPath` como
    `Join-Path $VbaAddinDir 'FinanceFmtTools.xlam'` (nome de arquivo fixo, igual ao do instalador
    VBA legado).

    Flag de estado: logo após a declaração existente `$script:TempExtractDir = $null`, adicionar
    `$script:VbaRemoved = $false` com um comentário curto explicando que só vira `$true` quando uma
    instalação VBA legada foi efetivamente detectada E removida do disco (consumido no relatório
    final do PASSO 4, adicionado na Task 2).

    Função `Remove-LegacyVbaAddin`: inserir logo após o fechamento da função `Assert-ExcelNotRunning`
    já existente (função relacionada de pré-instalação) e antes de `Test-PeMachine`. Estrutura exata:

    Primeiro, um guard de saída antecipada: se `Test-Path -LiteralPath $VbaXlamPath` for falso,
    `return` imediatamente — SEM abrir o Excel, sem qualquer automação COM. Se o arquivo existir,
    logar via `Write-Info` que uma instalação legada foi detectada (incluir o `$VbaXlamPath` na
    mensagem).

    Em seguida, declarar três variáveis locais inicializadas a `$null`: `$excel`, `$wb`,
    `$foundAddin`. Todo o bloco de automação COM deve ficar dentro de um `try { } finally { }`
    externo, com um `try { } catch { }` interno (aninhado dentro do `try` externo) envolvendo TODA a
    automação COM — nunca deixe uma exceção de COM escapar para fora desta função.

    Dentro do try interno: criar `$excel = New-Object -ComObject Excel.Application`, setar
    `$excel.Visible = $false` e `$excel.DisplayAlerts = $false`, criar
    `$wb = $excel.Workbooks.Add()` (necessário para acessar a coleção `AddIns`, mesmo padrão do
    instalador VBA legado — ver contexto). Depois, iterar com um for indexado clássico
    (`for ($i = 1; $i -le $excel.AddIns.Count; $i++)`), obtendo `$ai = $excel.AddIns.Item($i)` em
    cada iteração: se `$ai.Title -eq $VbaAddinTitle`, atribuir `$foundAddin = $ai` e `break` (NÃO
    liberar este objeto ainda — será liberado no finally); caso contrário, liberar imediatamente com
    `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($ai) | Out-Null` antes de continuar o
    loop (evita acumular RCWs de add-ins não-relacionados). Após o loop: se `$foundAddin` não for
    `$null`, setar `$foundAddin.Installed = $false` e logar sucesso via `Write-Ok` (mencionar
    `$VbaAddinTitle`); caso contrário, logar via `Write-Info` que nenhum add-in registrado com esse
    Title foi encontrado (arquivo será removido mesmo assim, sem tratar isso como erro).

    No catch interno (captura qualquer exceção da automação COM acima): chamar `Write-Warn2` com uma
    mensagem explicando que não foi possível desregistrar o add-in VBA legado via Excel COM
    (interpolar `$_.Exception.Message`) e que o arquivo será removido mesmo assim — NÃO relançar,
    NÃO chamar `exit`, apenas deixar a função continuar normalmente após o catch.

    No finally externo (deve rodar sempre, mesmo se o try/catch interno falhou no meio): fechar e
    liberar cada objeto COM na ordem Workbook → add-in encontrado (se houver) → Application, cada
    passo de fechamento/quit dentro de seu próprio `try { } catch { }` silencioso (sem logar, apenas
    evita que uma falha ao fechar impeça a liberação dos objetos seguintes) — ou seja: se `$wb` não
    for `$null`, tentar `$wb.Close($false)` num try/catch silencioso e então
    `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null`; se `$foundAddin`
    não for `$null`, `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($foundAddin) | Out-Null`;
    se `$excel` não for `$null`, tentar `$excel.Quit()` num try/catch silencioso e então
    `[System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null`. Ao final do
    finally, sempre chamar `[GC]::Collect()` seguido de `[GC]::WaitForPendingFinalizers()`
    (incondicional, roda mesmo se `$excel`/`$wb` nunca chegaram a ser criados).

    Depois do bloco try/finally (fora dele, ainda dentro da função, executa independentemente do
    resultado da automação COM): remover o arquivo com
    `Remove-Item -LiteralPath $VbaXlamPath -Force -ErrorAction SilentlyContinue`, logar sucesso via
    `Write-Ok` (mencionar o caminho), e setar `$script:VbaRemoved = $true`. NÃO tentar remover
    `$VbaAddinDir` (pasta do Office, não é do add-in gerenciar). NÃO usar `Set-StrictMode` (não é
    convenção deste arquivo, diferente do script VBA legado).
  </action>
  <verify>
    <automated>grep -c 'function Remove-LegacyVbaAddin' scripts/install.ps1 | grep -qx 1 && grep -q '\$VbaAddinTitle' scripts/install.ps1 && grep -q '\$VbaXlamPath' scripts/install.ps1 && grep -q '\$VbaAddinDir' scripts/install.ps1 && test "$(grep -c 'ReleaseComObject' scripts/install.ps1)" -ge 3 && grep -q 'GC\]::Collect' scripts/install.ps1 && grep -q 'WaitForPendingFinalizers' scripts/install.ps1 && grep -q '\$script:VbaRemoved = \$false' scripts/install.ps1 && printf '%s\n' 'param([string]$TargetPath)' '$content = Get-Content -Raw -LiteralPath $TargetPath' '$errs = $null' '$null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errs)' 'if ($errs -and $errs.Count -gt 0) { Write-Output ("PARSE ERRORS: {0}" -f $errs.Count); $errs | ForEach-Object { Write-Output $_.Message }; exit 1 }' 'Write-Output "SYNTAX OK"' 'exit 0' > /tmp/gsd-hbu-check-syntax.ps1 && SCRIPT_WIN=$(wslpath -w /tmp/gsd-hbu-check-syntax.ps1) && TARGET_WIN=$(wslpath -w scripts/install.ps1) && /mnt/c/windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_WIN" -TargetPath "$TARGET_WIN"</automated>
  </verify>
  <done>
    Constantes `$VbaAddinTitle`/`$VbaAddinDir`/`$VbaXlamPath` declaradas junto ao bloco de
    identidade; `$script:VbaRemoved` inicializada a `$false` junto a `$script:TempExtractDir`;
    `Remove-LegacyVbaAddin` retorna cedo (sem tocar em COM) quando o arquivo não existe; quando
    existe, abre o Excel, itera `AddIns`, desregistra o add-in casado por Title (ou loga que não
    encontrou, sem erro), nunca deixa uma exceção de COM escapar (try/catch interno sempre
    silencia), sempre libera todo objeto COM obtido em `finally` (incluindo itens não-casados da
    coleção AddIns) e sempre chama `GC.Collect`/`WaitForPendingFinalizers`, remove o arquivo e seta
    `$script:VbaRemoved = $true` somente após a remoção efetiva. A checagem de sintaxe
    (`PSParser::Tokenize` via `powershell.exe`) reporta `SYNTAX OK` sem erros de parse.
  </done>
</task>

<task type="auto">
  <name>Task 2: Chamar Remove-LegacyVbaAddin no PASSO 0, atualizar relatório final e comment-based help</name>
  <files>scripts/install.ps1</files>
  <action>
    No PASSO 0 (bloco que começa com `Write-Step 'Pré-instalação'`), imediatamente após a linha
    `Assert-ExcelNotRunning` já existente e ANTES do bloco de checagem informativa de bitness
    (comentário `# Bitness do Office...`), inserir duas linhas: `Write-Step 'Detectando instalação
    VBA legada'` seguida da chamada `Remove-LegacyVbaAddin` (sem parênteses, sem argumentos — função
    sem parâmetros). Isso garante que a detecção/remoção ocorre antes do PASSO 1 (resolução dos
    binários C#) e do PASSO 2 (registro HKCU), exatamente como especificado.

    No PASSO 4 (bloco "Relatório final", que já lista "O que foi instalado:" com os binários e
    chaves de registro, antes da seção "Próximos passos:"), inserir um bloco condicional
    `if ($script:VbaRemoved) { ... }` logo após a listagem de "O que foi instalado" e antes de
    "Próximos passos:". Dentro do bloco: uma linha em branco (`Write-Host ''`), um cabeçalho
    (`Write-Host 'Migração automática:' -ForegroundColor White`), e uma linha reportando que uma
    instalação legada VBA (.xlam) foi detectada e removida, interpolando `$VbaXlamPath` na mensagem
    no mesmo estilo das demais linhas do relatório (ex.:
    `Write-Host ("  - Instalação legada VBA (.xlam) detectada e removida: {0}" -f $VbaXlamPath)`).

    No comment-based help no topo do arquivo: atualizar `.SYNOPSIS` acrescentando uma frase
    mencionando que o script também detecta e remove automaticamente uma instalação legada da
    versão VBA (.xlam), se presente, antes de instalar a versão C#. Atualizar `.DESCRIPTION`, na
    lista numerada "FLUXO PRINCIPAL" (atualmente 5 itens: baixar zip, extrair, copiar arquivos,
    registrar HKCU, validar/limpar): renumerar os 5 itens existentes de 1-5 para 2-6, e inserir um
    novo item 1 descrevendo a nova primeira etapa — detecta uma instalação legada da versão VBA
    (`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`) e, se encontrada, desregistra-a do
    Excel via automação COM e remove o arquivo, antes de prosseguir com os passos de instalação da
    versão C# (agora itens 2-6). NÃO alterar o texto de `.PARAMETER`, `.EXAMPLE` ou `.NOTES` — esta
    migração não introduz nenhum parâmetro novo nem muda o modo de invocação do script.

    Depois de concluir as três edições acima, rode a checagem de sintaxe completa do arquivo final
    (mesmo mecanismo do `<verify>` abaixo) para confirmar que o script inteiro (constantes + função
    + call site + relatório + help) continua sendo PowerShell válido antes de considerar a task
    concluída.
  </action>
  <verify>
    <automated>test "$(grep -c 'Remove-LegacyVbaAddin' scripts/install.ps1)" -ge 2 && grep -q 'Detectando instala' scripts/install.ps1 && test "$(grep -c 'VbaRemoved' scripts/install.ps1)" -ge 2 && test "$(grep -ci 'vba' scripts/install.ps1)" -ge 5 && printf '%s\n' 'param([string]$TargetPath)' '$content = Get-Content -Raw -LiteralPath $TargetPath' '$errs = $null' '$null = [System.Management.Automation.PSParser]::Tokenize($content, [ref]$errs)' 'if ($errs -and $errs.Count -gt 0) { Write-Output ("PARSE ERRORS: {0}" -f $errs.Count); $errs | ForEach-Object { Write-Output $_.Message }; exit 1 }' 'Write-Output "SYNTAX OK"' 'exit 0' > /tmp/gsd-hbu-check-syntax.ps1 && SCRIPT_WIN=$(wslpath -w /tmp/gsd-hbu-check-syntax.ps1) && TARGET_WIN=$(wslpath -w scripts/install.ps1) && /mnt/c/windows/System32/WindowsPowerShell/v1.0/powershell.exe -NoProfile -ExecutionPolicy Bypass -File "$SCRIPT_WIN" -TargetPath "$TARGET_WIN"</automated>
  </verify>
  <done>
    `Remove-LegacyVbaAddin` é chamada exatamente uma vez no fluxo principal, dentro do PASSO 0,
    antes de qualquer resolução de binários C# (PASSO 1) ou registro HKCU (PASSO 2); o relatório
    final (PASSO 4) contém um bloco condicional que só imprime a linha de migração quando
    `$script:VbaRemoved` é `$true`; o comment-based help (.SYNOPSIS/.DESCRIPTION) documenta a nova
    capacidade de migração automática; a checagem de sintaxe completa do arquivo final
    (`PSParser::Tokenize` via `powershell.exe -ExecutionPolicy Bypass`) reporta `SYNTAX OK` sem
    erros de parse.
  </done>
</task>

</tasks>

<threat_model>
## Trust Boundaries

| Boundary | Description |
|----------|--------------|
| Local Excel COM session (`Remove-LegacyVbaAddin`) ↔ Excel process on the same machine | Not a network/remote boundary — same trust level as the already-shipped VBA-era installer's identical `Workbooks.Add()` + `AddIns` COM pattern; the file being removed is the user's own previously-installed add-in, not attacker-controlled input. |
| `install.ps1` main flow ↔ `Remove-LegacyVbaAddin` | Internal function boundary — a COM failure inside `Remove-LegacyVbaAddin` must never abort/exit the outer C# install flow. |

## STRIDE Threat Register

| Threat ID | Category | Component | Disposition | Mitigation Plan |
|-----------|----------|-----------|-------------|-----------------|
| T-quick-hbu-01 | Tampering | Excel.Application COM session opened by Remove-LegacyVbaAddin | accept | Opening `$excel.Workbooks.Add()` while a legacy add-in is already registered as a startup add-in will trigger that add-in's own auto-load VBA event — inherent to how Excel's add-in model already behaves whenever Excel starts with this add-in registered, not a new risk introduced by this script; mirrors the exact `Workbooks.Add()` pattern already used (and already shipped) by the archived VBA-era `Install-FinanceFmtTools.ps1`. |
| T-quick-hbu-02 | Denial of Service | Remove-LegacyVbaAddin COM object lifecycle (Excel.Application, Workbook, AddIns items) | mitigate | Outer `try/finally` unconditionally closes the workbook without saving, quits Excel, and calls `Marshal.ReleaseComObject` on every COM object obtained — including each non-matching `AddIns.Item()` released immediately during the enumeration loop — followed by `[GC]::Collect()`/`[GC]::WaitForPendingFinalizers()`, so no orphaned EXCEL.EXE process or leaked RCW survives any error path. |
| T-quick-hbu-03 | Denial of Service | install.ps1 main flow (PASSO 0 → PASSO 1/2) | mitigate | All Excel COM automation inside `Remove-LegacyVbaAddin` is wrapped in an inner `try/catch` that never re-throws — on any COM failure (Excel not installed, automation error) it logs `Write-Warn2` and returns normally, so the outer script (running under `$ErrorActionPreference = 'Stop'`) never aborts the C# installation because of a legacy-removal failure. |
| T-quick-hbu-04 | Repudiation / Information Disclosure | Legacy file removal (`Remove-Item` on `$VbaXlamPath`) | accept | Deletion uses `-Force -ErrorAction SilentlyContinue` on a fixed, well-known path the user's own account already owns (`%APPDATA%\Microsoft\AddIns\FinanceFmtTools.xlam`) — no path is derived from untrusted input, consistent with the rest of install.ps1's existing file operations under `$InstallDir`. |
</threat_model>

<verification>
1. `grep` structural checks confirm all required identifiers exist: `function Remove-LegacyVbaAddin` (exactly once), `Remove-LegacyVbaAddin` call site (total occurrences >= 2 across definition + call), `$VbaAddinTitle`, `$VbaAddinDir`, `$VbaXlamPath`, `$script:VbaRemoved` (set + read, >= 2 occurrences), `ReleaseComObject` (>= 3 occurrences: workbook, application, and iterated/found AddIns item), `GC]::Collect`, `WaitForPendingFinalizers`.
2. `[System.Management.Automation.PSParser]::Tokenize` (invoked via `powershell.exe -ExecutionPolicy Bypass -File`, per `<environment_note>`) reports `SYNTAX OK` with zero parse errors for the fully-edited `scripts/install.ps1`, both after Task 1 and after Task 2.
3. Manual review confirms: (a) no COM object leak on any error path — every `$excel`/`$wb`/`$foundAddin` obtained has a matching `ReleaseComObject`; (b) `$script:VbaRemoved` is only ever set to `$true` after the file is actually removed, never speculatively; (c) no undefined variable is referenced (e.g. `$foundAddin` always initialized to `$null` before the try block); (d) `Remove-LegacyVbaAddin` is called exactly once, before PASSO 1's binary resolution and PASSO 2's HKCU registration; (e) `scripts/uninstall.ps1` and `scripts/verify-environment.ps1` are untouched (`git diff --stat` shows only `scripts/install.ps1`).
4. Live-Excel end-to-end validation (creating a genuine test .xlam registered as a legacy VBA add-in via COM and confirming `install.ps1` removes it) is explicitly OUT of scope for this plan — deferred to the orchestrator per `<environment_note>`.
</verification>

<success_criteria>
- `scripts/install.ps1` detects `%APPDATA%\Microsoft\AddIns\FinanceFmtTools.xlam` and, when present, attempts to unregister it from Excel (matching by Title `'Finance Fmt Tools'`) before removing the file from disk — all before any C# binary resolution or HKCU registration.
- A total failure of Excel COM automation (Excel absent, COM error) never aborts or exits the script — only logs a warning and continues to remove the file and proceed with the C# install.
- No legacy .xlam present means zero Excel COM automation occurs (fast, side-effect-free path).
- Every COM object obtained is released in `finally`, including non-matching `AddIns` items enumerated during the search.
- PASSO 4's final report conditionally documents the legacy removal only when it happened.
- Comment-based help accurately documents the new automatic-migration capability.
- `scripts/uninstall.ps1` and `scripts/verify-environment.ps1` remain unmodified.
- Full-file PowerShell syntax check (`PSParser::Tokenize`) passes with zero errors.
</success_criteria>

<output>
Create `.planning/quick/260712-hbu-adicionar-deteccao-e-remocao-de-instalac/260712-hbu-SUMMARY.md` when done
</output>
