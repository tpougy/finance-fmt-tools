---
phase: quick-260713-jrg
plan: 01
type: execute
wave: 1
depends_on: []
files_modified:
  - README.md
autonomous: true
requirements: []
quick_id: 260713-jrg
description: "Mover a seção de instalação do README.md para o topo do arquivo (logo após título/descrição) e adicionar uma frase sobre a migração automática de VBA (v2.1.0)"
date: 2026-07-13

must_haves:
  truths:
    - "Ao abrir README.md, `## Instalação` é a primeira seção de nível 2 do arquivo — aparece logo após o bloco de título/descrição curta do projeto (linhas 1-7 de hoje: `# Finance Format Tools`, blockquote de descrição, linha `**Plataforma:** ...`, separador `---`), antes de `## Introdução` e de qualquer outra seção (formatos, ribbon, arquitetura, desenvolvimento etc.)"
    - "A subseção `### Atualizando da versão VBA`, dentro de `## Instalação`, agora deixa explícito — logo na primeira linha após o heading — que o instalador (`scripts/install.ps1`) detecta e remove automaticamente uma instalação legada em VBA antes de instalar a versão C#, sem exigir nenhuma ação manual do usuário"
    - "Todas as 9 seções de nível 2 (`##`) e as 11 subseções de nível 3 (`###`) que existiam no README antes da edição continuam existindo depois, com o mesmo texto — nenhum conteúdo foi removido; apenas a seção `## Instalação` foi reordenada e uma única frase nova foi adicionada dentro dela"
  artifacts:
    - path: "README.md"
      provides: "Seção de instalação no topo do arquivo (antes de Introdução/formatos/ribbon/desenvolvimento), com nota sobre migração automática de VBA da v2.1.0"
      contains: "detecta automaticamente uma instalação anterior da versão VBA"
  key_links:
    - from: "README.md — `## Instalação` > `### Atualizando da versão VBA`"
      to: "scripts/install.ps1 — função `Remove-LegacyVbaAddin` (comportamento real já implementado e liberado na v2.1.0)"
      via: "frase textual no README consistente com a descrição já usada em RELEASE_NOTES.md (seção 'Finance Fmt Tools v2.1.0' > 'Migração automática da versão VBA legada')"
      pattern: "detecta automaticamente uma instalação anterior da versão VBA"
---

<objective>
Editar exclusivamente `README.md`: (1) mover a seção `## Instalação` inteira (heading + todo o conteúdo, incluindo a subseção `### Atualizando da versão VBA`) para o topo do arquivo, logo após o bloco de título/descrição curta do projeto e antes de `## Introdução` e de qualquer outra seção; (2) dentro dessa seção já relocada, adicionar uma frase nova explicando que o instalador detecta e remove automaticamente uma instalação legada em VBA antes de instalar a versão C#, sem nenhuma ação manual — capacidade já implementada e lançada na v2.1.0 (`Remove-LegacyVbaAddin` em `scripts/install.ps1`, documentada em `RELEASE_NOTES.md`).

Purpose: o README hoje enterra as instruções de instalação depois de duas seções inteiras de features/formatos (`## Introdução`, que inclui a tabela "O que o add-in faz"), forçando quem só quer instalar a rolar bastante; e o texto atual da subseção "Atualizando da versão VBA" só descreve os passos manuais antigos (pré-v2.1.0), sem mencionar que isso já é automático desde a v2.1.0 — o que pode levar o usuário a fazer manualmente algo que o instalador já faz sozinho.

Output: `README.md` com a mesma quantidade de seções e o mesmo conteúdo de hoje, apenas com `## Instalação` reordenada para o topo e uma frase nova (blockquote) dentro de `### Atualizando da versão VBA`.
</objective>

<execution_context>
@$HOME/.claude/get-shit-done/workflows/execute-plan.md
@$HOME/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@README.md

# Estado atual do README.md (linhas de hoje, antes da edição — vão mudar de número após o Task 1,
# mas o CONTEÚDO e a ORDEM RELATIVA descritos aqui são a fonte da verdade):
#
# Linhas 1-7: bloco de título — "# Finance Format Tools" (H1), um blockquote de uma linha com a
# descrição curta do projeto, uma linha em negrito "**Plataforma:** ... COM add-in", e um separador
# horizontal "---". Este bloco NÃO se move e NÃO é alterado.
#
# Linhas 9-312: exatamente 9 seções de nível 2 (##), nesta ordem, cada uma separada da vizinha por
# um separador "---" (uma linha em branco, "---", uma linha em branco) — 9 separadores "---" no
# total no arquivo hoje (1 logo após o bloco de título + 8 entre as 9 seções; nenhum depois da
# última seção, que é o fim do arquivo):
#   1. ## Introdução                    (linha 9)  — inclui "### O que o add-in faz" (tabela)
#   2. ## Família Fin xD                (linha 39) — inclui 5 subseções ### + 4 sub-subseções ####
#   3. ## Outros formatos               (linha 139) — inclui ### Percentual, ### Spread em bps,
#                                        ### Datas, ### Texto
#   4. ## Instalação                    (linha 174) — inclui ### Atualizando da versão VBA
#                                        — ESTA É A SEÇÃO A MOVER PARA O TOPO (Task 1) —
#   5. ## Referência rápida do ribbon   (linha 220)
#   6. ## Preferências de sessão        (linha 249)
#   7. ## Arquitetura do projeto        (linha 260)
#   8. ## Desenvolvimento               (linha 291)
#   9. ## Licença                       (linha 309)
#
# A seção "## Instalação" hoje (linhas 174-217) tem este conteúdo, na íntegra, nesta ordem:
#   - heading "## Instalação"
#   - parágrafo "Abra o **PowerShell** no Windows e execute o comando abaixo (não é necessário
#     ser administrador):"
#   - bloco de código ps1 com o comando de instalação (irm .../scripts/install.ps1 | iex)
#   - blockquote sobre o comando não ter checksum/assinatura
#   - parágrafo "Isso baixa a release mais recente do GitHub, registra o add-in COM inteiramente em
#     `HKCU` (nenhuma chave em `HKLM`, nenhum `regasm`) e valida a instalação automaticamente."
#   - "**Requisitos:**" + lista de 3 itens (Windows+Excel, .NET Framework 4.8+, sem admin)
#   - "**Diagnóstico opcional antes de instalar**" + bloco de código ps1 (verify-environment.ps1)
#   - "**Para remover o add-in**" + bloco de código ps1 (uninstall.ps1)
#   - subseção "### Atualizando da versão VBA" com um parágrafo introdutório e uma lista numerada
#     de 4 passos manuais (Arquivo > Opções > Suplementos, etc.), terminando na frase "Depois disso,
#     rode o instalador acima normalmente."
#
# RELEASE_NOTES.md (seção "Finance Fmt Tools v2.1.0" > "Migração automática da versão VBA legada")
# descreve o comportamento real já implementado, para referência de linguagem/consistência:
# "O instalador (`scripts/install.ps1`) agora detecta sozinho uma instalação anterior da versão VBA
# (`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`), desregistra-a do Excel via automação COM
# ... e remove o arquivo do disco — tudo isso antes de instalar a versão C#, sem nenhuma ação manual
# do usuário."
#
# scripts/install.ps1 (função `Remove-LegacyVbaAddin`, chamada no fluxo principal do instalador)
# implementa exatamente esse comportamento: checa se `%APPDATA%\Microsoft\AddIns\FinanceFmtTools.xlam`
# existe: se não existir, retorna sem abrir o Excel; se existir, abre o Excel via COM, desregistra o
# add-in VBA (`Installed = $false`) e remove o arquivo do disco — tudo automaticamente, antes de
# prosseguir com a instalação da versão C#.
</context>

<tasks>

<task type="auto">
  <name>Task 1: Mover a seção "## Instalação" para o topo do README.md</name>
  <files>README.md</files>
  <action>
    Ler README.md por inteiro primeiro (Read tool, arquivo inteiro — tem 312 linhas hoje, cabe numa
    única leitura) para confirmar os limites exatos de cada seção antes de editar, já que qualquer
    edição anterior pode ter deslocado números de linha em relação ao que está descrito no bloco
    `<context>` acima (o CONTEÚDO e a ORDEM descritos lá são a fonte da verdade, não os números de
    linha).

    Recortar a seção `## Instalação` inteira: começando exatamente no heading `## Instalação` e
    terminando exatamente na última linha do seu próprio conteúdo, a frase "Depois disso, rode o
    instalador acima normalmente." (esse limite final inclui a subseção aninhada
    `### Atualizando da versão VBA` com sua lista numerada de 4 passos). Junto com esse recorte,
    remover também UM dos dois separadores `---` que hoje cercam a seção (o que vem imediatamente
    antes do heading `## Instalação`, ou o que vem imediatamente depois do fim do seu conteúdo —
    escolha qualquer um dos dois, mas remova exatamente um). Depois desse recorte, deve sobrar
    exatamente um separador `---` entre o fim do conteúdo de `## Outros formatos` (que termina na
    frase "...útil para CNPJs, códigos CETIP, CEPs e outros identificadores que começam com zero ou
    contêm caracteres que o Excel interpretaria como número ou data.") e o heading
    `## Referência rápida do ribbon` — nunca dois `---` seguidos ali, nunca zero.

    Colar a seção `## Instalação` recortada (heading + conteúdo completo, sem nenhuma alteração de
    texto nesta tarefa — isso é puro recorte-e-cola de reordenação) imediatamente depois do
    separador `---` que hoje fecha o bloco de título (o `---` que vem logo antes do heading
    `## Introdução`), fazendo `## Instalação` se tornar a primeira seção de nível 2 do arquivo,
    antes de `## Introdução`. Adicionar um novo separador `---` (linha em branco, `---`, linha em
    branco — no mesmo estilo usado em todo o resto do arquivo) logo depois do conteúdo colado de
    `## Instalação` e antes do heading `## Introdução`.

    Não alterar nenhum texto interno da seção `## Instalação` nesta tarefa (isso é feito na Task 2).
    Não tocar em nenhuma outra seção. A ordem relativa das outras 8 seções
    (Introdução, Família Fin xD, Outros formatos, Referência rápida do ribbon, Preferências de
    sessão, Arquitetura do projeto, Desenvolvimento, Licença) deve permanecer exatamente a mesma de
    hoje, apenas todas deslocadas uma posição depois da seção Instalação recolocada no topo.
  </action>
  <verify>
    <automated>cd /home/thomaz/pessoal/finance-fmt-tools &amp;&amp; diff &lt;(grep '^## ' README.md) &lt;(printf '%s\n' '## Instalação' '## Introdução' '## Família Fin xD' '## Outros formatos' '## Referência rápida do ribbon' '## Preferências de sessão' '## Arquitetura do projeto' '## Desenvolvimento' '## Licença') &amp;&amp; diff &lt;(grep '^### ' README.md) &lt;(printf '%s\n' '### Atualizando da versão VBA' '### O que o add-in faz' '### Quando usar cada formato' '### Exemplos de exibição' '### Comportamento de alinhamento (Force Align)' '### Comportamento do zero como traço (Zero Dash)' '### Strings de formato geradas' '### Percentual' '### Spread em bps' '### Datas' '### Texto') &amp;&amp; test "$(grep -c '^---$' README.md)" -eq 9 &amp;&amp; test "$(grep -c '^#### ' README.md)" -eq 4 &amp;&amp; echo TASK1_OK</automated>
  </verify>
  <done>
    `grep '^## ' README.md` retorna as 9 seções nesta ordem exata: Instalação, Introdução, Família
    Fin xD, Outros formatos, Referência rápida do ribbon, Preferências de sessão, Arquitetura do
    projeto, Desenvolvimento, Licença. `grep '^### ' README.md` retorna as 11 subseções nesta ordem
    exata, com "Atualizando da versão VBA" agora em primeiro lugar (por estar dentro de Instalação,
    que agora é a primeira seção). O arquivo continua com exatamente 9 separadores `---` e 4
    sub-subseções `####` (inalteradas, todas dentro de "Strings de formato geradas").
  </done>
</task>

<task type="auto">
  <name>Task 2: Adicionar a frase sobre migração automática de VBA e validar integridade final</name>
  <files>README.md</files>
  <action>
    Dentro da seção `## Instalação` já relocada para o topo (Task 1), localizar o heading da
    subseção `### Atualizando da versão VBA`. Imediatamente depois dessa linha de heading — antes da
    primeira linha de corpo já existente, "Se você já tinha o add-in antigo (`.xlam`, VBA)
    instalado, remova-o **antes** de instalar esta versão," — inserir um novo parágrafo em formato
    de blockquote Markdown (prefixo `&gt;`, mesmo estilo de nota lateral já usado em outros pontos
    deste README, por exemplo a nota sobre a migração VBA→C# em `## Introdução` e a nota sobre
    checksum logo abaixo do comando principal de instalação em `## Instalação`), com este texto,
    verbatim, sem parafrasear nem encurtar:

    "A partir da v2.1.0, o instalador (`scripts/install.ps1`) detecta automaticamente uma instalação
    anterior da versão VBA (`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`), desregistra-a do
    Excel via automação COM e remove o arquivo do disco — tudo isso antes de instalar a versão C#,
    sem nenhuma ação manual do usuário. Os passos abaixo permanecem documentados apenas como
    referência, caso a remoção automática não seja possível por algum motivo."

    Formatar como uma única linha de blockquote (`&gt; ` seguido do texto acima, com os nomes de
    arquivo/caminho em crases simples como já mostrado), separada por uma linha em branco antes e
    depois, no mesmo padrão de espaçamento dos outros blockquotes do arquivo. Não alterar nenhuma
    outra palavra da subseção "Atualizando da versão VBA" (a lista numerada de 4 passos manuais e a
    frase final "Depois disso, rode o instalador acima normalmente." continuam exatamente como
    estão, agora servindo de fallback/referência). Não alterar nenhuma outra parte do README.

    Depois de inserir a frase, rodar uma checagem final de integridade em todo o arquivo (via
    Bash, comandos no bloco `&lt;verify&gt;` abaixo) para confirmar que nenhum conteúdo de nenhuma
    seção foi perdido durante as Tasks 1 e 2 — cada seção deve preservar uma frase-âncora
    característica sua, e a contagem de headings deve continuar batendo com a Task 1.
  </action>
  <verify>
    <automated>cd /home/thomaz/pessoal/finance-fmt-tools &amp;&amp; grep -qF 'detecta automaticamente uma instalação anterior da versão VBA' README.md &amp;&amp; test "$(awk '/^### Atualizando da versão VBA$/{f=1;next} f &amp;&amp; NF {print; exit}' README.md)" = '&gt; A partir da v2.1.0, o instalador (`scripts/install.ps1`) detecta automaticamente uma instalação anterior da versão VBA (`FinanceFmtTools.xlam` em `%APPDATA%\Microsoft\AddIns`), desregistra-a do Excel via automação COM e remove o arquivo do disco — tudo isso antes de instalar a versão C#, sem nenhuma ação manual do usuário. Os passos abaixo permanecem documentados apenas como referência, caso a remoção automática não seja possível por algum motivo.' &amp;&amp; test "$(grep -c '^## ' README.md)" -eq 9 &amp;&amp; test "$(grep -c '^### ' README.md)" -eq 11 &amp;&amp; test "$(grep -c '^#### ' README.md)" -eq 4 &amp;&amp; grep -qF 'análises de debêntures, CRI/CRA, NTN-B e FIIs' README.md &amp;&amp; grep -qF 'negativos entre parênteses, separador de milhar' README.md &amp;&amp; grep -qF 'O Excel multiplica o valor por 100 automaticamente ao detectar' README.md &amp;&amp; grep -qF 'Set-ExecutionPolicy Bypass -Scope Process -Force; irm https://raw.githubusercontent.com/tpougy/finance-fmt-tools/main/scripts/install.ps1 | iex' README.md &amp;&amp; grep -qF '☐ Alinhar à direita (force)' README.md &amp;&amp; grep -qF 'essa é uma mudança deliberada de comportamento em relação à versão VBA' README.md &amp;&amp; grep -qF 'AccountingFormatBuilder, abstrações IExcelGateway/IRangeHandle/ILog' README.md &amp;&amp; grep -qF 'dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release' README.md &amp;&amp; grep -qF 'MIT — ver [`LICENSE`](./LICENSE).' README.md &amp;&amp; test "$(git status --porcelain -- . | grep -v '^?? .planning/' | awk '{print $2}')" = 'README.md' &amp;&amp; echo TASK2_OK</automated>
  </verify>
  <done>
    A subseção `### Atualizando da versão VBA` (dentro de `## Instalação`, já no topo do arquivo)
    começa, logo após o heading, com o novo parágrafo em blockquote explicando a migração automática
    de VBA da v2.1.0 — verbatim, incluindo a menção a `scripts/install.ps1`,
    `FinanceFmtTools.xlam` e `%APPDATA%\Microsoft\AddIns` — seguido, sem alteração, pelos 4 passos
    manuais antigos e pela frase final já existente. A contagem de headings (`##`=9, `###`=11,
    `####`=4) permanece idêntica à da Task 1. Frases-âncora de todas as 9 seções (Introdução,
    Família Fin xD, Outros formatos, Instalação, Referência rápida do ribbon, Preferências de
    sessão, Arquitetura do projeto, Desenvolvimento, Licença) continuam presentes no arquivo. `git
    status --porcelain` mostra apenas `README.md` como arquivo rastreado modificado (fora de
    `.planning/`).
  </done>
</task>

</tasks>

<threat_model>
## Trust Boundaries

| Boundary | Description |
|----------|--------------|
| README.md (documentação) ↔ leitor humano que segue as instruções de instalação | O texto do README é a única fonte que um usuário final consulta antes de rodar comandos PowerShell baixados da internet (`irm ... \| iex`); documentação imprecisa ou desatualizada pode levar o usuário a rodar passos manuais desnecessários (ou a pular um passo que ainda seria necessário). |

## STRIDE Threat Register

| Threat ID | Category | Component | Disposition | Mitigation Plan |
|-----------|----------|-----------|-------------|-----------------|
| T-quick-jrg-01 | Tampering (conteúdo/documentação) | README.md — reordenação da seção `## Instalação` | mitigate | Task 1 usa checagens automatizadas (`diff` da lista exata de headings `##`/`###`, contagem de separadores `---` e de headings `####`) para garantir que a reordenação não alterou nem removeu nenhum conteúdo — apenas mudou a posição de uma seção. |
| T-quick-jrg-02 | Information Disclosure/Repudiation (documentação desatualizada) | `### Atualizando da versão VBA` — texto legado descrevendo apenas os passos manuais pré-v2.1.0 | mitigate | Task 2 insere uma frase precisa e verbatim (extraída de `RELEASE_NOTES.md`/`scripts/install.ps1`, não inventada) deixando claro que a remoção da instalação VBA legada agora é automática, evitando que o usuário execute manualmente um passo que o instalador já faz sozinho. |
| T-quick-jrg-03 | Tampering (perda de conteúdo por erro de edição) | README.md — arquivo inteiro | accept | Escopo é só documentação (nenhum código executável é alterado); risco residual coberto pela checagem final de frases-âncora de todas as 9 seções na Task 2, que falha o gate automatizado se qualquer seção perder conteúdo característico. |

</threat_model>

<verification>
1. Task 1's automated check confirms `## ` heading order is exactly `Instalação, Introdução, Família Fin xD, Outros formatos, Referência rápida do ribbon, Preferências de sessão, Arquitetura do projeto, Desenvolvimento, Licença`, `### ` heading order starts with `Atualizando da versão VBA` followed by the original 10, `---` separator count stays at 9, and `#### ` count stays at 4.
2. Task 2's automated check confirms the new blockquote sentence about automatic VBA migration is the first non-blank line after the `### Atualizando da versão VBA` heading (exact string match), heading counts are unchanged from Task 1, nine distinct anchor phrases (one per original top-level section) are all still present verbatim in the file, and `git status --porcelain` shows only `README.md` as a modified tracked file.
3. Manual spot-check (optional, not gating): open README.md and confirm visually that `## Instalação` now reads naturally right after the title block, and that the "Atualizando da versão VBA" subsection makes sense with the new sentence followed by the legacy manual steps as fallback.
</verification>

<success_criteria>
- `README.md` is the only file modified (`git status --porcelain` shows no other tracked file changed).
- `## Instalação` is the first `##` section in the file, appearing immediately after the title/short-description block and before `## Introdução`.
- Every section and subsection that existed before the edit (9 `##`, 11 `###`, 4 `####`) still exists after the edit, with unchanged text, except for the one new sentence added in Task 2.
- The `### Atualizando da versão VBA` subsection now states, in its own words sourced from `RELEASE_NOTES.md`/`scripts/install.ps1`, that the installer automatically detects and removes a legacy VBA installation before installing the C# version, with no manual step required.
- No installation command, requirement, or step was invented — only existing README content was relocated, plus the one clarifying sentence.
</success_criteria>

<output>
Create `.planning/quick/260713-jrg-ajustar-readme-md-mover-a-se-o-de-instru/260713-jrg-SUMMARY.md` when done
</output>
