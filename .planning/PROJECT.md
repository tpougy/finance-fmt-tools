# Finance Fmt Tools

## What This Is

Add-in do Excel "Finance Fmt" que adiciona uma aba na Ribbon com atalhos de formatação (contábil, percentual, data, texto) para uso em planilhas financeiras. A implementação foi migrada de VBA (`.xlam`) para um add-in COM em C# (.NET Framework 4.8), preservando integralmente a experiência da Ribbon para o usuário final — mesmos grupos, mesmos botões, mesmos atalhos — com um fluxo de desenvolvimento e release moderno inspirado no projeto irmão `outlook-classic-delay-send`: testes automatizados (xUnit), instalador HKCU sem admin, pipeline de release via GitHub Actions e runbook `gh` CLI para release manual.

## Core Value

Aplicar formatos financeiros/contábeis padronizados a células do Excel com um clique — agora sobre uma base de código C# testável, com dev/build/release 100% via terminal (VS Code + dotnet CLI), sem depender de Visual Studio completo.

## Requirements

### Validated

- ✓ Add-in reimplementado em C# como COM add-in puro (`IDTExtensibility2` + `IRibbonExtensibility` + Ribbon XML), sem VSTO — v1.0
- ✓ Ribbon tab "Finance Fmt" com grupos Numérico/Percentual/Data/Texto/Info, mesmos botões/tooltips do VBA — v1.0 (código completo; smoke test em Excel real ainda `human_needed`)
- ✓ Paridade de funcionalidade com todos os 11 formatos do VBA atual (Fin 0D/2D/4D/8D, % 2D/4D, Spread bps, ISO/BR/BR Extenso, Texto) — v1.0, 40/40 testes xUnit passando
- ✓ Guarda de seleção inválida (Chart/Shape) mostra mensagem amigável em vez de quebrar o add-in (FMT-06) — v1.0
- ✓ Checkboxes "Alinhar à direita" (off por padrão) e "Zero contábil" (on por padrão) funcionam durante a sessão, sem persistência entre aberturas do Excel — v1.0
- ✓ Projeto compilável 100% via `dotnet` CLI (build/test), sem exigir Visual Studio completo — v1.0
- ✓ Testes automatizados (xUnit) cobrindo o format engine com abstrações (`IExcelGateway`/`IRangeHandle`) que isolam a API real do Excel — v1.0
- ✓ Instalador/desinstalador PowerShell (`scripts/install.ps1`/`uninstall.ps1`) que baixa a release do GitHub e registra o add-in via HKCU, sem admin, com proteção `DoNotDisableAddinList` contra soft-disable — v1.0 (código completo; execução em Windows+Excel real ainda `human_needed`)
- ✓ Pipeline de CI (GitHub Actions, `.github/workflows/release.yml`) disparado por tag `v*.*.*` que compila, testa, empacota e publica a release automaticamente — v1.0, disparo real confirmado em `windows-latest`: tag `v2.0.0` publicada em 2026-07-12, workflow verde (restore/build/test/package/release), asset `FinanceFmtTools.zip` verificado (7/7 arquivos, zip íntegro)
- ✓ Runbook (`RELEASE.md`) + comandos `gh` documentados para criar releases manualmente, com changelog (`RELEASE_NOTES.md`) por release — v1.0
- ✓ Código VBA legado arquivado na branch `archive/vba-legacy`, removido do fluxo principal (`main`) — v1.0, branch publicada em `origin` em 2026-07-12

### Active

<!-- Nada definido ainda para o próximo milestone. Candidatos abaixo vêm do que já estava documentado como fora de escopo do v1.0 (REQUIREMENTS.md v2). -->

- [ ] Rodar os 2 checklists `human_needed` restantes (live-Excel smoke test, live install/uninstall) em uma máquina Windows+Excel real e registrar o resultado
- [ ] Suporte a Excel 32-bit (detecção de bitness e registro condicional) — hoje o instalador só alerta, nunca bloqueia
- [ ] Novos formatos/botões além dos 11 existentes
- [ ] Internacionalização além de PT-BR

### Out of Scope

- Persistência das preferências de "Alinhar à direita" / "Zero contábil" — removida deliberadamente na migração v1.0 (simplificação pedida pelo usuário); reavaliar apenas se usuários reais sentirem falta
- VSTO / instalador ClickOnce/MSI — exige Visual Studio completo, contrário ao fluxo VS Code + dotnet CLI desejado; ainda válido
- Convivência VBA + C# em paralelo — a migração v1.0 foi uma substituição completa; o VBA fica arquivado só na branch `archive/vba-legacy`; ainda válido
- Suporte a Excel 32-bit no v1.0 — decisão explícita do usuário; reaberto como candidato de v2 acima (não é mais "fora de escopo para sempre", só para este milestone)

## Context

- **Codebase atual**: solução C# com 3 projetos (`FinanceFmtTools.Engine` net48+net8.0, `FinanceFmtTools.ComAddin` net48, `FinanceFmtTools.Engine.Tests` net8.0/xUnit), 40 testes passando, 0 Warnings/0 Errors em `dotnet build`. Add-in COM (`Connect.cs`/`AddInHost.cs`) implementa `IDTExtensibility2`+`IRibbonExtensibility`, GUID fixo `881EFDF3-424C-4240-BCA0-714DAC2B9CD7`/ProgId `FinanceFmtTools.Connect`.
- **Instalação**: `scripts/install.ps1`/`uninstall.ps1`/`verify-environment.ps1` — registro 100% HKCU, sem admin, sem `regasm`.
- **Release**: `.github/workflows/release.yml` (tag-triggered, `windows-latest`) + `RELEASE.md` (runbook manual `gh` CLI) + `RELEASE_NOTES.md` (changelog). Asset fixo `FinanceFmtTools.zip`.
- **VBA legado**: preservado integralmente na branch `archive/vba-legacy` (local, tip `cf2559b`), removido de `main`. `src/customUI14.xml` é a única exceção — continua ativo, embutido como `EmbeddedResource` no projeto C#.
- **Estado do repositório real**: `tpougy/finance-fmt-tools` é um repositório público real. Teve 2 releases VBA (`v1.0.0`/`v1.0.1`) e, em 2026-07-12, a primeira release C# (`v2.0.0`) — `main` e `archive/vba-legacy` publicados em `origin`, tag `v2.0.0` disparou `.github/workflows/release.yml` com sucesso em `windows-latest`. Push/release autorizado explicitamente pelo usuário.
- **Ambiente de desenvolvimento**: todo este milestone foi executado em um sandbox Linux/WSL sem Windows/Excel — 3 dos 5 fases (COM entry point, instalação, release) têm itens `human_needed` explícitos e documentados (ver STATE.md Deferred Items e `.planning/v1.0-MILESTONE-AUDIT.md`), nenhum foi simulado ou assumido como passando.
- **Projeto de inspiração**: `~/pessoal/outlook-classic-delay-send` — forneceu o template quase verbatim para `Connect.cs`, `install.ps1`/`uninstall.ps1`, e `release.yml`/`RELEASE.md`, adaptado de Outlook para Excel.

## Constraints

- **Plataforma**: Windows + Excel 2016+ — Why: manter compatibilidade com a base de usuários existente do add-in VBA
- **Tooling**: Desenvolvimento via VS Code + dotnet CLI, sem depender de Visual Studio completo — Why: pedido explícito do usuário, replicando o fluxo do `outlook-classic-delay-send`
- **Runtime**: .NET Framework 4.8, buildável com .NET 8 SDK — Why: COM interop com Excel exige .NET Framework clássico; o SDK moderno é só a ferramenta de build
- **Instalação**: Registro do add-in só em HKCU, sem exigir admin — Why: mesma UX do instalador atual e do projeto de inspiração
- **Compatibilidade de UX**: Mesmos botões/atalhos visíveis na Ribbon — Why: usuários já acostumados com o add-in atual não devem perceber regressão

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| COM add-in puro (IDTExtensibility2 + Ribbon XML) em vez de VSTO | VSTO exige Visual Studio completo + ClickOnce/MSI; foge do fluxo VS Code + dotnet CLI pedido | ✓ Good — shipped em `Connect.cs`, build 100% via `dotnet` |
| Remover persistência de "Alinhar à direita" / "Zero contábil" | Simplificação pedida pelo usuário — evita complexidade de config store para 2 flags de baixo valor | ✓ Good — `RibbonSessionConfig` é in-memory puro, sem `CustomXMLPart` |
| "Alinhar à direita" = off por padrão, "Zero contábil" = on por padrão | Definido explicitamente pelo usuário para a versão sem persistência | ✓ Good — shipped exatamente assim, deliberadamente diferente dos 2 defaults contraditórios do VBA original |
| CI completo (GitHub Actions) + runbook `gh` CLI para release | Usuário quer automação via tag push E um fluxo manual executável por IA, replicando o `outlook-classic-delay-send` | ✓ Good — `release.yml` + `RELEASE.md` shipped; disparo real ainda pendente (ver abaixo) |
| VBA legado arquivado na branch `archive/vba-legacy` | Preserva histórico/código sem manter dois fluxos de release ativos | ✓ Good |
| Aprovar 2 pacotes NuGet não-oficiais (`Microsoft.Office.Interop.Excel`/`MicrosoftOfficeCore16`, republicados por `CamronBute`) para referenciar tipos COM do Office sem instalação completa do Office/VSTO | Nenhum pacote oficial da Microsoft existe para isso; conteúdo verificado como genuíno via inspeção binária (`strings`/`.nupkg`) antes da aprovação autônoma | ✓ Good — build 100% verde, decisão documentada em STATE.md para revisão do usuário |
| Nunca executar `git push origin main`/`archive/vba-legacy`, tag real, ou `gh release create` de forma autônoma contra o remoto real, sem autorização explícita | Repositório é público e real; publicar ~100 commits e cortar uma release real são ações externamente visíveis e difíceis de reverter — exigem autorização humana explícita, mesmo dentro de um fluxo "100% autônomo" | ✓ Good — usuário autorizou explicitamente em 2026-07-12 ("Execute os testes e crie você mesmo o release"); push + tag `v2.0.0` + release executados e verificados |

## Evolution

This document evolves at phase transitions and milestone boundaries.

**After each phase transition** (via `/gsd-transition`):
1. Requirements invalidated? → Move to Out of Scope with reason
2. Requirements validated? → Move to Validated with phase reference
3. New requirements emerged? → Add to Active
4. Decisions to log? → Add to Key Decisions
5. "What This Is" still accurate? → Update if drifted

**After each milestone** (via `/gsd-complete-milestone`):
1. Full review of all sections
2. Core Value check — still the right priority?
3. Audit Out of Scope — reasons still valid?
4. Update Context with current state

---
*Last updated: 2026-07-11 after v1.0 milestone*
