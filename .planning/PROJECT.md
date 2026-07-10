# Finance Fmt Tools

## What This Is

Add-in do Excel "Finance Fmt" que adiciona uma aba na Ribbon com atalhos de formatação (contábil, percentual, data, texto) para uso em planilhas financeiras. Hoje é implementado em VBA (`.xlam`), distribuído via GitHub Releases com um instalador PowerShell. Este milestone migra a implementação para C# (COM add-in), preservando a experiência da Ribbon para o usuário final, com um fluxo de desenvolvimento e release moderno inspirado no projeto irmão `outlook-classic-delay-send`.

## Core Value

Aplicar formatos financeiros/contábeis padronizados a células do Excel com um clique — agora sobre uma base de código C# testável, com dev/build/release 100% via terminal (VS Code + dotnet CLI), sem depender de Visual Studio completo.

## Requirements

### Validated

<!-- Inferido do mapeamento do codebase VBA atual (.planning/codebase/). -->

- ✓ Ribbon tab "Finance Fmt" com grupos Numérico/Percentual/Data/Texto/Info — existing
- ✓ Botões de formatação contábil: Fin 2D/4D/8D — existing
- ✓ Botões de percentual: Pct 2D/4D, e Spread (bps) — existing
- ✓ Botões de data: Date ISO, Date BR, Date BR Long — existing
- ✓ Botões utilitários: Integer, Text — existing
- ✓ Checkboxes "Alinhar à direita" (ForceAlign) e "Zero contábil" (ZeroDash) alterando o formato aplicado — existing
- ✓ Botão "Sobre" / link de documentação — existing
- ✓ Instalador PowerShell que baixa a última release do GitHub e registra o add-in no Excel — existing
- ✓ Distribuição via GitHub Releases (binário como asset, sem admin) — existing

### Active

<!-- Escopo deste milestone: migração VBA → C#. -->

- [ ] Add-in reimplementado em C# como COM add-in puro (IDTExtensibility2 + Ribbon XML), sem VSTO
- [ ] Paridade de funcionalidade com todos os botões/formatos do VBA atual
- [ ] Checkboxes "Alinhar à direita" e "Zero contábil" continuam funcionando durante a sessão, mas sem persistência entre aberturas do Excel — "Alinhar à direita" inicia desligado, "Zero contábil" inicia ligado
- [ ] Projeto compilável 100% via `dotnet` CLI (sem exigir Visual Studio completo), desenvolvimento em VS Code
- [ ] Testes automatizados (xUnit) cobrindo a lógica de negócio (format engine), com abstrações que isolam a API do Excel
- [ ] Instalador PowerShell (`.ps1`) que baixa a release do GitHub e registra o add-in via HKCU (sem admin) — inspirado no `Install-OutlookUndoSend.ps1`
- [ ] Pipeline de CI (GitHub Actions) disparado por tag `v*.*.*` que compila, testa, empacota e publica a release automaticamente
- [ ] Runbook + comandos `gh` documentados para criar releases manualmente (executável por um agente de IA), com changelog por release
- [ ] Código VBA legado arquivado na branch `archive/vba-legacy`, removido do fluxo principal (`main`)

### Out of Scope

- Persistência das preferências de "Alinhar à direita" / "Zero contábil" — removida deliberadamente nesta migração (simplificação pedida pelo usuário)
- VSTO / instalador ClickOnce/MSI — exige Visual Studio completo, contrário ao fluxo VS Code + dotnet CLI desejado
- Convivência VBA + C# em paralelo — a migração é uma substituição completa; o VBA fica arquivado só na branch `archive/vba-legacy`
- Novos formatos de número ou funcionalidades além das já existentes no VBA — este milestone troca a implementação, não adiciona features

## Context

- Codebase VBA mapeado em `.planning/codebase/` (`STACK.md`, `ARCHITECTURE.md`, `STRUCTURE.md`, `INTEGRATIONS.md`): 4 módulos `.bas` + 1 Ribbon XML, arquitetura em camadas (Ribbon UI → Callbacks → Format Engine → Config/Utils), estado persistido via `CustomXMLPart` dentro do próprio `.xlam`.
- Projeto de inspiração: `~/pessoal/outlook-classic-delay-send` — add-in COM para Outlook Classic em C# 9 / .NET Framework 4.8 (buildado com .NET 8 SDK), arquitetura em camadas (`Abstractions/Domain/Services/Ui`), 25 testes xUnit, registro HKCU sem admin, instalador PowerShell one-liner, pipeline GitHub Actions (`windows-latest`) disparado por tag `v*.*.*`, e runbook `RELEASE.md` com comandos `gh` para release manual/assistida por IA.
- Este repositório (`finance-fmt-tools`) já é distribuído publicamente via GitHub Releases, com instalador `Install-FinanceFmtTools.ps1` existente — o novo instalador deve substituir esse fluxo mantendo a mesma facilidade de uso (`irm ... | iex`).
- Branch `archive/vba-legacy` já criada a partir do HEAD anterior a este milestone, preservando o código VBA antes do início da migração.

## Constraints

- **Plataforma**: Windows + Excel 2016+ — Why: manter compatibilidade com a base de usuários existente do add-in VBA
- **Tooling**: Desenvolvimento via VS Code + dotnet CLI, sem depender de Visual Studio completo — Why: pedido explícito do usuário, replicando o fluxo do `outlook-classic-delay-send`
- **Runtime**: .NET Framework 4.8, buildável com .NET 8 SDK — Why: COM interop com Excel exige .NET Framework clássico; o SDK moderno é só a ferramenta de build
- **Instalação**: Registro do add-in só em HKCU, sem exigir admin — Why: mesma UX do instalador atual e do projeto de inspiração
- **Compatibilidade de UX**: Mesmos botões/atalhos visíveis na Ribbon — Why: usuários já acostumados com o add-in atual não devem perceber regressão

## Key Decisions

| Decision | Rationale | Outcome |
|----------|-----------|---------|
| COM add-in puro (IDTExtensibility2 + Ribbon XML) em vez de VSTO | VSTO exige Visual Studio completo + ClickOnce/MSI; foge do fluxo VS Code + dotnet CLI pedido | — Pending |
| Remover persistência de "Alinhar à direita" / "Zero contábil" | Simplificação pedida pelo usuário — evita complexidade de config store para 2 flags de baixo valor | — Pending |
| "Alinhar à direita" = off por padrão, "Zero contábil" = on por padrão | Definido explicitamente pelo usuário para a versão sem persistência | — Pending |
| CI completo (GitHub Actions) + runbook `gh` CLI para release | Usuário quer automação via tag push E um fluxo manual executável por IA, replicando o `outlook-classic-delay-send` | — Pending |
| VBA legado arquivado na branch `archive/vba-legacy` | Preserva histórico/código sem manter dois fluxos de release ativos | ✓ Good |

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
*Last updated: 2026-07-10 after initialization*
