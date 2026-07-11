# Requirements: Finance Fmt Tools

**Defined:** 2026-07-10
**Core Value:** Aplicar formatos financeiros/contábeis padronizados a células do Excel com um clique — agora sobre uma base de código C# testável, com dev/build/release 100% via terminal.

## v1 Requirements

Full-parity migration from VBA to C#. No staged rollout — every requirement below is in scope for this milestone.

### Format Engine (Domain)

- [x] **FMT-01**: Botões "Fin 0D/2D/4D/8D" aplicam formato contábil idêntico ao VBA para as 16 combinações de decimals (0/2/4/8) × Alinhar à direita × Zero contábil
- [x] **FMT-02**: Botões "Pct 0,00%" e "Pct 0,0000%" aplicam o formato percentual correspondente
- [x] **FMT-03**: Botão "Spread (bps)" aplica o formato de spread em basis points
- [x] **FMT-04**: Botões "Date ISO", "Date BR" e "Date BR Longa" aplicam os formatos de data correspondentes, com meses em português independente do idioma da interface do Excel
- [x] **FMT-05**: Botões "Integer" e "Text" aplicam os formatos correspondentes
- [ ] **FMT-06**: Aplicar um formato com uma seleção inválida (Chart/Shape em vez de Range) mostra uma mensagem amigável em vez de quebrar o add-in
- [x] **FMT-07**: O format engine (equivalente ao `AccountingFmt`) tem cobertura de testes xUnit para as 16 combinações, executável via `dotnet test` sem Excel instalado

### Ribbon & Checkboxes

- [ ] **RIB-01**: A aba "Finance Fmt" aparece na Ribbon com os mesmos grupos, botões e tooltips da versão VBA
- [ ] **RIB-02**: Checkbox "Alinhar à direita" funciona durante a sessão (afeta os formatos aplicados), inicia sempre desligado ao abrir o Excel, sem persistência entre sessões
- [ ] **RIB-03**: Checkbox "Zero contábil" funciona durante a sessão (afeta os formatos aplicados), inicia sempre ligado ao abrir o Excel, sem persistência entre sessões
- [ ] **RIB-04**: Botão "Sobre" e o link de documentação funcionam a partir da Ribbon

### Build & Test

- [x] **DEV-01**: O projeto compila e roda os testes 100% via `dotnet` CLI (build/test), sem exigir Visual Studio completo

### Instalação (64-bit only)

- [ ] **INST-01**: Instalador PowerShell one-liner (`irm ... | iex`) baixa a última release do GitHub e registra o add-in via HKCU para Excel 64-bit, sem exigir admin
- [ ] **INST-02**: Script de desinstalação remove o registro HKCU e os arquivos instalados
- [ ] **INST-03**: O instalador grava a chave `DoNotDisableAddinList` para evitar que o Excel desative o add-in silenciosamente após um erro transiente (Resiliency)

### CI/CD & Release

- [ ] **REL-01**: Um push de tag `v*.*.*` dispara um workflow do GitHub Actions que compila, testa, empacota e publica a release automaticamente
- [ ] **REL-02**: Existe um runbook documentado com comandos `gh` para criar uma release manualmente (executável por uma pessoa ou por um agente de IA), sem depender do CI
- [ ] **REL-03**: Cada release publicada inclui notas de changelog descrevendo o que mudou

### Legado VBA

- [ ] **LEGACY-01**: O código-fonte VBA está preservado na branch `archive/vba-legacy` e removido do fluxo de release ativo do `main`
- [ ] **LEGACY-02**: README e instruções de instalação apontam apenas para o novo add-in em C#

## v2 Requirements

Deferred to future release. Not tracked in this milestone's roadmap.

### Possíveis expansões futuras

- **FUT-01**: Suporte a Excel 32-bit (detecção de bitness e registro condicional) — explicitamente adiado; este milestone assume Excel 64-bit
- **FUT-02**: Novos formatos/botões além dos 12 existentes
- **FUT-03**: Internacionalização além de PT-BR

## Out of Scope

Explicitly excluded. Documented to prevent scope creep.

| Feature | Reason |
|---------|--------|
| Persistência das preferências "Alinhar à direita" / "Zero contábil" | Simplificação pedida pelo usuário — removida deliberadamente nesta migração |
| VSTO / instalador ClickOnce/MSI | Exige Visual Studio completo; contrário ao fluxo VS Code + dotnet CLI desejado |
| Convivência VBA + C# em paralelo no `main` | Migração é substituição completa; VBA fica só na branch `archive/vba-legacy` |
| `NumberFormatLocal` em vez de `NumberFormat` | Quebraria formatação em Excel com UI em outro idioma; VBA já usa a abordagem invariante correta |
| Suporte a Excel 32-bit | Decisão explícita do usuário — só 64-bit neste milestone (ver FUT-01) |

## Traceability

Which phases cover which requirements. Updated during roadmap creation.

| Requirement | Phase | Status |
|-------------|-------|--------|
| FMT-01 | Phase 1 - Format Engine Core | Complete |
| FMT-02 | Phase 1 - Format Engine Core | Complete |
| FMT-03 | Phase 1 - Format Engine Core | Complete |
| FMT-04 | Phase 1 - Format Engine Core | Complete |
| FMT-05 | Phase 1 - Format Engine Core | Complete |
| FMT-06 | Phase 2 - Abstractions & Orchestration | Pending |
| FMT-07 | Phase 1 - Format Engine Core | Complete |
| RIB-01 | Phase 3 - COM Entry Point & Real Excel Integration | Pending |
| RIB-02 | Phase 3 - COM Entry Point & Real Excel Integration | Pending |
| RIB-03 | Phase 3 - COM Entry Point & Real Excel Integration | Pending |
| RIB-04 | Phase 3 - COM Entry Point & Real Excel Integration | Pending |
| DEV-01 | Phase 1 - Format Engine Core | Complete |
| INST-01 | Phase 4 - Installation & Registration | Pending |
| INST-02 | Phase 4 - Installation & Registration | Pending |
| INST-03 | Phase 4 - Installation & Registration | Pending |
| REL-01 | Phase 5 - CI/CD Pipeline & Release Runbook | Pending |
| REL-02 | Phase 5 - CI/CD Pipeline & Release Runbook | Pending |
| REL-03 | Phase 5 - CI/CD Pipeline & Release Runbook | Pending |
| LEGACY-01 | Phase 5 - CI/CD Pipeline & Release Runbook | Pending |
| LEGACY-02 | Phase 5 - CI/CD Pipeline & Release Runbook | Pending |

**Coverage:**
- v1 requirements: 20 total (corrected — Format Engine section has 7 items, FMT-01..07; previous count of 19 was an off-by-one error)
- Mapped to phases: 20
- Unmapped: 0 ✓

---
*Requirements defined: 2026-07-10*
*Last updated: 2026-07-10 after roadmap creation (traceability filled, coverage count corrected from 19 to 20)*
