# Roadmap: Finance Fmt Tools

## Overview

This milestone ports the "Finance Fmt" Excel Ribbon add-in from VBA to a pure C# COM add-in, layer by layer, bottom-up. Work starts with the format engine — the pure, COM-free business logic that decides which `NumberFormat` string to apply — and locks its behavior down with `dotnet test` before any Excel/COM code exists. It then builds the fakeable orchestration layer (`IExcelGateway`, `FormatEngine`, `RibbonController`) so guard-clause and selection-invalid behavior is unit-tested with fakes. Only then does real `Microsoft.Office.Interop.Excel` code, the Ribbon XML, and the COM entry point (`Connect`) get written and verified against a live Excel session. Installation/registration (HKCU-only, no admin, bitness-aware) is its own phase because it carries distinct risks (WOW64 redirection, Resiliency auto-disable) that have nothing to do with the add-in's own code. The milestone closes with the CI/CD pipeline, the `gh`-CLI release runbook, and confirmation that the VBA legacy code is fully archived off `main`.

## Phases

**Phase Numbering:**
- Integer phases (1, 2, 3): Planned milestone work
- Decimal phases (2.1, 2.2): Urgent insertions (marked with INSERTED)

Decimal phases appear between their surrounding integers in numeric order.

- [x] **Phase 1: Format Engine Core** - Pure C# format-string logic (Fin/Pct/Spread/Date/Integer/Text), fully covered by `dotnet test`, zero Excel/COM dependency
- [x] **Phase 2: Abstractions & Orchestration** - `IExcelGateway`/`IRangeHandle` seam plus `FormatEngine`/`RibbonController` orchestration, unit-tested against fakes
- [ ] **Phase 3: COM Entry Point & Real Excel Integration** - Real `Microsoft.Office.Interop.Excel` wiring, Ribbon XML, `Connect.cs`; the Ribbon tab, buttons, checkboxes, and About/docs link work in a live Excel session
- [ ] **Phase 4: Installation & Registration** - HKCU-only, no-admin PowerShell installer/uninstaller with Resiliency protection
- [ ] **Phase 5: CI/CD Pipeline & Release Runbook** - Tag-triggered GitHub Actions release pipeline, manual `gh` CLI runbook, changelog, and VBA legacy cleanup confirmation

## Phase Details

### Phase 1: Format Engine Core
**Goal**: The format engine (equivalent to VBA's `modFormatEngine.bas`) exists as pure C#, with zero Excel/COM references, and its output is proven byte-for-byte correct against the VBA original via automated tests — buildable and testable using only the `dotnet` CLI.
**Depends on**: Nothing (first phase)
**Requirements**: FMT-01, FMT-02, FMT-03, FMT-04, FMT-05, FMT-07, DEV-01
**Success Criteria** (what must be TRUE):
  1. `dotnet test` passes for all 16 accounting-format combinations (decimals in {0,2,4,8} x Alinhar à direita x Zero contábil), matching the VBA `AccountingFmt` output exactly (FMT-01, FMT-07)
  2. `dotnet test` confirms the format registry produces the correct percentual strings for "Pct 0,00%"/"Pct 0,0000%" and the correct basis-points string for "Spread (bps)" (FMT-02, FMT-03)
  3. `dotnet test` confirms "Date ISO", "Date BR", and "Date BR Longa" produce the correct format strings with Portuguese month names, independent of any Excel UI-locale setting (FMT-04)
  4. `dotnet test` confirms "Integer" and "Text" produce their corresponding format strings (FMT-05)
  5. The solution (add-in project + test project) builds and all tests run 100% via `dotnet build`/`dotnet test`, with no Visual Studio installation required (DEV-01)
**Plans**: 3 plans

Plans:
- [x] 01-01-PLAN.md — Bootstrap solution/projects (net48;net8.0 Engine + net8.0 xUnit Tests) and shared contract types (FormatKeys, FormatCategory, CellAlignment, FormatDef)
- [x] 01-02-PLAN.md — AccountingFormatBuilder: port and prove the 16-combination accounting format matrix (FMT-01, FMT-07)
- [x] 01-03-PLAN.md — FormatRegistry: wire all 11 format keys (literal Pct/Spread/Date/Text entries + Fin/Integer family via AccountingFormatBuilder) (FMT-02, FMT-03, FMT-04, FMT-05)

### Phase 2: Abstractions & Orchestration
**Goal**: The seam between business logic and real Excel COM objects (`IExcelGateway`/`IRangeHandle`) exists as interfaces, and the orchestration logic that applies a format to a selection — including the invalid-selection guard — is fully exercised by `dotnet test` using fakes, with no real Excel instance involved.
**Depends on**: Phase 1
**Requirements**: FMT-06
**Success Criteria** (what must be TRUE):
  1. `dotnet test` confirms `FormatEngine` resolves a format key via the Phase 1 registry and applies the resulting `NumberFormat`/alignment through a fake Excel gateway, with no real COM types referenced anywhere in the tested code path
  2. `dotnet test` confirms that when the current selection is not a Range (e.g. a Chart/Shape is selected), `FormatEngine` logs a warning and returns without throwing — the C# equivalent of VBA's `SafeSelection()` guard, proving the friendly-message behavior at the orchestration level before any live Excel exists (FMT-06)
  3. `dotnet test` confirms `RibbonController` loads the embedded Ribbon XML resource and answers session-state queries (checkbox pressed/unpressed) against an in-memory config object defaulting to "Alinhar à direita" off / "Zero contábil" on, using a fake rather than a live `IRibbonUI`
**Plans**: 2 plans

Plans:
- [x] 02-01-PLAN.md — IExcelGateway/IRangeHandle/ILog seam + FormatEngine.Apply/ApplyToSelection orchestration, including the FMT-06 invalid-selection guard clause
- [x] 02-02-PLAN.md — RibbonSessionConfig (RIB-02/RIB-03 authoritative defaults) + RibbonController (in-memory checkbox state, embedded customUI14.xml resource loading)

### Phase 3: COM Entry Point & Real Excel Integration
**Goal**: The add-in runs inside a real, live Excel session — the Ribbon tab renders with full parity to the VBA version, every button applies its format, both checkboxes behave correctly for the session, and the About/docs actions work — verified by manual smoke test, not unit tests alone.
**Depends on**: Phase 2
**Requirements**: RIB-01, RIB-02, RIB-03, RIB-04
**Success Criteria** (what must be TRUE):
  1. Loading the add-in in a live Excel session shows the "Finance Fmt" Ribbon tab with the same groups (Numérico/Percentual/Data/Texto/Info), buttons, and tooltips as the VBA version (RIB-01)
  2. Clicking each of the 12 format buttons in a live Excel session applies the correct number format to the selected range, and selecting a Chart/Shape then clicking a format button shows a friendly message instead of crashing the add-in — the live confirmation of the Phase 1/2 format engine and guard clause working end-to-end
  3. Toggling "Alinhar à direita" changes the alignment of subsequently-applied formats during the session, starts unchecked every time Excel opens, and never persists across Excel restarts (RIB-02)
  4. Toggling "Zero contábil" changes the accounting format's zero-display behavior during the session, starts checked every time Excel opens, and never persists across Excel restarts (RIB-03)
  5. Clicking "Sobre" on the Ribbon shows an About message with version info, and the documentation-link button opens the correct docs URL (RIB-04)
**Plans**: TBD
**UI hint**: yes

### Phase 4: Installation & Registration
**Goal**: A non-admin user can install and uninstall the add-in on 64-bit Excel with a single PowerShell command, and the installed add-in survives transient errors without being silently disabled by Excel.
**Depends on**: Phase 3
**Requirements**: INST-01, INST-02, INST-03
**Success Criteria** (what must be TRUE):
  1. Running the installer one-liner (`irm .../install.ps1 | iex`) on a 64-bit Excel machine downloads the latest GitHub release, registers the add-in entirely under `HKCU` with no admin prompt, and the "Finance Fmt" tab appears the next time Excel opens (INST-01)
  2. Running the uninstall script removes the `HKCU` registration keys and the installed files, and the Ribbon tab no longer appears after Excel restarts (INST-02)
  3. After installation, the `DoNotDisableAddinList` registry key is present for the add-in's ProgID, so a transient runtime error does not cause Excel to silently disable the add-in (INST-03)
  4. Running the installer twice in a row, or running the uninstaller when the add-in was never installed, completes without error (idempotency)
**Plans**: TBD

### Phase 5: CI/CD Pipeline & Release Runbook
**Goal**: Releases are fully automated from a tag push, a documented manual fallback exists for a person or AI agent to cut a release without CI, and the VBA legacy code/docs are completely out of the active `main` flow.
**Depends on**: Phase 4
**Requirements**: REL-01, REL-02, REL-03, LEGACY-01, LEGACY-02
**Success Criteria** (what must be TRUE):
  1. Pushing a `v*.*.*` tag triggers a GitHub Actions workflow on `windows-latest` that builds, tests, packages, and runs `gh release create` to publish a release with the add-in zip (DLL + vendored interop DLLs + installer scripts) as an asset, with no manual steps (REL-01)
  2. A documented runbook (e.g. `RELEASE.md`) with exact `gh` CLI commands lets a person or an AI agent create a release manually, without depending on the CI workflow (REL-02)
  3. Every published release includes changelog notes describing what changed in that version (REL-03)
  4. The VBA source exists only on the `archive/vba-legacy` branch and is absent from `main`'s active build/release flow (LEGACY-01)
  5. The README and installation instructions on `main` reference only the new C# add-in and its installer, with no remaining mention of the old VBA/`.xlam` flow (LEGACY-02)
**Plans**: TBD

## Progress

**Execution Order:**
Phases execute in numeric order: 1 → 2 → 3 → 4 → 5

| Phase | Plans Complete | Status | Completed |
|-------|----------------|--------|-----------|
| 1. Format Engine Core | 3/3 | Complete | 2026-07-11 |
| 2. Abstractions & Orchestration | 2/2 | Complete | 2026-07-11 |
| 3. COM Entry Point & Real Excel Integration | 0/TBD | Not started | - |
| 4. Installation & Registration | 0/TBD | Not started | - |
| 5. CI/CD Pipeline & Release Runbook | 0/TBD | Not started | - |
