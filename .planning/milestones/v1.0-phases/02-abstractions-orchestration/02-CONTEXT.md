# Phase 2: Abstractions & Orchestration - Context

**Gathered:** 2026-07-11
**Status:** Ready for planning
**Mode:** Auto-generated (discuss skipped via workflow.skip_discuss)

<domain>
## Phase Boundary

The seam between business logic and real Excel COM objects (`IExcelGateway`/`IRangeHandle`) exists as interfaces, and the orchestration logic that applies a format to a selection — including the invalid-selection guard — is fully exercised by `dotnet test` using fakes, with no real Excel instance involved.

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions to guide decisions. Key constraints carried over from Phase 1: zero real `Microsoft.Office.Interop.Excel` references in the code under test in this phase (that's Phase 3's job) — `IExcelGateway`/`IRangeHandle` must be pure C# interfaces with a hand-written fake implementation for tests, still buildable/testable on this Linux dev environment via `dotnet test`. `RibbonController`'s session-state defaults must match the VBA behavior: "Alinhar à direita" off by default, "Zero contábil" on by default (see src/modConfig.bas and src/modRibbon.bas for the exact VBA semantics being ported). The FMT-06 guard clause (invalid selection → friendly warning, never throws/crashes) is this phase's other core deliverable.

</decisions>

<code_context>
## Existing Code Insights

Phase 1 delivered `FinanceFmtTools.Engine` (net48;net8.0 class library) with `FormatRegistry.TryGetFormatDef`, `AccountingFormatBuilder`, `FormatKeys`, `FormatCategory`, `CellAlignment`, `FormatDef` — see `.planning/phases/01-format-engine-core/01-SUMMARY.md` and `01-03-SUMMARY.md`. This phase's `FormatEngine` orchestrator calls `FormatRegistry.TryGetFormatDef` directly. VBA source of truth for orchestration behavior: `src/modFormatEngine.bas` (`ApplyFormat`/`ApplyFormatToSelection`), `src/modUtils.bas` (`SafeSelection` — the selection-guard pattern being ported), `src/modRibbon.bas` (Ribbon → engine wiring, `IRibbonUI` handle pattern), `src/modConfig.bas` (`CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` defaults).

</code_context>

<specifics>
## Specific Ideas

No specific requirements — discuss phase skipped. Refer to ROADMAP phase description and success criteria (FMT-06 in REQUIREMENTS.md).

</specifics>

<deferred>
## Deferred Ideas

None — discuss phase skipped.

</deferred>
