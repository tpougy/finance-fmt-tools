# Phase 3: COM Entry Point & Real Excel Integration - Context

**Gathered:** 2026-07-11
**Status:** Ready for planning
**Mode:** Auto-generated (discuss skipped via workflow.skip_discuss)

<domain>
## Phase Boundary

The add-in runs inside a real, live Excel session — the Ribbon tab renders with full parity to the VBA version, every button applies its format, both checkboxes behave correctly for the session, and the About/docs actions work — verified by manual smoke test, not unit tests alone.

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions to guide decisions.

### Critical environment constraint (not discretionary — must be planned around)
This phase's own success criteria explicitly require a **live Excel session** for full verification (RIB-01 through RIB-04), and this development environment is Linux/WSL with no Windows, no Excel, and no COM runtime available. There is no way to execute or interactively smoke-test real `Microsoft.Office.Interop.Excel` COM code from here. The plan MUST still produce the real, buildable C# COM add-in code (Connect.cs entry point, Ribbon XML wiring via the already-embedded `customUI14.xml` from Phase 2, real `Microsoft.Office.Interop.Excel` implementations of Phase 2's `IExcelGateway`/`IRangeHandle` interfaces) — compiled and unit-testable pieces should be proven via `dotnet build`/`dotnet test` as far as possible, but the actual "run inside live Excel and click every button" verification is expected to come back as `human_needed` and will be deferred to the user, who has a real Windows+Excel machine. Do not attempt to fake or skip this reality — plan the real COM code, build it, and clearly scope what can vs. cannot be verified in this environment.

</decisions>

<code_context>
## Existing Code Insights

Phase 1 delivered `FinanceFmtTools.Engine` with `FormatRegistry`, `AccountingFormatBuilder`, format value types. Phase 2 delivered the `IExcelGateway`/`IRangeHandle`/`ILog` interfaces, `FormatEngine.Apply`/`ApplyToSelection` (with the FMT-06 no-throw guard), `RibbonController`/`RibbonSessionConfig` (in-memory checkbox state, embedded `customUI14.xml` XML loading via suffix-match resource resolution). See `.planning/phases/02-abstractions-orchestration/02-01-SUMMARY.md` and `02-02-SUMMARY.md`. This phase's job is the real COM implementations of those interfaces plus the actual Excel Add-in entry point (VBA equivalent: `src/ThisWorkbook.bas`'s addin lifecycle, `src/modRibbon.bas`'s ribbon callbacks and `IRibbonUI` handling, `src/customUI14.xml`'s `onAction`/`getPressed` bindings). VBA source of truth for exact Ribbon button-to-callback wiring: `src/customUI14.xml` (already linked as an embedded resource in Phase 2's csproj) and `src/modRibbon.bas`.

</code_context>

<specifics>
## Specific Ideas

No specific requirements — discuss phase skipped. Refer to ROADMAP phase description and success criteria (RIB-01, RIB-02, RIB-03, RIB-04 in REQUIREMENTS.md). Note ROADMAP.md marks this phase with `**UI hint**: yes`, but `workflow.ui_phase` is disabled project-wide (this is an Excel Ribbon add-in replicating an existing, fully-specified VBA UI — not a web/app frontend needing a new design contract).

</specifics>

<deferred>
## Deferred Ideas

None — discuss phase skipped.

</deferred>
