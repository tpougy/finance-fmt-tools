# Phase 1: Format Engine Core - Context

**Gathered:** 2026-07-10
**Status:** Ready for planning
**Mode:** Auto-generated (discuss skipped via workflow.skip_discuss)

<domain>
## Phase Boundary

The format engine (equivalent to VBA's `modFormatEngine.bas`) exists as pure C#, with zero Excel/COM references, and its output is proven byte-for-byte correct against the VBA original via automated tests — buildable and testable using only the `dotnet` CLI.

</domain>

<decisions>
## Implementation Decisions

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions (VBA source in `src/modFormatEngine.bas`, `src/modConfig.bas`) to guide decisions. Preserve the VBA behavior exactly: the `AccountingFmt` 3-section format logic (positive;negative;zero), `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` semantics, and all `FMT_*` format keys must have a 1:1 C# equivalent. Since the .NET Framework 4.8 target cannot execute tests on this Linux dev environment, prefer a project layout where the pure-logic library and its test project can also multi-target `net8.0` (or the test project targets `net8.0` only) so `dotnet test` actually runs here — net48 build correctness can still be verified via `dotnet build` using the `Microsoft.NETFramework.ReferenceAssemblies` NuGet package cross-platform.

</decisions>

<code_context>
## Existing Code Insights

Source of truth for behavior parity: `src/modFormatEngine.bas` (format registry `GetFormatDef`, `ApplyFormat`, `AccountingFmt`), `src/modConfig.bas` (format-key constants, `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`). No C# code exists yet in this repo — this phase creates the initial solution/project structure.

</code_context>

<specifics>
## Specific Ideas

No specific requirements — discuss phase skipped. Refer to ROADMAP phase description and success criteria (FMT-01 through FMT-05, FMT-07, DEV-01 in REQUIREMENTS.md).

</specifics>

<deferred>
## Deferred Ideas

None — discuss phase skipped.

</deferred>
