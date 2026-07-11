---
phase: 02-abstractions-orchestration
plan: 02
subsystem: format-engine
tags: [dotnet, csharp, xunit, tdd, ribbon, embedded-resource]

# Dependency graph
requires:
  - phase: 02-abstractions-orchestration (plan 01)
    provides: "IExcelGateway/IRangeHandle/ILog seam + FormatEngine.Apply/ApplyToSelection orchestration, extending the existing FinanceFmtTools.Engine project"
provides:
  - "RibbonSessionConfig — in-memory checkbox session state with the authoritative RIB-02/RIB-03 defaults (ForceAlign=false, ZeroDash=true), no persistence"
  - "RibbonController — instance class owning a RibbonSessionConfig Config property and GetCustomUiXml() embedded-resource loader"
  - "EmbeddedResource Link in FinanceFmtTools.Engine.csproj pointing at src/customUI14.xml (no physical duplication), applying to both net48 and net8.0"
  - "4 new xUnit tests (RibbonControllerTests), 39/39 total project tests passing"
affects: [phase-3-com-entry-point]

# Tech tracking
tech-stack:
  added: []
  patterns: ["Embedded-resource loading via suffix match (EndsWith), not a hardcoded logical resource name — avoids silent resource-name drift", "Instance-based stateful service (RibbonController owning RibbonSessionConfig) vs. Phase 1's static, parameterized FormatRegistry/AccountingFormatBuilder convention"]

key-files:
  created:
    - src/FinanceFmtTools.Engine/RibbonSessionConfig.cs
    - src/FinanceFmtTools.Engine/RibbonController.cs
    - src/FinanceFmtTools.Engine.Tests/RibbonControllerTests.cs
  modified:
    - src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj

key-decisions:
  - "No REFACTOR commits for either TDD task — both GREEN implementations matched the plan's target shape (from 02-RESEARCH.md Pattern 2/3) exactly on first pass, consistent with Phase 1 and 02-01's precedent of skipping REFACTOR when nothing is genuinely warranted."
  - "RibbonSessionConfig defaults (ForceAlign=false, ZeroDash=true) deliberately do NOT match either of src/modConfig.bas's or src/modUtils.bas's two mutually contradictory VBA defaults — per 02-RESEARCH.md Pitfall 1, the authoritative source is REQUIREMENTS.md RIB-02/RIB-03, a considered behavior change for this migration."
  - "src/customUI14.xml is linked into the C# project via MSBuild EmbeddedResource Link rather than duplicated — keeps a single source of truth for the Ribbon XML (and its onAction/getPressed callback names) during the VBA-to-C# transition."

patterns-established:
  - "RibbonController.cs pattern: instance class with a get-only Config property (RibbonSessionConfig, mutable in-memory) and a GetCustomUiXml() resource loader — deliberately narrow scope (no IRibbonUI caching, no InvalidateControl, no image loading) per 02-CONTEXT.md's resolved open question; Phase 3 adds the real COM seam on top of this unchanged."

requirements-completed: []

# Metrics
duration: 15min
completed: 2026-07-11
---

# Phase 2 Plan 2: RibbonController Summary

**`RibbonSessionConfig` (in-memory checkbox state with the authoritative RIB-02/RIB-03 defaults) and `RibbonController` (owns that config, loads the embedded `customUI14.xml` Ribbon resource via suffix-matched `Assembly.GetManifestResourceNames`) — proven by 4 new xUnit tests (39/39 total) with zero `IRibbonUI`/`stdole`/`Microsoft.Office.Core`/persistence references anywhere in the tested path. This is the last plan in Phase 2 — the phase is now complete.**

## Performance

- **Duration:** 15 min
- **Started:** 2026-07-11T04:35:00Z
- **Completed:** 2026-07-11T04:50:00Z
- **Tasks:** 2 completed (2 TDD RED/GREEN pairs, no REFACTOR)
- **Files modified:** 4 (3 created, 1 modified)

## Accomplishments
- Added `RibbonSessionConfig.cs` — a plain mutable class with `ForceAlign` (default `false`) and `ZeroDash` (default `true`), matching REQUIREMENTS.md's RIB-02/RIB-03 wording exactly, not either of the two contradictory VBA defaults found in `src/modConfig.bas`/`src/modUtils.bas`. No persistence logic of any kind.
- Added `RibbonController.cs` — a sealed instance class with a get-only `Config` property, a parameterless constructor delegating to `this(new RibbonSessionConfig())`, and a constructor accepting an injected `RibbonSessionConfig` (throws `ArgumentNullException` on `null`).
- Added `RibbonController.GetCustomUiXml()` — enumerates `typeof(RibbonController).Assembly.GetManifestResourceNames()`, resolves the resource whose name `EndsWith("customUI14.xml", StringComparison.OrdinalIgnoreCase)`, and returns its full text (or `string.Empty` if not found) — suffix-match resolution rather than a hardcoded exact resource name, per 02-RESEARCH.md Pattern 3/Pitfall 3.
- Added an unconditioned `<EmbeddedResource Include="../customUI14.xml" Link="Resources/customUI14.xml" />` `ItemGroup` to `FinanceFmtTools.Engine.csproj`, embedding the real `src/customUI14.xml` (no physical duplication) for both `net48` and `net8.0`.
- Added `RibbonControllerTests.cs` with 4 `[Fact]` tests: default session-state values, in-memory mutation, constructor-injected config honored, and `GetCustomUiXml()` returning the real file contents (asserting the literal `tabFinanceFmt` tab id and `onLoad="OnRibbonLoad"` attribute are present).
- Full test project now passes 39/39 tests (35 before this plan + 4 new `RibbonControllerTests`); `dotnet build src/FinanceFmtTools.sln -c Release` remains 0 Warning(s)/0 Error(s) on both `net48` and `net8.0`.
- Verified zero `IRibbonUI`/`stdole`/`Microsoft.Office.Core` references and zero `CustomXMLPart`/`File.Read`/`File.Write`/`Registry.` references anywhere in `RibbonSessionConfig.cs`/`RibbonController.cs` via grep.
- **Phase 2 (Abstractions & Orchestration) is now fully complete: 2/2 plans done, all 3 roadmap success criteria proven by `dotnet test`.**

## Task Commits

Each task was committed atomically (TDD RED → GREEN pairs):

1. **Task 1: RibbonSessionConfig defaults + RibbonController checkbox session state**
   - RED: `f678e68` (test) — failing `RibbonControllerTests.cs`, `RibbonController`/`RibbonSessionConfig` don't exist yet
   - GREEN: `c957de5` (feat) — `RibbonSessionConfig`/`RibbonController` implemented, 3/3 new tests pass, 38/38 total project tests pass
2. **Task 2: RibbonController.GetCustomUiXml — embedded Ribbon XML resource**
   - RED: `d058398` (test) — failing extended `RibbonControllerTests.cs`, `GetCustomUiXml` doesn't exist yet
   - GREEN: `602fa4d` (feat) — `EmbeddedResource` `Link` added to csproj + `GetCustomUiXml()` implemented, 4/4 new tests pass, 39/39 total project tests pass, `dotnet build` 0 Warning(s)/0 Error(s) on both TFMs

No REFACTOR commits — both GREEN implementations matched the plan's specified target shape (from 02-RESEARCH.md Pattern 2/3) verbatim, with nothing genuinely warranted to extract at this size.

## Files Created/Modified
- `src/FinanceFmtTools.Engine/RibbonSessionConfig.cs` - `ForceAlign`/`ZeroDash` auto-properties, defaults `false`/`true` per RIB-02/RIB-03
- `src/FinanceFmtTools.Engine/RibbonController.cs` - `Config` property, two constructors, `GetCustomUiXml()` suffix-match resource loader
- `src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj` - added unconditioned `EmbeddedResource Include="../customUI14.xml" Link="Resources/customUI14.xml"` `ItemGroup`
- `src/FinanceFmtTools.Engine.Tests/RibbonControllerTests.cs` - 4 `[Fact]` tests covering defaults, mutation, constructor injection, and embedded-XML loading

## Decisions Made
- Skipped REFACTOR commits for both TDD tasks since the first GREEN implementation already matched the plan's specified shape verbatim.
- Confirmed `RibbonSessionConfig`'s defaults deliberately diverge from both VBA source files' contradictory defaults, per 02-RESEARCH.md Pitfall 1 and REQUIREMENTS.md RIB-02/RIB-03 — this is a considered migration behavior change, not a VBA-parity bug.
- Linked (not duplicated) `src/customUI14.xml` into the C# project, keeping a single source of truth for the Ribbon XML and its callback names during the VBA→C# transition, per 02-RESEARCH.md Pattern 3 / Open Question 2 resolution.

## Deviations from Plan
None — plan executed exactly as written. All acceptance criteria (build greps, test filters, RED-before-GREEN commit ordering, COM-reference/persistence bans) passed on first verification for every task.

## Issues Encountered
None.

## User Setup Required
None - no external service configuration required.

## Next Phase Readiness
- Phase 2 (Abstractions & Orchestration) is fully complete: `IExcelGateway`/`IRangeHandle`/`ILog`, `FormatEngine.Apply`/`ApplyToSelection` (FMT-06 guard), and now `RibbonSessionConfig`/`RibbonController` (checkbox state + embedded Ribbon XML loading) all exist and are proven via `dotnet test` with zero real Excel/COM references.
- Phase 3 (COM Entry Point & Real Excel Integration) can now build the real `Microsoft.Office.Interop.Excel`-backed `IExcelGateway`/`IRangeHandle` implementations, wire a live `IRibbonUI` against `RibbonController.Config`, implement `Connect.cs` (`IDTExtensibility2`/`IRibbonExtensibility.GetCustomUI` returning `RibbonController.GetCustomUiXml()`), and add the actual user-facing `MessageBox`/friendly-dialog behavior for the FMT-06 guard clause.
- No blockers for Phase 3.

---
*Phase: 02-abstractions-orchestration*
*Completed: 2026-07-11*
