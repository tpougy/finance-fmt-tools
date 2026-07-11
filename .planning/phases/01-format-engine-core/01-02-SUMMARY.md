---
phase: 01-format-engine-core
plan: 02
subsystem: format-engine
tags: [dotnet, csharp, xunit, tdd, accounting-format]

# Dependency graph
requires: ["01-01-bootstrap-solution"]
provides:
  - "AccountingFormatBuilder.Build(int decimals, bool forceAlign, bool zeroDash) — pure port of VBA's private AccountingFmt"
  - "17 xUnit tests (16-combination matrix + negative-decimals guard) proving byte-for-byte VBA parity"
affects: [01-03-format-registry]

# Tech tracking
tech-stack:
  added: []
  patterns: ["TDD RED->GREEN commit pair per feature", "new string('0', decimals) as the natural guard against negative input (ArgumentOutOfRangeException) instead of an explicit check"]

key-files:
  created:
    - src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs
    - src/FinanceFmtTools.Engine.Tests/AccountingFormatBuilderTests.cs
  modified: []

key-decisions:
  - "No REFACTOR commit needed — the GREEN implementation already matched the plan's target shape (explicit two-branch structure, no static mutable state, public static pure function) with no excessive duplication to extract."

patterns-established:
  - "AccountingFormatBuilder.cs pattern: pure static function ported 1:1 from a VBA private Function, preserving the source's branch structure even when a 'cleaner' unified formula would be tempting — correctness-by-construction over elegance when porting legacy business logic."

requirements-completed: [FMT-01, FMT-07]

# Metrics
duration: unknown (session interrupted by API rate limit; resumed and closed out in a follow-up session)
completed: 2026-07-11
---

# Phase 1 Plan 2: AccountingFormatBuilder Summary

**Pure C# port of VBA's private `AccountingFmt` helper, proven byte-for-byte identical across all 16 decimals x forceAlign x zeroDash combinations via a TDD RED->GREEN xUnit `[Theory]`, plus a guard test for negative input.**

## Accomplishments
- `AccountingFormatBuilder.Build(int decimals, bool forceAlign, bool zeroDash)` ported from `src/modFormatEngine.bas:188-222`, preserving the VBA algorithm's two-branch structure (general N-decimals case, then an explicit `decimals == 0` override) rather than unifying into one formula (per 01-RESEARCH.md Pitfall #4).
- `new string('0', decimals)` throws `ArgumentOutOfRangeException` natively for negative `decimals`, satisfying the defensive-guard requirement without a redundant explicit check.
- 17/17 tests pass under `dotnet test` (16-row `[Theory]` matrix + 1 `[Fact]` guard test), all values matching 01-RESEARCH.md's pre-computed table exactly.
- Full solution still builds 0 Warning(s)/0 Error(s) across both `net48` and `net8.0` legs after this plan's additions.
- No static/mutable fields introduced — `AccountingFormatBuilder` is a pure, stateless function (verified by grep in acceptance criteria).

## Task Commits
1. **RED — failing test for AccountingFormatBuilder.Build** - `5f9255f` (test)
2. **GREEN — implement AccountingFormatBuilder.Build** - `e075fb6` (feat)

No REFACTOR commit — implementation already matched the target shape on first GREEN pass.

## Files Created
- `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs` - pure static `Build` function, two-branch structure mirroring VBA exactly
- `src/FinanceFmtTools.Engine.Tests/AccountingFormatBuilderTests.cs` - 16-row `[Theory]`/`[InlineData]` matrix + 1 `[Fact]` guard test

## Decisions Made
- Relied on .NET's native `ArgumentOutOfRangeException` from `new string('0', decimals)` for the negative-decimals guard instead of writing a separate explicit range check — fewer lines, same observable behavior, and it's exactly what the plan's `<action>` specified.

## Deviations from Plan
None — implementation matches the plan's specified behavior, file list, and acceptance criteria exactly.

## Issues Encountered
- This plan's execution was interrupted mid-session by a subagent API rate limit after the RED and GREEN commits landed but before SUMMARY.md/STATE.md/ROADMAP.md closeout. Resumed in a follow-up session: re-verified all 17 tests still pass, re-ran the full solution build (0/0 warnings/errors), confirmed all acceptance-criteria greps, then completed the closeout (this SUMMARY.md + STATE.md/ROADMAP.md/REQUIREMENTS.md updates). No code changes were needed during resume — the RED/GREEN work was already correct and complete.

## User Setup Required
None.

## Next Phase Readiness
- `AccountingFormatBuilder.Build` is ready for 01-03's `FormatRegistry` to call directly for all Fin-family (Fin 0D/2D/4D/8D) registry entries.
- No blockers.

---
*Phase: 01-format-engine-core*
*Completed: 2026-07-11*
