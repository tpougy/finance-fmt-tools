---
phase: quick-260712-glm-refatorar-accountingformatbuilder
plan: 1
subsystem: engine
tags: [csharp, refactor, accounting-format, net48, net8.0]

# Dependency graph
requires:
  - phase: 01-02 (AccountingFormatBuilder.Build implementation)
    provides: original working two-branch Build implementation with inline string literals
provides:
  - FormatTokens.cs — named constants for every combinable piece of the 3-section Excel accounting number-format string
  - AccountingFormatBuilder.Build recomposed entirely from FormatTokens.* references, byte-for-byte identical output
affects: [any future plan touching AccountingFormatBuilder.cs or extending accounting-format tokens]

# Tech tracking
tech-stack:
  added: []
  patterns:
    - "Named-token extraction: string-composition logic sources every literal piece from a dedicated static-const class instead of inline literals"

key-files:
  created: [src/FinanceFmtTools.Engine/FormatTokens.cs]
  modified: [src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs]

key-decisions:
  - "Kept VBA's original two-branch structure (general N-decimals case + explicit decimals==0 override) unchanged — pure token extraction, no logic unification"
  - "FormatTokens constants are plain public const fields (string/char) — no records/init-only properties, to guarantee net48 compilation matches FormatDef.cs's established constraint"

patterns-established:
  - "Pattern: extract magic-string tokens for domain-specific mini-languages (Excel number-format) into a single named-constant class colocated with the consumer, each constant documented with a one-line rationale comment"

requirements-completed: []

# Metrics
duration: ~15min
completed: 2026-07-12
---

# Quick Task 260712-glm: Refactor AccountingFormatBuilder to use FormatTokens Summary

**Extracted all inline Excel accounting-format string literals in `AccountingFormatBuilder.Build` into 11 named constants in a new `FormatTokens` class — pure refactor, byte-for-byte identical output.**

## Performance

- **Duration:** ~15 min
- **Completed:** 2026-07-12
- **Tasks:** 1
- **Files modified:** 2 (1 created, 1 modified)

## Accomplishments
- Created `src/FinanceFmtTools.Engine/FormatTokens.cs` with 11 named `public const` fields (`FillAlignPrefix`, `OpenParenPad`, `CloseParenPad`, `OpenParen`, `CloseParen`, `PadHyphen`, `DigitsBase`, `DecimalPoint`, `ZeroDashLiteral`, `ZeroDigit`, `SectionSeparator`), each with a one-line rationale comment explaining its meaning in Excel's number-format mini-language.
- Recomposed `AccountingFormatBuilder.Build` so every string-concatenation piece is sourced from `FormatTokens.*` — same two-branch control-flow shape (general N-decimals case, then explicit `decimals == 0` override), same variable names (`pos`/`neg`/`zer`), same negative-decimals throw behavior via `new string(FormatTokens.ZeroDigit, decimals)`.
- Verified byte-for-byte parity: all 17 `AccountingFormatBuilderTests` (16-row Theory matrix + negative-decimals Fact) pass unchanged, file untouched.
- Verified `FormatTokens.cs` compiles cleanly on both `net48` and `net8.0` legs (`dotnet build src/FinanceFmtTools.sln -c Release` → 0 Warning(s)/0 Error(s)).
- Confirmed via grep that zero raw duplicate token literals (`" * "`, `"_("`, `"_)"`, `"_-"`, `"#,##0"`, bare `";"`, raw `"-"` zero-dash) remain in `Build`'s method body — all sourced from `FormatTokens.*` (16 references).

## Task Commits

Each task was committed atomically:

1. **Task 1: Extract format tokens into FormatTokens.cs, recompose AccountingFormatBuilder.Build, and verify byte-for-byte parity** - `afe34c7` (refactor)

## Files Created/Modified
- `src/FinanceFmtTools.Engine/FormatTokens.cs` - New static class with 11 named constants for every combinable Excel accounting-format token
- `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs` - `Build` method recomposed from `FormatTokens.*` references; same public signature, same two-branch structure, comments updated to reference the new token source

## Decisions Made
- Kept the VBA-mirrored two-branch structure exactly as-is (no unification into a single formula) — this was an explicit plan constraint, matching the existing header comment's rationale (`01-RESEARCH.md` Common Pitfalls #4).
- Used plain `const` fields (not records/readonly static properties) in `FormatTokens.cs` to avoid the CS0518 risk documented in `FormatDef.cs`'s header comment for net48 targets — verified by building both `net48` and `net8.0` legs rather than assuming safety.

## Deviations from Plan

None - plan executed exactly as written.

## Issues Encountered

None. `dotnet` was not on `PATH` by default in this shell session; resolved by adding `$HOME/.dotnet` to `PATH` for the build/test commands (environment quirk, not a code change).

## User Setup Required

None - no external service configuration required.

## Next Phase Readiness
- `AccountingFormatBuilder.cs` and `FormatTokens.cs` are ready for any future plan that needs to tweak a single accounting-format token (e.g., swapping hyphen-padding for parenthesis-padding) — the change now touches one named constant instead of requiring surgery on string-concatenation expressions.
- No blockers or concerns for downstream work.

---
*Task: quick-260712-glm-refatorar-accountingformatbuilder*
*Completed: 2026-07-12*

## Self-Check: PASSED

- FOUND: src/FinanceFmtTools.Engine/FormatTokens.cs
- FOUND: src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs
- FOUND: .planning/quick/260712-glm-refatorar-accountingformatbuilder-para-u/260712-glm-SUMMARY.md
- FOUND commit: afe34c7
