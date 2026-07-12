---
phase: quick-260712-glm-refatorar-accountingformatbuilder
plan: 1
type: execute
wave: 1
depends_on: []
autonomous: true
files_modified:
  - src/FinanceFmtTools.Engine/FormatTokens.cs
  - src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs
must_haves:
  truths:
    - "AccountingFormatBuilder.Build(decimals, forceAlign, zeroDash) returns byte-for-byte identical output to the pre-refactor implementation for every combination already covered by AccountingFormatBuilderTests.cs (16-row Theory matrix + negative-decimals guard Fact)."
    - "No inline string literal for an accounting-format token (fill-prefix, parenthesis padding, literal parentheses, hyphen padding, digit-group pattern, decimal point, zero-dash literal, section separator) remains directly in AccountingFormatBuilder.Build — every such piece is a named constant referenced from FormatTokens."
    - "AccountingFormatBuilderTests.cs is not modified and all its tests still pass."
  artifacts:
    - path: "src/FinanceFmtTools.Engine/FormatTokens.cs"
      provides: "Named C# constants for every combinable piece of the 3-section Excel accounting number-format string"
      contains: "public static class FormatTokens"
    - path: "src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs"
      provides: "Build method composed entirely from FormatTokens constants, same two-branch (general + decimals==0) structure as before"
      contains: "FormatTokens."
  key_links:
    - from: "src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs"
      to: "src/FinanceFmtTools.Engine/FormatTokens.cs"
      via: "FormatTokens.<ConstantName> references inside Build's string concatenation"
      pattern: "FormatTokens\\."
---

<objective>
Refactor `AccountingFormatBuilder.Build` (`src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs`) so every
loose string literal used to assemble the 3-section Excel accounting number-format string is replaced by a
named constant from a new `FormatTokens` class (`src/FinanceFmtTools.Engine/FormatTokens.cs`). This is a PURE
refactor — no behavior change, no new dependency, no build-time codegen. It exists so future tweaks to a
single token (e.g. swapping the hyphen-padding token for a parenthesis-padding token) touch one obvious
constant instead of requiring surgery on string-concatenation expressions.

Purpose: Make the accounting-format assembly logic self-documenting and easy to tweak — each token gets a
name and a one-line rationale comment, instead of the reader having to reverse-engineer what `"_)"` or
`" * "` mean from Excel's number-format mini-language.

Output: `src/FinanceFmtTools.Engine/FormatTokens.cs` (new file, named constants only) and
`src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs` (modified — same public signature and same
two-branch structure, now composed from `FormatTokens` constants). `AccountingFormatBuilderTests.cs` is
read-only context — it must keep passing completely unchanged.
</objective>

<execution_context>
@$HOME/.claude/get-shit-done/workflows/execute-plan.md
@$HOME/.claude/get-shit-done/templates/summary.md
</execution_context>

<context>
@.planning/STATE.md
@./CLAUDE.md
@src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs
@src/FinanceFmtTools.Engine/FormatDef.cs
@src/FinanceFmtTools.Engine.Tests/AccountingFormatBuilderTests.cs
</context>

<tasks>

<task type="auto">
  <name>Task 1: Extract format tokens into FormatTokens.cs, recompose AccountingFormatBuilder.Build, and verify byte-for-byte parity</name>
  <files>src/FinanceFmtTools.Engine/FormatTokens.cs, src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs</files>
  <action>
Create `src/FinanceFmtTools.Engine/FormatTokens.cs` as a new `public static class FormatTokens` in the
`FinanceFmtTools.Engine` namespace (plain `public const` fields only — no records, no init-only properties,
this file compiles under net48 same as `FormatDef.cs`). Define exactly these named constants, each with a
one-line comment explaining what the token means in Excel's number-format mini-language (mirror the
rationale already given inline in the current `AccountingFormatBuilder.cs` header comment, do not invent new
behavior):
- `FillAlignPrefix` = `" * "` — the fill-character token that forces right alignment when CFG_FORCE_ALIGN
  (forceAlign) is on.
- `OpenParenPad` = `"_("` — padding the width of an opening parenthesis without printing one; used to open
  the positive/zero sections so their digits line up under the negative section's real `(`.
- `CloseParenPad` = `"_)"` — padding the width of a closing parenthesis without printing one; closes the
  positive/zero sections to match the negative section's real `)`.
- `OpenParen` = `"("` — literal opening parenthesis wrapping the negative section.
- `CloseParen` = `")"` — literal closing parenthesis wrapping the negative section.
- `PadHyphen` = `"_-"` — padding the width of a hyphen without printing one; appended at the end of every
  section.
- `DigitsBase` = `"#,##0"` — the thousands-grouped digit pattern shared by the decimals==0 and decimals&gt;0
  cases (the latter appends `DecimalPoint` + a run of zeros).
- `DecimalPoint` = `"."` — decimal point separator, only used when decimals &gt; 0.
- `ZeroDashLiteral` = `"-"` — the literal character shown in the zero section when CFG_ZERO_DASH (zeroDash)
  is on.
- `ZeroDigit` = `'0'` (a `char`, not `string`) — the fractional-digit character repeated `decimals` times via
  `new string(FormatTokens.ZeroDigit, decimals)`.
- `SectionSeparator` = `";"` — joins the positive/negative/zero sections into the final 3-section format
  string.

Then rewrite `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs`'s `Build` method so every one of the
literals above is replaced by the matching `FormatTokens.<Name>` reference, while preserving the exact
existing control-flow shape used today: keep the `string dec = new string(<zero-digit-token>, decimals)`
line first (unchanged position — this is what makes `Build(-1, ...)` throw `ArgumentOutOfRangeException`
natively, must keep throwing for negative decimals), keep the `if (forceAlign) { ... } else { ... }` block
computing the general-case `pos`/`neg`/`zer` (with `DigitsBase + DecimalPoint + dec` as the digit pattern),
and keep the subsequent `if (decimals == 0) { if (forceAlign) { ... } else { ... } }` override block exactly
as today (same nesting, same variable names `pos`/`neg`/`zer`) — just using `DigitsBase` alone (no
`DecimalPoint`/`dec`) as the digit pattern in this override. In every branch, `zer` is `pos` when `zeroDash`
is false, and `(forceAlign prefix or none) + OpenParenPad + ZeroDashLiteral + CloseParenPad + PadHyphen` when
`zeroDash` is true — identical in both the general and decimals==0 branches, matching today's behavior (the
literal `" * _(-_)_-"` / `"_(-_)_-"` values do not depend on `decimals`). Build the final return value as
`pos + FormatTokens.SectionSeparator + neg + FormatTokens.SectionSeparator + zer`. Do not unify the
two-branch structure into a single formula — this mirrors the VBA source deliberately (see the existing
header comment referencing `01-RESEARCH.md` Common Pitfalls #4 and `01-02-SUMMARY.md`'s
patterns-established note); update the header comment and the decimals==0 inline comment so they describe
the new token-composed code accurately (no stale references to raw literals that no longer appear in this
file). Do not touch `FormatRegistry.cs`, `FormatDef.cs`, or `AccountingFormatBuilderTests.cs`.

After the rewrite, prove the extraction is real and nothing regressed: run the existing test suite (must
stay green, unchanged file), build the full solution to confirm `FormatTokens.cs` compiles cleanly on both
`net48` and `net8.0` (plain `const` fields carry no CS0518 risk — that only applies to C# 9
record/init-only syntax per `FormatDef.cs`'s header comment, but prove it rather than assume it), and grep
`AccountingFormatBuilder.cs`'s `Build` method body to confirm every token literal was actually replaced
(non-zero count of `FormatTokens.` references) rather than merely duplicated alongside the new constants
(zero remaining raw occurrences of `" * "`, `"_("`, `"_)"`, `"_-"`, a bare `";"` section separator, or the
raw `"-"` zero-dash literal inside `Build`'s body — filter out the class's header/doc comments, which
describe these characters in prose, so a comment does not produce a false negative).
  </action>
  <verify>
    <automated>dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release &amp;&amp; dotnet build src/FinanceFmtTools.sln -c Release</automated>
  </verify>
  <done>
`dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release` reports all 17
tests passing (16-row Theory matrix + negative-decimals Fact), with `AccountingFormatBuilderTests.cs`
untouched (`git diff --stat` shows zero changes to that file). `dotnet build src/FinanceFmtTools.sln -c
Release` completes with 0 Warning(s)/0 Error(s) across both `net48` and `net8.0` legs. `FormatTokens.cs`
exists with all eleven named constants and one-line rationale comments; `AccountingFormatBuilder.cs`'s
`Build` method body contains no remaining raw duplicate of any token literal — every section-string piece is
sourced from a `FormatTokens.*` constant. `git diff --stat` shows only these two files touched — no changes
to `FormatRegistry.cs`, `FormatDef.cs`, or the test file.
  </done>
</task>

</tasks>

<threat_model>
## Trust Boundaries

| Boundary | Description |
|----------|--------------|
| None | This is an internal, pure-function string-composition refactor with no external input, no I/O, no new dependency, and no change to any trust boundary already present in the codebase. |

## STRIDE Threat Register

| Threat ID | Category | Component | Disposition | Mitigation Plan |
|-----------|----------|-----------|-------------|-----------------|
| T-quick-01 | Tampering | `FormatTokens.cs` constants | accept | Compile-time `const` fields, not user-controllable at runtime; covered by existing `AccountingFormatBuilderTests.cs` byte-for-byte assertions, so any accidental token change is caught by the test suite before merge. |
</threat_model>

<verification>
1. `dotnet test src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release` — all 17 tests pass, `AccountingFormatBuilderTests.cs` diff is empty.
2. `dotnet build src/FinanceFmtTools.sln -c Release` — 0 Warning(s)/0 Error(s), both `net48` and `net8.0` legs build.
3. `git diff --stat` — only `src/FinanceFmtTools.Engine/FormatTokens.cs` (new) and
   `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs` (modified) appear; no test file, `FormatRegistry.cs`,
   or `FormatDef.cs` changes.
</verification>

<success_criteria>
- `FormatTokens.cs` exists with named constants for every combinable piece of the accounting-format string
  (fill-align prefix, parenthesis padding pair, literal parenthesis pair, hyphen padding, digit-group base,
  decimal point, zero-dash literal, zero digit char, section separator).
- `AccountingFormatBuilder.Build` composes its return value entirely from `FormatTokens.*` references — zero
  inline string literals for these tokens remain in the method body.
- All 17 existing tests in `AccountingFormatBuilderTests.cs` pass unchanged — output is byte-for-byte
  identical to the pre-refactor implementation for every tested combination.
- Full solution builds clean (0/0 warnings/errors) on both `net48` and `net8.0`.
</success_criteria>

<output>
Create `.planning/quick/260712-glm-refatorar-accountingformatbuilder-para-u/260712-glm-SUMMARY.md` when done
</output>
