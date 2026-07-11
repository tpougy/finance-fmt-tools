---
status: issues
files_reviewed: 11
depth: standard
findings:
  critical: 0
  warning: 0
  info: 3
  total: 3
fixed:
  - "WR-1: stale 'wired in a later task' comment in FormatRegistry.cs — reworded to past tense, rebuilt clean (0 warnings/0 errors)"
---

## Summary

Reviewed the full Phase 1 format-engine surface — `AccountingFormatBuilder`, `CellAlignment`, `FormatCategory`, `FormatDef`, `FormatKeys`, `FormatRegistry`, both test files, and both `.csproj` files — against `src/modFormatEngine.bas` and `src/modConfig.bas` as the source of truth.

Transcription fidelity is excellent: all 11 `FMT_*` key constants match verbatim (`FormatKeys.cs` vs `modConfig.bas:19-29`), all 7 literal registry entries (`Pct4D`, `Pct2D`, `SpreadBps`, `DateIso`, `DateBr`, `DateBrLong`, `Text`) match their VBA `NumberFmt`/`DisplayName`/`Category` byte-for-byte including the escaped-quote decoding for `SpreadBps` (`#,##0.0" bps"`), and the `AccountingFormatBuilder.Build` port reproduces VBA's `AccountingFmt` two-branch structure (general N-decimals case, then the explicit `decimals == 0` override) exactly, including the order-of-operations subtlety where `zer` in the zero-override block is computed from the *newly reassigned* `pos`, not the stale one. I hand-verified all 16 `[InlineData]` rows in `AccountingFormatBuilderTests.cs` against the VBA algorithm and found no transcription errors. The registry's decision to hard-code `CellAlignment.General` for every entry (including the Fin/Integer family) is correct and was clearly verified directly against the VBA source (`modFormatEngine.bas`'s `GetFormatDef` never assigns `f.Alignment` in any `Case` branch) — notably this correctly diverges from an incorrect `CellAlignment.Right` suggestion in `01-RESEARCH.md`'s illustrative snippet, which is good engineering, not a bug.

No critical or security issues found — this code has no I/O, no external input parsing, and no mutable state; it's pure string composition over value types. One warning (a stale comment) and three low-severity info items are noted below, none of which affect current correctness.

### CR / none

No critical findings.

### WR-1: Stale "wired in a later task" comment in FormatRegistry.cs contradicts the file's own content

**File:** `src/FinanceFmtTools.Engine/FormatRegistry.cs:6`

**Description:** The top-of-file comment reads: *"The VBA source never assigns f.Alignment in any Case branch, so every entry (including the Fin family, **wired in a later task**) carries the General alignment value..."* This phrasing was accurate when the file was first created in commit `8f0c027` (only the 7 literal entries existed, and a companion paragraph explicitly said `"Interim state: only the 7 literal (non-Fin) entries are wired so far... the next task adds explicit cases for them"`). Commit `5443f21` wired the Fin/Integer family into this same file and correctly deleted the "Interim state" paragraph — but left the inline parenthetical "(including the Fin family, wired in a later task)" untouched. Now that all 11 keys (including Integer/Fin2D/Fin4D/Fin8D) are wired directly in this file, the phrase is factually wrong and would mislead a future reader into thinking the Fin family is still pending elsewhere.

**Suggested fix:** Drop the parenthetical or reword to past tense, e.g.: *"...so every entry, including the Fin family below, carries the General alignment value..."*

### IN-1: FormatDef constructor performs no null-argument validation

**File:** `src/FinanceFmtTools.Engine/FormatDef.cs:16-23`

**Description:** The constructor assigns `key`, `displayName`, and `numberFormat` directly with no null/empty guard. Every current call site in `FormatRegistry.cs` passes non-null string literals, so this isn't exercised today, but a `FormatDef` with a `null` `NumberFormat` would silently propagate until something downstream (e.g., a future Excel COM `Range.NumberFormat = null` write in Phase 2/3) throws a much less informative error far from the actual mistake.

**Suggested fix:** Optional for Phase 1 given the closed set of call sites; worth a lightweight `ArgumentNullException`-style guard (or at least an XML-doc note on the non-null contract) before this type is consumed by Phase 2/3's COM-facing code.

### IN-2: `Build`'s `int decimals` has a much wider valid range than VBA's 16-bit `Integer`, diverging in failure mode for out-of-range input

**File:** `src/FinanceFmtTools.Engine/AccountingFormatBuilder.cs:13,18`

**Description:** VBA's `AccountingFmt(ByVal decimals As Integer, ...)` uses a 16-bit `Integer` (-32768..32767); passing an out-of-range value there is caught by VBA's own type coercion long before `String(decimals, "0")` runs. The C# port takes a 32-bit `int`, so `new string('0', decimals)` will happily attempt to allocate a multi-gigabyte string for a large positive value (risking `OutOfMemoryException`/`ArgumentOutOfRangeException` from the runtime's max-string-length guard instead of VBA's immediate, small-threshold error) before throwing for negative values as documented. This is purely theoretical today — `FormatRegistry` only ever calls `Build` with `0`, `2`, `4`, or `8` — but if `Build` is ever exposed as a public API called with untrusted/derived input in a later phase, the failure boundary won't match VBA's.

**Suggested fix:** No action needed while call sites are limited to the 4 known constants; if `Build` becomes part of a wider public surface, consider clamping/validating `decimals` against a sane upper bound (e.g., reject anything outside VBA's `Integer` range) to keep failure semantics aligned.

### IN-3: `net8.0`-only test project never exercises the `net48` build of the Engine library at runtime

**File:** `src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj:4`, `src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj:4`

**Description:** `FinanceFmtTools.Engine` multi-targets `net48;net8.0`, but `FinanceFmtTools.Engine.Tests` targets `net8.0` only, so all 31 tests run exclusively against the `net8.0` build. The `net48` TFM is verified solely by `dotnet build` succeeding (0 warnings/errors), never by `dotnet test`. Given this phase's stated approach (Linux dev environment, `net48` compiled via `Microsoft.NETFramework.ReferenceAssemblies` reference-assembly-only compilation, no CLR available to actually run `net48` binaries on Linux) this is a reasonable, deliberate tradeoff — the code here has no TFM-conditional logic or culture-sensitive APIs that would plausibly diverge — but it's worth flagging explicitly as a coverage boundary: a regression that only manifests on the real .NET Framework CLR (e.g., a subtle string/allocation behavior difference) would not be caught by the test suite as currently wired.

**Suggested fix:** No change required for Phase 1. If real Excel/COM testing on Windows becomes available in a later phase, consider adding a `net48` leg to CI (even just `dotnet test -f net48` on a Windows runner) for full parity assurance.
