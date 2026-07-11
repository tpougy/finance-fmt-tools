---
status: issues
files_reviewed: 13
depth: standard
findings:
  critical: 0
  warning: 0
  info: 1
  total: 1
fixed:
  - "WR-1: FormatEngine.Apply now guards range==null (mirrors VBA's ApplyFormat Nothing check), logs a warning, never throws — regression test Apply_RangeNulo_LogaAvisoENaoLanca added"
  - "WR-2: RibbonController.GetCustomUiXml() now throws InvalidOperationException on missing embedded resource instead of silently returning empty string"
  - "IN-2: RibbonController constructor now uses nameof(config) instead of a string literal"
  - "IN-3: Reworded the misleading IRangeHandle.cs comment about namespace lookup"
skipped:
  - "IN-1: FormatEngine.Apply's alignment-assignment branch stays untested dead code — deferred per the review's own suggestion, until the first real non-General format entry exists"
---

## Summary

Reviewed all 13 Phase 2 files (`IExcelGateway`, `ILog`, `IRangeHandle`, `FormatEngine`, `RibbonController`, `RibbonSessionConfig`, the test doubles, and the four test classes) plus the `.csproj`. Also cross-referenced the VBA source of truth (`src/modFormatEngine.bas`, `src/modUtils.bas`, `src/modConfig.bas`) and empirically verified claims by building/running the test suite (`dotnet build` / `dotnet test`, 39/39 passing, 0 warnings/0 errors on net8.0) and by running two scratch experiments (see CR notes below), both removed afterward with a clean `git status`.

Confirmed clean on the four specifically-requested checks, with one caveat noted in WR-1:
- **No accidental COM/Interop/WinForms references**: grep across `FinanceFmtTools.Engine` and `FinanceFmtTools.Engine.Tests` found zero hits for `Microsoft.Office.Interop`, `IRibbonUI`, `System.Windows.Forms`, `MessageBox`, `Excel.Application` outside of generated `obj/project.assets.json` build metadata (not source). `IRangeHandle`/`IExcelGateway` are pure C# abstractions as intended.
- **`RibbonSessionConfig` defaults**: verified exactly `ForceAlign = false`, `ZeroDash = true` — correctly diverges from both of VBA's contradictory defaults (`modConfig.bas`'s uninitialized-bool default of `False/False`, and `modUtils.bas`'s `LoadConfig`/XML-literal default of `True/False`), matching RIB-02/RIB-03 and the tests in `RibbonControllerTests.cs`.
- **`GetCustomUiXml()` suffix-match**: resolves the single embedded `customUI14.xml` resource correctly (confirmed the SDK-generated logical resource name is unique for this suffix, and the test asserting `tabFinanceFmt` + `onLoad="OnRibbonLoad"` content passes). Robustness gap on the "resource missing" branch is flagged as WR-2 below.
- **FMT-06 guard never throws for an invalid selection**: `FormatEngine.ApplyToSelection` correctly never throws when `gateway.TryGetSelectedRange` returns `false` — confirmed by `FormatEngineSelectionGuardTests`. However, the lower-level `FormatEngine.Apply(range, ...)` overload itself has no equivalent null-guard, which is a genuine parity gap from the VBA source (flagged as WR-1, confirmed empirically to throw).

## Findings

### WR-1: `FormatEngine.Apply` throws `NullReferenceException` on a null range, unlike its VBA source

- **File**: `src/FinanceFmtTools.Engine/FormatEngine.cs:9-24`
- **Description**: VBA's `ApplyFormat` (`src/modFormatEngine.bas:24-31`) explicitly guards `If rng Is Nothing Then Log ... : Exit Sub` before touching the range — a defense-in-depth check independent of `SafeSelection()`'s own guard. The C# port only reproduces this guard at the `ApplyToSelection` level (via `IExcelGateway.TryGetSelectedRange`'s `out bool` pattern); the public `Apply(IRangeHandle range, ILog log, string formatKey, bool forceAlign, bool zeroDash)` overload has no `range == null` check at all. Since `Apply` is a public static entry point that Phase 3 (or any other caller) can invoke directly with a range obtained by other means, a null range reaches `range.NumberFormat = def.NumberFormat;` (line 17) unguarded.
- **Failure scenario**: Verified empirically — calling `FormatEngine.Apply(null, log, FormatKeys.Fin2D, false, true)` throws `System.NullReferenceException: Object reference not set to an instance of an object.` at `FormatEngine.cs:17`, instead of logging a warning and returning, as the "guard clause never throws" contract (and VBA parity) would imply. This contradicts the explicit FMT-06 intent that invalid-range scenarios are handled by logging, not exceptions, at the orchestration layer.
- **Suggested fix**: Add the missing guard at the top of `Apply`, mirroring the VBA source:
  ```csharp
  if (range == null)
  {
      log.Warn("FormatEngine.Apply: range é null — abortando '" + formatKey + "'.");
      return;
  }
  ```
  Add a regression test (e.g. in `FormatEngineTests.cs`) asserting `Record.Exception(() => FormatEngine.Apply(null, log, ...))` is null, mirroring the existing unknown-key test.

### WR-2: `RibbonController.GetCustomUiXml()` fails silently and untested when the embedded resource is missing

- **File**: `src/FinanceFmtTools.Engine/RibbonController.cs:22-41`
- **Description**: When no manifest resource name ends with `"customUI14.xml"`, the method returns `string.Empty` with no logging, no exception, and no diagnostic of any kind (`RibbonController` has no `ILog` dependency at all). This satisfies "won't throw an unhelpful exception," but goes to the opposite extreme: a genuine packaging regression (e.g., the `EmbeddedResource` item in `FinanceFmtTools.Engine.csproj:18` gets removed, renamed, or the relative path `../customUI14.xml` breaks) would silently produce an empty ribbon XML string. In Phase 3, feeding an empty string to Excel's `GetCustomUI` COM callback would render no Finance Fmt ribbon tab at all, with zero error trail pointing back to "resource not found" — a hard-to-diagnose production regression. Additionally, the "no match found" branch (and the "more than one match found — first one silently wins" scenario, however unlikely today with a single `EmbeddedResource` entry) has zero test coverage; only the happy path (`GetCustomUiXml_CarregaRecursoEmbutido_ContemTabFinanceFmt`) is tested.
- **Failure scenario**: Someone edits `FinanceFmtTools.Engine.csproj` (e.g., changes the `Link` path or accidentally deletes the `EmbeddedResource` entry) during a future refactor. `dotnet build`/`dotnet test` still pass because nothing asserts the string is well-formed XML content-independent of the "happy path" test data, and in a CI run without an assertion on `GetCustomUiXml()` returning non-empty for the shipped `.dll`, this regression ships silently until a user reports "the ribbon tab disappeared."
- **Suggested fix**: At minimum, add a test that simulates/asserts the empty-string contract explicitly (e.g., extract the suffix-match loop into a testable helper that accepts an `IEnumerable<string>` of resource names, so both the "found" and "not found" — and "multiple matches" — branches can be unit-tested without needing a second real assembly). Consider also throwing a descriptive `InvalidOperationException` ("Embedded resource 'customUI14.xml' not found in assembly.") instead of returning empty, since a missing ribbon resource is an unrecoverable build/packaging defect, not a normal runtime condition like FMT-06's invalid selection.

### IN-1: `FormatEngine.Apply`'s alignment-assignment branch is currently dead code and untested

- **File**: `src/FinanceFmtTools.Engine/FormatEngine.cs:18-21`
- **Description**: `if (def.Alignment != CellAlignment.General) { range.HorizontalAlignment = def.Alignment; }` can never be true today: every `FormatRegistry.TryGetFormatDef` case constructs its `FormatDef` with `CellAlignment.General` (per `FormatRegistry.cs`'s own comment, faithfully porting the VBA source where `f.Alignment` is never assigned in any `Case` branch). This is intentional VBA parity, not a bug, but it means no test in `FormatEngineTests.cs`/`FormatEngineSelectionGuardTests.cs` exercises the `range.HorizontalAlignment` setter path at all.
- **Failure scenario**: If a future format entry is added with `CellAlignment.Right`/`Left` (a very plausible near-term change, since the interface and enum already exist for exactly this purpose), a typo or logic inversion in this branch would go unnoticed — there's no existing test that would catch a regression here.
- **Suggested fix**: Add one test using a hypothetical/temporary alignment-carrying `FormatDef` (or defer this until the first real non-General format is added) to lock in the branch's behavior before it becomes load-bearing.

### IN-2: `RibbonController` constructor uses a string literal instead of `nameof(config)`

- **File**: `src/FinanceFmtTools.Engine/RibbonController.cs:15`
- **Description**: `if (config == null) throw new ArgumentNullException("config");` hardcodes the parameter name as a string. `nameof(config)` is available (project targets `LangVersion 9.0`) and avoids silent drift if the parameter is ever renamed during a refactor.
- **Suggested fix**: `throw new ArgumentNullException(nameof(config));`

### IN-3: Misleading comment in `IRangeHandle.cs` about C# namespace resolution

- **File**: `src/FinanceFmtTools.Engine/Abstractions/IRangeHandle.cs:1-2`
- **Description**: The comment states: "Needs `using FinanceFmtTools.Engine;` for `CellAlignment` — child namespaces do not automatically see the parent namespace's types in C#." This is empirically false for this codebase's namespace layout. Verified by removing the `using FinanceFmtTools.Engine;` line and rebuilding both `net48` and `net8.0` targets — the project still compiles with 0 errors/0 warnings (change was reverted immediately afterward; `git status` confirms the file is unmodified). C#'s simple-name lookup for a dotted namespace declaration (`namespace FinanceFmtTools.Engine.Abstractions { ... }`) does walk outward through the enclosing `FinanceFmtTools.Engine` namespace across the whole compilation (this is also why `FormatEngineTests.cs` and `FormatEngineSelectionGuardTests.cs` reference `FormatEngine`/`FormatKeys` — both declared in the `FinanceFmtTools.Engine` namespace in other files — without any `using FinanceFmtTools.Engine;` directive of their own).
- **Suggested fix**: Either remove the now-redundant `using` and the comment, or correct the comment's technical claim (e.g., "kept explicit for readability/clarity, even though it's not strictly required by C#'s namespace lookup rules") so it doesn't propagate an incorrect mental model to future contributors.
