# Architecture Research

**Domain:** C# COM Excel add-in (IDTExtensibility2 + IRibbonExtensibility), migrated from a layered VBA add-in
**Researched:** 2026-07-10
**Confidence:** HIGH (pattern verified against a working sibling implementation — `~/pessoal/outlook-classic-delay-send` — plus current official Microsoft docs/NuGet for the COM interop pieces)

## Standard Architecture

### System Overview

The existing VBA add-in is a 4-layer callback-driven design (Ribbon XML → thin callbacks → Format Engine → Config/Utils). The sibling project `outlook-classic-delay-send` already solved "how do you build this exact shape in C# as a pure COM add-in, testable with xUnit, buildable with only the `dotnet` CLI" for Outlook. Its solution generalizes directly to Excel — same COM interfaces (`IDTExtensibility2`, `Office.IRibbonExtensibility`), same problem (Office callbacks are invoked by name through `IDispatch`, business logic must not touch COM directly to be testable). The target architecture below is that pattern re-applied to Excel's object model (`Application`/`Range` instead of `Application`/`MailItem`/`NameSpace`).

```
┌─────────────────────────────────────────────────────────────────────┐
│                     Ribbon XML (declarative, embedded resource)      │
│   tab "Finance Fmt" → groups → buttons (tag=FMT_KEY) + 2 checkboxes  │
│   onAction / getPressed / onLoad → named methods on Connect          │
├───────────────────────────────┬───────────────────────────────────────┤
│  COM ENTRY POINT: Connect     │  Excel invokes these BY NAME via      │
│  [ComVisible][Guid][ProgId]   │  IDispatch.Invoke — requires          │
│  : IDTExtensibility2,          │  ClassInterfaceType.AutoDispatch      │
│    Office.IRibbonExtensibility │  (see Pattern 1)                     │
├───────────────────────────────┴───────────────────────────────────────┤
│                     Composition: AddInHost                            │
│   Wires everything from the `object Application` param of             │
│   OnConnection. Owns lifecycle (Wire/Teardown). No business logic.    │
├───────────┬───────────────────────────────┬───────────────────────────┤
│ RibbonController │        FormatEngine (Services)  │   AppConfig      │
│ (Ribbon-UI only) │  resolves + applies a format key │  (in-memory,    │
│ caches IRibbonUI,│  via IExcelGateway                │  no persist.    │
│ GetCustomUiXml,  │                                    │  per this      │
│ GetPressed       │                                    │  milestone)    │
├───────────┴───────────────────┬───────────────────────┴───────────────┤
│         Domain (pure, no COM, no Excel reference at all)              │
│   FormatKeys, FormatDefinition, FormatRegistry, AccountingFormat      │
│   Builder — 1:1 port of modFormatEngine.bas's GetFormatDef/AccountingFmt│
├─────────────────────────────────────────────────────────────────────┤
│         Abstractions: IExcelGateway / IRangeHandle / ILog             │
│   The ONLY seam between Domain/Services and real Excel COM objects    │
├─────────────────────────────────────────────────────────────────────┤
│         Services: ExcelGateway (the ONLY class that touches           │
│         Excel.Application / Excel.Range COM types)                    │
└───────────────────────────────┬───────────────────────────────────────┘
                                 ▼
                     Excel Application / user's selected Range
```

### Component Responsibilities

| Component | Responsibility | Typical Implementation |
|-----------|----------------|------------------------|
| Ribbon XML (embedded resource) | Declares tab/groups/buttons/checkboxes; carries the format key as a `tag` attribute so one generic callback can serve all buttons | `Ribbon/ribbon.xml`, `<EmbeddedResource>` in the `.csproj`, loaded via `Assembly.GetManifestResourceStream` |
| `Connect` (COM entry point) | The object Excel instantiates via CLSID/ProgID; implements `IDTExtensibility2` + `Office.IRibbonExtensibility`; exposes every Ribbon-XML-referenced callback as a plain public method (so `IDispatch.GetIDsOfNames` can find it); delegates every call to `AddInHost` in a try/catch that **swallows and logs** (never let an exception escape a COM boundary call) | `[ComVisible(true)] [Guid("…")] [ProgId("FinanceFmtTools.Connect")] [ClassInterface(ClassInterfaceType.AutoDispatch)] sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility` |
| `AddInHost` (Composition) | Manual DI/wiring, created once by `Connect`'s constructor, populated in `OnConnection`/`Wire(object application)`; owns the lifecycle (`Wire`, `Teardown`) and exposes `FormatEngine`, `Ribbon`, `Config`, `Log` to `Connect` | `Composition/AddInHost.cs` — plain class, not itself unit-tested (it's wiring, same as the Outlook reference) |
| `RibbonController` | Ribbon-UI-only concerns: load the embedded XML string, cache `IRibbonUI` on `onLoad`, answer `getPressed` for the two checkboxes from the current `AppConfig`, invalidate controls after a toggle | `Services/RibbonController.cs` implementing `Abstractions/IRibbonController.cs` — **unit-testable** (XML-resource loading + getPressed logic need no live Excel) |
| `FormatEngine` | Orchestrates "apply format key X to the current selection": resolves the format via the pure `Domain.FormatRegistry`, gets the selection via `IExcelGateway`, applies `NumberFormat`/`HorizontalAlignment`, wrapped in screen-updating suppression | `Services/FormatEngine.cs` — depends only on `IExcelGateway` (interface), **unit-testable with a fake gateway**, no real Excel needed |
| `Domain.FormatRegistry` / `FormatDefinition` / `AccountingFormatBuilder` / `FormatKeys` | 1:1 port of `modFormatEngine.bas`'s `GetFormatDef`/`AccountingFmt` and `modConfig.bas`'s `FMT_*` constants — pure C#, zero COM references, zero interfaces needed to test | `Domain/` — plain POCOs + static/pure classes, tested directly by xUnit like any ordinary C# logic |
| `Domain.AppConfig` | In-memory session state for `ForceAlign`/`ZeroDash` (no persistence this milestone — see Constraints); defaults `ForceAlign=false`, `ZeroDash=true` | `Domain/AppConfig.cs` — mutable POCO, owned by `AddInHost`, read by `FormatEngine`/`RibbonController` via a `Func<AppConfig>` provider |
| `IExcelGateway` / `IRangeHandle` (Abstractions) | The **single interface boundary** between business logic and Excel COM — analogous to `IOutlookGateway` in the sibling project | `Abstractions/IExcelGateway.cs`, `Abstractions/IRangeHandle.cs` — interfaces only |
| `ExcelGateway` (Services) | The **only** class that touches `Excel.Application`/`Excel.Range`/`Excel.Window` COM types: resolves `Application.Selection`, guards it's a `Range` (equivalent of `SafeSelection()`), sets `NumberFormat`/`HorizontalAlignment`, toggles `Application.ScreenUpdating` | `Services/ExcelGateway.cs` implementing `IExcelGateway` — **not** unit tested directly (would require live Excel); verified by manual/live smoke test only |
| `ILog` / `FileLog` | Cross-cutting logging, equivalent of `modUtils.bas`'s `Log`/`HandleError` | `Abstractions/ILog.cs`, `Infrastructure/FileLog.cs` — file-based (`%LocalAppData%\FinanceFmtTools\logs\`), same pattern as the sibling project (replaces VBA's Immediate-window/hidden-sheet logging, which has no equivalent outside the VBE) |
| About/Docs actions | Equivalent of `modUtils.bas`'s `ShowAbout`/`OpenDocsURL` | A small `Services/AboutPresenter.cs` (`MessageBox.Show` + `Process.Start(url)`) — trivial, not worth a WinForms `Form` unless the About box needs more than static text |

## Recommended Project Structure

```
src/
├── FinanceFmt.sln                        # solution: 2 projects (add-in + tests)
│
├── FinanceFmt/                           # PROJECT 1 — the COM add-in (product)
│   ├── FinanceFmt.csproj                 # SDK-style, net48, AnyCPU, buildable via `dotnet build`
│   │
│   ├── Connect.cs                        # COM entry point: IDTExtensibility2 + IRibbonExtensibility
│   │                                      # + one method per Ribbon-XML callback name (AutoDispatch)
│   │
│   ├── Composition/
│   │   └── AddInHost.cs                  # manual DI/wiring + lifecycle (Wire/Teardown)
│   │
│   ├── Abstractions/                     # interfaces — decouple Domain/Services from Excel COM
│   │   ├── IExcelGateway.cs              # the ONE boundary with the Excel Object Model
│   │   ├── IRangeHandle.cs               # wraps the one Range the gateway hands back
│   │   ├── IRibbonController.cs
│   │   └── ILog.cs
│   │
│   ├── Domain/                           # PURE logic — no COM, no Excel reference — testable by CLI
│   │   ├── FormatKeys.cs                 # port of modConfig.bas's FMT_* constants
│   │   ├── FormatDefinition.cs           # port of the FormatDef UDT
│   │   ├── FormatRegistry.cs             # port of GetFormatDef's Select Case registry
│   │   ├── AccountingFormatBuilder.cs    # port of AccountingFmt (3-section format string)
│   │   └── AppConfig.cs                  # ForceAlign/ZeroDash in-memory state + defaults
│   │
│   ├── Services/                         # IMPLEMENTATIONS (the ones that touch COM live here)
│   │   ├── ExcelGateway.cs               # impl. of IExcelGateway — ONLY class touching Excel.Range/Application
│   │   ├── FormatEngine.cs               # orchestrates Domain.FormatRegistry + IExcelGateway
│   │   ├── RibbonController.cs           # impl. of IRibbonController (embedded XML + getPressed)
│   │   └── AboutPresenter.cs             # About dialog + docs-link opener
│   │
│   ├── Infrastructure/
│   │   └── FileLog.cs                    # impl. of ILog (file-based, %LocalAppData%)
│   │
│   ├── Ribbon/
│   │   └── ribbon.xml                    # EmbeddedResource — customUI 2009/2010 schema
│   │
│   └── Properties/
│       └── AssemblyInfo.cs               # [assembly: ComVisible(false)] + version (ComVisible(true) only on Connect)
│
└── FinanceFmt.Tests/                      # PROJECT 2 — xUnit tests (not shipped/installed)
    ├── FinanceFmt.Tests.csproj            # net48, xUnit, PackageReference to Microsoft.NET.Test.Sdk
    ├── FakeExcelGateway.cs                # in-memory fake of IExcelGateway/IRangeHandle
    ├── FormatRegistryTests.cs             # every FMT_* key produces the expected NumberFormat string
    ├── AccountingFormatBuilderTests.cs    # ForceAlign/ZeroDash permutations (parity with VBA's AccountingFmt)
    ├── FormatEngineTests.cs               # "no Range selected" guard, screen-updating suppression, alignment
    └── RibbonControllerTests.cs           # embedded XML loads + contains expected control ids; GetPressed logic
```

### Structure Rationale

- **`Domain/`:** Exactly the part of the VBA architecture (`modFormatEngine.bas` + the format-key constants half of `modConfig.bas`) that the milestone requires to be independently unit-testable. Nothing in this folder may reference `Microsoft.Office.Interop.Excel` — that constraint is what makes `dotnet test` runnable on a CI agent that doesn't even have Excel installed.
- **`Abstractions/`:** Interfaces only. `IExcelGateway` is the direct Excel-flavored equivalent of the sibling project's `IOutlookGateway` — "the fronteira única com o Object Model; nenhuma outra classe toca COM" (its own doc comment, verified in `~/pessoal/outlook-classic-delay-send/src/UndoSend/Abstractions/IOutlookGateway.cs:8-11`). Keeping it a separate folder from `Domain/` signals "this is the seam you fake in tests," matching the sibling project's `FakeOutlookGateway.cs` pattern.
- **`Services/`:** Where COM is actually touched (`ExcelGateway`) sits next to orchestration logic that depends only on interfaces (`FormatEngine`, `RibbonController`). This mirrors the sibling project's `Services/` folder exactly (`OutlookGateway.cs` + `SendInterceptor.cs` + `RibbonController.cs` side by side), and is why `FormatEngine`/`RibbonController` can be tested with fakes while `ExcelGateway` cannot.
- **`Composition/`:** Isolates "wiring code" (which cannot reasonably be unit tested — it constructs real COM-touching objects from an `object Application` parameter) from everything else, so the untested surface area is as small and clearly marked as possible.
- **`FinanceFmt.Tests/`:** A second project, not installed, referencing `FinanceFmt.csproj` via `<ProjectReference>` — same pattern as `UndoSend.Tests.csproj` in the sibling repo (confirmed: `IsPackable=false`, `Microsoft.NET.Test.Sdk` + `xunit` + `xunit.runner.visualstudio` package references, `net48` target so it runs against the exact same runtime as the add-in).

## Architectural Patterns

### Pattern 1: COM entry point stays thin; `IDispatch`/`AutoDispatch` is required, not optional

**What:** `Connect` is `[ComVisible(true)] [Guid("…")] [ProgId("…")] [ClassInterface(ClassInterfaceType.AutoDispatch)]`, implementing `IDTExtensibility2` (lifecycle) and `Office.IRibbonExtensibility` (`GetCustomUI`). Every method referenced by name in the Ribbon XML (`onAction`, `getPressed`, `onLoad`, `getEnabled`, etc.) must ALSO be a plain public method on `Connect`, because Excel calls Ribbon callbacks by name through `IDispatch.GetIDsOfNames`/`Invoke` — they are not part of `IRibbonExtensibility` itself, which only defines `GetCustomUI`.

**When to use:** Always, for this add-in shape (pure COM add-in without VSTO). This is not Excel-specific — it is how every Fluent-Ribbon Office host (Word, Excel, Outlook, PowerPoint) resolves Ribbon callbacks, confirmed both by the sibling project's own diagnosed incident (`Connect.cs:21-27`: *"AutoDispatch (NÃO None): … Com ClassInterfaceType.None a classe não expõe um dispinterface, então esses métodos ficam invisíveis e o Office não consegue chamá-los (botão nunca habilita / clique não age)"* — this was live-diagnosed, not theoretical) and by current documentation of `IRibbonExtensibility`/`IDispatch` (confirmed via WebSearch — MEDIUM/HIGH confidence, matches sibling project's live-validated finding).

**Trade-offs:** `AutoDispatch` means every `public` method on `Connect` becomes COM-callable — keep `Connect` to *only* the methods Excel needs to call (lifecycle + ribbon callbacks) and push everything else to `AddInHost`, so the COM-visible surface stays small and intentional.

**Example (Excel-flavored, adapted from `Connect.cs`):**
```csharp
[ComVisible(true)]
[Guid("<NEW-GUID-GENERATED-FOR-THIS-ADDIN>")]
[ProgId("FinanceFmtTools.Connect")]
[ClassInterface(ClassInterfaceType.AutoDispatch)]
public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
{
    private readonly AddInHost _host = new AddInHost();

    public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
    {
        try { _host.Wire(Application); } catch (Exception ex) { TryLog(ex); }
    }

    public string GetCustomUI(string RibbonID) => _host.Ribbon?.GetCustomUiXml() ?? string.Empty;

    public void OnRibbonLoad(Office.IRibbonUI ribbonUi) => _host.Ribbon?.CacheRibbon(ribbonUi);

    // ONE generic callback for every format button (tag carries the FMT_* key) —
    // mirrors modFormatEngine.bas's key-based registry more closely than 12 near-duplicate
    // VBA-style RibbonFin2D/RibbonPct4D/... methods would.
    public void OnApplyFormat(Office.IRibbonControl control)
    {
        try { _host.FormatEngine?.ApplyFormatToSelection(control?.Tag ?? string.Empty); }
        catch (Exception ex) { TryLog(ex); }
    }

    public void OnToggleConfig(Office.IRibbonControl control, bool pressed)
        => _host.SetConfigFlag(control?.Tag, pressed);

    public bool GetConfigPressed(Office.IRibbonControl control)
        => _host.Ribbon?.GetPressed(control?.Tag ?? string.Empty) ?? false;
}
```

### Pattern 2: Business logic behind a single Excel gateway interface

**What:** `IExcelGateway` (+ `IRangeHandle`) is the only type in the codebase that any code outside `Services/ExcelGateway.cs` is allowed to know maps to `Excel.Application`/`Excel.Range`. `FormatEngine` depends only on `IExcelGateway`; tests construct a `FakeExcelGateway`/`FakeRangeHandle` instead of touching real Excel — directly analogous to `FakeOutlookGateway` in the sibling project's test suite (`UndoSend.Tests/FakeOutlookGateway.cs`).

**When to use:** Any COM object that (a) needs to be null-checked/guarded (VBA's `SafeSelection()`), (b) needs COM lifetime management (`Marshal.ReleaseComObject`), or (c) needs its state read/written from business logic that should be unit-testable.

**Trade-offs:** Slightly more boilerplate (an interface + a fake per COM surface) than calling `Application.Selection` directly, but it is the entire reason the format engine can be covered by `dotnet test` without Excel installed on the build agent — which is an explicit milestone requirement.

**Example:**
```csharp
// Abstractions/IExcelGateway.cs
public interface IExcelGateway
{
    // Equivalent of modUtils.bas's SafeSelection(): returns false (and logs) if the
    // current selection is not a Range (e.g. a Shape/Chart is selected).
    bool TryGetSelectionRange(out IRangeHandle range);

    // Equivalent of the Application.ScreenUpdating = False/True wrap in ApplyFormat.
    void WithScreenUpdatingSuspended(Action action);
}

public interface IRangeHandle
{
    void SetNumberFormat(string numberFormat);
    void SetHorizontalAlignment(CellHorizontalAlignment alignment); // small enum, NOT Excel.XlHAlign,
                                                                     // so Domain never references Excel types
}

// Services/FormatEngine.cs — depends ONLY on the interface, fully unit-testable
public sealed class FormatEngine
{
    private readonly IExcelGateway _gateway;
    private readonly Func<AppConfig> _config;
    private readonly ILog _log;

    public void ApplyFormatToSelection(string formatKey)
    {
        FormatDefinition def = FormatRegistry.Resolve(formatKey, _config());
        if (!_gateway.TryGetSelectionRange(out IRangeHandle range))
        {
            _log.Warn($"FormatEngine: no Range selected, ignoring '{formatKey}'.");
            return;
        }
        _gateway.WithScreenUpdatingSuspended(() =>
        {
            range.SetNumberFormat(def.NumberFormat);
            if (def.Alignment.HasValue) range.SetHorizontalAlignment(def.Alignment.Value);
        });
    }
}
```

### Pattern 3: Registry-of-pure-functions for the format catalog

**What:** `Domain.FormatRegistry.Resolve(key, config) -> FormatDefinition` is a direct, mechanical port of `GetFormatDef`'s `Select Case` (`src/modFormatEngine.bas:81-170`) plus `AccountingFmt` (`src/modFormatEngine.bas:188-222`). It takes the format key (string or a `FormatKeys` constant) and the live `AppConfig` (for `ForceAlign`/`ZeroDash`) and returns a value object — no side effects, no COM, no `IExcelGateway` dependency at all.

**When to use:** For every format-string-generation rule that the VBA README documents as "the extension point" (`README.md:237-241`) — adding a new format becomes "add a `FormatKeys` constant + a `case` in `FormatRegistry` + a `<button tag="...">` in `ribbon.xml`", no `Connect.cs` change needed because of Pattern 1's generic `OnApplyFormat` callback.

**Trade-offs:** None significant — this is strictly more testable than the VBA original, at the cost of needing a `FormatDefinition` value type (VBA's `FormatDef` UDT) instead of a bag of module-level globals.

## Data Flow

### Request Flow — "apply a format to the selected cells" (mirrors VBA's Primary Request Path)

```
User clicks "Fin 2D" button (Ribbon/ribbon.xml, onAction="OnApplyFormat" tag="FIN_2D")
    ↓ (Excel → IDispatch.Invoke on Connect, by name)
Connect.OnApplyFormat(IRibbonControl control)
    ↓ delegates, no logic
AddInHost.FormatEngine.ApplyFormatToSelection(control.Tag)
    ↓
FormatEngine:
  1. Domain.FormatRegistry.Resolve("FIN_2D", currentConfig)   ← PURE, no COM
       → reads AppConfig.ForceAlign / AppConfig.ZeroDash (mirrors CFG_FORCE_ALIGN/CFG_ZERO_DASH)
       → Domain.AccountingFormatBuilder builds the 3-section NumberFormat string
       → returns FormatDefinition { NumberFormat, Alignment? }
  2. IExcelGateway.TryGetSelectionRange(out range)             ← the ONLY COM touch so far
       → false + log warning if Selection isn't a Range (mirrors SafeSelection())
  3. IExcelGateway.WithScreenUpdatingSuspended(() => {
       range.SetNumberFormat(def.NumberFormat);
       if (def.Alignment.HasValue) range.SetHorizontalAlignment(def.Alignment.Value);
     })
    ↓
Services/ExcelGateway (real impl.): Application.ScreenUpdating = false;
  ((Excel.Range)_selection).NumberFormat = ...; ((Excel.Range)_selection).HorizontalAlignment = ...;
  Application.ScreenUpdating = true;
    ↓
Excel repaints the selected cells with the new format
```

### Configuration Change Flow — "toggle a checkbox" (no persistence, per this milestone)

```
User toggles "Zero contábil" checkbox (ribbon.xml, onAction="OnToggleConfig" tag="ZERO_DASH", getPressed="GetConfigPressed")
    ↓
Connect.OnToggleConfig(control, pressed) → AddInHost.SetConfigFlag("ZERO_DASH", pressed)
    ↓
AddInHost mutates the single in-memory AppConfig instance shared by FormatEngine/RibbonController
    ↓
Next "apply format" call reads the updated AppConfig.ZeroDash live (same pattern as VBA reading the
CFG_ZERO_DASH global at AccountingFmt-build time — no separate "reload" step needed)
    ↓
(Optional) AddInHost.Ribbon.InvalidateControl(control.Id) → Excel re-queries getPressed immediately
```

There is **no** analogue to `modUtils.bas`'s `LoadConfig`/`SaveConfig`/`CustomXMLPart` persistence or `ThisWorkbook.bas`'s `Workbook_BeforeClose` fallback save — both are deliberately out of scope for this milestone (`PROJECT.md` Out of Scope: *"Persistência das preferências … removida deliberadamente"*). `AppConfig` is created fresh with its documented defaults (`ForceAlign=false`, `ZeroDash=true`) every time `AddInHost.Wire()` runs (i.e., every Excel session), and simply discarded on `OnDisconnection`.

### Key Data Flows

1. **Format application (Ribbon button → `Range.NumberFormat`):** Connect → AddInHost.FormatEngine → Domain.FormatRegistry (pure) → IExcelGateway (COM boundary) → real `Excel.Range`. This is the flow the milestone requires to be unit-testable up to (but not including) the `IExcelGateway` boundary.
2. **Ribbon bootstrap (`onLoad`):** Connect.OnRibbonLoad → AddInHost.Ribbon.CacheRibbon(ribbonUi) — caches the `IRibbonUI` handle needed later for `InvalidateControl`/`Invalidate`, exactly like VBA's `mRibbon As IRibbonUI` module-level singleton (`src/modRibbon.bas:10`).
3. **Config toggle (checkbox → in-memory flag):** Connect.OnToggleConfig → AddInHost mutates `AppConfig` → next format application reads it. No disk I/O anywhere in this milestone (a deliberate simplification vs. VBA's `CustomXMLPart` persistence).
4. **About / docs link:** Connect.OnAbout / Connect.OnOpenDocs → AddInHost.AboutPresenter — no format-engine or config involvement, mirrors VBA's `ShowAbout`/`OpenDocsURL` being simple leaf actions in `modUtils.bas`.

## Scaling Considerations

This is a single-user, single-process Office add-in — there is no multi-tenant/multi-user scaling axis. The only "scale" that matters is the number of format definitions and Ribbon controls over time.

| Scale | Architecture Adjustments |
|-------|--------------------------|
| Today (11 formats, 2 checkboxes) | `FormatRegistry` as a single `switch`/dictionary is fine, exactly like VBA's `Select Case`. |
| Adding several new formats/categories | Still fine as a `switch`/dictionary in `FormatRegistry` — do not introduce a plugin/reflection-based registry; that would be over-engineering for a catalog this size (VBA's own README explicitly documents the flat `Select Case` as the intended extension point, and there's no evidence the catalog will grow beyond a few dozen entries). |
| Hypothetical: format catalog becomes data-driven / user-editable | Only then consider externalizing `FormatRegistry` to a JSON/XML resource loaded at startup — out of scope for this migration (`PROJECT.md` Out of Scope: *"Novos formatos de número ou funcionalidades além das já existentes"*). |

### Scaling Priorities

1. **Not a bottleneck concern for this project** — the "scale" axis that actually matters is developer/maintainer scale (how easy is it to add a format, keep the Ribbon and the registry in sync, and keep tests fast). The layered structure above already optimizes for that.
2. **Startup latency** — `AddInHost.Wire()` runs synchronously inside `OnConnection`, which Excel calls on its main STA thread at Excel startup; keep it fast (no network calls, no blocking I/O) exactly as VBA's `OnRibbonLoad` does today (in-memory config only, no file/registry read in the hot path).

## Anti-Patterns

### Anti-Pattern 1: Business logic ("format engine") referencing `Microsoft.Office.Interop.Excel` types directly

**What people do:** Put the `NumberFormat`/`HorizontalAlignment`-string-building logic straight into a class that also holds an `Excel.Range` field, "because it's simpler."
**Why it's wrong:** The moment `Domain`/format-resolution code references `Excel.Range` or `Excel.Application`, it can no longer be exercised by `dotnet test` without a real Excel COM object, which defeats the explicit milestone goal ("abstractions that isolate the Excel API" so tests don't need a real Excel instance). It also silently reintroduces the exact limitation the VBA codebase always had (no unit tests possible) that this migration is meant to fix.
**Do this instead:** Keep `Domain/` 100% free of any `using Microsoft.Office.Interop.Excel;` (or `Office.*`) — verify this in CI with `dotnet test` running with **no Excel installed** on the runner (confirms the boundary is real, not just aspirational). All COM access funnels through `IExcelGateway`.

### Anti-Pattern 2: `Connect` (COM entry point) containing conditional/business logic

**What people do:** Put `if (control.Tag == "FIN_2D") { ... build format string here ... }`-style logic directly in the Ribbon callback methods on `Connect`, because that's the first place the callback fires.
**Why it's wrong:** This is exactly the VBA anti-pattern the milestone is explicitly trying to avoid re-creating — `modRibbon.bas`'s own documented convention is that ribbon callbacks are "exactly one line of logic" (`src/modRibbon.bas:5-7`), delegating to the engine. If `Connect` grows business logic, it becomes untestable (COM classes can't easily be instantiated/exercised by xUnit without a COM host) and duplicates responsibility that `FormatEngine`/`FormatRegistry` already own.
**Do this instead:** Every `Connect` method must be a one-line delegation to `AddInHost` (mirrors the sibling project's `Connect.cs`, where literally every method body is `try { _host.X(...); } catch { TryLog(...); }`).

### Anti-Pattern 3: Letting exceptions escape `Connect`'s methods

**What people do:** Assume normal .NET exception propagation is fine because "the caller (Excel) will just show an error."
**Why it's wrong:** An unhandled exception thrown from a COM `IDTExtensibility2`/`IRibbonExtensibility`/Ribbon-callback method can cause Excel to disable the add-in (demote its `LoadBehavior`) or silently fail to render Ribbon controls, with no useful diagnostic to the end user — this is exactly the failure mode the sibling project's `Connect.cs` comments call out (*"R5: não rebaixar LoadBehavior 3→2"*).
**Do this instead:** Every `Connect` method wraps its call to `AddInHost`/services in `try/catch`, logging via `ILog` and returning a safe default (`string.Empty`, `false`, or simply returning) rather than propagating.

## Integration Points

### External Services

| Service | Integration Pattern | Notes |
|---------|---------------------|-------|
| Excel (host application) | COM in-process server, registered via `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins\<ProgId>` (`LoadBehavior=3`) + `HKEY_CLASSES_ROOT\CLSID\{guid}` — same HKCU-only, no-admin registration model already used by the sibling project's installer | The installer (PowerShell, separate milestone concern) must write registry values for the **same fixed CLSID/ProgID** the `Connect` class declares — this is a hard contract between `Connect.cs`'s attributes and the installer script, analogous to `BUILD.md §5`'s "GUID do add-in fixado" table in the sibling project. |
| Excel Object Model (COM) | `Application`/`Range`/`Window` accessed exclusively through `IExcelGateway`/`ExcelGateway` | Confirmed pattern: sibling project's `IOutlookGateway`/`OutlookGateway` is the literal template — "ÚNICA classe que fala com o Outlook Object Model" (`OutlookGateway.cs:10-15`); the Excel equivalent inherits the same COM-lifetime discipline (release intermediate COM objects touched during a call; `Range` handles returned to callers are comparatively short-lived per format application, unlike Outlook's long-held folder/session references, so the COM-release burden is lighter but the single-gateway discipline still applies). |
| Excel/Office PIAs (build-time dependency) | `Microsoft.Office.Interop.Excel` is available as an **official Microsoft NuGet package** (`Microsoft.Office.Interop.Excel`, e.g. 16.0.18925.20022 on nuget.org, confirmed via WebSearch) — a materially easier path than the sibling project had for Outlook (no official Outlook PIA NuGet package exists, hence its `lib/`-vendored-GAC-DLL approach). `stdole` similarly has an official NuGet package. `Microsoft.Office.Core` (needed for `Office.IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`) and `Extensibility` (`IDTExtensibility2`) do **not** have an equally clear-cut official NuGet package — the proven fallback (already working in the sibling repo) is to vendor `OFFICE.DLL`/`Extensibility.dll` into this repo's own `lib/` folder and reference them via `<Reference HintPath="…">`, `SpecificVersion=false`, `Private=true` — these two files are **not Outlook-specific** (Office Core and the Add-In Designer's `IDTExtensibility2` are shared across every Office host app), so they can plausibly be copied verbatim from `~/pessoal/outlook-classic-delay-send/lib/` rather than re-obtained from a machine with Excel installed. Verify version/GUID compatibility before relying on this (flagged MEDIUM confidence — worth a quick confirmation in Phase 1/setup rather than assumed). |
| GitHub Actions (`windows-latest`) CI | `dotnet restore`/`build`/`test` on the solution; since `Domain/` and `Services/FormatEngine.cs`+`RibbonController.cs` tests never instantiate real Excel COM objects, `dotnet test` can run on a CI runner **without Excel installed**, exactly as the sibling project's `UndoSend.Tests` runs without Outlook installed | This is the concrete reason the milestone's CI requirement ("Pipeline de CI … que compila, testa, empacota") is achievable without a self-hosted runner that has Office installed. |

### Internal Boundaries

| Boundary | Communication | Notes |
|----------|---------------|-------|
| Ribbon XML ↔ `Connect` | `onAction`/`getPressed`/`onLoad` attribute values are method **names**, resolved by Excel via `IDispatch` at runtime — no compile-time link between the XML and `Connect`; a typo in either place fails silently (control never enables / click does nothing) | Mirrors VBA exactly (same failure mode existed there); mitigate with `RibbonControllerTests` that assert the embedded XML contains the expected control ids/callback names (as the sibling project already does in `RibbonControllerTests.cs`). |
| `Connect` ↔ `AddInHost` | Direct method calls, one-line delegation, wrapped in try/catch at the `Connect` side | `Connect` never touches `IExcelGateway`/`Domain` directly — always through `AddInHost`. |
| `AddInHost` ↔ `FormatEngine`/`RibbonController`/`AppConfig` | Direct method calls; `AppConfig` shared via closures (`Func<AppConfig>`) rather than re-fetched from a store, since there is no persistence layer this milestone | If persistence is reintroduced in a future milestone, this is the seam to insert an `IConfigStore` (the sibling project's `IConfigStore`/`ConfigService` is the ready-made template — not needed now per `PROJECT.md` Out of Scope). |
| `FormatEngine` ↔ `Domain.FormatRegistry` | Direct method call, pure function, no interface needed (it's not something you'd ever want to fake in a test — you want to exercise the real registry) | Contrast with `IExcelGateway`, which you always want to fake in `FormatEngine` tests. |
| `FormatEngine`/`RibbonController` ↔ `IExcelGateway` | Interface call — this is the one boundary with two implementations: `ExcelGateway` (production) and `FakeExcelGateway` (`FinanceFmt.Tests`) | This is the single most important boundary in the whole architecture — it is what makes the milestone's testability requirement achievable. |

## Suggested Build Order (dependency-driven, informs roadmap phasing)

1. **`Domain/`** (`FormatKeys`, `FormatDefinition`, `FormatRegistry`, `AccountingFormatBuilder`, `AppConfig`) — zero dependencies on anything Excel/COM. Buildable and 100% unit-testable with `dotnet test` before a single line of COM interop exists. This is where format-string parity with the VBA original (`AccountingFmt`'s `ForceAlign`/`ZeroDash` permutations, date/percent/text formats) should be locked down first, since it's the highest-value, lowest-risk-to-test slice and the milestone's core testability requirement.
2. **`Abstractions/`** (`IExcelGateway`, `IRangeHandle`, `IRibbonController`, `ILog`) — interfaces only, depend only on `Domain` types (e.g., a small `CellHorizontalAlignment` enum so `Domain`/`Abstractions` never reference `Excel.XlHAlign`). No implementation yet.
3. **`Services/FormatEngine.cs`** + **`Services/RibbonController.cs`** (business orchestration against the interfaces) — depends on `Domain` + `Abstractions` only, **not** on any real Excel COM type. Fully testable here with `FakeExcelGateway`/a fake `AppConfig` provider — this is the second wave of unit tests (guard-clause behavior: "no Range selected", screen-updating suppression called, alignment only set when defined; embedded-XML-loads-and-contains-expected-ids for the Ribbon controller).
4. **`Services/ExcelGateway.cs`** (the real COM implementation) — first point in the build where `Microsoft.Office.Interop.Excel`/`Microsoft.Office.Core` types are referenced; requires the interop assemblies to be resolvable (NuGet package + vendored `lib/` DLLs, see Integration Points). Not unit tested; verified by manual/live smoke test in real Excel.
5. **`Ribbon/ribbon.xml`** (embedded resource) — can be authored in parallel with step 3/4 (it's a straight XML port of `customUI14.xml` with `onAction`/`tag` reshuffled per Pattern 1/3); only needs to exist before step 6's manual smoke test.
6. **`Composition/AddInHost.cs`** — depends on everything above; wires real implementations together from the `object Application` parameter received in `OnConnection`. Not unit tested (pure wiring), same as the sibling project.
7. **`Connect.cs`** (COM entry point, attributes + `IDTExtensibility2`/`IRibbonExtensibility` + named Ribbon callbacks) — depends on `AddInHost`; requires a **freshly generated GUID** (never reuse the Outlook sibling's) and a fixed `ProgId`, which becomes a hard contract for the (separate-milestone) PowerShell installer. Verified only by registering + loading in a real Excel instance.
8. **`FinanceFmt.Tests` project** — can and should be scaffolded alongside step 1 (not after everything else), since steps 1–3 are exactly what it exists to cover; steps 4–7 are deliberately outside its scope (verified manually/live instead).

**Why this order matters for phasing:** it lets the roadmap put "format engine + config, fully unit tested" in an early phase that never requires a Windows machine with Excel installed to validate (only `dotnet test`), and defer "real COM wiring + live Excel verification" to a later phase where manual/live smoke testing is expected and budgeted for — mirroring how the sibling project's own phase plan separated "Fase 3 (Implementation, CLI-verifiable)" from "Fase 5/6 (registro + Outlook ao vivo)" in `~/pessoal/outlook-classic-delay-send/.planning/ARCHITECTURE-CSHARP.md` (§ "Confirmados na Fase 3" vs. "Diferidos para Fase 5/6").

## Sources

- Existing VBA architecture map: `.planning/codebase/ARCHITECTURE.md`, `.planning/codebase/STRUCTURE.md` (this repository) — the layering being mirrored.
- Reference implementation (same problem, solved for Outlook, HIGH confidence — working, live-validated code): `~/pessoal/outlook-classic-delay-send/src/UndoSend/Connect.cs`, `Composition/AddInHost.cs`, `Abstractions/IOutlookGateway.cs`, `Services/OutlookGateway.cs`, `Services/RibbonController.cs`, `Abstractions/IRibbonController.cs`, `Ribbon/ribbon.xml`, `Domain/SendInterceptorLogic.cs`, `Domain/AppConfig.cs`, `Abstractions/IConfigStore.cs`; test patterns: `UndoSend.Tests/FakeOutlookGateway.cs`, `UndoSend.Tests/RibbonControllerTests.cs`, `UndoSend.Tests/SendInterceptorLogicTests.cs`; project files: `UndoSend/UndoSend.csproj`, `UndoSend.Tests/UndoSend.Tests.csproj`.
- Reference project's own architecture/build docs (design rationale, live-diagnosed pitfalls): `~/pessoal/outlook-classic-delay-send/.planning/ARCHITECTURE-CSHARP.md` (§1–§4), `~/pessoal/outlook-classic-delay-send/BUILD.md` (§3–§6, PIA/interop referencing strategy).
- [IRibbonExtensibility Interface (Microsoft.Office.Core)](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.core.iribbonextensibility?view=office-pia) — confirms `GetCustomUI` is the only member of the interface itself; Ribbon callbacks are resolved separately via `IDispatch`. MEDIUM-HIGH confidence (official docs).
- [IRibbonUI object (Office VBA reference)](https://learn.microsoft.com/en-us/office/vba/api/office.iribbonui) — confirms `onLoad`/`IRibbonUI` caching pattern is the standard mechanism across Office hosts, not Outlook-specific. HIGH confidence.
- [NuGet Gallery — Microsoft.Office.Interop.Excel](https://www.nuget.org/packages/microsoft.office.interop.excel/) — confirms an official Microsoft-signed NuGet package exists for the Excel PIA (16.0.18925.20022 at time of research), a materially better path than the sibling project's Outlook-PIA vendoring workaround. MEDIUM confidence (dependency/TFM details not independently build-verified in this research pass — recommend a quick spike in the setup/stack phase).
- [NuGet Gallery — stdole](https://www.nuget.org/packages/stdole) — confirms an official NuGet package for `stdole` also exists. MEDIUM confidence.
- WebSearch: "C# Excel COM add-in IDTExtensibility2 IRibbonExtensibility ClassInterfaceType AutoDispatch ribbon callback IDispatch" — corroborates (not Excel-exclusive, generic Office COM add-in pattern) that `ClassInterfaceType.AutoDispatch` + implementing callback methods as plain public members is the standard shape; consistent with, and less authoritative than, the sibling project's own live-diagnosed finding. LOW-MEDIUM confidence on its own, raised to MEDIUM-HIGH by agreement with the sibling project's live-validated result.

---
*Architecture research for: C# COM Excel add-in migration (finance-fmt-tools)*
*Researched: 2026-07-10*
