# Phase 2: Abstractions & Orchestration - Research

**Researched:** 2026-07-11
**Domain:** C# seam design (pure interfaces over Excel COM) + orchestration logic, unit-tested with hand-written fakes, zero `Microsoft.Office.Interop.Excel` references
**Confidence:** HIGH (grounded directly in this repo's Phase 1 code/VBA source, plus a structurally analogous sibling project explicitly cited as this project's workflow inspiration)

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
None locked — discuss phase was skipped (`workflow.skip_discuss`, full autonomous `/gsd-autonomous` run per `02-CONTEXT.md`).

### Claude's Discretion
All implementation choices are at Claude's discretion — discuss phase was skipped per user setting (full autonomous run, `/gsd-autonomous`). Use ROADMAP phase goal, success criteria, and codebase conventions to guide decisions. Key constraints carried over from Phase 1: zero real `Microsoft.Office.Interop.Excel` references in the code under test in this phase (that's Phase 3's job) — `IExcelGateway`/`IRangeHandle` must be pure C# interfaces with a hand-written fake implementation for tests, still buildable/testable on this Linux dev environment via `dotnet test`. `RibbonController`'s session-state defaults must match the VBA behavior: "Alinhar à direita" off by default, "Zero contábil" on by default (see src/modConfig.bas and src/modRibbon.bas for the exact VBA semantics being ported). The FMT-06 guard clause (invalid selection → friendly warning, never throws/crashes) is this phase's other core deliverable.

### Deferred Ideas (OUT OF SCOPE)
None — discuss phase skipped.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| FMT-06 | Aplicar um formato com uma seleção inválida (Chart/Shape em vez de Range) mostra uma mensagem amigável em vez de quebrar o add-in | `IExcelGateway.TryGetSelectedRange(out IRangeHandle)` Try-pattern guard (Pattern 1) + `FormatEngine.ApplyToSelection`'s no-throw, logged-warning guard clause (Pattern 2, Pitfall 2) — proven via `FakeExcelGateway`/`SpyLog` in `dotnet test` without any live Excel/COM instance. The actual user-facing "friendly message" (MsgBox-equivalent) is explicitly deferred to Phase 3 (see Pitfall 2); this phase proves the orchestration-level contract only. |
</phase_requirements>

## Project Constraints (from CLAUDE.md)

These directives apply to all work in this phase and were checked against every recommendation below:

- **Platform**: Windows + Excel 2016+ only — this phase adds no platform-specific code (still zero COM), so it remains buildable/testable on Linux, consistent with Phase 1.
- **Tooling**: VS Code + `dotnet` CLI only, no Visual Studio dependency — all new code in this phase must build via `dotnet build`/`dotnet test` exactly as Phase 1 does; no `.sln`-only or VS-IDE-only project features.
- **Runtime**: .NET Framework 4.8, buildable with the .NET 8 SDK — new files must live in (or extend) a project multi-targeting `net48;net8.0` exactly like `FinanceFmtTools.Engine`, not a net8.0-only project (that constraint applies to the Tests project only, per Phase 1 precedent).
- **Installation**: HKCU-only, no admin — not directly relevant to this phase's scope (Phase 4's concern), but reinforces that no COM registration code belongs in this phase.
- **UX compatibility**: Same buttons/shortcuts visible on the Ribbon — this phase must not change `src/customUI14.xml`'s control IDs, labels, or callback names, since Phase 3 will wire real callbacks to those exact names; Phase 2 reuses the file byte-for-byte (see Pattern 3).
- **GSD Workflow Enforcement** (CLAUDE.md): file-changing work must go through a GSD command (`/gsd-execute-phase` for planned phase work) — noted for the planner/executor, not applicable to this research-only step.

## Summary

Phase 2 adds one new orchestration layer on top of Phase 1's pure `FinanceFmtTools.Engine` format-string logic: a COM-free `IExcelGateway`/`IRangeHandle` seam, a `FormatEngine` orchestrator that resolves a format key via Phase 1's `FormatRegistry` and applies it through that seam (including the FMT-06 invalid-selection guard), and a `RibbonController` that owns in-memory checkbox session state and loads the embedded Ribbon XML resource. None of this requires any real COM type — it is proven entirely by `dotnet test` using hand-written fakes, exactly as Phase 1 was.

The most important finding is a **default-value trap**: this repo's own VBA source contains *two different, mutually contradictory* answers for what "Alinhar à direita"/"Zero contábil" default to (see Pitfall 1). Neither VBA reading matches the values this phase must actually implement. The authoritative values are stated explicitly in `REQUIREMENTS.md` (RIB-02/RIB-03) and `02-CONTEXT.md`: **ForceAlign defaults to `false` (off), ZeroDash defaults to `true` (on)** — these come from the persistence layer being deliberately deleted in this migration, not from either VBA file. The planner must not grep-copy a "VBA default" without checking this.

The second major finding is a scope boundary between Phase 2 and Phase 3's `RibbonController`/`Connect.cs` split. `RIB-01..04` (the actual live Ribbon behavior) are traced to Phase 3, not Phase 2 — Phase 2's `RibbonController` success criterion is deliberately narrow: load an embedded XML resource (proven testable with zero COM, verified empirically in this session) and answer in-memory checkbox-state queries. No live `IRibbonUI` caching, no `InvalidateControl`, no image loading is needed in this phase.

A local sibling project (`outlook-classic-delay-send`, the same one CLAUDE.md cites as this project's dev/build/release inspiration) contains a directly analogous, already-shipped `IOutlookGateway`/`IRibbonController`/`ILog` seam with hand-written fakes and constructor-injected dependencies. It is not an authoritative spec for this project, but it is a concrete, working precedent built by the same author for the same class of problem (COM Add-in without VSTO, `dotnet test`-first), and its patterns (Try-pattern gateway methods, `ILog.Warn/Info/Error`, resource-loaded-by-suffix Ribbon XML, instance-based stateful services vs. static pure-domain classes) transfer directly.

**Primary recommendation:** Add `FormatEngine.cs`, `Abstractions/IExcelGateway.cs`, `Abstractions/IRangeHandle.cs`, `Abstractions/ILog.cs`, and `RibbonController.cs` to the *existing* `FinanceFmtTools.Engine` project (still `net48;net8.0`, still zero COM references) rather than creating a new project — Phase 3 is what introduces the first COM-referencing project. Keep `FormatEngine`/`FormatRegistry`/`AccountingFormatBuilder` as static, parameterized, dependency-free functions (Phase 1's established convention); make `RibbonController` an instance class holding session state (it must not be static/shared, both for correctness and for `dotnet test`'s default cross-class parallelism).

## Architectural Responsibility Map

This project has no web/API/DB tiers — it is a single-process desktop COM add-in. Tiers are adapted accordingly:

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Format key → `NumberFormat` string resolution | Domain / Format Engine Core (Phase 1, done) | — | Pure string derivation, already proven COM-free |
| Selection-validity guard (Range vs. Chart/Shape) | Orchestration (Phase 2, this phase) | Excel COM Host (Phase 3 real impl) | The guard's *logic* ("no valid range → warn, don't throw") is a COM-free business rule; only the real "what is `Selection`" check needs COM, and that's hidden behind `IExcelGateway` |
| Applying `NumberFormat`/alignment to a range | Orchestration (Phase 2, via `IRangeHandle`) | Excel COM Host (Phase 3 real impl) | Orchestration decides *what* to set; the gateway implementation decides *how* (fake in Phase 2, real `Range` in Phase 3) |
| Ribbon tab/group/button/checkbox declaration | Ribbon UI (XML, done — `src/customUI14.xml`) | — | Static declarative XML, no logic, already written for the VBA version and reusable verbatim |
| Checkbox session-state defaults + mutation | Orchestration (`RibbonController`, Phase 2, this phase) | Ribbon UI (Phase 3 `getPressed`/`onAction` wiring) | State (and its defaults) lives in `RibbonController`; the live Ribbon only queries/mutates it in Phase 3 |
| Embedded Ribbon XML loading (`GetCustomUiXml`-equivalent) | Orchestration (Phase 2, this phase) | Ribbon UI | Loading a `.NET` embedded resource is 100% COM-free and `dotnet test`-provable now; Phase 3 feeds the same string into the real `IRibbonExtensibility.GetCustomUI` COM method |
| Live `IRibbonUI` caching / `InvalidateControl` | Excel COM Host (Phase 3) | — | Requires a real `IRibbonUI` reference that doesn't exist until Excel loads the add-in; not needed in Phase 2 at all (see Open Questions) |
| About dialog / docs link | Ribbon UI callback (Phase 3, RIB-04) | — | Simple pass-through, explicitly out of Phase 2's requirement scope (traced to Phase 3 in REQUIREMENTS.md) |

## Standard Stack

No new packages are introduced in this phase. Phase 2 reuses Phase 1's already-verified toolchain:

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| xUnit | 2.9.3 | Test framework | Already pinned and proven in Phase 1 [VERIFIED: src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj] |
| xunit.runner.visualstudio | 3.1.5 | Test discovery/runner | Already pinned in Phase 1 [VERIFIED: same csproj] |
| Microsoft.NET.Test.Sdk | 17.14.1 | Test SDK | Already pinned in Phase 1 [VERIFIED: same csproj] |
| Microsoft.NETFramework.ReferenceAssemblies | 1.0.3 | net48 build-only reference assemblies | Already pinned in Phase 1, build-time only, no runtime footprint [VERIFIED: src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj] |

### Supporting
None. This phase deliberately adds **zero** new NuGet packages — the `ILog`/`IExcelGateway`/`IRangeHandle` abstractions are plain hand-rolled interfaces (see Don't Hand-Roll section for why a full logging framework would be the wrong call here).

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Hand-rolled `ILog` interface | `Microsoft.Extensions.Logging.Abstractions` | Would be the *first* runtime NuGet dependency in `FinanceFmtTools.Engine`; adds an abstraction-only package for a single `Warn`/`Info`/`Error` call site — unjustified for this phase's scope. Revisit only if Phase 3+ needs structured logging across many components. |

**Installation:** None required — no new packages this phase. Confirm existing packages still resolve with:
```bash
dotnet build src/FinanceFmtTools.sln -c Release
```
Verified in this session: build succeeds, 0 Warning(s)/0 Error(s), on the same `dotnet 8.0.422` install used in Phase 1 [VERIFIED: ran `dotnet build` in this research session].

## Package Legitimacy Audit

Not applicable — this phase installs no new external packages. All packages in use were already audited/pinned during Phase 1 (see `.planning/phases/01-format-engine-core/01-RESEARCH.md`).

## Architecture Patterns

### System Architecture Diagram

```
                         ┌─────────────────────────────────────────┐
                         │   Excel Ribbon Click (Phase 3, live)     │
                         │   e.g. onAction="RibbonFin8D"            │
                         └───────────────────┬───────────────────────┘
                                             │ (not built until Phase 3)
                                             ▼
┌──────────────────────────────────────────────────────────────────────────┐
│ PHASE 2 — Orchestration (this phase, zero COM, dotnet-test-provable)      │
│                                                                            │
│  RibbonController                        FormatEngine                    │
│  ┌───────────────────────┐               ┌─────────────────────────┐    │
│  │ RibbonSessionConfig    │  reads        │ ApplyToSelection(       │    │
│  │  ForceAlign = false ───┼──────────────▶│   IExcelGateway,        │    │
│  │  ZeroDash   = true     │               │   ILog, formatKey,      │    │
│  │ GetCustomUiXml()       │               │   forceAlign, zeroDash) │    │
│  └───────────────────────┘               └───────────┬─────────────┘    │
│                                                        │                  │
│                                    1. gateway.TryGetSelectedRange(out r)  │
│                                       false → log.Warn(...); return       │
│                                    2. FormatRegistry.TryGetFormatDef(...) │
│                                       false → log.Warn(...); return       │
│                                    3. r.NumberFormat = def.NumberFormat   │
│                                       (r.HorizontalAlignment if needed)   │
│                                                        │                  │
│                                                        ▼                  │
│                              ┌─────────────────────────────────────┐     │
│                              │ IExcelGateway / IRangeHandle (seam) │     │
│                              │  Phase 2: FakeExcelGateway/          │     │
│                              │           FakeRangeHandle (tests)   │     │
│                              │  Phase 3: real Range/Application     │     │
│                              │           wrapper (not built yet)   │     │
│                              └─────────────────────────────────────┘     │
└──────────────────────────────────────────────────────────────────────────┘
                                                        │
                                                        ▼ (calls, unchanged from Phase 1)
┌──────────────────────────────────────────────────────────────────────────┐
│ PHASE 1 — Format Engine Core (done): FormatRegistry.TryGetFormatDef      │
│  → AccountingFormatBuilder.Build → FormatDef (Key/DisplayName/            │
│    NumberFormat/Category/Alignment)                                       │
└──────────────────────────────────────────────────────────────────────────┘
```

A reader can trace the primary use case end-to-end: a (future, Phase 3) Ribbon click → `FormatEngine.ApplyToSelection` → gateway selection guard → `FormatRegistry` lookup → `IRangeHandle` mutation. In Phase 2's tests, the top box doesn't exist; tests call `FormatEngine.ApplyToSelection` directly with a `FakeExcelGateway`.

### Recommended Project Structure

Extend the existing `FinanceFmtTools.Engine` project — do **not** create a new project for this phase (see rationale below and Open Questions for the one genuinely ambiguous alternative).

```
src/FinanceFmtTools.Engine/
├── FormatKeys.cs                  # Phase 1 (unchanged)
├── FormatCategory.cs               # Phase 1 (unchanged)
├── CellAlignment.cs                # Phase 1 (unchanged) — already documented as the
│                                   #   COM-free stand-in Phase 2/3 will map to XlHAlign
├── FormatDef.cs                    # Phase 1 (unchanged)
├── AccountingFormatBuilder.cs      # Phase 1 (unchanged)
├── FormatRegistry.cs               # Phase 1 (unchanged)
├── FormatEngine.cs                 # NEW — static orchestrator (Apply / ApplyToSelection)
├── RibbonController.cs             # NEW — instance class, owns checkbox session state
├── RibbonSessionConfig.cs          # NEW — plain mutable class: ForceAlign, ZeroDash
├── Abstractions/
│   ├── IExcelGateway.cs            # NEW — bool TryGetSelectedRange(out IRangeHandle)
│   ├── IRangeHandle.cs             # NEW — NumberFormat, HorizontalAlignment, Address
│   └── ILog.cs                     # NEW — Warn/Info/Error, never throws
└── Resources/
    └── (Link only — see below; no physical copy needed)

src/FinanceFmtTools.Engine.Tests/
├── ... (Phase 1 files, unchanged)
├── FormatEngineTests.cs            # NEW — success criterion 1
├── FormatEngineSelectionGuardTests.cs  # NEW — success criterion 2 / FMT-06
├── RibbonControllerTests.cs        # NEW — success criterion 3
├── FakeExcelGateway.cs             # NEW — test double
├── FakeRangeHandle.cs              # NEW — test double
└── SpyLog.cs                       # NEW — test double, records Warn/Info calls for assertion
```

**Why extend `FinanceFmtTools.Engine` instead of a new project:** Phase 1's own `FormatRegistry.cs` doc comment already frames itself as "the key -> FormatDef lookup that every Ribbon button click resolves through," directly continuing VBA's `modFormatEngine.bas` (which itself contains both the registry *and* `ApplyFormat`/`ApplyFormatToSelection` — the orchestration Phase 2 is porting). Phase 1's own SUMMARY explicitly anticipates this: *"FormatRegistry.TryGetFormatDef's no-throw, bool-returning contract for unrecognized keys is exactly the shape Phase 2's FormatEngine/RibbonController orchestration layer needs to build its FMT-06... guard clause on top of"* [VERIFIED: `.planning/phases/01-format-engine-core/01-03-SUMMARY.md`]. Both Phase 1 and Phase 2 share the identical constraint (zero COM, `net48;net8.0`, `dotnet test`-provable) — there is no technical reason to split them, and doing so only pays off once Phase 3 needs a COM-referencing project, which is a different constraint entirely (see below).

### Pattern 1: Try-pattern gateway seam (no exceptions, no nulls-as-signal)

**What:** `IExcelGateway.TryGetSelectedRange(out IRangeHandle range)` returns `bool`, mirroring `FormatRegistry.TryGetFormatDef`'s existing convention in this codebase.
**When to use:** Any COM-adjacent "might not exist / might be the wrong type" query — this collapses VBA's two separate guards (`SafeSelection()` returning `Nothing`, and `ApplyFormat`'s own `rng Is Nothing` check) into a single guard in C#, since both VBA cases map to the same "cannot proceed" outcome.
**Example:**
```csharp
// src/FinanceFmtTools.Engine/Abstractions/IExcelGateway.cs
// Pattern source: this repo's own FormatRegistry.TryGetFormatDef (src/FinanceFmtTools.Engine/FormatRegistry.cs)
namespace FinanceFmtTools.Engine.Abstractions
{
    public interface IExcelGateway
    {
        // false => current selection is not a Range (e.g. Chart/Shape selected, or nothing selected).
        // Mirrors VBA's SafeSelection()'s TypeName(Selection) <> "Range" check (src/modUtils.bas:74-89),
        // but collapsed into one no-throw boolean query instead of a Nothing-returning function.
        bool TryGetSelectedRange(out IRangeHandle range);
    }
}
```
```csharp
// src/FinanceFmtTools.Engine/Abstractions/IRangeHandle.cs
namespace FinanceFmtTools.Engine.Abstractions
{
    public interface IRangeHandle
    {
        string NumberFormat { get; set; }
        CellAlignment HorizontalAlignment { get; set; }   // maps to XlHAlign in the Phase 3 real impl
        string Address { get; }                            // for parity with VBA's rng.Address(External:=True) logging
    }
}
```

### Pattern 2: Static orchestrator, instance-based stateful service

**What:** Keep `FormatEngine` a static class with all dependencies passed as parameters (matching Phase 1's `FormatRegistry`/`AccountingFormatBuilder` convention exactly); make `RibbonController` an instance class (it owns mutable per-session state).
**When to use:** Static + parameterized for pure/stateless orchestration (nothing to hold between calls, trivially safe under xUnit's default cross-class parallel execution [CITED: xunit.net "Running Tests in Parallel" docs — different test classes run in parallel by default; state on a static type would leak across them]). Instance-based for anything that must hold state across calls (checkbox pressed/unpressed) — VBA's own `Public CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` global mutable booleans are exactly the kind of shared static state that would make parallel `dotnet test` runs flaky if replicated literally in C#.
**Example:**
```csharp
// src/FinanceFmtTools.Engine/FormatEngine.cs
// Ports VBA's ApplyFormat/ApplyFormatToSelection (src/modFormatEngine.bas:24-77).
// Static + parameterized, consistent with FormatRegistry.cs/AccountingFormatBuilder.cs.
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine
{
    public static class FormatEngine
    {
        public static void ApplyToSelection(
            IExcelGateway gateway, ILog log, string formatKey, bool forceAlign, bool zeroDash)
        {
            if (!gateway.TryGetSelectedRange(out IRangeHandle range))
            {
                // FMT-06: friendly-message behavior at the orchestration level.
                // The actual user-facing MsgBox is Phase 3's job (needs live Excel/WinForms UI);
                // this phase only proves the no-throw, logged-warning contract.
                log.Warn($"FormatEngine.ApplyToSelection: seleção atual não é um intervalo válido — abortando '{formatKey}'.");
                return;
            }

            Apply(range, log, formatKey, forceAlign, zeroDash);
        }

        public static void Apply(
            IRangeHandle range, ILog log, string formatKey, bool forceAlign, bool zeroDash)
        {
            if (!FormatRegistry.TryGetFormatDef(formatKey, forceAlign, zeroDash, out FormatDef def))
            {
                log.Warn($"FormatEngine.Apply: chave de formato desconhecida '{formatKey}'.");
                return;
            }

            range.NumberFormat = def.NumberFormat;
            if (def.Alignment != CellAlignment.General)
            {
                range.HorizontalAlignment = def.Alignment;
            }

            log.Info($"FormatEngine.Apply: aplicado '{def.DisplayName}' em {range.Address}.");
        }
    }
}
```
```csharp
// src/FinanceFmtTools.Engine/RibbonSessionConfig.cs
// Authoritative defaults per REQUIREMENTS.md RIB-02/RIB-03 and 02-CONTEXT.md — see Pitfall 1
// for why this does NOT match either VBA source file's literal default.
namespace FinanceFmtTools.Engine
{
    public sealed class RibbonSessionConfig
    {
        public bool ForceAlign { get; set; } = false;  // "Alinhar à direita" starts OFF
        public bool ZeroDash { get; set; } = true;      // "Zero contábil" starts ON
    }
}
```

### Pattern 3: Embedded-resource loading via suffix match, not exact name

**What:** Resolve the embedded Ribbon XML resource name by suffix (`EndsWith(...)`) rather than a hardcoded expected full name.
**When to use:** Any time an embedded resource's logical name depends on `RootNamespace`/folder layout that could drift.
**Why:** The sibling project hit exactly this bug in production — its `RibbonControllerTests.cs` carries a regression-test comment documenting it: *"O bug original era o nome do recurso... divergir do RootNamespace real... fazendo GetCustomUI devolver vazio e o botão NÃO aparecer na Ribbon. A resolução agora é por sufixo."* [VERIFIED: `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend.Tests/RibbonControllerTests.cs:8-11` and `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend/Services/RibbonController.cs:179-226`]
**Example:**
```csharp
// src/FinanceFmtTools.Engine/RibbonController.cs (excerpt)
private string LoadEmbeddedXml()
{
    var asm = typeof(RibbonController).Assembly;
    string resourceName = null;
    foreach (var name in asm.GetManifestResourceNames())
    {
        if (name.EndsWith("customUI14.xml", System.StringComparison.OrdinalIgnoreCase))
        {
            resourceName = name;
            break;
        }
    }
    if (resourceName == null) { _log.Error("RibbonController: embedded customUI14.xml not found."); return string.Empty; }

    using var stream = asm.GetManifestResourceStream(resourceName);
    using var reader = new System.IO.StreamReader(stream);
    return reader.ReadToEnd();
}
```

**Embedding the existing `src/customUI14.xml` without duplicating it — verified working:**
```xml
<!-- src/FinanceFmtTools.Engine/FinanceFmtTools.Engine.csproj -->
<ItemGroup>
  <EmbeddedResource Include="../customUI14.xml" Link="Resources/customUI14.xml" />
</ItemGroup>
```
This was empirically verified in this research session (a scratch `net8.0` console project embedding a file one directory above the project via `Link`, then enumerating `Assembly.GetManifestResourceNames()`) — it produces a predictable logical resource name (`{RootNamespace}.Resources.customUI14.xml`) and requires no physical file duplication [VERIFIED: empirical test run in this session, `dotnet 8.0.422`, output: `test.Resources.shared.xml`]. Reusing the existing VBA-era `src/customUI14.xml` verbatim (rather than copying it into the C# project tree) keeps a single source of truth for the Ribbon XML during the transition and costs nothing extra, since the button/checkbox `onAction`/`getPressed` callback *names* in that file (`RibbonFin8D`, `RibbonGetForceAlign`, etc.) are exactly what Phase 3's `Connect.cs` will need to implement as matching C# method names anyway.

### Anti-Patterns to Avoid
- **Static mutable session state (VBA-literal port):** VBA's `Public CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` globals are exactly the pattern to *not* copy verbatim into C# as `static` fields — it would make `RibbonController`'s state shared across every test class, unsafe under xUnit's default parallel-across-classes execution, and unnecessarily hard to reset between tests. Use an instance (`RibbonSessionConfig` owned by a `RibbonController` instance).
- **Showing a real message box from `FormatEngine`:** VBA's `SafeSelection()`/`ApplyFormat` call `MsgBox` directly. Do not attempt to reproduce this in Phase 2 — there is no live Excel/WinForms host to show a dialog against during `dotnet test`, and a blocking `MessageBox.Show` call in an automated test run would hang. Phase 2 proves only the logged-warning, no-throw contract; the actual user-facing dialog is Phase 3's job once real UI is available.
- **Hardcoding the embedded resource's exact expected name:** see Pattern 3 — resolve by suffix.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Structured/leveled logging across the whole add-in | A full logging framework (`Microsoft.Extensions.Logging`, `Serilog`, etc.) | A minimal hand-rolled `ILog` interface (`Warn`/`Info`/`Error`, never throws) | This phase needs exactly one warning-log call site proven by a test spy. Pulling in a logging framework now is the *inverted* mistake — it would be over-engineering for the actual need, and would be this project's first runtime NuGet dependency in `FinanceFmtTools.Engine`. The sibling project reached the same conclusion independently [VERIFIED: `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend/Abstractions/ILog.cs`] — revisit only if a later phase needs file rotation, structured fields, or multiple sinks. |
| Excel-selection type checking | Reflection-based `TypeName(Selection)`-style duck typing in C# | `IExcelGateway.TryGetSelectedRange(out IRangeHandle)` — push the "is this actually a Range" decision entirely into the (fake, then real) gateway implementation | Keeps `FormatEngine` COM-free and testable; the real Phase 3 implementation is the only place that ever needs to ask Excel "what type is `Application.Selection`" |

**Key insight:** Every "Don't Hand-Roll" temptation in this phase points the same direction — resist adding infrastructure (frameworks, reflection, custom COM-detection code) that the phase's actual, narrow success criteria don't require. The three success criteria are fully satisfiable with plain interfaces, one static orchestrator class, and one small stateful controller class.

## Common Pitfalls

### Pitfall 1: The VBA source has two contradictory answers for the checkbox defaults — neither is what this phase must implement

**What goes wrong:** A planner/implementer greps `modConfig.bas` or `modUtils.bas` for "the VBA default" and copies whichever value they find first.
**Why it happens:** This repo's own VBA source is internally inconsistent:
- `src/modConfig.bas:15-16` declares `Public CFG_FORCE_ALIGN As Boolean` / `Public CFG_ZERO_DASH As Boolean` with no initializer — VBA booleans default to `False`/`False` [VERIFIED: `src/modConfig.bas`].
- `src/modUtils.bas:4-7` (`CFG_XML_ROOT`) hardcodes the *very first* persisted XML as `<ForceAlign>true</ForceAlign><ZeroDash>false</ZeroDash>` — i.e. `True`/`False` — and `LoadConfig` (`src/modUtils.bas:116-136`) uses matching fallback defaults (`ReadBoolNode(..., "ForceAlign", defaultVal:=True)`, `ReadBoolNode(..., "ZeroDash", defaultVal:=False)`) [VERIFIED: `src/modUtils.bas`].
- Neither of those two VBA readings (`False/False` raw-global, or `True/False` persisted-XML) matches what the **new C# add-in** must actually do.
**How to avoid:** The authoritative source for Phase 2/3's behavior is `REQUIREMENTS.md` RIB-02/RIB-03 and `02-CONTEXT.md`'s locked decision, which explicitly states: **ForceAlign off by default, ZeroDash on by default**, with **no persistence at all** (persistence is explicitly listed as Out of Scope in `REQUIREMENTS.md`'s "Out of Scope" table: *"Persistência das preferências... removida deliberadamente nesta migração"*). This is a deliberate, considered behavior change for the migration — not a VBA-parity bug. `RibbonSessionConfig` must be hardcoded to `ForceAlign = false`, `ZeroDash = true`, full stop, and must never read/write any file or `CustomXMLPart`-equivalent.
**Warning signs:** Any test asserting `ForceAlign` defaults to `true`, or any code path attempting to load/save these values from disk/resource/registry, is wrong for this migration.

### Pitfall 2: Conflating "log a warning" (Phase 2) with "show the user a message" (Phase 3)

**What goes wrong:** Trying to make `FormatEngine`'s guard clause show a real dialog (`MessageBox.Show`/`Application.Excel.MsgBox`) so it "looks like" VBA's `SafeSelection()`.
**Why it happens:** VBA's `SafeSelection()` (`src/modUtils.bas:74-89`) does show a `MsgBox` directly, and it's tempting to port that 1:1.
**How to avoid:** Roadmap Phase 2 success criterion 2 explicitly says *"logs a warning and returns without throwing"* — it does not require a message box. Roadmap Phase 3 success criterion 2 explicitly assigns *"shows a friendly message"* to Phase 3 (live Excel session, RIB-01..04 scope). Keep `FormatEngine`/`RibbonController` free of any UI-dialog dependency in this phase; that avoids a `dotnet test` run trying to pop a blocking dialog.
**Warning signs:** Any `System.Windows.Forms` or `MessageBox`/`MsgBox` reference appearing in `FinanceFmtTools.Engine` in this phase is a scope violation (and would also require `<UseWindowsForms>true</UseWindowsForms>`, which Phase 1's csproj deliberately does not set).

### Pitfall 3: Embedded resource logical name drift

**What goes wrong:** Hardcoding an expected resource name (`"FinanceFmtTools.Engine.customUI14.xml"`) that silently stops matching if the project's `RootNamespace`, folder, or `Link` path changes later, causing `GetCustomUiXml()` to return an empty string with no build error.
**Why it happens:** SDK-style `EmbeddedResource` logical names are computed from `RootNamespace` + relative path, which is easy to get subtly wrong, especially when linking a file from outside the project directory (as recommended here for `src/customUI14.xml`).
**How to avoid:** Resolve by suffix match, not exact name (see Pattern 3). This is not a hypothetical risk — the sibling project shipped this exact bug and had to fix it, with a regression test now guarding against recurrence [VERIFIED: `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend.Tests/RibbonControllerTests.cs:8-11`].
**Warning signs:** `GetCustomUiXml()` (or equivalent) returning `string.Empty` in a test that expects real XML content.

### Pitfall 4: `RIB-*` requirements are not this phase's job — don't over-build `RibbonController`

**What goes wrong:** Implementing live `IRibbonUI` caching, `InvalidateControl`, image-loading (`stdole.IPictureDisp`), or About/docs-link wiring inside Phase 2's `RibbonController`, because the sibling project's `IRibbonController` has all of these.
**Why it happens:** The sibling project's `RibbonController` (read in full during this research) is a fuller, COM-referencing implementation because that project didn't stage its COM boundary across two phases — it built the whole thing in one project from the start.
**How to avoid:** `REQUIREMENTS.md`'s own traceability table maps **RIB-01, RIB-02, RIB-03, RIB-04 all to Phase 3**, not Phase 2 — Phase 2's only requirement is **FMT-06**. Phase 2's `RibbonController` success criterion (roadmap criterion 3) only asks for: (a) loading the embedded XML resource, (b) answering checkbox pressed/unpressed queries against in-memory state with the correct defaults. No `IRibbonUI`/image/invalidate logic is needed until Phase 3 wires the real COM callbacks. This project also has **no custom icons at all** — `src/customUI14.xml` uses only built-in `imageMso="..."` references, so unlike the sibling project, there is no `GetImage`/`IPictureDisp` concern to port at all, ever [VERIFIED: `src/customUI14.xml` — every `<button>` uses `imageMso`, none reference a custom image resource].
**Warning signs:** Any `stdole`, `Microsoft.Office.Core`, or `IRibbonUI` reference appearing anywhere in this phase's diff is a scope violation — it would also break the "zero real COM types" constraint from `02-CONTEXT.md`.

## Code Examples

### Test doubles (fakes), matching Phase 1's per-file convention

```csharp
// src/FinanceFmtTools.Engine.Tests/FakeRangeHandle.cs
using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FakeRangeHandle : IRangeHandle
    {
        public string NumberFormat { get; set; } = "General";
        public CellAlignment HorizontalAlignment { get; set; } = CellAlignment.General;
        public string Address { get; set; } = "$A$1";
    }
}
```
```csharp
// src/FinanceFmtTools.Engine.Tests/FakeExcelGateway.cs
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FakeExcelGateway : IExcelGateway
    {
        // Test control switch: simulates a Chart/Shape being selected instead of a Range.
        public bool SelectionIsRange { get; set; } = true;
        public FakeRangeHandle SelectedRange { get; set; } = new FakeRangeHandle();

        public bool TryGetSelectedRange(out IRangeHandle range)
        {
            if (!SelectionIsRange)
            {
                range = null;
                return false;
            }
            range = SelectedRange;
            return true;
        }
    }
}
```
```csharp
// src/FinanceFmtTools.Engine.Tests/SpyLog.cs
using System.Collections.Generic;
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class SpyLog : ILog
    {
        public List<string> Warnings { get; } = new List<string>();
        public List<string> Infos { get; } = new List<string>();

        public void Warn(string message) => Warnings.Add(message);
        public void Info(string message) => Infos.Add(message);
        public void Error(string message) { }
    }
}
```

### FMT-06 guard-clause test (success criterion 2)

```csharp
// src/FinanceFmtTools.Engine.Tests/FormatEngineSelectionGuardTests.cs
using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FormatEngineSelectionGuardTests
    {
        [Fact]
        public void ApplyToSelection_seleção_não_é_range_loga_aviso_e_não_lança()
        {
            var gateway = new FakeExcelGateway { SelectionIsRange = false };
            var log = new SpyLog();

            var ex = Record.Exception(() =>
                FormatEngine.ApplyToSelection(gateway, log, FormatKeys.Fin2D, forceAlign: false, zeroDash: true));

            Assert.Null(ex);
            Assert.Single(log.Warnings);
        }
    }
}
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| VBA global mutable state (`Public CFG_FORCE_ALIGN`) read/written from anywhere | Instance-held state (`RibbonSessionConfig` owned by a `RibbonController` instance) | This phase (Phase 2) | Removes cross-test/cross-session state leakage risk; matches xUnit's default parallel-across-classes execution model |
| Static-only functions everywhere (Phase 1) | Static for stateless orchestration (`FormatEngine`), instance for stateful services (`RibbonController`) | This phase (Phase 2) | Matches the sibling project's already-proven split between `Domain/` (static, pure) and `Services/` (instance, DI-friendly) |

**Deprecated/outdated:** VBA's `CustomXMLPart`-based persistence (`src/modUtils.bas:102-259`, `LoadConfig`/`SaveConfig`) — explicitly out of scope for the C# migration per `REQUIREMENTS.md`'s "Out of Scope" table. Do not port any part of this persistence mechanism.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `FormatEngine`/`RibbonController`/`Abstractions/*` should live in the existing `FinanceFmtTools.Engine` project rather than a new project | Recommended Project Structure | Low — both layouts are technically valid; if the planner picks a new project instead, only the `.csproj`/`ProjectReference` wiring changes, not the class designs themselves. Flagged as Claude's discretion, not a locked decision (CONTEXT.md grants full discretion here). |
| A2 | `RibbonController` should expose a `RibbonSessionConfig` object (rather than two flat bool properties) so Phase 3's Ribbon callbacks can read `ForceAlign`/`ZeroDash` from one place before calling `FormatEngine` | Pattern 2 | Low — purely an internal design convenience; either shape satisfies success criterion 3 as literally worded |
| A3 | Phase 2's `RibbonController` needs no `IRibbonUI`-caching/`InvalidateControl` abstraction at all (deferred entirely to Phase 3) | Pitfall 4, Open Questions | Medium — if the planner disagrees and wants a `IRibbonUiHandle`-style seam built now (mirroring `IExcelGateway`), that's a reasonable alternative reading of criterion 3's "using a fake rather than a live IRibbonUI" phrase; see Open Questions for both readings |
| A4 | The embedded Ribbon XML resource should link to the *existing* `src/customUI14.xml` (via MSBuild `Link`) rather than duplicating its content into the C# project tree | Pattern 3 | Low — empirically verified to work mechanically; the only risk is if the planner prefers a physical copy for stronger project-directory isolation, which is also valid but risks XML drift during the VBA→C# transition period |

**None of the above are training-data guesses** — all are grounded in direct reads of this repo's own VBA source, this repo's own Phase 1 code/summaries, `REQUIREMENTS.md`/`CONTEXT.md`, an empirical test run in this session, or a concrete local sibling codebase. They are logged here because they represent **design choices left to Claude's discretion** by `02-CONTEXT.md`, not because the underlying facts are unverified.

## Open Questions

1. **Does `RibbonController` need an `IRibbonUiHandle`-style seam interface in Phase 2, or is checkbox state + XML loading sufficient?**
   - What we know: `REQUIREMENTS.md` traces RIB-01..04 entirely to Phase 3; Phase 2's only requirement is FMT-06. Roadmap success criterion 3 for Phase 2 only asks for XML-resource loading + checkbox-state queries "using a fake rather than a live IRibbonUI."
   - What's unclear: whether "using a fake" is (a) simply describing the overall no-live-COM testing philosophy of this phase (my recommendation — no new interface needed), or (b) an implicit instruction to build a small `IRibbonUiHandle`/`CacheRibbon(object ribbonUi)`-style seam now (mirroring the sibling project's `IRibbonController.CacheRibbon(object ribbonUi)`, which notably also uses `object` rather than a COM type in its signature specifically so it stays fakeable) so that Phase 3 only has to *implement* it rather than *design* it.
   - Recommendation: Start with the narrower reading (no `IRibbonUiHandle` in Phase 2) since nothing in this add-in currently needs `InvalidateControl` (VBA's own `mRibbon` is captured in `OnRibbonLoad` but never actually invoked anywhere in `modRibbon.bas` — it appears to be unused/defensive-only in the original VBA too [VERIFIED: `src/modRibbon.bas`, full-file read, no `mRibbon.` call site found]). If the planner wants forward-compatibility insurance, a one-method `object CacheRibbonHandle(object ribbonUi)` stub (matching the sibling's `object`-typed signature trick) is a low-cost addition that doesn't violate the zero-COM-type constraint.

2. **Should `RibbonController`'s embedded XML resource be linked from `src/customUI14.xml` (shared with the still-active VBA `.xlam`) or copied into the new project?**
   - What we know: Linking works mechanically (verified empirically). The VBA source and its `.xlam` build are still the active, released artifact until Phase 5's `LEGACY-01`/`LEGACY-02` archive it.
   - What's unclear: whether the planner wants to guard against someone editing `src/customUI14.xml` for a VBA-side reason mid-migration and unknowingly affecting the in-progress C# build's embedded resource too.
   - Recommendation: Link (Pattern 3) — the callback names must stay identical between the two versions anyway during the transition, so shared-source is actually the safer default, not a risk.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| .NET 8 SDK | `dotnet build`/`dotnet test` (DEV-01) | ✓ | 8.0.422 | — |
| `FinanceFmtTools.sln` build (net48 + net8.0 legs) | All of Phase 2's new code | ✓ | Confirmed via `dotnet build src/FinanceFmtTools.sln -c Release` in this session — 0 Warning(s)/0 Error(s) | — |

No new external dependencies are introduced by this phase (no new tools, services, or packages). Phase 2 is purely additive C# source within the already-working Phase 1 toolchain.

## Security Domain

This is a local, single-user, single-process Excel COM add-in with no network calls, no authentication surface, and no persisted secrets in this phase (persistence itself is out of scope — see Pitfall 1). Most ASVS web/API categories do not meaningfully apply; the table below reflects that honestly rather than force-fitting web-app controls onto a desktop add-in.

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | No | No auth surface — single local user, no login |
| V3 Session Management | No | No session concept beyond in-process Excel state |
| V4 Access Control | No | No multi-user/permission model |
| V5 Input Validation | Yes (narrow) | The FMT-06 guard clause itself *is* this phase's input-validation control: never trust that `Selection` is a `Range` before mutating it; validate via `IExcelGateway.TryGetSelectedRange` before use, never via `try`/`catch` around a COM cast |
| V6 Cryptography | No | No secrets, no crypto in this phase |

### Known Threat Patterns for this stack

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| Unvalidated selection type causing an unhandled COM exception (add-in crash, or worse, Excel silently disabling the add-in per Resiliency — a Phase 4 concern) | Denial of Service (of the add-in itself) | The Try-pattern gateway guard (this phase) plus never allowing `FormatEngine`/`RibbonController` to let an exception escape to the (future) Ribbon callback boundary |
| Format-key string used as a `switch` key without a default no-throw case | Denial of Service | Already solved in Phase 1 — `FormatRegistry.TryGetFormatDef`'s `default: def = null; return false;` case, reused unchanged by `FormatEngine.Apply`'s own guard |

## Sources

### Primary (HIGH confidence)
- `src/modFormatEngine.bas` (this repo) — `ApplyFormat`/`ApplyFormatToSelection`/`GetFormatDef`/`AccountingFmt`, the exact VBA logic being ported
- `src/modUtils.bas` (this repo) — `SafeSelection`, `LoadConfig`/`SaveConfig`, `CustomXMLPart` persistence (being deliberately dropped)
- `src/modRibbon.bas`, `src/modConfig.bas`, `src/customUI14.xml` (this repo) — Ribbon callback wiring, config defaults, actual Ribbon XML content
- `src/FinanceFmtTools.Engine/*.cs` (this repo, Phase 1) — `FormatRegistry`, `AccountingFormatBuilder`, `FormatDef`, `FormatKeys`, `FormatCategory`, `CellAlignment` — read in full
- `.planning/phases/01-format-engine-core/01-01-SUMMARY.md`, `01-02-SUMMARY.md`, `01-03-SUMMARY.md`, `.planning/STATE.md`, `.planning/ROADMAP.md`, `.planning/REQUIREMENTS.md`, `.planning/phases/02-abstractions-orchestration/02-CONTEXT.md` (this repo)
- `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend/Abstractions/*.cs`, `Services/OutlookGateway.cs`, `Services/RibbonController.cs`, `Services/SendInterceptor.cs`, `Domain/SendInterceptorLogic.cs`, `UndoSend.Tests/*.cs`, `.planning/ARCHITECTURE-CSHARP.md` (local sibling project explicitly cited in this repo's own `CLAUDE.md` as the dev/build/release workflow inspiration) — read in full for gateway/controller/logging conventions
- Empirical test run in this session: `dotnet 8.0.422`, `EmbeddedResource Include="../shared.xml" Link="Resources/shared.xml"` → confirmed logical resource name resolution works across a project-directory boundary

### Secondary (MEDIUM confidence)
- [Running Tests in Parallel — xUnit.net](https://xunit.net/docs/running-tests-in-parallel) — confirms different test classes run in parallel by default, same-class tests run serially; informs the static-vs-instance design recommendation

### Tertiary (LOW confidence)
None — every claim in this document traces to a file read in this session, an empirical test run in this session, or a WebSearch-confirmed official xUnit doc.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — no new packages, reusing Phase 1's already-verified toolchain
- Architecture: HIGH — derived directly from this repo's own Phase 1 code/summaries, the VBA source being ported, and a concrete working sibling implementation by the same author for the same class of problem
- Pitfalls: HIGH — the default-value contradiction (Pitfall 1) and the resource-naming pitfall (Pitfall 3) are both grounded in direct file reads, not speculation

**Research date:** 2026-07-11
**Valid until:** No external time pressure — this is internal architecture with no third-party version drift risk; valid until Phase 2 is planned/executed (no expiry needed for a same-session handoff)
