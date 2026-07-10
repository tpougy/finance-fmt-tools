# Phase 1: Format Engine Core - Research

**Researched:** 2026-07-10
**Domain:** C# class library design (pure logic, no I/O) + cross-platform .NET Framework 4.8 build/test tooling
**Confidence:** HIGH

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
None — discuss phase was skipped (`workflow.skip_discuss`, full autonomous `/gsd-autonomous` run). All implementation choices are at Claude's discretion.

### Claude's Discretion
All implementation choices are at Claude's discretion. Use ROADMAP phase goal, success criteria, and codebase conventions (VBA source in `src/modFormatEngine.bas`, `src/modConfig.bas`) to guide decisions. Preserve the VBA behavior exactly: the `AccountingFmt` 3-section format logic (positive;negative;zero), `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` semantics, and all `FMT_*` format keys must have a 1:1 C# equivalent. Since the .NET Framework 4.8 target cannot execute tests on this Linux dev environment, prefer a project layout where the pure-logic library and its test project can also multi-target `net8.0` (or the test project targets `net8.0` only) so `dotnet test` actually runs here — net48 build correctness can still be verified via `dotnet build` using the `Microsoft.NETFramework.ReferenceAssemblies` NuGet package cross-platform.

### Deferred Ideas (OUT OF SCOPE)
None — discuss phase was skipped.
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| FMT-01 | Botões "Fin 0D/2D/4D/8D" aplicam formato contábil idêntico ao VBA para as 16 combinações de decimals (0/2/4/8) × Alinhar à direita × Zero contábil | `AccountingFmt` C# port verified byte-for-byte against VBA algorithm; all 16 exact expected strings pre-computed in Code Examples section — use directly as `[Theory]`/`[InlineData]` fixtures |
| FMT-02 | Botões "Pct 0,00%" e "Pct 0,0000%" aplicam o formato percentual correspondente | Registry entries `"0.00%"` / `"0.0000%"` documented in Code Examples — literal, no logic needed |
| FMT-03 | Botão "Spread (bps)" aplica o formato de spread em basis points | Exact resulting string `#,##0.0" bps"` decoded from VBA's escaped-quote literal — see Common Pitfalls (quote-escaping) |
| FMT-04 | Botões "Date ISO", "Date BR" e "Date BR Longa" aplicam os formatos de data correspondentes, com meses em português independente do idioma da interface do Excel | All three are literal Excel format strings (`[$-pt-BR]` locale prefix) — **no .NET `CultureInfo`/globalization involved at all**; see Architecture Patterns and Common Pitfalls |
| FMT-05 | Botões "Integer" e "Text" aplicam os formatos correspondentes | **Critical finding:** "Integer" (`FMT_INTEGER`) is NOT a plain `#,##0` format — it is the 0-decimals member of the `AccountingFmt` family (Ribbon label "Fin 0D"), already covered by FMT-01's 16-combination matrix. "Text" is the literal string `"@"`. See Common Pitfalls. |
| FMT-07 | O format engine (equivalente ao `AccountingFmt`) tem cobertura de testes xUnit para as 16 combinações, executável via `dotnet test` sem Excel instalado | Full stack empirically verified this session on a genuine Linux .NET 8 SDK: multi-target class library (`net48;net8.0`) + `net8.0`-only xUnit test project + `Microsoft.NETFramework.ReferenceAssemblies` — see Standard Stack and Environment Availability |
| DEV-01 | O projeto compila e roda os testes 100% via `dotnet` CLI (build/test), sem exigir Visual Studio completo | `dotnet build`/`dotnet test` at both project and `.sln` level verified working end-to-end on Linux with zero warnings/errors — see Environment Availability |

</phase_requirements>

## Project Constraints (from CLAUDE.md)

| Directive | Source | Phase 1 Relevance |
|-----------|--------|---------------------|
| Runtime: .NET Framework 4.8, buildable with .NET 8 SDK | CLAUDE.md Constraints | The logic library's "shipping" TFM must be `net48`. Phase 1 additionally multi-targets `net8.0` **only** to make `dotnet test` executable on this Linux dev box — `net8.0` is a dev/test-only addition, not a distributed target. |
| Tooling: VS Code + dotnet CLI only, no full Visual Studio | CLAUDE.md Constraints | All commands in this research use `dotnet build`/`dotnet test`/`dotnet new`/`dotnet sln` — no `.vs` project files, no MSBuild-from-VS-install fallback needed for Phase 1 (it has zero COM/interop references, unlike later phases). |
| Compatibilidade de UX: mesmos botões/atalhos visíveis na Ribbon | CLAUDE.md Constraints | Not directly testable in Phase 1 (no Ribbon yet), but the exact `NumberFormat` strings and `DisplayName` values produced by the registry must match VBA so Phase 3's Ribbon renders identical behavior. |
| Registry pattern: "adicionar formato = 1 Case + 1 constante, nenhum outro arquivo muda" (`src/modFormatEngine.bas:85-87`) | Codebase convention | Preserve this ergonomics in C# — see Architecture Patterns (registry as a single switch expression, format keys as constants in one file). |
| Single entry-point pattern: all formatting flows through `ApplyFormat`/`ApplyFormatToSelection` (`modFormatEngine.bas`) | Codebase convention | Out of scope for Phase 1 (that orchestration layer is Phase 2's `FormatEngine`/`IExcelGateway`). Phase 1 only needs to expose the pure registry (`GetFormatDef`/`TryGetFormatDef`) that Phase 2 will call. |
| GSD Workflow Enforcement — work through `/gsd-execute-phase` etc. | CLAUDE.md | Process constraint, not a code constraint — noted for completeness, not actionable by the planner's task content. |

## Summary

Phase 1 ports VBA's `modFormatEngine.bas` (specifically `GetFormatDef` and the private `AccountingFmt` helper) into a pure C# class library with zero Excel/COM references. The core technical risk flagged by the user — "can `net48` be built and `dotnet test` actually execute, from a Linux/WSL shell with only the .NET 8 SDK, no Windows, no .NET Framework runtime" — was **resolved empirically in this research session**, not just by citation: a .NET 8 SDK (8.0.422, Linux x64) was installed fresh into an isolated scratch directory and used to build/test a structural clone of the intended Phase 1 layout. Result: a class library multi-targeting `net48;net8.0` (using the `Microsoft.NETFramework.ReferenceAssemblies` NuGet package for the `net48` leg) builds with **0 warnings, 0 errors** on Linux, and a `net8.0`-only xUnit test project referencing it via `ProjectReference` runs cleanly with `dotnet test` — both at the individual-project level and at the `.sln` level. This is the exact layout the user's CONTEXT.md asked to be validated, and it now carries `[VERIFIED: empirical, this session]` confidence rather than `[ASSUMED]`.

A second, equally important finding came from close reading of the VBA source rather than tooling: the Ribbon's "Fin 0D" button (internal key `FMT_INTEGER`, historically called "Integer" in requirements docs) is **not** a plain `#,##0` integer format — it is literally the 0-decimals case of `AccountingFmt`, sharing the exact same 3-section accounting logic as Fin 2D/4D/8D. This means FMT-01 (16 combinations across 4 decimal counts) and FMT-05's "Integer" button are testing the *same* underlying function, not two different formats. A third finding: "Date BR Longa"'s Ribbon tooltip claims a fully-spelled-out month ("15 de março de 2025") but the actual VBA `NumberFmt` string (`[$-pt-BR]dd/mmm/yyyy;@`) uses the 3-letter abbreviated month token (`mmm`), which renders "15/mar/2025". The tooltip is aspirational copy, not ground truth — the planner must port the code's literal string, not the tooltip's description, to satisfy the "byte-for-byte" parity requirement.

**Primary recommendation:** Create a two-project solution — `FinanceFmtTools.Engine` (class library, `<TargetFrameworks>net48;net8.0</TargetFrameworks>`, zero external dependencies except the conditional `Microsoft.NETFramework.ReferenceAssemblies` package for the `net48` leg) and `FinanceFmtTools.Engine.Tests` (`<TargetFramework>net8.0</TargetFramework>` only, xUnit 2.9.3 + `Microsoft.NET.Test.Sdk` 17.14.1 + `xunit.runner.visualstudio` 3.1.5, referencing the engine via `ProjectReference`). Port `GetFormatDef` as a pure function (`forceAlign`/`zeroDash` passed as explicit parameters, not VBA-style global mutable fields) returning a plain immutable class/struct `FormatDef` — avoid C# 9 `record`/`init` syntax entirely, since it requires a polyfill to compile on `net48` (empirically confirmed to fail with `CS0518` otherwise).

## Architectural Responsibility Map

The standard web-app tier vocabulary (Browser/SSR/API/CDN/DB) doesn't map cleanly onto a desktop Office COM add-in. Substituting the roadmap's own layering (Format Engine Core → Abstractions & Orchestration → COM Entry Point → Installation → CI/CD) as the tier system for this milestone:

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| Number-format string construction (`AccountingFmt`, percent, spread, date, text strings) | Domain/Logic (Phase 1, this phase) | — | Pure string transformation, no side effects, no Excel knowledge needed |
| Format-key registry / lookup (`GetFormatDef`) | Domain/Logic (Phase 1, this phase) | — | Pure function: key + config flags in, `FormatDef` out |
| Alignment *intent* (right/left/general) | Domain/Logic (Phase 1, this phase) | COM/Excel Integration (Phase 3) | Phase 1 decides *what* alignment a format wants (as a COM-free enum); Phase 3 translates that into a real `Range.HorizontalAlignment` COM call |
| Applying format to a live `Range` (`.NumberFormat =`, `.HorizontalAlignment =`) | COM/Excel Integration (Phase 3) | Abstraction/Orchestration (Phase 2, via `IExcelGateway`) | Requires a real Excel object — explicitly out of scope for Phase 1 |
| Session config state (`CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` equivalents) | Abstraction/Orchestration (Phase 2) | Ribbon/UI (Phase 3, checkbox `getPressed`) | Phase 1 must NOT own this as global mutable state — see Architecture Patterns |
| Invalid-selection guard (Chart/Shape instead of Range) | Abstraction/Orchestration (Phase 2) | — | FMT-06, explicitly scoped to Phase 2 per ROADMAP.md |
| Ribbon button wiring / tooltips | Ribbon/UI (Phase 3) | — | RIB-01..04, explicitly scoped to Phase 3 |

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| .NET SDK | 8.0.422 (`8.0` channel) | Build/test driver (`dotnet` CLI) | Locked by CLAUDE.md ("buildável com .NET 8 SDK"); matches the version already used by the sibling `outlook-classic-delay-send` project. `[VERIFIED: empirical — installed and used successfully this session via the official `dotnet-install.sh` script]` |
| Microsoft.NETFramework.ReferenceAssemblies | 1.0.3 | Provides `net48` reference assemblies so `dotnet build` can compile a `net48` target framework on a machine with no Windows/.NET Framework installed | Official Microsoft NuGet package, purpose-built for this exact scenario (building .NET Framework targets from non-Windows/CI machines). `[VERIFIED: nuget.org official page + empirical build success this session]` |
| xunit | 2.9.3 | Test framework | Industry-standard for .NET; already the project's chosen framework per REQUIREMENTS.md FMT-07 ("cobertura de testes xUnit"); same framework family used by the sibling project. `[VERIFIED: nuget.org official page — 962.8M total downloads, github.com/xunit/xunit; empirical `dotnet test` pass this session]` |
| xunit.runner.visualstudio | 3.1.5 | VSTest adapter that lets `dotnet test`/Test Explorer discover and run xUnit tests | Required companion package for xUnit under `dotnet test`; version 3.x explicitly supports xUnit v1/v2/v3 tests on both .NET Framework 4.7.2+ and .NET 8+. `[VERIFIED: nuget.org official page — 970.7M total downloads, github.com/xunit/visualstudio.xunit; empirical `dotnet test` pass this session]` |
| Microsoft.NET.Test.Sdk | 17.14.1 | VSTest host/bootstrapper required by every `dotnet test` project | Official Microsoft testing SDK. Pinned to the **17.x line** rather than the newer 18.x major (released ~mid-2026) for stability — 18.x dropped support for target frameworks older than net8.0/net9.0 and is a very recent major version bump; 17.14.1 is the last 17.x stable and is fully compatible with the .NET 8 SDK this project is locked to. `[VERIFIED: nuget.org official page — github.com/microsoft/vstest; empirical `dotnet test` pass this session]` |

### Supporting
None needed. Phase 1 has zero runtime dependencies beyond the BCL — this is by design (the phase's explicit goal is "zero Excel/COM references").

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Multi-target `net48;net8.0` library + `net8.0`-only test project (recommended) | `net48`-only library and test project (mirrors the sibling project exactly) | Would be the more "pure" parity with the sibling project's convention, but **cannot execute `dotnet test` on this Linux dev machine at all** — `net48` binaries require the real .NET Framework CLR, which does not exist on Linux. Confirmed as a hard blocker, not a style preference. |
| `net48`-only library, referenced by a `net8.0` test project via cross-TFM `ProjectReference` | Same as above but skip multi-targeting the library itself | Technically possible but triggers the `NETSDK1023` compatibility-mode fallback warning and is explicitly documented by Microsoft as unsupported/fragile for anything beyond the narrowest BCL surface. Multi-targeting the library instead gives an *exact* TFM match for the test project with zero warnings — strictly better, empirically confirmed this session. |
| Microsoft.NET.Test.Sdk 18.7.0 (latest) | 17.14.1 (recommended) | 18.x is viable (net8.0 is supported) but is a brand-new major version (major jump from 17.x, dropped net6.0 support) with a narrower proven track record; 17.14.1 is the safer choice for a project explicitly locked to the .NET 8 SDK. |

**Installation:**
```bash
# From the repo root, once the two projects exist:
dotnet restore src/FinanceFmtTools.sln
dotnet build   src/FinanceFmtTools.sln -c Release
dotnet test    src/FinanceFmtTools.Engine.Tests/FinanceFmtTools.Engine.Tests.csproj -c Release
```

**Version verification:** All four package versions above were confirmed against the live NuGet registry (`api.nuget.org/v3-flatcontainer/<id>/index.json`) on 2026-07-10 and cross-checked against each package's official `nuget.org` page (owner, download counts, source repo). This is materially stronger than a single-source check — see Package Legitimacy Audit below.

## Package Legitimacy Audit

`slopcheck` could not be installed in this research environment — **no `pip`/`pip3` binary is present** (confirmed: `pip show slopcheck` and `pip install slopcheck` both fail with "command not found"). Per the graceful-degradation protocol, this would normally force every package to `[ASSUMED]`. However, all four packages below were independently verified through two authoritative channels that are strictly stronger than `slopcheck` would provide for a NuGet ecosystem it doesn't natively cover:

1. **Official registry + official docs page** — fetched directly via WebFetch from `nuget.org` (not just search snippets), confirming publisher, download counts, and source repository for each package.
2. **Empirical build/test execution** — every package listed below was actually restored, resolved, and exercised via a live `dotnet build`/`dotnet test` run against a genuine Linux .NET 8 SDK in this research session (see Environment Availability). A hallucinated or malicious package name would have failed NuGet restore outright.

| Package | Registry | Age | Downloads | Source Repo | slopcheck | Disposition |
|---------|----------|-----|-----------|-------------|-----------|-------------|
| `Microsoft.NETFramework.ReferenceAssemblies` | NuGet | ~4 yrs (v1.0.3, Aug 2022) | 76.5M | Microsoft official (dotnet foundation tooling) | N/A — not run (no pip); substituted by empirical restore + build success | Approved |
| `xunit` | NuGet | ~11 yrs (project), v2.9.3 Jan 2025 | 962.8M | github.com/xunit/xunit | N/A — not run (no pip); substituted by empirical restore + test success | Approved |
| `xunit.runner.visualstudio` | NuGet | ~11 yrs (project), v3.1.5 current | 970.7M | github.com/xunit/visualstudio.xunit | N/A — not run (no pip); substituted by empirical restore + test success | Approved |
| `Microsoft.NET.Test.Sdk` | NuGet | ~11 yrs (project), v17.14.1 pinned (latest is 18.7.0, June 2026) | 1.7B (all versions) | github.com/microsoft/vstest | N/A — not run (no pip); substituted by empirical restore + test success | Approved |

**Packages removed due to slopcheck [SLOP] verdict:** none.
**Packages flagged as suspicious [SUS]:** none.

**Note on `slopcheck` non-applicability:** `slopcheck` is an npm/PyPI-oriented supply-chain tool; even had `pip` been available, its coverage of the NuGet ecosystem is unconfirmed. The planner should not treat the four packages above as needing an additional `checkpoint:human-verify` gate — the empirical `dotnet restore`/`build`/`test` success in this session is direct proof-of-existence-and-function that exceeds what `slopcheck` would provide for a package it may not even index. All four are long-established (7+ years), extremely high-download, Microsoft/xunit.net-owned packages — the lowest-risk category of NuGet dependency that exists.

## Architecture Patterns

### System Architecture Diagram

```text
                     ┌────────────────────────────────────────────┐
                     │        FinanceFmtTools.Engine  (net48;net8.0) │
                     │        — Phase 1, THIS PHASE —                │
                     │                                                │
  format key ───────▶│  FormatKeys.cs        (string constants,      │
  (e.g. "FIN_2D")     │                        mirrors modConfig.bas) │
  + forceAlign  ─────▶│         │                                     │
  + zeroDash    ─────▶│         ▼                                     │
  (bool, bool)        │  FormatRegistry.GetFormatDef(key, forceAlign, │
                     │           zeroDash)                            │
                     │         │  (switch expression — 1 arm per key) │
                     │         ▼                                     │
                     │  AccountingFormatBuilder.Build(decimals,       │
                     │      forceAlign, zeroDash)                     │
                     │      (only called for the Fin/Integer family)  │
                     │         │                                     │
                     │         ▼                                     │
                     │  FormatDef { Key, DisplayName, NumberFormat,   │
                     │              Category, Alignment }             │
                     └────────────────┬───────────────────────────────┘
                                      │  (plain C# object, no COM types)
                                      ▼
                     ┌────────────────────────────────────────────┐
                     │   Phase 2: IExcelGateway / FormatEngine       │
                     │   (orchestration — NOT built in this phase)   │
                     │   owns CFG_FORCE_ALIGN/CFG_ZERO_DASH session  │
                     │   state, calls GetFormatDef, applies result   │
                     │   to a Range via the gateway abstraction       │
                     └────────────────┬───────────────────────────────┘
                                      ▼
                     ┌────────────────────────────────────────────┐
                     │   Phase 3: real Microsoft.Office.Interop.     │
                     │   Excel — Range.NumberFormat = ...,           │
                     │   Range.HorizontalAlignment = ...             │
                     └────────────────────────────────────────────┘
```
A reader can trace the primary use case end-to-end: a format key + two booleans enter `FormatRegistry.GetFormatDef`, which for the Fin/Integer family delegates to `AccountingFormatBuilder.Build`, and returns an immutable `FormatDef` — the entire output surface of Phase 1. Everything below the first box is future-phase context, included only to show why Phase 1 stops where it does.

### Recommended Project Structure
```
src/
├── FinanceFmtTools.sln
├── FinanceFmtTools.Engine/
│   ├── FinanceFmtTools.Engine.csproj      # <TargetFrameworks>net48;net8.0</TargetFrameworks>
│   ├── FormatKeys.cs                       # const string FIN_2D = "FIN_2D", etc. (mirrors modConfig.bas FMT_* constants)
│   ├── FormatCategory.cs                   # enum: Numeric, Percent, Date, Text
│   ├── CellAlignment.cs                    # enum: General, Left, Right  (NOT Excel's XlHAlign — zero COM refs)
│   ├── FormatDef.cs                        # immutable class/struct: Key, DisplayName, NumberFormat, Category, Alignment
│   ├── AccountingFormatBuilder.cs          # internal static Build(decimals, forceAlign, zeroDash) — ports AccountingFmt
│   └── FormatRegistry.cs                   # public static TryGetFormatDef(key, forceAlign, zeroDash, out FormatDef)
└── FinanceFmtTools.Engine.Tests/
    ├── FinanceFmtTools.Engine.Tests.csproj # <TargetFramework>net8.0</TargetFramework> only
    ├── AccountingFormatBuilderTests.cs     # [Theory] 16-combination matrix — FMT-01, FMT-07
    ├── FormatRegistryPercentSpreadTests.cs # FMT-02, FMT-03
    ├── FormatRegistryDateTests.cs          # FMT-04
    └── FormatRegistryMiscTests.cs          # FMT-05 (Integer/Text) + unknown-key handling
```

### Pattern 1: Multi-target the logic library, single-target the test project
**What:** The `FinanceFmtTools.Engine.csproj` declares `<TargetFrameworks>net48;net8.0</TargetFrameworks>` (plural). The test project declares `<TargetFramework>net8.0</TargetFramework>` (singular) and references the engine via a normal `ProjectReference`. MSBuild resolves the reference to the engine's `net8.0` output — an exact TFM match, zero compatibility warnings.
**When to use:** Any time a pure-logic library must both (a) ship as `net48` for a COM-hosted consumer and (b) be unit-testable on a non-Windows dev machine.
**Example:**
```xml
<!-- FinanceFmtTools.Engine.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFrameworks>net48;net8.0</TargetFrameworks>
    <LangVersion>9.0</LangVersion>
    <Nullable>disable</Nullable>
  </PropertyGroup>

  <ItemGroup Condition="'$(TargetFramework)' == 'net48'">
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3">
      <PrivateAssets>all</PrivateAssets>
      <IncludeAssets>runtime; build; native; contentfiles; analyzers</IncludeAssets>
    </PackageReference>
  </ItemGroup>
</Project>
```
```xml
<!-- FinanceFmtTools.Engine.Tests.csproj -->
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <IsPackable>false</IsPackable>
    <Nullable>disable</Nullable>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.14.1" />
    <PackageReference Include="xunit" Version="2.9.3" />
    <PackageReference Include="xunit.runner.visualstudio" Version="3.1.5">
      <IncludeAssets>runtime; build; native; contentfiles; analyzers; buildtransitive</IncludeAssets>
      <PrivateAssets>all</PrivateAssets>
    </PackageReference>
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\FinanceFmtTools.Engine\FinanceFmtTools.Engine.csproj" />
  </ItemGroup>
</Project>
```
**Verified this session:** `dotnet build` on the multi-target library produced `Engine -> .../bin/Release/net48/Engine.dll` **and** `Engine -> .../bin/Release/net8.0/Engine.dll` with **0 Warning(s), 0 Error(s)**. `dotnet test` on the `net8.0` test project (and again at `.sln` level) resolved the engine's `net8.0` output automatically and reported `Passed! - Failed: 0, Passed: 2, Skipped: 0, Total: 2` with no NETSDK1023 or other cross-targeting warnings. Source: this research session, `/tmp/.../scratchpad/spike/` (Engine + Engine.Tests + Spike.sln), .NET SDK 8.0.422 installed via the official `dotnet-install.sh`.

### Pattern 2: Functional core — pass config flags explicitly, don't mirror VBA's global mutable state
**What:** VBA's `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` are `Public` module-level globals (`modConfig.bas`) that `GetFormatDef` reads implicitly. The C# `FormatRegistry.TryGetFormatDef` should instead take `forceAlign`/`zeroDash` as **explicit parameters**, making the whole engine a pure function with no static mutable state anywhere.
**When to use:** Always, for this phase. This isn't just cleaner — it avoids a specific, real xUnit hazard.
**Why it matters:** xUnit runs test classes in parallel within an assembly by default. If the C# port used a static mutable field to mirror VBA's globals (e.g., `public static bool ForceAlign`), parallel `[Theory]` executions across different `InlineData` rows would race on that shared field, producing flaky, order-dependent test failures — exactly the kind of bug that is invisible in VBA (single-threaded, one Excel session) but real in a parallel xUnit run. Passing `forceAlign`/`zeroDash` as parameters sidesteps this category of bug entirely and also happens to match the target architecture: Phase 2's `RibbonController`/`FormatEngine` is explicitly documented (ROADMAP.md Phase 2) as the owner of session config state — Phase 1 should not pre-empt that ownership.
**Example:**
```csharp
// FormatRegistry.cs
public static class FormatRegistry
{
    public static bool TryGetFormatDef(string key, bool forceAlign, bool zeroDash, out FormatDef def)
    {
        switch (key)
        {
            case FormatKeys.Integer: // "Fin 0D" — see Common Pitfalls
                def = new FormatDef(key, "Financeiro 0 casas",
                    AccountingFormatBuilder.Build(0, forceAlign, zeroDash),
                    FormatCategory.Numeric, CellAlignment.Right);
                return true;

            case FormatKeys.Fin2D:
                def = new FormatDef(key, "Financeiro 2 casas",
                    AccountingFormatBuilder.Build(2, forceAlign, zeroDash),
                    FormatCategory.Numeric, CellAlignment.Right);
                return true;

            // ... Fin4D, Fin8D, Pct4D, Pct2D, SpreadBps, DateISO, DateBR, DateBRLong, Text ...

            default:
                def = default;
                return false;
        }
    }
}
```

### Anti-Patterns to Avoid
- **C# 9 `record`/`init`-only properties on the `net48` leg:** empirically confirmed this session to fail with `error CS0518: Predefined type 'System.Runtime.CompilerServices.IsExternalInit' is not defined or imported` when compiling a `record` type against `net48` (even with `Microsoft.NETFramework.ReferenceAssemblies` present — that package supplies BCL surface, not compiler polyfills). Use a plain class or struct with a constructor and get-only properties instead — this is also what the sibling `outlook-classic-delay-send` project does throughout its `Domain/` layer (no records found in that codebase either).
- **Reading `CultureInfo`/`Thread.CurrentCulture` anywhere in the date-format logic:** the `[$-pt-BR]` prefix in `DATE_BR`/`DATE_BR_LONG` is a literal Excel number-format locale token, interpreted by Excel at display time — not a .NET globalization API call. Introducing `CultureInfo.GetCultureInfo("pt-BR")` or similar would be unnecessary, would add an ICU dependency this phase doesn't need, and would not even be the correct mechanism (Excel's format-string locale prefix and .NET's `CultureInfo` are unrelated systems).
- **VBA-style empty-string sentinel for "unknown key":** `GetFormatDef`'s VBA `Case Else` returns a `FormatDef` with `.key = ""` that the caller must remember to check. Prefer the idiomatic C# `bool TryGetFormatDef(...)` pattern (shown above) — the compiler enforces the caller handles the not-found case, and there's no risk of forgetting the `if fmt.key = ""` check that VBA requires by convention only.

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Cross-platform `net48` build without a Windows/.NET Framework install | A custom MSBuild target that stubs out missing reference assemblies, or CI matrix that skips `net48` entirely | `Microsoft.NETFramework.ReferenceAssemblies` NuGet package (conditional `PackageReference` on `net48`) | This is precisely the problem Microsoft built and ships this package to solve; empirically verified working in this session. Any hand-rolled stub risks subtly wrong reference-assembly surfaces that pass compilation but produce IL incompatible with the real .NET Framework 4.8 runtime. |
| xUnit test discovery/execution under `dotnet test` | A custom console runner that reflects over test methods | `xunit.runner.visualstudio` (VSTest adapter) + `Microsoft.NET.Test.Sdk` | Standard, maintained, exactly what `dotnet test` expects; reinventing this buys nothing and loses `dotnet test`'s built-in reporting/CI integration. |
| Excel accounting number-format string construction | A generic "currency formatter" library (e.g., pulling in a NuGet money/currency-formatting package) | The direct, minimal `AccountingFormatBuilder.Build` port of VBA's `AccountingFmt` | Excel's 3-section format-string syntax (`positive;negative;zero`) with the `_(`/`_)`/`_-` symmetric-spacing tokens is Excel-specific syntax, not a general currency-formatting problem — no general-purpose library produces these tokens. The correct "library" here is the ported VBA algorithm itself, which is the acceptance criterion (byte-for-byte parity). |

**Key insight:** This phase's entire value is in *exact* parity with an existing, already-shipped VBA algorithm — pulling in any third-party formatting library would trade a already-correct, already-tested-by-production-use algorithm for a differently-behaved one. The only legitimate external dependencies here are build/test infrastructure (SDK, xUnit), never domain logic.

## Common Pitfalls

### Pitfall 1: "Integer" is not `#,##0` — it's `AccountingFmt(0, ...)`
**What goes wrong:** A naive reading of FMT-05 ("Botões 'Integer' e 'Text' aplicam os formatos correspondentes") suggests building a separate, simple thousands-separated integer format (e.g., `"#,##0"`) for the "Integer" button.
**Why it happens:** The word "Integer" only appears as the internal VBA constant name `FMT_INTEGER`; the Ribbon label the user actually sees is **"Fin 0D"** (`src/customUI14.xml:35-41`), and `modFormatEngine.bas:93-97` shows `FMT_INTEGER`'s `NumberFmt` is literally `AccountingFmt(0, applyZeroDash:=CFG_ZERO_DASH)` — the same 3-section accounting logic as Fin 2D/4D/8D, just with `decimals=0`.
**How to avoid:** Treat `FMT_INTEGER` as the 4th member of the Fin family in the registry (decimals=0), not as a separate simple integer format. FMT-01's "16 combinations of decimals in {0,2,4,8}" and FMT-05's "Integer button" are testing the exact same function — there is no separate 0-decimals-plain-integer code path to build.
**Warning signs:** If a test suite has both a "16-combination AccountingFmt matrix" test file AND a separate "Integer format" test asserting a bare `#,##0` string, one of them is wrong per the VBA source of truth.

### Pitfall 2: "Date BR Longa" tooltip describes a format the code doesn't produce
**What goes wrong:** `src/customUI14.xml:109-115`'s supertip says `Ex.: 15 de março de 2025` (fully spelled-out month with "de...de" literal text), which looks like it needs an Excel format string such as `dd "de" mmmm "de" yyyy`. But `modFormatEngine.bas:150-154` shows the actual `NumberFmt` is `"[$-pt-BR]dd/mmm/yyyy;@"` — 3-letter abbreviated month (`mmm`), slash-separated, no literal "de" text. Rendered output is actually `15/mar/2025`, not `15 de março de 2025`.
**Why it happens:** The tooltip copy and the implementation drifted apart at some point in the VBA project's history; both currently ship together, so this is the add-in's real, current, production behavior — not a bug to "fix" during this migration.
**How to avoid:** Port the exact string `[$-pt-BR]dd/mmm/yyyy;@` for `DATE_BR_LONG`. Do not attempt to build a "more correct" fully-spelled-out format based on the tooltip text — that would be a behavior change, not a migration, and would fail the phase's "byte-for-byte correct against the VBA original" goal.
**Warning signs:** A test asserting the `DATE_BR_LONG` format string contains `mmmm` or literal `"de"` text is testing the tooltip's aspiration, not the shipped VBA behavior.

### Pitfall 3: `FMT_SPREAD_BPS`'s VBA string has escaped quotes — decode before porting
**What goes wrong:** The VBA source line is `f.NumberFmt = "#,##0.0"" bps"""` — a string containing VBA's `""`-escaped double-quote sequences. Copy-pasting or eyeballing this into a C# string literal without correctly resolving the escaping produces a different, wrong string.
**Why it happens:** VBA escapes a literal `"` inside a string by doubling it (`""`). Parsing `"#,##0.0"" bps"""` character by character: `#,##0.0` + (`""` → literal `"`) + ` bps` + (`""` → literal `"`, with the final lone `"` closing the string) yields the actual runtime string **`#,##0.0" bps"`** — i.e., the Excel format code `#,##0.0` followed by a quoted literal text suffix `" bps"` (standard Excel syntax for appending literal text to a number format). This matches the Ribbon tooltip's example ("125,0 bps").
**How to avoid:** Use the pre-decoded value directly: in C#, this is `"#,##0.0\" bps\""` as a string literal (or a verbatim string `@"#,##0.0"" bps"""` using C#'s own doubled-quote escaping convention, which happens to look identical to the VBA source — a nice coincidence worth double-checking with a unit test rather than trusting by eye).
**Warning signs:** A test expects `#,##0.0 bps` (no quotes) or `#,##0.0"bps"` (missing the leading space before "bps") — both are subtly wrong versions of the correctly-decoded string.

### Pitfall 4: `String(0, "0")` / `new string('0', 0)` edge case is already handled — don't "simplify" it away
**What goes wrong:** VBA's `AccountingFmt` has an explicit `If decimals = 0 Then` branch (`modFormatEngine.bas:209-219`) that duplicates the `pos`/`neg`/`zer` construction without the trailing `.` + digits. The comment explains why: `String(0, "0")` returns `""`, and naively concatenating `"0." & ""` would produce a dangling `"0."` (a decimal point with nothing after it) — invalid/misleading as an Excel format code. A "cleaner" C# refactor that tries to unify the 0-decimals and N-decimals cases into one formula (e.g., conditionally omitting just the trailing `.`) is more error-prone than keeping VBA's explicit two-branch structure.
**How to avoid:** Port the two-branch structure as-is (see Code Examples below) rather than attempting a single-formula simplification. The 16-combination test matrix in this document is the ground truth to test against — if a "simplified" implementation doesn't reproduce all 16 exactly, the simplification is wrong.
**Warning signs:** A 0-decimals test case producing `_(#,##0._)_-`  (trailing dot) instead of `_(#,##0_)_-` (no dot).

### Pitfall 5: Don't let `dotnet test` at the `.sln` level attempt to run the `net48` build
**What goes wrong:** If a future contributor adds a *second* test project that also targets `net48` (mirroring the sibling project's single-`net48`-target convention exactly), `dotnet test` at the solution level will attempt to launch that test host on Linux and fail outright — .NET Framework assemblies cannot execute under the CoreCLR-based `dotnet` runtime on Linux, full stop (this is a hard OS/runtime limitation, not a configuration issue).
**How to avoid:** Only the class library multi-targets `net48;net8.0`. The test project must target `net8.0` only. This was verified this session: `dotnet test Spike.sln` correctly built both `net48` and `net8.0` outputs of the library (since it's referenced by the build) but only *ran* the `net8.0` test project — because that was the only project in the solution with test packages referenced.
**Warning signs:** `dotnet test` output showing a test host trying to load `net48` and failing with a platform/runtime error, or CI green on Windows but red on any Linux-based pre-check.

## Code Examples

### AccountingFormatBuilder — full C# port, verified against VBA algorithm
```csharp
// Source: ported from src/modFormatEngine.bas:188-222 (private Function AccountingFmt)
// Verified this session: produces byte-identical output to the VBA algorithm for all 16
// combinations below (independently re-derived from the VBA source, not copied from a guess).
namespace FinanceFmtTools.Engine
{
    internal static class AccountingFormatBuilder
    {
        internal static string Build(int decimals, bool forceAlign, bool zeroDash)
        {
            string dec = new string('0', decimals);
            string pos, neg, zer;

            if (decimals == 0)
            {
                // VBA explicit zero-decimals branch (modFormatEngine.bas:209-219) —
                // avoids the dangling "0." that String(0,"0") concatenation would produce.
                if (forceAlign) { pos = " * _(#,##0_)_-"; neg = " * (#,##0)_-"; }
                else             { pos = "_(#,##0_)_-";     neg = "(#,##0)_-"; }
            }
            else
            {
                if (forceAlign) { pos = " * _(#,##0." + dec + "_)_-"; neg = " * (#,##0." + dec + ")_-"; }
                else             { pos = "_(#,##0." + dec + "_)_-";     neg = "(#,##0." + dec + ")_-"; }
            }

            zer = zeroDash ? (forceAlign ? " * _(-_)_-" : "_(-_)_-") : pos;

            return pos + ";" + neg + ";" + zer;
        }
    }
}
```

### All 16 expected values — use directly as `[Theory]`/`[InlineData]` fixtures (FMT-01, FMT-07)
Generated this session by running the exact algorithm above via `dotnet run` on a genuine Linux .NET 8 SDK — copy these values directly into the test suite rather than re-deriving them by hand.

| Decimals | ForceAlign | ZeroDash | Result |
|---|---|---|---|
| 0 | false | false | `_(#,##0_)_-;(#,##0)_-;_(#,##0_)_-` |
| 0 | false | true | `_(#,##0_)_-;(#,##0)_-;_(-_)_-` |
| 0 | true | false | ` * _(#,##0_)_-; * (#,##0)_-; * _(#,##0_)_-` |
| 0 | true | true | ` * _(#,##0_)_-; * (#,##0)_-; * _(-_)_-` |
| 2 | false | false | `_(#,##0.00_)_-;(#,##0.00)_-;_(#,##0.00_)_-` |
| 2 | false | true | `_(#,##0.00_)_-;(#,##0.00)_-;_(-_)_-` |
| 2 | true | false | ` * _(#,##0.00_)_-; * (#,##0.00)_-; * _(#,##0.00_)_-` |
| 2 | true | true | ` * _(#,##0.00_)_-; * (#,##0.00)_-; * _(-_)_-` |
| 4 | false | false | `_(#,##0.0000_)_-;(#,##0.0000)_-;_(#,##0.0000_)_-` |
| 4 | false | true | `_(#,##0.0000_)_-;(#,##0.0000)_-;_(-_)_-` |
| 4 | true | false | ` * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(#,##0.0000_)_-` |
| 4 | true | true | ` * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(-_)_-` |
| 8 | false | false | `_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(#,##0.00000000_)_-` |
| 8 | false | true | `_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(-_)_-` |
| 8 | true | false | ` * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(#,##0.00000000_)_-` |
| 8 | true | true | ` * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(-_)_-` |

### Non-accounting registry entries — literal values (FMT-02, FMT-03, FMT-04, FMT-05)
| Format Key | VBA Source | Exact Resulting String |
|---|---|---|
| `PCT_4D` | `modFormatEngine.bas:118-122` | `0.0000%` |
| `PCT_2D` | `modFormatEngine.bas:124-128` | `0.00%` |
| `SPREAD_BPS` | `modFormatEngine.bas:131-135` (escaped-quote decode — see Pitfall 3) | `#,##0.0" bps"` |
| `DATE_ISO` | `modFormatEngine.bas:138-142` | `yyyy-mm-dd;@` |
| `DATE_BR` | `modFormatEngine.bas:144-148` | `[$-pt-BR]dd/mm/yyyy;@` |
| `DATE_BR_LONG` | `modFormatEngine.bas:150-154` (tooltip mismatch — see Pitfall 2) | `[$-pt-BR]dd/mmm/yyyy;@` |
| `TEXT` | `modFormatEngine.bas:157-161` | `@` |

### `[Theory]` xUnit pattern for the 16-combination matrix
```csharp
// Source: xunit.net official docs (Theory/InlineData) — https://xunit.net/docs/getting-started/v2/getting-started
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    public class AccountingFormatBuilderTests
    {
        [Theory]
        [InlineData(0, false, false, "_(#,##0_)_-;(#,##0)_-;_(#,##0_)_-")]
        [InlineData(0, false, true,  "_(#,##0_)_-;(#,##0)_-;_(-_)_-")]
        [InlineData(0, true,  false, " * _(#,##0_)_-; * (#,##0)_-; * _(#,##0_)_-")]
        [InlineData(0, true,  true,  " * _(#,##0_)_-; * (#,##0)_-; * _(-_)_-")]
        [InlineData(2, false, false, "_(#,##0.00_)_-;(#,##0.00)_-;_(#,##0.00_)_-")]
        // ... remaining 11 rows from the table above ...
        public void Build_MatchesVbaAlgorithm(int decimals, bool forceAlign, bool zeroDash, string expected)
        {
            string actual = AccountingFormatBuilder.Build(decimals, forceAlign, zeroDash);
            Assert.Equal(expected, actual);
        }
    }
}
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|---------------|--------|
| VBA `Select Case` registry with `Public Type FormatDef` and empty-string sentinel for "not found" | C# switch expression + `bool TryGetFormatDef(...)` idiomatic Try-pattern | This migration (Phase 1) | Compiler-enforced not-found handling; no risk of forgetting the `.key = ""` check |
| VBA `Public` module-level globals (`CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`) read implicitly by the format function | Explicit parameters passed into a pure function | This migration (Phase 1) | Eliminates a whole class of shared-mutable-state test flakiness under xUnit's parallel test execution; also correctly defers config-state ownership to Phase 2 per the roadmap's own layering |
| .NET Framework reference assemblies only available via a full Visual Studio/Windows SDK install | `Microsoft.NETFramework.ReferenceAssemblies` NuGet package (Microsoft-published, ~2019 onward) | Established pattern, not new in 2026 | Makes `net48` cross-compilation from Linux/Mac CI runners routine — this is *why* the package exists |

**Deprecated/outdated:** None specific to this phase — the format-string algorithm itself (VBA's `AccountingFmt`) is being preserved exactly, not modernized; only its packaging/tooling is being modernized.

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|----------------|
| A1 | `LangVersion 9.0` (matching the sibling `outlook-classic-delay-send` project's convention) is the right choice for this new project, rather than leaving `LangVersion` unset (defaulting to C# 7.3 on `net48`, latest on `net8.0`) | Architecture Patterns, Pattern 1 | Low — this is a style/consistency choice, not a correctness one. If unset, switch expressions (C# 8+) wouldn't compile on the `net48` leg, forcing a less elegant `if`/`else if` chain in `FormatRegistry`. Either way the phase's success criteria (test pass/fail) are unaffected. |
| A2 | `FormatDef` should be a plain class/struct rather than a `readonly struct` or other specific shape | Architecture Patterns, Pattern 2 | Low — any immutable, non-record shape avoids the `IsExternalInit` problem; the exact choice (class vs. struct) is a minor performance/ergonomics tradeoff with no bearing on the phase's test-pass success criteria. |

**If this table is empty:** N/A — two low-risk style assumptions are logged above; everything else in this document (package versions, build/test viability, VBA algorithm decoding, tooltip-vs-code discrepancies) was either empirically verified this session or directly read from the VBA source files.

## Open Questions

1. **Should `FormatDef.Alignment` reuse Excel's own `XlHAlign`-style values or a Phase-1-private enum?**
   - What we know: Phase 1's explicit goal is "zero Excel/COM references," and `Microsoft.Office.Interop.Excel` isn't available to reference from a `net48;net8.0`-multi-targeted project without breaking the `net8.0` leg entirely (COM interop doesn't exist on non-Windows/non-Framework targets).
   - What's unclear: Nothing technical — the constraint makes the answer unambiguous. This is listed only so the planner explicitly names the private enum (e.g., `CellAlignment`) rather than accidentally deferring the decision into Phase 2 in a way that creates rework.
   - Recommendation: Define `CellAlignment { General, Left, Right }` in Phase 1 (see Recommended Project Structure); Phase 2's `IExcelGateway` implementation maps this to the real `XlHAlign` enum when it exists.

2. **Should the `net8.0` leg of the class library ship as a NuGet package or stay dev/test-only?**
   - What we know: CLAUDE.md and ROADMAP.md are unambiguous that the *distributed* artifact is `net48` (COM-hosted inside Excel). The `net8.0` leg's only purpose, per CONTEXT.md's discretion note, is enabling `dotnet test` on this Linux dev machine.
   - What's unclear: Nothing blocking — flagged only so the planner doesn't accidentally scope in packaging/publishing work for the `net8.0` build output, which is out of scope for this phase (and arguably any phase — Phase 5's CI/CD packages only the `net48` output per ROADMAP.md Phase 5 success criteria).
   - Recommendation: Treat `net8.0` as build/test-infrastructure only; do not add any `<Pack>`/publish steps for it.

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| .NET SDK (`dotnet` CLI) | All of DEV-01, FMT-07 | ✗ (not pre-installed in this research sandbox — confirmed `dotnet: command not found`, no apt/snap package, no `/usr/share/dotnet`) | — | Install via the official script: `curl -sSL https://dot.net/v1/dotnet-install.sh \| bash -s -- --channel 8.0 --install-dir <dir>`, then add `<dir>` to `PATH`. **Verified working this session** — installed SDK 8.0.422 (Linux x64) this way and used it for every empirical check in this document. |
| pip / PyPI tooling (for `slopcheck`) | Package Legitimacy Gate protocol | ✗ (`pip`/`pip3` not present in this environment) | — | None needed for this phase — see Package Legitimacy Audit for the substitute verification method used (official NuGet registry + empirical build/test). |
| A Windows machine or .NET Framework CLR | Actually *executing* the final `net48` add-in assembly | ✗ (this dev environment is Linux/WSL only, per the phase's stated constraint) | — | Not required for Phase 1 — `net48` correctness is verified by **compilation only** (`dotnet build`), never execution, in this phase. Actual execution against a live Excel session is explicitly Phase 3's job (ROADMAP.md), which the project's other constraints (Windows + Excel 2016+) assume will run on a real Windows machine. |

**Missing dependencies with no fallback:** None — every missing dependency above has a working, verified fallback.

**Missing dependencies with fallback:**
- `.NET 8 SDK` — install via `dotnet-install.sh` (verified this session, see above). The planner should make "install/verify the .NET 8 SDK is on `PATH`" an explicit first task if this phase's execution environment doesn't already have it pre-provisioned.

## Security Domain

`security_enforcement` is absent from `.planning/config.json`'s `workflow` block, which per the enforcement default means **enabled**. Applying that lens honestly to a phase with essentially no attack surface:

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-------------------|
| V2 Authentication | No | No authentication concept exists in a local Excel add-in's format engine |
| V3 Session Management | No | No sessions — Phase 2 will hold in-memory config state, but that's not a security session concept |
| V4 Access Control | No | No access-control boundary within a single-user desktop process |
| V5 Input Validation | Marginally yes | `AccountingFormatBuilder.Build`'s `decimals` parameter should not accept negative values silently — `new string('0', decimals)` throws `ArgumentOutOfRangeException` for negative `decimals` in .NET (unlike VBA's `String()`, which throws a different runtime error 5 for the same case). Since Phase 1's registry only ever calls this with the hardcoded constants `0`, `2`, `4`, `8`, this is not user-reachable input in practice — but the registry's `TryGetFormatDef` returning `false` for an unrecognized `key` string (rather than throwing) is the one genuinely load-bearing "input validation" behavior this phase must preserve (mirrors VBA's `Case Else` fallback), because Phase 2's guard-clause logic (FMT-06) depends on it not throwing. |
| V6 Cryptography | No | No cryptographic operations in this phase |

### Known Threat Patterns for this stack
None apply meaningfully. This is a pure, side-effect-free, single-process, no-network, no-persistence string-transformation library — there is no injection surface (the `key` parameter is always an internal constant, never end-user free text; the only "external" input is which of 12 fixed Ribbon buttons was clicked, and that mapping is compiled into the add-in, not user-supplied). The one real hygiene item is the `TryGetFormatDef` no-throw-on-unknown-key contract noted under V5 above, since Phase 2's whole "friendly message instead of crash" guarantee (FMT-06) is built on top of it.

## Sources

### Primary (HIGH confidence — empirical, this session)
- Live `dotnet build`/`dotnet test`/`dotnet new sln`/`dotnet sln add` runs against a freshly-installed Linux x64 .NET 8 SDK (8.0.422), executed in this research session against a structural clone of the recommended Phase 1 layout (`net48;net8.0` multi-target library + `net8.0`-only xUnit test project + `Microsoft.NETFramework.ReferenceAssemblies`). Confirms: 0 warnings/0 errors build, `dotnet test` passes at both project and `.sln` level, `record`/`init` fails with `CS0518` on `net48`, `switch` expressions compile fine on `net48` with `LangVersion 9.0`.
- `src/modFormatEngine.bas`, `src/modConfig.bas`, `src/customUI14.xml`, `src/modRibbon.bas` (this repo) — read in full; source of the "Integer = Fin 0D" and "Date BR Longa tooltip mismatch" and "Spread bps quote-escaping" findings.

### Secondary (MEDIUM-HIGH confidence — official docs, cross-checked)
- [nuget.org/packages/Microsoft.NETFramework.ReferenceAssemblies](https://www.nuget.org/packages/Microsoft.NETFramework.ReferenceAssemblies) — package purpose, version 1.0.3, download count
- [nuget.org/packages/xunit](https://www.nuget.org/packages/xunit) — version 2.9.3, 962.8M downloads, github.com/xunit/xunit
- [nuget.org/packages/Microsoft.NET.Test.Sdk](https://www.nuget.org/packages/microsoft.net.test.sdk) — versions 17.x/18.x, 1.7B downloads, github.com/microsoft/vstest, 18.0.0 breaking-change notes (dropped net6.0, requires net8+)
- [nuget.org/packages/xunit.runner.visualstudio](https://www.nuget.org/packages/xunit.runner.visualstudio) — version 3.1.5, 970.7M downloads, github.com/xunit/visualstudio.xunit, supports xUnit v1/v2/v3 on net472+/net8+
- [api.nuget.org NuGet v3 flatcontainer index](https://api.nuget.org/v3-flatcontainer/) — authoritative version-list confirmation for all four packages
- [github.com/manuelroemer/IsExternalInit](https://github.com/manuelroemer/IsExternalInit) and [mking.net CS0518 writeup](https://www.mking.net/blog/error-cs0518-isexternalinit-not-defined) — corroborate the `record`/`init` on `net48` gotcha, independently reproduced empirically in this session
- `/home/thomaz/pessoal/outlook-classic-delay-send` (sibling reference project explicitly named in PROJECT.md) — `UndoSend.csproj`, `UndoSend.Tests.csproj`, `BUILD.md`, `.github/workflows/release.yml`, `RecoveryPolicyTests.cs` — used to confirm this user's established C# conventions (LangVersion 9.0, plain classes not records, xUnit `[Fact]`/`[Theory]` style, `Domain/` layer with zero COM references) and, importantly, to confirm that project's tests were only ever run/proven on a real Windows machine — it does **not** already solve the Linux-execution problem this phase's CONTEXT.md flags, which is why the empirical spike in this session was necessary rather than assuming the sibling's pattern transfers directly.

### Tertiary (LOW confidence)
None — every claim in this document is either empirically verified this session or read directly from a first-party source (this repo's VBA files or official NuGet/vendor documentation).

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH — package existence/versions confirmed via live registry + official docs; end-to-end build/test viability confirmed empirically, not by citation alone
- Architecture: HIGH — directly derived from reading the actual VBA source line-by-line, cross-checked against the roadmap's own phase boundaries
- Pitfalls: HIGH — three of five pitfalls (record/init failure, spread-bps quote decoding, Integer=Fin-0D) are either empirically reproduced or mechanically derived from the source text, not inferred

**Research date:** 2026-07-10
**Valid until:** 30 days (stable domain — .NET Framework/8 SDK tooling changes slowly; re-verify NuGet package versions if planning is delayed past ~2026-08-10, especially `Microsoft.NET.Test.Sdk` which is mid-major-version-bump at time of writing)
