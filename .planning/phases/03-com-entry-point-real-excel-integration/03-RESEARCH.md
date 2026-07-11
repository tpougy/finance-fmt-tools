# Phase 3: COM Entry Point & Real Excel Integration - Research

**Researched:** 2026-07-11
**Domain:** Classic (non-VSTO) .NET Framework 4.8 COM "Shared Add-in" for Excel — `IDTExtensibility2` + `IRibbonExtensibility`, real `Microsoft.Office.Interop.Excel` COM interop, cross-compiled from Linux with zero Windows/Excel access
**Confidence:** HIGH — the core technical risk (can this even compile without Windows/Office?) was resolved **empirically, in this session**, using the project's actual Phase 1/2 code, not a synthetic clone.

<user_constraints>
## User Constraints (from CONTEXT.md)

### Locked Decisions
None locked as user decisions — discuss phase was skipped (`workflow.skip_discuss`, full autonomous `/gsd-autonomous` run). However, one constraint is explicitly **non-discretionary** per 03-CONTEXT.md:

> **Critical environment constraint (not discretionary — must be planned around):** This phase's own success criteria explicitly require a live Excel session for full verification (RIB-01 through RIB-04), and this development environment is Linux/WSL with no Windows, no Excel, and no COM runtime available. There is no way to execute or interactively smoke-test real `Microsoft.Office.Interop.Excel` COM code from here. The plan MUST still produce the real, buildable C# COM add-in code (Connect.cs entry point, Ribbon XML wiring via the already-embedded customUI14.xml from Phase 2, real Microsoft.Office.Interop.Excel implementations of Phase 2's IExcelGateway/IRangeHandle interfaces) — compiled and unit-testable pieces should be proven via dotnet build/dotnet test as far as possible, but the actual "run inside live Excel and click every button" verification is expected to come back as human_needed and will be deferred to the user, who has a real Windows+Excel machine. Do not attempt to fake or skip this reality — plan the real COM code, build it, and clearly scope what can vs. cannot be verified in this environment.

### Claude's Discretion
All implementation choices (project structure, GUID/ProgID values, exact `ILog`/gateway implementation shape, whether to hand-roll or NuGet-source the `Extensibility` assembly, `IRibbonUI` caching strategy) are at Claude's discretion. Use ROADMAP phase goal, success criteria, codebase conventions (VBA source, Phase 1/2 code, the sibling `outlook-classic-delay-send` project explicitly named in CLAUDE.md as this project's dev/build/release inspiration), and this research to guide decisions.

### Deferred Ideas (OUT OF SCOPE)
None — discuss phase skipped. Explicitly **not** this phase's job (confirmed against ROADMAP.md and REQUIREMENTS.md):
- Actual COM registration / `regasm` / HKCU registry writes / installer script — **Phase 4** (INST-01/02/03).
- Persisting `ForceAlign`/`ZeroDash` across sessions — explicitly out of scope for the whole migration (REQUIREMENTS.md "Out of Scope" table).
- 32-bit Excel support — explicitly out of scope (FUT-01).
</user_constraints>

<phase_requirements>
## Phase Requirements

| ID | Description | Research Support |
|----|-------------|------------------|
| RIB-01 | A aba "Finance Fmt" aparece na Ribbon com os mesmos grupos, botões e tooltips da versão VBA | Reuse Phase 2's embedded `customUI14.xml` (unmodified) via `Connect.GetCustomUI` returning `RibbonController.GetCustomUiXml()`. The XML itself is untouched, so group/button/tooltip parity is structural, not something Phase 3 needs to re-derive — see Architecture Patterns Pattern 1 and the `ClassInterfaceType.AutoDispatch` pitfall (if this is wrong, buttons render but the tab shows blank/errors). |
| RIB-02 | Checkbox "Alinhar à direita" funciona durante a sessão, inicia sempre desligado, sem persistência | Wire `RibbonChkForceAlign`/`RibbonGetForceAlign` directly to Phase 2's `RibbonController.Config.ForceAlign` (already defaults `false`, already has zero persistence) — Phase 3 adds no new state, just a live COM/IDispatch pass-through. See Code Examples. |
| RIB-03 | Checkbox "Zero contábil" funciona durante a sessão, inicia sempre ligado, sem persistência | Same mechanism as RIB-02, using `RibbonController.Config.ZeroDash` (defaults `true`). |
| RIB-04 | Botão "Sobre" e o link de documentação funcionam a partir da Ribbon | `RibbonAbout` → `MessageBox.Show` (WinForms, needs `UseWindowsForms=true`); `RibbonFinInfo` → `Process.Start(new ProcessStartInfo(url){UseShellExecute=true})`. Both empirically verified to compile in this session — see Code Examples and Common Pitfalls (`Process.Start` default). |

Additionally (from ROADMAP.md's fuller success-criteria list, not separately ID'd but load-bearing for this phase): clicking each of the 12 format buttons must apply the correct format via the real Excel object model, and selecting a Chart/Shape must show a friendly message instead of crashing — this is the live-COM completion of Phase 1/2's `FormatRegistry`/`FormatEngine`, requiring a real `IExcelGateway`/`IRangeHandle` implementation over `Microsoft.Office.Interop.Excel` (see Architecture Patterns Pattern 3 and Pattern 4).
</phase_requirements>

## Project Constraints (from CLAUDE.md)

| Directive | Source | Phase 3 Relevance |
|-----------|--------|---------------------|
| Platform: Windows + Excel 2016+ | CLAUDE.md Constraints | This is the FIRST phase where the "Windows + Excel" constraint becomes structurally unavoidable in the code itself (real COM types) — but the *build* must still succeed on this Linux dev box, per the environment constraint above. Execution/manual verification is deferred to the user's real Windows+Excel machine. |
| Tooling: VS Code + dotnet CLI, no full Visual Studio | CLAUDE.md Constraints | Empirically confirmed this session: `dotnet build` (no MSBuild-from-VS fallback needed) successfully compiles a net48 project referencing real Excel/Office COM interop types + `ComVisible`/`Guid`/`ClassInterface` attributes, on Linux. No `<COMReference>`/`ResolveComReference`/`tlbimp` used anywhere (see Pitfall 1) — that MSBuild task is what would have forced a Visual Studio/Windows-only build, and it is deliberately avoided. |
| Runtime: .NET Framework 4.8, buildable with .NET 8 SDK | CLAUDE.md Constraints | The new COM-entry-point project targets **`net48` only** (NOT multi-targeted with `net8.0` — COM Shared-Add-in hosting via `mscoree.dll`/regasm is a .NET-Framework-only mechanism; see Pitfall 5). Built via `dotnet build` using .NET 8 SDK 8.0.422, already installed in this environment at `/home/thomaz/.dotnet` — confirmed working. |
| Instalação: HKCU-only, sem admin | CLAUDE.md Constraints | Not this phase's job to *implement* (Phase 4), but this phase's `Connect.cs` MUST declare the `[Guid]`/`[ProgId]`/`[ComVisible(true)]` values that Phase 4's installer will read/reuse verbatim — see "What Phase 3 Must Expose for Phase 4" below. |
| Compatibilidade de UX: mesmos botões/atalhos na Ribbon | CLAUDE.md Constraints | Directly satisfied by reusing Phase 2's embedded, unmodified `customUI14.xml` — Phase 3 must not edit that file's control ids/labels/callback names. |
| GSD Workflow Enforcement | CLAUDE.md | Process constraint, not directly actionable in task content. |

## Summary

Phase 3 adds the **first COM-referencing project** to the solution: a `net48`-only class library that is the actual, loadable Excel add-in — a classic "Shared COM Add-in" (pre-VSTO pattern) implementing `Extensibility.IDTExtensibility2` (lifecycle) and `Microsoft.Office.Core.IRibbonExtensibility` (Ribbon XML), with one public method per Ribbon callback (`RibbonFin2D`, `RibbonChkForceAlign`, `RibbonGetForceAlign`, `RibbonAbout`, etc.) matching `src/customUI14.xml`'s `onAction`/`getPressed` attributes exactly. This project references the existing, unchanged `FinanceFmtTools.Engine` (Phase 1/2) via a normal `ProjectReference` and implements Phase 2's `IExcelGateway`/`IRangeHandle`/`ILog` interfaces against the real `Microsoft.Office.Interop.Excel` object model.

**The single biggest open risk flagged in STATE.md — "can real COM interop code, including `IDTExtensibility2`/`IRibbonExtensibility`/`Microsoft.Office.Interop.Excel`, actually be *compiled* (not run) via `dotnet build` on this Linux/WSL box with zero Windows/Office access" — was resolved empirically in this research session, not by citation.** A full mini version of Phase 3's architecture (hand-rolled `Extensibility.IDTExtensibility2` shim + NuGet-sourced `Microsoft.Office.Interop.Excel`/`Microsoft.Office.Core` + a real `IExcelGateway`/`IRangeHandle` implementation + `Connect : IDTExtensibility2, Office.IRibbonExtensibility` + `MessageBox.Show` + `Process.Start` + a live `ProjectReference` to the **actual, unmodified** `FinanceFmtTools.Engine` project, calling the **actual** `FormatEngine.ApplyToSelection`/`FormatKeys.Fin2D`/`RibbonController.GetCustomUiXml()`) was built and compiled with `dotnet build -c Release`: **0 Warnings, 0 Errors**. This is materially stronger evidence than Phase 1's own spike (which used a structural clone) — this session's spike used the project's real, already-committed code.

The second major finding is a genuine, non-obvious sourcing gap: there is **no official Microsoft NuGet package for the classic `Extensibility` assembly** (the tiny `IDTExtensibility2`/`ext_ConnectMode`/`ext_DisconnectMode` types, historically shipped as `Extensibility.dll`/`msaddndr.dll`). Two real options were found and empirically verified:
1. **Hand-declare the interface locally** (recommended) — `IDTExtensibility2` is a stable, 5-method, publicly documented COM interface with a fixed IID (`B65AD801-ABAF-11D0-BB8B-00A0C90F2744`, confirmed against Microsoft Learn's own published attribute); COM resolves interfaces by GUID, not by which assembly declares the .NET type, so a local `[ComImport, Guid("B65AD801-...")]` declaration is functionally identical to referencing the real `Extensibility.dll`. Verified empirically to compile and produces a much smaller distributable (no 2MB+ extra dependency).
2. **`Microsoft.VisualStudio.Interop` NuGet package (official Microsoft publisher)** — confirmed via direct binary inspection (`strings` on the downloaded `.dll`) to physically contain the `Extensibility` namespace with the exact matching GUID `B65AD801-ABAF-11D0-BB8B-00A0C90F2744`. This is a legitimate, if surprising, official source (VS add-ins and pre-VSTO Office add-ins historically shared the same "Add-in Designer" COM library), but it drags in a heavy (~2.1MB) VS-specific assembly plus one transitive dependency, which is confusing to find inside an Excel add-in's dependency list.

The third finding, requiring explicit user awareness: the two remaining Excel-specific interop packages needed (`Microsoft.Office.Interop.Excel` for the Excel Object Model, and `MicrosoftOfficeCore16`/`MicrosoftOfficeCore` for `Microsoft.Office.Core.IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`) are **not** officially published by Microsoft — both are third-party repackages (publisher `CamronBute` and `kep` respectively) of the real, licensed Office PIA binaries. Their `.nuspec` files self-report `<authors>Microsoft</authors>` but the actual NuGet *owner* account is not a verified Microsoft publisher, and both explicitly state **"there is no license since it is a repackaging of Office assemblies."** They are empirically real and functional (confirmed via direct binary content inspection — genuine `Excel.Range`/`Workbook`/`Worksheet`/`XlHAlign` and `Office.IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl` types are present, matching official Microsoft Learn documentation), extremely widely used (30.2M downloads for `Microsoft.Office.Interop.Excel`, and it is literally the answer Microsoft's own Q&A support forum gives to "what NuGet package do I use for Excel interop"), and are the only practical way to compile this project **without a Windows+Office machine to vendor real PIA DLLs from** (the alternative the sibling project used — see Pitfall 1). This is a genuine licensing/support-ambiguity risk distinct from "is this a real/safe package" — flagged for explicit user confirmation, not a hallucination risk.

A fourth finding, load-bearing and specific to this exact class of add-in: the sibling `outlook-classic-delay-send` project's `Connect.cs` carries a **first-hand, empirically-diagnosed production bug fix** directly transferable to this project — `[ClassInterface(ClassInterfaceType.None)]` (a common, "more correct-looking" COM interop choice) silently breaks **all Ribbon callbacks that are not part of `IDTExtensibility2`/`IRibbonExtensibility`** (i.e. every `onAction`/`getPressed` method: `RibbonFin2D`, `RibbonChkForceAlign`, `RibbonGetForceAlign`, etc.), because Office invokes those by **name** via late-bound `IDispatch.GetIDsOfNames`, which `ClassInterfaceType.None` does not expose. The fix, diagnosed in a live Excel-family (Outlook) smoke test by the same author, is `[ClassInterface(ClassInterfaceType.AutoDispatch)]`. This cannot be independently re-verified in this session (no live Excel available), but it is exactly the same failure mode this project's `Connect.cs` would hit, and the fix is a one-attribute change with no downside for this project's scope (a private, non-reusable, single-purpose add-in class — the usual "why not AutoDispatch" versioning concerns for public COM libraries don't apply).

**Primary recommendation:** Add a new `FinanceFmtTools.ComAddin` project (`net48` only, `UseWindowsForms=true`, `AnyCPU`) referencing `FinanceFmtTools.Engine` via `ProjectReference` plus three interop dependencies (hand-rolled `Extensibility` shim + NuGet `Microsoft.Office.Interop.Excel` + NuGet `MicrosoftOfficeCore16`), implementing `Connect : IDTExtensibility2, Office.IRibbonExtensibility` with `[ClassInterface(ClassInterfaceType.AutoDispatch)]`, a small `AddInHost` composition class (mirroring the sibling's proven pattern) wiring a real `ExcelGateway`/`ExcelRangeHandle` + a `TraceLog` + the **unmodified** Phase 2 `RibbonController`. Build-verify via `dotnet build` in this environment; defer all actual Excel-session behavior verification to a `checkpoint:human-verify`/`human_needed` task on the user's real Windows+Excel machine.

## Architectural Responsibility Map

| Capability | Primary Tier | Secondary Tier | Rationale |
|------------|-------------|----------------|-----------|
| COM entry point / lifecycle (`IDTExtensibility2`) | COM/Excel Integration (Phase 3, this phase) | — | Only a real Windows+Excel host can instantiate this via `mscoree.dll`; must be a thin, non-throwing shell (see Pattern 2) |
| Ribbon XML delivery (`IRibbonExtensibility.GetCustomUI`) | COM/Excel Integration (Phase 3) | Ribbon/UI declaration (Phase 2's embedded XML, unchanged) | `GetCustomUI` is a 1-line pass-through to `RibbonController.GetCustomUiXml()` — no new XML authored this phase |
| Live `IRibbonUI` caching / `InvalidateControl` | COM/Excel Integration (Phase 3) | — | Requires a real `IRibbonUI` object that only exists once Excel calls `OnRibbonLoad`; see Pattern 5 for why this phase should cache+invalidate even though VBA never did |
| Real `IExcelGateway`/`IRangeHandle` implementation | COM/Excel Integration (Phase 3) | Abstraction/Orchestration (Phase 2's interfaces, unchanged) | Phase 2 defined the seam; Phase 3 is the only phase that can fill it with real `Excel.Application`/`Excel.Range` |
| Format-key → `NumberFormat` resolution | Domain/Logic (Phase 1, unchanged) | — | `FormatRegistry`/`AccountingFormatBuilder` — Phase 3 calls this through `FormatEngine`, never re-implements it |
| Session checkbox state (`ForceAlign`/`ZeroDash`) | Abstraction/Orchestration (Phase 2's `RibbonSessionConfig`, unchanged) | Ribbon/UI (Phase 3's `getPressed`/`onAction` wiring) | Phase 3 adds zero new state — it is a live pass-through to the already-complete `RibbonController.Config` |
| FMT-06 friendly message (`MessageBox`) | COM/Excel Integration (Phase 3) | Orchestration (Phase 2's no-throw guard, unchanged) | Phase 2 proved "logs a warning, never throws"; Phase 3 adds the actual user-facing dialog on top, without modifying Phase 2's tested code (see Pattern 4) |
| About dialog / docs link (RIB-04) | COM/Excel Integration (Phase 3) | — | Pure `System.Windows.Forms`/`System.Diagnostics.Process` calls, no Excel object needed at all |
| COM registration (HKCU keys, ProgID/CLSID writes) | Installer (Phase 4, not this phase) | COM/Excel Integration (Phase 3 only *declares* the values via attributes) | Phase 3 fixes the GUID/ProgID/AssemblyName in code; Phase 4 consumes them verbatim in `install.ps1` |

## Standard Stack

### Core
| Library | Version | Purpose | Why Standard |
|---------|---------|---------|--------------|
| Microsoft.NETFramework.ReferenceAssemblies | 1.0.3 | net48 build-only reference assemblies | Already approved/pinned in Phase 1; reused unchanged for the new project's `net48` target. `[VERIFIED: nuget.org + empirical dotnet build success this session]` |
| Hand-rolled `Extensibility` shim (local source file, no package) | n/a (local code, IID matches the real `Extensibility.IDTExtensibility2`) | Provides `IDTExtensibility2`/`ext_ConnectMode`/`ext_DisconnectMode` for the COM entry point | No lightweight official NuGet package exists for this (search returned zero dedicated results). The interface is small (5 methods, 2 enums), stable since Office 97/VS pre-VSIX era, and COM resolves it by GUID, not assembly identity. `[VERIFIED: empirical — compiled this session with the exact documented IID `B65AD801-ABAF-11D0-BB8B-00A0C90F2744` sourced from Microsoft Learn's own published attribute on `Extensibility.IDTExtensibility2`]` |
| Microsoft.Office.Interop.Excel | 16.0.18925.20022 | Excel Object Model (`Application`, `Range`, `Workbook`, `Worksheet`, `XlHAlign`, ...) | The de facto standard NuGet package for this exact scenario (30.2M downloads; Microsoft's own Q&A forum recommends it). **Not an officially Microsoft-published package** (owner: `CamronBute`) — `[ASSUMED — see Package Legitimacy Audit]`. Content verified via direct binary inspection this session (`strings` confirms real `Application`/`Range`/`Workbook`/`Worksheet`/`XlHAlign` type names) and empirical `dotnet build` success. |
| MicrosoftOfficeCore16 | 16.0.16626.20000 | `Microsoft.Office.Core` (`IRibbonExtensibility`, `IRibbonUI`, `IRibbonControl`) | Same publisher family as the Excel package above (consistency of trust profile), current versioning matching Office 16.0. **Not officially Microsoft-published** — `[ASSUMED — see Package Legitimacy Audit]`. Content verified via `strings` this session: contains `IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`/`GetCustomUI`/`InvalidateControl`. |

### Supporting
| Library | Version | Purpose | When to Use |
|---------|---------|---------|-------------|
| `System.Windows.Forms` (net48 framework assembly, via `<UseWindowsForms>true</UseWindowsForms>`) | (part of .NET Framework 4.8) | `MessageBox.Show` for the FMT-06 friendly message and the RIB-04 "Sobre" dialog | Simplest, zero-extra-dependency way to show a modal message from a COM add-in — matches VBA's `MsgBox` conceptually. Verified to compile this session. |
| `System.Diagnostics.Process` (BCL, no package) | (part of .NET Framework 4.8) | `RibbonFinInfo` → open the docs URL in the default browser | Replaces VBA's `ThisWorkbook.FollowHyperlink`; no Excel object needed. Verified to compile this session (see Pitfall 4 for the `UseShellExecute` nuance). |

### Alternatives Considered
| Instead of | Could Use | Tradeoff |
|------------|-----------|----------|
| Hand-rolled `Extensibility` shim (recommended) | `Microsoft.VisualStudio.Interop` NuGet package (official Microsoft publisher) | Officially sanctioned, zero hand-rolled COM code — but adds a ~2.1MB assembly plus one transitive package (`Microsoft.VisualStudio.Imaging.Interop.14.0.DesignTime`) purely to obtain a 5-method interface, and is architecturally confusing (a Visual Studio SDK package referenced from an Excel Ribbon add-in). Empirically verified this session (`strings` confirms the exact matching GUID is physically present in the shipped `.dll`). Switch to this if the hand-rolled shim ever causes a live-Excel loading problem that can't otherwise be diagnosed — see Assumptions Log A2. |
| `Microsoft.Office.Interop.Excel` NuGet (unofficial repackage) | Vendor real `Microsoft.Office.Interop.Excel.dll`/`office.dll`/`stdole.dll` from the GAC of a real, licensed Windows+Office installation (the approach the sibling `outlook-classic-delay-send` project actually used, since its dev environment *had* Windows+Office) | Removes the licensing ambiguity entirely (genuine, licensed Microsoft binaries, matching the target machine's exact installed Office build) — but **requires access to a Windows machine with Office installed**, which this dev sandbox does not have. This is the cleaner long-term choice if/when the user builds on their own Windows+Excel machine; not available in this Linux/WSL research session. |
| `MicrosoftOfficeCore16` (CamronBute, 16.0.16626.20000, low download count ~15K) | `MicrosoftOfficeCore` (kep, 15.0.0, 2.2M downloads, `net35` only, last updated 2017) | Far more downloads/longer track record, but stale (2017) and `net35`-targeted (still forward-compatible with `net48`, but unmaintained). The Ribbon interfaces (`IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`) have not changed since Office 2007, so either works functionally — `MicrosoftOfficeCore16` was chosen for version-family consistency with the Excel Interop package, not because it is objectively safer. |

**Installation:**
```bash
dotnet add src/FinanceFmtTools.ComAddin package Microsoft.NETFramework.ReferenceAssemblies --version 1.0.3
dotnet add src/FinanceFmtTools.ComAddin package Microsoft.Office.Interop.Excel --version 16.0.18925.20022
dotnet add src/FinanceFmtTools.ComAddin package MicrosoftOfficeCore16 --version 16.0.16626.20000
# Extensibility.IDTExtensibility2 shim: no package — add Extensibility.cs directly (see Code Examples)
```

**Version verification:** All versions above were confirmed against the live NuGet registry (`nuget.org`) and cross-checked by downloading the actual `.nupkg` and inspecting file contents with `strings` in this research session (not by web-search citation alone).

## Package Legitimacy Audit

`slopcheck` could not be installed in this research environment (no `pip`/`pip3` binary present — same finding as Phase 1). Per the graceful-degradation protocol, packages would normally all be marked `[ASSUMED]`. As in Phase 1, official-Microsoft-owned packages here were independently verified through two channels stronger than `slopcheck` would provide for NuGet: (1) direct nuget.org owner/metadata inspection, and (2) **downloading the actual `.nupkg` and running `strings` against the contained `.dll`** to confirm the exact expected types/GUIDs are physically present — not just that a package with this name exists. The two Office-interop repackages are genuinely functional and non-malicious (confirmed the same way) but carry a distinct, real risk: **no official Microsoft publisher, no clear license** — flagged accordingly, not as a hallucination risk.

| Package | Registry | Age | Downloads | Source Repo | slopcheck | Disposition |
|---------|----------|-----|-----------|-------------|-----------|-------------|
| `Microsoft.NETFramework.ReferenceAssemblies` | NuGet | ~4 yrs (v1.0.3, Aug 2022) | 76.5M | Microsoft official (.NET Foundation tooling) | N/A (no pip) — substituted by empirical build success this session | Approved (already used in Phase 1) |
| Hand-rolled `Extensibility` shim | n/a (local source, no registry) | n/a | n/a | n/a — IID sourced from Microsoft Learn's own published `Extensibility.IDTExtensibility2` docs page | N/A | Approved — not an external package; GUID cross-checked against official Microsoft Learn documentation, and against the byte-identical GUID physically found inside the official `Microsoft.VisualStudio.Interop` package this session |
| `Microsoft.Office.Interop.Excel` | NuGet | Uncertain original publish date; current version 16.0.18925.20022 published 2025-10-17 | 30.2M (all versions) | None (repackaging of Microsoft Office binaries; publisher `CamronBute`, not a verified Microsoft account) | N/A (no pip) — content-verified via `strings` (`Application`/`Range`/`Workbook`/`Worksheet`/`XlHAlign` types physically present) | **Flagged — planner must add `checkpoint:human-verify`** (licensing ambiguity — nuspec states "there is no license since it is a repackaging of Office assemblies" — not a functionality risk) |
| `MicrosoftOfficeCore16` | NuGet | Current version 16.0.16626.20000, published 2025-10-17; very low download count (~15K) for this specific package | ~15K (this package); publisher (`CamronBute`) has 45.3M downloads across all 10 published packages | None (same repackaging pattern) | N/A (no pip) — content-verified via `strings` (`IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`/`GetCustomUI`/`InvalidateControl` physically present) | **Flagged — planner must add `checkpoint:human-verify`** (same licensing ambiguity + lower download count than its sibling package) |

**Packages removed due to slopcheck [SLOP] verdict:** none — both flagged packages are real, functional, non-malicious (independently confirmed via binary content inspection), not hallucinated or dangerously new.
**Packages flagged as suspicious [SUS]:** `Microsoft.Office.Interop.Excel`, `MicrosoftOfficeCore16` — both for the same reason (no official Microsoft publisher badge, no clear license), not for signs of being fake or malicious. The planner should insert a `checkpoint:human-verify` step before these are added to the project, explaining the licensing tradeoff plainly (accept the widely-used community convention, or defer this phase's actual package restore until the user can vendor real PIAs from their own Windows+Office machine — see Alternatives Considered).

## Architecture Patterns

### System Architecture Diagram

```text
┌───────────────────────────────────────────────────────────────────────────┐
│  Live Excel Session (human_needed — not executable in this dev sandbox)   │
│                                                                             │
│  Excel loads add-in via HKCU registration (Phase 4) ──► mscoree.dll shim  │
│                            │                                               │
│                            ▼                                               │
│              new FinanceFmtTools.ComAddin.Connect()                        │
│                            │ OnConnection(Application, ...)                │
│                            ▼                                               │
│                    AddInHost.Wire(application)                             │
│           ┌────────────────┼─────────────────────────┐                    │
│           ▼                ▼                         ▼                    │
│   RealExcelGateway   TraceLog (ILog)      RibbonController (Phase 2,       │
│   (wraps Excel.App)                        UNCHANGED — Config +           │
│           │                                 GetCustomUiXml())             │
│           │                                       │                       │
│           │         GetCustomUI(RibbonID) ◄────────┘                      │
│           │                │                                              │
│           │                ▼                                              │
│           │      Excel renders "Finance Fmt" tab (RIB-01)                 │
│           │                │                                              │
│           │     user clicks "Fin 2D" button (onAction="RibbonFin2D")      │
│           │                ▼                                              │
│           │      Connect.RibbonFin2D(control)                             │
│           │                │ (delegates, 1 line, matches modRibbon.bas)   │
│           │                ▼                                              │
│           │      AddInHost.ApplyFormat("FIN_2D")                          │
│           │                │                                              │
│           │   1. gateway.TryGetSelectedRange(out range)                   │
│           │      false → MessageBox.Show(...) + log.Warn (RIB — friendly  │
│           │      message, FMT-06 live completion)                         │
│           │      true  ▼                                                  │
│           │   2. FormatEngine.Apply(range, log, "FIN_2D",                 │
│           │        ribbon.Config.ForceAlign, ribbon.Config.ZeroDash)      │
│           │        (Phase 1/2 code, UNCHANGED)                            │
│           │                │                                              │
│           └───────────────▶│  range.NumberFormat = "..."                  │
│                             │  range.HorizontalAlignment = ...             │
│                             ▼                                              │
│                    Real Excel.Range mutated (visible to user)              │
└───────────────────────────────────────────────────────────────────────────┘
```
A reader can trace the primary use case end-to-end: Excel activates `Connect` via COM → `Connect` delegates to a thin `AddInHost` → the host's real `IExcelGateway` guards the selection (showing a message on failure) → the **unmodified** Phase 1/2 `FormatEngine`/`FormatRegistry` resolve and apply the format → the real `Excel.Range` is mutated. Everything below "Live Excel Session" that isn't reachable from Linux is explicitly out of this research's ability to execute — only to build and reason about.

### Recommended Project Structure
```
src/
├── FinanceFmtTools.sln                          # add the new project via `dotnet sln add`
├── FinanceFmtTools.Engine/                      # Phase 1-2, UNCHANGED
├── FinanceFmtTools.Engine.Tests/                # Phase 1-2, UNCHANGED
└── FinanceFmtTools.ComAddin/                    # NEW — Phase 3, net48 ONLY, not multi-targeted
    ├── FinanceFmtTools.ComAddin.csproj
    ├── Connect.cs                                # COM entry point — IDTExtensibility2 + IRibbonExtensibility + all RibbonXxx callbacks (1 line each)
    ├── Extensibility.cs                          # hand-rolled IDTExtensibility2 / ext_ConnectMode / ext_DisconnectMode shim
    ├── AddInHost.cs                              # composition root — wires RealExcelGateway + TraceLog + (unchanged) RibbonController
    ├── RealExcelGateway.cs                       # IExcelGateway impl over Excel.Application.Selection
    ├── RealRangeHandle.cs                        # IRangeHandle impl over Excel.Range
    ├── TraceLog.cs                                # ILog impl via System.Diagnostics.Trace
    └── Properties/
        └── AssemblyInfo.cs                        # [assembly: ComVisible(false)] — only Connect is marked ComVisible(true)
```

### Pattern 1: Thin COM entry point delegating to a composition-root host
**What:** `Connect` (the `[ComVisible(true)][Guid(...)][ProgId(...)]` class Excel actually instantiates) contains **zero business logic** — every method is a 1-3 line delegation, each wrapped in `try/catch` that swallows and logs, never rethrows.
**When to use:** Always, for the COM entry point specifically — this is the exact pattern the sibling `outlook-classic-delay-send` project uses (`Connect.cs` → `AddInHost`), and it directly reflects `src/modRibbon.bas`'s own documented convention ("cada callback tem exatamente 1 linha de lógica").
**Why it matters:** A COM entry-point constructor or method that throws can cause Excel to silently disable the add-in (Resiliency, a Phase 4 concern) or, worse, prevent it from loading at all with an opaque error. Every method must be defensive.
**Example:**
```csharp
// Source: pattern adapted from /home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend/Connect.cs (read in full this session)
using System;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;

namespace FinanceFmtTools.ComAddin
{
    [ComVisible(true)]
    [Guid("PASTE-A-FRESH-GUID-HERE")]   // generate via `uuidgen` (verified available on this Linux box)
    [ProgId("FinanceFmtTools.Connect")]
    // AutoDispatch (NOT None) — see Pitfall 2. Ribbon callbacks (RibbonFin2D, RibbonChkForceAlign, ...)
    // are invoked by NAME via IDispatch, not via IDTExtensibility2/IRibbonExtensibility — ClassInterfaceType.None
    // would make them invisible to Excel (diagnosed as a live production bug in the sibling project's Gate 6 smoke test).
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        private readonly AddInHost _host = new AddInHost();

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try { _host.Wire(Application); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try { _host.Teardown(); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            try { return _host.Ribbon?.GetCustomUiXml() ?? string.Empty; }
            catch (Exception ex) { TryLog(ex); return string.Empty; }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUi)
        {
            try { _host.CacheRibbon(ribbonUi); }
            catch (Exception ex) { TryLog(ex); }
        }

        // One method per customUI14.xml onAction/getPressed name — see Code Examples for the full 12+4 set.
        public void RibbonFin2D(Office.IRibbonControl control) => _host.ApplyFormat(FormatKeys.Fin2D);
        public bool RibbonGetForceAlign(Office.IRibbonControl control) => _host.Ribbon?.Config.ForceAlign ?? false;
        public void RibbonChkForceAlign(Office.IRibbonControl control, bool pressed) => _host.SetForceAlign(pressed);
        public void RibbonAbout(Office.IRibbonControl control) { try { _host.ShowAbout(); } catch (Exception ex) { TryLog(ex); } }
        public void RibbonFinInfo(Office.IRibbonControl control) { try { _host.OpenDocs(); } catch (Exception ex) { TryLog(ex); } }

        private void TryLog(Exception ex) { try { _host.Log?.Error(ex.ToString()); } catch { /* last line of defense */ } }
    }
}
```

### Pattern 2: Real `IExcelGateway`/`IRangeHandle` — proven via ProjectReference to the actual Phase 1/2 code
**What:** Implement Phase 2's interfaces against `Microsoft.Office.Interop.Excel.Application`/`Range`, exactly the same shape the fakes already proved out in `dotnet test`.
**Empirically verified this session** (`dotnet build -c Release`, 0 Warnings/0 Errors) by copying the **actual, current** `FinanceFmtTools.Engine` project into a scratch solution and building this exact implementation against it via `ProjectReference` — MSBuild correctly resolved the multi-targeted Engine project's `net48` leg automatically.
```csharp
// Source: this research session's empirical spike — compiled successfully against the real
// src/FinanceFmtTools.Engine/Abstractions/IExcelGateway.cs and IRangeHandle.cs (unmodified)
using System.Runtime.InteropServices;
using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinanceFmtTools.ComAddin
{
    public sealed class RealRangeHandle : IRangeHandle
    {
        private readonly Excel.Range _range;
        public RealRangeHandle(Excel.Range range) { _range = range; }

        public string NumberFormat
        {
            get => (string)_range.NumberFormat;
            set => _range.NumberFormat = value;
        }

        public CellAlignment HorizontalAlignment
        {
            get
            {
                var v = (Excel.XlHAlign)_range.HorizontalAlignment;
                if (v == Excel.XlHAlign.xlHAlignRight) return CellAlignment.Right;
                if (v == Excel.XlHAlign.xlHAlignLeft) return CellAlignment.Left;
                return CellAlignment.General;
            }
            set
            {
                _range.HorizontalAlignment =
                    value == CellAlignment.Right ? Excel.XlHAlign.xlHAlignRight :
                    value == CellAlignment.Left ? Excel.XlHAlign.xlHAlignLeft :
                    Excel.XlHAlign.xlHAlignGeneral;
            }
        }

        // External:=True mirrors VBA's rng.Address(External:=True) exactly (src/modFormatEngine.bas logging).
        public string Address => _range.Address[External: true];
    }

    public sealed class RealExcelGateway : IExcelGateway
    {
        private readonly Excel.Application _app;
        public RealExcelGateway(Excel.Application app) { _app = app; }

        public bool TryGetSelectedRange(out IRangeHandle range)
        {
            object sel = _app.Selection;   // typed `object` in the real PIA — Selection can be Range/Chart/Shape/etc.
            if (sel is Excel.Range r)
            {
                range = new RealRangeHandle(r);
                return true;
            }
            range = null;
            if (Marshal.IsComObject(sel)) Marshal.ReleaseComObject(sel);  // release the non-Range object we don't need
            return false;
        }
    }
}
```

### Pattern 3: Don't modify Phase 2's `FormatEngine` to add the friendly message — call it one level lower
**What:** Rather than adding a `MessageBox`-showing callback parameter to `FormatEngine.ApplyToSelection` (which would touch already-tested Phase 2 code), call `IExcelGateway.TryGetSelectedRange` directly in the new project's composition root, show the dialog on `false`, and call `FormatEngine.Apply` (the lower-level method, already public) on `true`.
**When to use:** Whenever a later phase needs to add a side effect (like a UI dialog) that an earlier, already-verified phase's orchestration function deliberately does not have.
**Why it matters:** Keeps Phase 2's `dotnet test`-proven contract (`ApplyToSelection` logs a warning, never shows UI, never throws) completely untouched — zero risk of regressing 39/39 passing tests — while still satisfying this phase's actual requirement (a real dialog).
**Example:**
```csharp
// AddInHost.cs (excerpt)
public void ApplyFormat(string formatKey)
{
    if (!_gateway.TryGetSelectedRange(out IRangeHandle range))
    {
        MessageBox.Show(
            "Selecione um intervalo de células antes de aplicar a formatação.",
            "Finance Fmt Tools", MessageBoxButtons.OK, MessageBoxIcon.Information);
        _log.Warn($"AddInHost.ApplyFormat: seleção inválida para '{formatKey}'.");
        return;
    }
    FormatEngine.Apply(range, _log, formatKey, Ribbon.Config.ForceAlign, Ribbon.Config.ZeroDash);
}
```

### Pattern 4: Cache `IRibbonUI` and call `InvalidateControl` on checkbox toggle, even though VBA never did
**What:** `OnRibbonLoad(IRibbonUI ribbonUi)` caches the handle; the two checkbox `onAction` handlers call `_ribbonUi?.InvalidateControl(control.Id)` after mutating `RibbonSessionConfig`.
**Why it matters:** `src/modRibbon.bas`'s VBA `OnRibbonLoad` also captures `mRibbon` but **never calls it anywhere** — the VBA checkbox handlers just mutate the global and rely on Excel's native checkbox widget to reflect its own click state. Whether that "just works" without an explicit invalidate is genuinely ambiguous from documentation/forum research (mixed answers found — some say a checkbox toggles its own displayed state on click regardless, others say `getPressed` is only re-queried after an explicit `InvalidateControl`). Rather than gamble on undocumented Excel behavior, cache the handle and call `InvalidateControl` defensively — this exactly matches the sibling project's own already-proven-in-a-live-Office-session pattern (`RibbonController.InvalidateUndoButton`/`InvalidateAll`) and costs nothing if it turns out to be unnecessary.
**Verification status:** This specific detail (does the checkbox visually update correctly) is one of the few behaviors this research could not resolve to HIGH confidence without a live session — flag as part of the `human_needed` smoke test checklist regardless of which implementation choice is made.

### Pattern 5: What Phase 3 must expose for Phase 4 (do not implement registration itself)
Phase 4's installer needs these exact, fixed values from this phase's code — document them plainly (e.g. in a short comment block on `Connect.cs`, mirroring the sibling's own `BUILD.md` §5 table) but do **not** write any registry/`regasm` code in this phase:

| Item | Where fixed in Phase 3 | Phase 4 consumes it as |
|---|---|---|
| GUID (CLSID) | `[Guid("...")]` on `Connect` — generate fresh via `uuidgen` (confirmed available on this Linux box) | `InprocServer32\{GUID}` key |
| ProgID | `[ProgId("FinanceFmtTools.Connect")]` | `HKCU\Software\Classes\FinanceFmtTools.Connect` + `...\Outlook`-style `Addins\FinanceFmtTools.Connect` discovery key (Excel equivalent: `HKCU\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect`) |
| AssemblyName | csproj `<AssemblyName>` (recommend `FinanceFmtTools.ComAddin`) | `Assembly=` value in `InprocServer32` (`FinanceFmtTools.ComAddin, Version=1.0.0.0, Culture=neutral, PublicKeyToken=null`) |
| Version | `AssemblyInfo`/csproj `<Version>` (recommend `1.0.0.0`, matching `modConfig.bas`'s `CFG_ADDIN_VERSION`) | Must match the registered `Assembly=` string exactly |
| Distributable file set | Build output of `FinanceFmtTools.ComAddin.csproj` | `FinanceFmtTools.ComAddin.dll` + `FinanceFmtTools.Engine.dll` + `Microsoft.Office.Interop.Excel.dll` + `office.dll` (4 files total — empirically confirmed this session; no `stdole.dll` needed since `customUI14.xml` uses only built-in `imageMso` icons, never a custom `loadImage` callback) |

### Anti-Patterns to Avoid
- **`[ClassInterface(ClassInterfaceType.None)]` on `Connect`:** breaks every named Ribbon callback (see Pitfall 2) — a subtle, hard-to-diagnose failure mode (the tab may still render since `GetCustomUI` is part of a real interface, but every button click silently no-ops).
- **Multi-targeting the new COM project with `net8.0`:** unlike `FinanceFmtTools.Engine`, this project cannot sensibly multi-target — classic Shared-Add-in COM hosting via `mscoree.dll`/regasm has no supported .NET 8/Core equivalent for this architecture. `net48` only.
- **Adding UI-dialog logic inside `FormatEngine`/`RibbonController` (the Engine project):** would reintroduce a COM/WinForms dependency into the project Phase 1/2 explicitly kept COM-free and `dotnet test`-provable on Linux — see Pattern 3.
- **Assuming `Process.Start(url)` (single-string overload) works identically to `Process.Start(new ProcessStartInfo(url){UseShellExecute=true})` across all target frameworks:** on `net48` the single-string overload does work (default `UseShellExecute` is `true` there), but always use the explicit `ProcessStartInfo` form regardless — it removes any ambiguity and matches modern guidance without any downside on `net48` (see Pitfall 4).

## Don't Hand-Roll

| Problem | Don't Build | Use Instead | Why |
|---------|-------------|-------------|-----|
| Excel Object Model types (`Range`, `Workbook`, `XlHAlign`, ...) | A custom late-binding/reflection-based Excel wrapper to avoid the PIA licensing question | `Microsoft.Office.Interop.Excel` NuGet package (flagged for human sign-off, not avoided) | Reflection-based COM automation (no PIA at all, pure `dynamic`/late-binding against `Type.GetTypeFromProgID("Excel.Application")`) is a real alternative some add-ins use specifically to dodge PIA licensing, but it sacrifices compile-time type safety for the entire Excel object model, contradicting this project's whole "testable, safe C# port" premise. Not recommended unless the user explicitly rejects the CamronBute package after reviewing the Package Legitimacy Audit. |
| `IDTExtensibility2` interface | A generic reflection-based "any add-in host" abstraction layer | The narrow, hand-rolled 3-type `Extensibility.cs` shim (or the official `Microsoft.VisualStudio.Interop` package) | The interface is 5 methods and 2 enums, fixed since the 1990s Add-in Designer typelib — there is no "smarter" abstraction to build; matching the documented GUID exactly is the entire job. |
| COM object lifetime management | A generic "COM object pool"/tracking framework | `Marshal.ReleaseComObject` in the one place it's actually needed (the rejected non-Range `Selection` object) + relying on RCW finalization for the single `Range` used per click | This project never enumerates COM collections (unlike the sibling's `Outlook.Folders`/`Outlook.Items` iteration, which genuinely needs rigorous per-object release) — a single `Range` per Ribbon click is exactly the low-risk case VBA itself never manually released either. |

**Key insight:** Everything genuinely novel in this phase (real COM types, `IDTExtensibility2`, Ribbon callback wiring) is dictated by Excel's own, decades-old add-in extensibility model — there is no library that abstracts this away without re-introducing VSTO (explicitly forbidden by CLAUDE.md). The only real "build vs. buy" decision in this phase is the tiny `Extensibility` shim, and even that is a justified, narrow exception (see Standard Stack).

## Common Pitfalls

### Pitfall 1: `<COMReference WrapperTool="tlbimp">` cannot be built via `dotnet build` at all
**What goes wrong:** The "textbook" SDK-style way to reference Office COM interop (`<COMReference Include="...">`, resolved by the registered TypeLib on the build machine) is what the sibling project's own architecture originally specified — and it **failed** there too, with `error MSB4803` ("the task `ResolveComReference` is not supported on the .NET Core / dotnet build MSBuild host").
**Why it happens:** `ResolveComReference` is an MSBuild task that only runs under the full .NET Framework MSBuild (Visual Studio or Build Tools), never under `dotnet build`'s cross-platform MSBuild host — and it additionally requires `TlbImp.exe`/`AxImp.exe` (NETFX SDK Tools) to be physically installed, which they were not even on the sibling project's real Windows dev machine.
**How to avoid:** Never use `<COMReference>` in this project. Reference the interop assemblies directly via `<PackageReference>` (this phase) or `<Reference HintPath="...">` pointing at vendored DLLs (an alternative if the user later builds on a real Windows+Office machine) — both approaches produce `EmbedInteropTypes=false`-equivalent output (interop DLLs copied to `bin/`, distributed alongside the add-in), with no `ResolveComReference` dependency at all. Verified empirically this session: `dotnet build` resolves everything cleanly with plain `<PackageReference>`.
**Warning signs:** `error MSB4803` or `error MSB3086` (`AxImp.exe`/`TlbImp.exe` not found) during `dotnet build`.

### Pitfall 2: `ClassInterfaceType.None` silently breaks every Ribbon callback
**What goes wrong:** The `Connect` class is marked `[ClassInterface(ClassInterfaceType.None)]` (a common "more correct" choice for versionable public COM libraries) — Excel loads the add-in, the Ribbon tab may even render (since `GetCustomUI` is part of a real declared interface, `IRibbonExtensibility`), but **every button click and every checkbox does nothing** — no exception, no log entry, just silent inaction.
**Why it happens:** Methods like `RibbonFin2D`, `RibbonChkForceAlign`, `RibbonGetForceAlign` are **not** members of `IDTExtensibility2` or `Office.IRibbonExtensibility` — they exist only because the Ribbon XML's `onAction`/`getPressed` attributes name them, and Office invokes them by calling `IDispatch.GetIDsOfNames("RibbonFin2D")` on the live COM object at runtime (late-bound, dynamic dispatch — exactly how VBA itself calls procedures by name). `ClassInterfaceType.None` means the class exposes **no** default dispinterface at all, so `GetIDsOfNames` fails to find these methods, and Office has no way to invoke them.
**How to avoid:** Use `[ClassInterface(ClassInterfaceType.AutoDispatch)]` (exposes all public members via `IDispatch` automatically — no separate interface authoring needed for a private, single-purpose add-in class like this one).
**Warning signs:** Tab renders, tooltips work, but clicking any button does nothing and no log entry appears. This is diagnosed, first-hand, in the sibling `outlook-classic-delay-send` project's own `Connect.cs` code comment, referencing "the live smoke test at Gate 6" — the exact same failure class applies here since Excel and Outlook share the identical Ribbon/IDispatch callback mechanism.

### Pitfall 3: No official, lightweight NuGet package exists for the `Extensibility` assembly
**What goes wrong:** Searching NuGet for "Extensibility" + "IDTExtensibility2" returns zero relevant results — it is easy to conclude (wrongly) that vendoring a real `Extensibility.dll` from a Windows GAC is the only option, which is impossible in this Linux dev environment.
**Why it happens:** The classic "Add-in Designer" typelib (`msaddndr.dll`) predates the NuGet ecosystem entirely and was never repackaged as its own standalone NuGet package by Microsoft or any third party under an obvious name.
**How to avoid:** Either hand-declare the tiny interface locally (5 methods, 2 enums, fixed IID `B65AD801-ABAF-11D0-BB8B-00A0C90F2744` — confirmed against Microsoft Learn's own published attribute) or reference the official `Microsoft.VisualStudio.Interop` NuGet package, which was confirmed this session (via direct binary inspection) to physically contain the exact same `Extensibility.IDTExtensibility2`/`ext_ConnectMode`/`ext_DisconnectMode` types with the matching GUID — a genuine, if non-obvious, official source.
**Warning signs:** Spending time trying to vendor `Extensibility.dll` from a Windows machine that doesn't exist in this workflow, or concluding the phase is blocked on missing tooling.

### Pitfall 4: `Process.Start(url)` behavior differs by target framework — always be explicit
**What goes wrong:** Much of the current (post-.NET-Core) guidance for opening a URL says "you must set `UseShellExecute = true` explicitly, since the default changed to `false`." Blindly applying only the newer pattern's *reasoning* (without checking) could lead to skipping the explicit flag on the assumption "this is legacy code, it must already default correctly" — or, conversely, assuming the flag is *mandatory* and adding unnecessary defensive code elsewhere.
**Why it happens:** `ProcessStartInfo.UseShellExecute` defaults to `true` on .NET Framework (including this project's `net48` target) and `false` on .NET Core/5+ — this project's COM add-in is `net48`, so the default is actually already correct without the explicit flag, but relying on an implicit default for something as visible as "does the docs button work" is fragile.
**How to avoid:** Always construct `Process.Start(new ProcessStartInfo(url) { UseShellExecute = true })` explicitly, regardless of which framework's default happens to be correct. Verified to compile this session.
**Warning signs:** A `Win32Exception: The system cannot find the file specified` at runtime when clicking the docs button (this is the classic .NET-Core-era symptom that would NOT actually occur here on `net48`, but is worth guarding against explicitly rather than relying on target-framework trivia).

### Pitfall 5: This new project cannot be multi-targeted or unit-tested in any environment reachable by this repo's current CI/dev setup
**What goes wrong:** Following Phase 1's own precedent (`net48;net8.0` multi-targeting to enable `dotnet test` on Linux) and attempting the same for the new COM project.
**Why it happens:** Phase 1's multi-targeting trick works because `FinanceFmtTools.Engine` has **zero COM/interop dependencies** — its `net8.0` leg is a completely valid, runnable class library. This new project's entire purpose is COM interop (`ComVisible`, `Guid`, `ClassInterface`, real `Microsoft.Office.Interop.Excel` types) — none of this has a meaningful `net8.0` equivalent for a classic Shared COM Add-in hosted via `mscoree.dll`/regasm. There is nothing to multi-target.
**How to avoid:** Target `net48` only. Accept that this project is **build-verified only** in this environment (and likely even on a real Windows dev machine without a live Excel instance) — its actual behavior can only be confirmed by a human running the compiled add-in inside real Excel. This matches the ROADMAP's own phrasing: "verified by manual smoke test, not unit tests alone."
**Warning signs:** Attempting `dotnet test` against this project and getting a platform/runtime error, or spending planning effort trying to design fakes for something that has already been fully faked in Phase 2 (do not re-fake `IExcelGateway` here — Phase 2 already did that; this phase supplies the *real* implementation only).

## Code Examples

### Full Ribbon callback set — matches `src/customUI14.xml` exactly
```csharp
// Source: control ids/onAction/getPressed names read directly from src/customUI14.xml this session.
// Each is a 1-line delegation, per src/modRibbon.bas's own documented convention.
public void RibbonInteger(Office.IRibbonControl control)   => _host.ApplyFormat(FormatKeys.Integer);
public void RibbonFin2D(Office.IRibbonControl control)     => _host.ApplyFormat(FormatKeys.Fin2D);
public void RibbonFin4D(Office.IRibbonControl control)     => _host.ApplyFormat(FormatKeys.Fin4D);
public void RibbonFin8D(Office.IRibbonControl control)     => _host.ApplyFormat(FormatKeys.Fin8D);
public void RibbonPct4D(Office.IRibbonControl control)     => _host.ApplyFormat(FormatKeys.Pct4D);
public void RibbonPct2D(Office.IRibbonControl control)     => _host.ApplyFormat(FormatKeys.Pct2D);
public void RibbonSpreadBps(Office.IRibbonControl control) => _host.ApplyFormat(FormatKeys.SpreadBps);
public void RibbonDateISO(Office.IRibbonControl control)   => _host.ApplyFormat(FormatKeys.DateIso);
public void RibbonDateBR(Office.IRibbonControl control)    => _host.ApplyFormat(FormatKeys.DateBr);
public void RibbonDateBRLong(Office.IRibbonControl control)=> _host.ApplyFormat(FormatKeys.DateBrLong);
public void RibbonText(Office.IRibbonControl control)      => _host.ApplyFormat(FormatKeys.Text);

public void RibbonChkForceAlign(Office.IRibbonControl control, bool pressed) => _host.SetForceAlign(pressed);
public bool RibbonGetForceAlign(Office.IRibbonControl control) => _host.Ribbon?.Config.ForceAlign ?? false;
public void RibbonChkZeroDash(Office.IRibbonControl control, bool pressed)   => _host.SetZeroDash(pressed);
public bool RibbonGetZeroDash(Office.IRibbonControl control)  => _host.Ribbon?.Config.ZeroDash ?? true;

public void RibbonFinInfo(Office.IRibbonControl control) => _host.OpenDocs();   // "Documentação" button
public void RibbonAbout(Office.IRibbonControl control)   => _host.ShowAbout();  // "Sobre" button
```
Note the C# `getPressed` signature is a **return-value function** (`bool GetPressed(IRibbonControl)`), not VBA's `ByRef returnValue As Variant` convention — confirmed via multiple independent sources this session (Microsoft forums, MrExcel community threads, and the sibling project's own `GetUndoEnabled(IRibbonControl) => bool`, its exact `getEnabled` equivalent).

### Hand-rolled `Extensibility.cs` shim (recommended primary approach)
```csharp
// Source: this research session — GUID cross-checked against Microsoft Learn's published
// [System.Runtime.InteropServices.Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")] attribute on
// Extensibility.IDTExtensibility2, and independently confirmed byte-identical inside the official
// Microsoft.VisualStudio.Interop NuGet package via `strings` this session.
using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }
}
```

### GUID generation on Linux (no PowerShell `[guid]::NewGuid()` available)
```bash
# Confirmed available in this environment:
uuidgen
# e.g. d36a425a-c9f7-4eaf-ae07-0e42cdb637bc — paste into [Guid("...")] on Connect
```

### csproj skeleton — empirically verified this session (`dotnet build -c Release`, 0 Warnings/0 Errors)
```xml
<Project Sdk="Microsoft.NET.Sdk">

  <PropertyGroup>
    <TargetFramework>net48</TargetFramework>
    <PlatformTarget>AnyCPU</PlatformTarget>
    <UseWindowsForms>true</UseWindowsForms>
    <LangVersion>9.0</LangVersion>
    <Nullable>disable</Nullable>
    <AssemblyName>FinanceFmtTools.ComAddin</AssemblyName>
  </PropertyGroup>

  <ItemGroup>
    <PackageReference Include="Microsoft.NETFramework.ReferenceAssemblies" Version="1.0.3">
      <PrivateAssets>all</PrivateAssets>
      <IncludeAssets>runtime; build; native; contentfiles; analyzers</IncludeAssets>
    </PackageReference>
    <PackageReference Include="Microsoft.Office.Interop.Excel" Version="16.0.18925.20022" />
    <PackageReference Include="MicrosoftOfficeCore16" Version="16.0.16626.20000" />
  </ItemGroup>

  <ItemGroup>
    <ProjectReference Include="..\FinanceFmtTools.Engine\FinanceFmtTools.Engine.csproj" />
  </ItemGroup>

</Project>
```

## State of the Art

| Old Approach | Current Approach | When Changed | Impact |
|--------------|------------------|--------------|--------|
| VBA `MsgBox`/`ThisWorkbook.FollowHyperlink` | `System.Windows.Forms.MessageBox.Show` / `System.Diagnostics.Process.Start` | This migration (Phase 3) | Same user-facing behavior, no VBA-specific APIs |
| `<COMReference WrapperTool="tlbimp">` resolved by a registered TypeLib | Direct `<PackageReference>`/`<Reference HintPath>` to interop DLLs, `EmbedInteropTypes=false` | Established pattern for `dotnet build`-only workflows (not new in 2026, but still not the "textbook" SDK-style default) | Makes the whole project buildable via `dotnet build` with zero MSBuild-from-VS/Windows dependency — the exact requirement CLAUDE.md locks in |
| VBA `mRibbon` captured but never invoked | `IRibbonUI` cached and `InvalidateControl` called defensively on every checkbox toggle | This migration (Phase 3), following the sibling project's already-proven pattern | Removes reliance on undocumented/ambiguous native-checkbox-self-refresh behavior |

**Deprecated/outdated:** VSTO-style add-ins (not used here per CLAUDE.md's explicit "no VSTO" constraint) and the old `<COMReference>`-first approach for cross-platform CLI builds (superseded by direct interop-package referencing for this exact use case).

## Assumptions Log

| # | Claim | Section | Risk if Wrong |
|---|-------|---------|---------------|
| A1 | `Microsoft.Office.Interop.Excel` (CamronBute) and `MicrosoftOfficeCore16` (CamronBute) are acceptable to depend on despite lacking an official Microsoft publisher badge and explicit license | Standard Stack, Package Legitimacy Audit | Medium — legally low-risk for internal/personal tooling (functionally verified, extremely widely used, Microsoft's own support forum recommends the Excel one), but not zero-risk; if the user rejects this, the fallback is vendoring real PIA DLLs from their own licensed Windows+Office machine (documented as an Alternative), which changes this phase's build story materially and should be re-planned if chosen |
| A2 | Hand-rolling the `Extensibility.IDTExtensibility2` shim (rather than depending on the official `Microsoft.VisualStudio.Interop` package) is architecturally preferable | Standard Stack, Pitfall 3 | Low — both were empirically proven to compile identically in this session; if the hand-rolled shim ever causes a live-Excel COM activation problem that can't otherwise be diagnosed (untestable from this environment), switching to the official NuGet package is a same-session, low-risk fallback |
| A3 | Caching `IRibbonUI` and calling `InvalidateControl` on every checkbox toggle is necessary/beneficial, even though the original VBA never did this | Pattern 4 | Low — worst case this is a harmless no-op extra COM call; but if the *opposite* is true (VBA's "do nothing" approach was actually required and `InvalidateControl` causes a re-entrant `getPressed` issue), this would only surface in the live-Excel smoke test, not before |
| A4 | `ClassInterfaceType.AutoDispatch` (not `None`) is required for this project too, based on a bug diagnosed in a *different* Office host (Outlook, not Excel) | Pitfall 2 | Low — Excel and Outlook share the identical classic Ribbon/`IDispatch` callback mechanism (both are "Shared COM Add-ins" using the same `customUI14.xml` schema and the same late-bound callback dispatch), so this is a well-grounded cross-application inference, but it was not independently re-diagnosed against Excel specifically in this session (no live Excel available) |
| A5 | No `stdole`/`IPictureDisp` reference is needed anywhere in this phase | Pattern 5, "What Phase 3 Must Expose" table | Low — directly verified by reading `src/customUI14.xml` in full: every button uses `imageMso="..."` (built-in Office icons), there is no `loadImage` attribute on `<customUI>` and no custom `image="..."` anywhere, so `OnLoadImage`/`GetImage` callbacks are simply never wired — if a future button adds a custom icon, this assumption would need revisiting |

**If this table is empty:** N/A — five assumptions logged above; all directly investigable by the user reviewing the licensing tradeoff (A1) or by the live-Excel smoke test (A2-A5).

## Open Questions (RESOLVED)

1. **RESOLVED — Does the checkbox visually refresh correctly without `InvalidateControl` (matching VBA's actual, already-shipped behavior), or is `InvalidateControl` genuinely required?**
   - RESOLVED: Implement the defensive version (cache `IRibbonUI` + call `InvalidateControl` on checkbox toggle, Pattern 4) — costs nothing, removes ambiguity. Included explicitly in the human_needed smoke-test checklist regardless.
   - What we know: VBA's `modRibbon.bas` never calls `mRibbon.InvalidateControl` anywhere, and the VBA add-in has shipped as `v1.0.0`/`v1.0.1` — implying the native checkbox toggle either self-refreshes or the discrepancy was never noticed/reported.
   - What's unclear: Community forum sources genuinely disagree on whether Excel's checkbox control needs an explicit invalidate to reflect its own `getPressed` state after its own `onAction` fires.
   - Recommendation: Implement the defensive version (cache + invalidate, Pattern 4) since it costs nothing and removes the ambiguity — but include this specific behavior explicitly in the `human_needed` smoke-test checklist regardless.

2. **RESOLVED — Will the `Microsoft.Office.Interop.Excel`/`MicrosoftOfficeCore16` (CamronBute) NuGet packages actually load correctly at runtime against the user's specific installed Excel version, given they are unofficial repackages?**
   - RESOLVED: Treated as part of the human_needed live-Excel verification scope, not a build-time blocker — this phase's plans proceed with these packages for compilation.
   - What we know: Content verified this session to contain the genuine, complete Excel/Office Core Object Model (matching official Microsoft Learn API documentation for `Application`/`Range`/`IRibbonExtensibility`/etc.), and the packages are versioned to match Office 16.0 (Excel 2016+, matching CLAUDE.md's stated minimum version).
   - What's unclear: Whether there are any subtle binary-compatibility differences between this repackaged PIA and the "real" PIA that would only manifest at runtime against a specific installed Excel build — this cannot be tested without a live Excel session.
   - Recommendation: Treat this as part of the `human_needed` live-Excel verification scope, not a build-time concern.
</output>

## Environment Availability

| Dependency | Required By | Available | Version | Fallback |
|------------|------------|-----------|---------|----------|
| .NET 8 SDK (`dotnet` CLI) | All build/restore work this phase | ✓ | 8.0.422 (already installed at `/home/thomaz/.dotnet`, confirmed working) | — |
| `uuidgen` | Generating the add-in's fixed GUID (no PowerShell `[guid]::NewGuid()` on Linux) | ✓ | (Linux `util-linux` package, confirmed present) | `python3 -c "import uuid; print(uuid.uuid4())"` also confirmed working as a fallback |
| Windows OS | Actually *running*/registering/activating the COM add-in | ✗ | — | None for execution — this is the phase's documented, accepted limitation. Build-only verification via `dotnet build` is the ceiling of what's achievable here; live behavior is `human_needed` on the user's real Windows+Excel machine. |
| Microsoft Excel 2016+ | Live smoke test (RIB-01..04, and the 12-button/Chart-Shape success criteria) | ✗ | — | None — explicitly deferred to the user, per 03-CONTEXT.md's non-discretionary constraint. |
| pip / PyPI tooling (for `slopcheck`) | Package Legitimacy Gate protocol | ✗ | — | Same as Phase 1 — substituted with direct NuGet registry inspection + `.nupkg` content verification via `strings` (see Package Legitimacy Audit) |

**Missing dependencies with no fallback:** None that block *this phase's own deliverable* (real, buildable COM add-in code) — the Windows+Excel gap is explicitly accepted and scoped as `human_needed`, not a blocker to producing the code and proving it compiles.

**Missing dependencies with fallback:**
- `pip`/`slopcheck` — substituted with stronger, more direct verification (registry + binary content inspection + empirical build), consistent with Phase 1's precedent.

## Security Domain

`security_enforcement` is absent from `.planning/config.json`'s workflow block (defaults to enabled). Applying that lens honestly to a phase whose only new attack surface is "opens one hardcoded HTTPS URL to this project's own GitHub repo":

### Applicable ASVS Categories

| ASVS Category | Applies | Standard Control |
|---------------|---------|-----------------|
| V2 Authentication | No | No authentication concept in a local, single-user Excel add-in |
| V3 Session Management | No | "Session" here means an in-memory Excel session, not a security session |
| V4 Access Control | No | No multi-user/permission boundary |
| V5 Input Validation | Marginally yes | `RibbonFinInfo`'s docs URL is a hardcoded constant (`CFG_DOCS_URL` equivalent), never derived from user input or external data — no injection surface. `TryGetSelectedRange`'s type-check (`sel is Excel.Range`) is itself the input-validation control carried over from Phase 2 (FMT-06). |
| V6 Cryptography | No | No cryptographic operations; `Process.Start` opens `https://` (TLS is the browser's/OS's concern, not this add-in's) |

### Known Threat Patterns for this stack

| Pattern | STRIDE | Standard Mitigation |
|---------|--------|---------------------|
| Unhandled exception escaping a COM entry-point method, causing Excel's Resiliency system to silently disable the add-in after a transient error | Denial of Service (of the add-in itself) | Every `Connect` method wraps its body in `try/catch`, logs, never rethrows (Pattern 1) — this is a Phase 4 concern too (`DoNotDisableAddinList`, INST-03), but Phase 3's own defensive coding is the first line of defense |
| A future contributor adds a custom Ribbon icon and wires `loadImage` without realizing `stdole.IPictureDisp` conversion needs `System.Drawing`/`AxHost`-derived plumbing (as the sibling project needed) | Information Disclosure / functional regression (icon fails to load, logged or silent) | Not applicable to this phase's actual scope (see Assumption A5) — flagged here only so a future phase doesn't assume image-loading is free to add without the extra plumbing the sibling project required |

## Sources

### Primary (HIGH confidence — empirical, this session)
- Live `dotnet build -c Release` runs against a `net48` project referencing real `Microsoft.Office.Interop.Excel`/`MicrosoftOfficeCore16` NuGet packages, a hand-rolled `Extensibility.IDTExtensibility2` shim, `ComVisible`/`Guid`/`ProgId`/`ClassInterface` attributes, `MessageBox.Show`, `Process.Start`, **and a live `ProjectReference` to the actual, unmodified `src/FinanceFmtTools.Engine` project** — confirmed 0 Warnings/0 Errors, using the already-installed .NET 8 SDK 8.0.422 at `/home/thomaz/.dotnet`.
- Direct `.nupkg` download + `unzip`/`strings` binary content inspection of `Microsoft.VisualStudio.Interop` (confirms `Extensibility.IDTExtensibility2`/`ext_ConnectMode`/`ext_DisconnectMode` and the exact GUID `B65AD801-ABAF-11D0-BB8B-00A0C90F2744` are physically present), `Microsoft.Office.Interop.Excel` (confirms `Application`/`Range`/`Workbook`/`Worksheet`/`XlHAlign`), and `MicrosoftOfficeCore16` (confirms `IRibbonExtensibility`/`IRibbonUI`/`IRibbonControl`/`GetCustomUI`/`InvalidateControl`).
- `src/customUI14.xml`, `src/modRibbon.bas`, `src/modUtils.bas`, `src/modConfig.bas`, `src/ThisWorkbook.bas` (this repo) — read in full; exact `onAction`/`getPressed` names, VBA `SafeSelection`/`ShowAbout`/`OpenDocsURL`/`LoadConfig`/`SaveConfig` behavior.
- `src/FinanceFmtTools.Engine/**/*.cs` (this repo, Phase 1/2, read in full) — `IExcelGateway`, `IRangeHandle`, `ILog`, `FormatEngine`, `RibbonController`, `RibbonSessionConfig`, `FormatKeys`, `CellAlignment`.
- `/home/thomaz/pessoal/outlook-classic-delay-send/src/UndoSend/Connect.cs`, `Composition/AddInHost.cs`, `Services/RibbonController.cs`, `Services/OutlookGateway.cs`, `Abstractions/IRibbonController.cs`, `Abstractions/IOutlookGateway.cs`, `UndoSend.csproj`, `Properties/AssemblyInfo.cs`, `BUILD.md` — a directly analogous, already-shipped, live-Excel/Outlook-family COM add-in by the same author, explicitly named in this project's own CLAUDE.md as the workflow inspiration. Read in full this session, including the `ClassInterfaceType.AutoDispatch` bug-fix comment (Gate 6 live smoke test finding) and the `<COMReference>`→direct-`<Reference>` interop-sourcing deviation documented in `BUILD.md` §6.
- `.planning/phases/01-format-engine-core/01-RESEARCH.md`, `.planning/phases/02-abstractions-orchestration/02-RESEARCH.md`, `02-01-SUMMARY.md`, `02-02-SUMMARY.md`, `.planning/STATE.md`, `.planning/ROADMAP.md`, `.planning/REQUIREMENTS.md`, `03-CONTEXT.md` (this repo).

### Secondary (MEDIUM confidence — official docs, cross-checked)
- [learn.microsoft.com/.../extensibility.idtextensibility2](https://learn.microsoft.com/en-us/dotnet/api/extensibility.idtextensibility2?view=visualstudiosdk-2022) — exact IID (`B65AD801-ABAF-11D0-BB8B-00A0C90F2744`) and method signatures, cross-verified against the binary content of the official `Microsoft.VisualStudio.Interop` package this session.
- [nuget.org/packages/Microsoft.Office.Interop.Excel](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel), [nuget.org/packages/MicrosoftOfficeCore16](https://www.nuget.org/packages/MicrosoftOfficeCore16), [nuget.org/packages/MicrosoftOfficeCore](https://www.nuget.org/packages/MicrosoftOfficeCore), [nuget.org/packages/Microsoft.VisualStudio.Interop](https://www.nuget.org/packages/Microsoft.VisualStudio.Interop) — version, publisher, download counts, dependencies, nuspec metadata.
- Microsoft Q&A: ["What NuGet package to I use for Microsoft.Office.Interop.Excel?"](https://learn.microsoft.com/en-us/answers/questions/1663943/what-nuget-package-to-i-use-for-microsoft-office-i) — confirms `Microsoft.Office.Interop.Excel` is the community/Microsoft-support-recommended answer despite not being an official Microsoft-published package.
- Community forum discussion (MrExcel, Microsoft Q&A) on C# `getPressed` callback signature (`bool GetPressed(IRibbonControl)`, not VBA's `ByRef` convention) and on whether `InvalidateControl` is required for checkbox self-refresh (genuinely mixed answers — see Open Question 1).

### Tertiary (LOW confidence)
None specific to a load-bearing claim — the one genuinely LOW-confidence item (checkbox self-refresh without `InvalidateControl`) is explicitly flagged as an Open Question rather than stated as fact.

## Metadata

**Confidence breakdown:**
- Standard stack: HIGH for the build/compile question (empirically proven, this session, using the real Phase 1/2 code) — MEDIUM for the two unofficial Office-interop packages' long-term licensing/support status (flagged for explicit human sign-off, not a technical uncertainty)
- Architecture: HIGH — directly derived from this repo's own VBA source, Phase 1/2 code, and a concrete, working, same-author sibling implementation for the identical class of problem (non-VSTO Office Shared COM Add-in)
- Pitfalls: HIGH for the `<COMReference>`/`dotnet build` and `ClassInterfaceType` findings (either empirically reproduced this session or first-hand documented in the sibling project's own code/BUILD.md) — MEDIUM for the `InvalidateControl` checkbox-refresh question (genuinely ambiguous across sources, explicitly flagged as an Open Question rather than asserted)

**Research date:** 2026-07-11
**Valid until:** 30 days for the NuGet package versions (re-verify if planning is delayed past ~2026-08-10); the architectural findings (COM entry point shape, `ClassInterfaceType`, `<COMReference>` build limitation) are stable, decades-old Office extensibility mechanics with no expiry.
