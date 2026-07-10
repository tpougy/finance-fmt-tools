# Project Research Summary

**Project:** Finance Fmt Tools — VBA → C# COM Add-in Migration
**Domain:** C# COM add-in for Excel (IDTExtensibility2 + Ribbon XML), migrating a shipped VBA `.xlam` add-in
**Researched:** 2026-07-10
**Confidence:** HIGH

## Executive Summary

This is a technology port, not a greenfield build: an already-shipped VBA Excel Ribbon add-in ("Finance Fmt") must be re-implemented as a pure C# COM add-in (`IDTExtensibility2` + `IRibbonExtensibility`, no VSTO), targeting `.NET Framework 4.8` (built with the .NET 8 SDK), with byte-for-byte behavioral parity for 12 Ribbon buttons and 2 session-only checkboxes. The project has an exceptionally strong reference point: `~/pessoal/outlook-classic-delay-send`, a working, shipped sibling C# COM add-in for Outlook that already solved ~90% of this exact problem (COM lifecycle, Ribbon plumbing, HKCU-only no-admin registration, `dotnet`-CLI-only build/test, GitHub Actions CI). The recommended approach is to reuse that project's architecture (`Abstractions/Domain/Services/Composition` layering), its host-agnostic PIA DLLs (`Extensibility.dll`, `OFFICE.DLL`, `stdole.dll`) unchanged, and only newly source `Microsoft.Office.Interop.Excel.dll`.

The core technical risk is not "can this be built" (all four research files converge on a proven, low-ambiguity pattern) but "will the port silently regress behavior that VBA's dynamic/forgiving runtime masked." Four risk classes dominate: (1) transcription risk in `AccountingFmt`'s exact 16-case format-string output (quote-escaping, padding tokens), (2) calling-convention mistranslation of VBA's `Sub`/`ByRef` Ribbon `getPressed` idiom into C#'s function-return idiom, (3) COM-specific failure modes that have no VBA equivalent at all — `ClassInterfaceType.None` silently hiding callbacks, unhandled exceptions triggering Excel's Resiliency auto-disable, and HKCU registry-view (WOW64) bitness redirection breaking 32-bit Excel installs — and (4) COM lifetime/threading discipline (RCW leaks, STA violations, stale cached `Range`/`Selection` references) that C#'s ordinary idioms actively encourage getting wrong. All four are well-understood and preventable per the research; several were live-diagnosed and fixed in the sibling Outlook project, giving high confidence in the mitigations.

Recommended approach: build the pure, COM-free `Domain` layer (format engine) first and lock down parity with xUnit before any COM code exists; introduce the `IExcelGateway` abstraction seam before writing the real COM implementation; treat Ribbon/COM wiring and installer/registration as separate phases each requiring a live Excel smoke test (not just unit tests) as their definition of done, since the most dangerous pitfalls (silent callback failures, bitness mismatches, Resiliency auto-disable) are invisible to abstraction-layer tests by design.

## Key Findings

### Recommended Stack

Target `.NET Framework 4.8` (`net48`) as a hard technical ceiling — Microsoft has confirmed the classic Office COM/VSTO add-in host cannot load modern .NET 5+ assemblies, so this is not a preference but a platform constraint. Build/test tooling is the .NET 8 SDK (`dotnet build`/`test`/`pack`) — the SDK only compiles the `net48` target, it does not change the runtime. C# language version must be explicitly pinned to 9.0 (SDK-style `net48` projects default to C# 7.3 otherwise). Ribbon XML uses the same `customUI` 2009/07 schema Excel has used since 2010 — no schema migration needed, only its delivery mechanism changes (embedded resource + `GetCustomUI()` instead of an OOXML part auto-loaded from the `.xlam`).

**Core technologies:**
- `.NET Framework 4.8` (`net48`) — the only runtime Office's COM add-in host (`mscoree.dll`) can load; confirmed via Microsoft Learn as a hard ceiling, not a style choice
- `.NET 8 SDK` (build tool only) — compiles/tests/packs the `net48` project; proven by the sibling project's CI
- `Extensibility.dll` / `OFFICE.DLL` / `stdole.dll` — host-agnostic COM interface assemblies; reuse verbatim from the sibling Outlook project's `lib/` folder, no Excel-specific variant exists
- `Microsoft.Office.Interop.Excel.dll` (v15.0.0.0) — the one genuinely new PIA; frozen at major version 15.0.0.0 since Office 2013 for backward compatibility across 2016/2019/2021/365; source from a real Office GAC install (preferred, matches sibling's proven approach) or an official-looking NuGet package (fallback)
- `xUnit 2.9.x` + `Microsoft.NET.Test.Sdk` — proven test stack from the sibling project, runs `net48` via `dotnet test` with no Excel/Office installed

### Expected Features

This is a parity port with no staged feature rollout — "table stakes" means "must reproduce byte-for-byte," and the deliverable is complete only when full parity is achieved. There is no competitor-feature landscape to weigh; the governing document is a VBA→C# Parity Risk Matrix instead.

**Must have (table stakes, full parity required):**
- `AccountingFmt`-equivalent pure function — 16 exhaustive cases (`decimals in {0,2,4,8} x forceAlign x zeroDash`), highest transcription risk in the whole port, exact strings documented in FEATURES.md
- All 12 Ribbon buttons (Fin 0D/2D/4D/8D, Pct 0.0000%/0.00%, Spread bps, Date ISO/BR/BR-Long, Text, Integer) wired via `onAction`
- 2 checkboxes ("Alinhar à direita", "Zero contábil") wired via `onAction`/`getPressed`, now session-only (no persistence)
- Selection-type guard (`SafeSelection()` equivalent) before any format application
- About dialog + docs-link button

**Deliberately simplified vs. VBA (already decided, not open for debate):**
- Checkbox state is session-only: `ForceAlign` defaults `false`, `ZeroDash` defaults `true`, reset every Excel session, no persistence layer at all

**Explicitly out of scope (do not add during this milestone):**
- Any new formats/buttons beyond the existing 12
- Re-adding persistence (registry, JSON, `CustomXMLPart`-equivalent) for the two checkboxes
- VSTO/ClickOnce/MSI installer path
- Localization beyond pt-BR
- `NumberFormatLocal` (always use invariant `NumberFormat`)

### Architecture Approach

Directly reuse the sibling Outlook project's layered pattern: `Connect` (thin COM entry point) -> `AddInHost` (composition/wiring) -> `RibbonController`/`FormatEngine` (Services, orchestration against interfaces) -> `Domain` (pure logic, zero COM references) with `Abstractions` (`IExcelGateway`/`IRangeHandle`) as the single seam between testable logic and real Excel COM objects. This structure is what makes `dotnet test` runnable on a CI agent with no Excel installed — the explicit milestone requirement.

**Major components:**
1. `Domain/` (`FormatKeys`, `FormatDefinition`, `FormatRegistry`, `AccountingFormatBuilder`, `AppConfig`) — pure C#, zero Excel/COM reference, 1:1 port of `modFormatEngine.bas`; fully unit-testable, buildable before any COM interop code exists
2. `Abstractions/` (`IExcelGateway`, `IRangeHandle`, `IRibbonController`, `ILog`) — interfaces only; the fake-able seam for tests
3. `Services/` (`FormatEngine`, `RibbonController` orchestration + `ExcelGateway` the only class touching real `Excel.Application`/`Excel.Range`) — orchestration is testable with fakes, the gateway itself is verified only by live smoke test
4. `Composition/AddInHost` — manual DI/wiring from the `object Application` param of `OnConnection`; not unit tested, kept intentionally small
5. `Connect` (COM entry point) — `[ComVisible][Guid][ProgId][ClassInterface(AutoDispatch)]`, implements `IDTExtensibility2` + `IRibbonExtensibility`; every method is a one-line delegation wrapped in try/catch, never contains business logic

### Critical Pitfalls

1. **COM registration bitness mismatch (WOW64 registry redirection)** — `HKCU\Software\Classes\CLSID` is officially WOW64-redirected; writing registration from 64-bit PowerShell makes it invisible to 32-bit Excel. Avoid by detecting Excel's bitness at install time and writing registry keys via `reg.exe /reg:32` or `/reg:64` explicitly (the sibling project only warns for Outlook's x64-only baseline — this project cannot assume a single bitness since "Excel 2016+" doesn't pin one).
2. **`ClassInterfaceType.None` silently hides Ribbon callbacks** — Office resolves `onAction`/`getPressed`/`onLoad` by name via late-bound `IDispatch`; `None` hides all public methods from late binding — buttons render but do nothing, no exception anywhere. Always use `[ClassInterface(ClassInterfaceType.AutoDispatch)]`. This is not theoretical — the sibling project live-diagnosed and fixed this exact bug.
3. **Unhandled exceptions trigger Excel's Resiliency auto-disable** — any uncaught exception from a Ribbon callback or `IDTExtensibility2` method causes Excel to silently disable the add-in after repeated failures, with no obvious diagnostic. Wrap every Office-reachable method in try/catch + log + safe default, and proactively write the `DoNotDisableAddinList` registry key at install time.
4. **`getPressed` calling-convention mistranslation** — VBA's Ribbon `getPressed` idiom is a `Sub` with `ByRef returnValue As Variant`; a literal port to a `void`/`out` C# method compiles but Excel's Ribbon engine won't bind it. The correct C# shape is a function returning `bool` directly.
5. **RCW/COM lifetime and stale-selection caching** — chaining property access, `foreach` over Excel collections, and caching `Workbook`/`Range`/`Selection` fields across callback invocations are all C#-idiomatic habits that leak COM references or apply formats to the wrong (non-active) workbook. Resolve `Application.Selection` fresh on every button click, release every `Range` in a `finally`, never `foreach` over a raw COM collection.

## Implications for Roadmap

Based on combined research (Architecture's "Suggested Build Order," Features' prioritization matrix, and Pitfalls' phase mapping all converge on the same dependency-driven order), suggested phase structure:

### Phase 1: Format Engine Core (Domain, COM-free)
**Rationale:** Zero dependency on Excel/COM; the highest-value, lowest-risk-to-test slice, and the exact place the highest transcription risk (`AccountingFmt`'s 16 cases) must be locked down before anything else is built on top of it. Can be fully validated with `dotnet test` with no Windows/Excel machine required.
**Delivers:** `FormatKeys`, `FormatDefinition`, `FormatRegistry`, `AccountingFormatBuilder`, `AppConfig` — pure C# port of `modFormatEngine.bas`, plus the `FinanceFmtTools.Tests` project scaffolded alongside it.
**Addresses:** All static/computed format-string table stakes (Fin family, Pct, Spread, Date ISO/BR/BR-Long, Text/Integer).
**Avoids:** Business logic referencing Excel types directly — enforce via CI running `dotnet test` on a runner with no Office installed.

### Phase 2: Abstractions + Orchestration (Services, fakeable)
**Rationale:** `IExcelGateway`/`IRangeHandle` must exist as interfaces before any real COM implementation is written, so `FormatEngine`/`RibbonController` orchestration logic (guard clauses, screen-updating suspension, `getPressed` state reads) can be fully unit-tested against a `FakeExcelGateway` — directly mirroring the sibling project's `FakeOutlookGateway` pattern.
**Delivers:** `Abstractions/` interfaces, `Services/FormatEngine.cs`, `Services/RibbonController.cs`, and their xUnit test coverage using fakes.
**Uses:** xUnit + fake-gateway pattern; the single-gateway-interface architecture pattern.
**Implements:** The `IExcelGateway`/`FormatEngine`/`RibbonController` components — the layer that makes the milestone's testability requirement achievable without touching real Excel.

### Phase 3: COM Entry Point + Real Excel Integration
**Rationale:** This is the first point where real Excel COM types (`Microsoft.Office.Interop.Excel`) enter the build, and where the pitfalls with no VBA equivalent become live risks. Must be its own phase with a live Excel smoke test as its definition of done — unit tests alone cannot catch `ClassInterfaceType.None` or Resiliency auto-disable failures.
**Delivers:** Vendored/sourced `Microsoft.Office.Interop.Excel.dll` + reused `Extensibility.dll`/`OFFICE.DLL`/`stdole.dll`; `Services/ExcelGateway.cs` (real COM implementation); `Connect.cs` (COM entry point with fresh GUID, `AutoDispatch`, try/catch on every method); `Ribbon/ribbon.xml` (embedded resource, ported from `customUI14.xml`); `Composition/AddInHost.cs`.
**Avoids:** `ClassInterfaceType.None` hiding callbacks, unhandled exceptions/Resiliency auto-disable, RCW leaks, STA thread violations, caching selection across calls.

### Phase 4: Installation & Registration
**Rationale:** Registration is architecturally separate from the add-in's own code and carries its own distinct, non-obvious risk class (WOW64 bitness redirection) that has no equivalent in the sibling Outlook project's proven pattern (which only warns for non-x64). Needs explicit design, not a copy-paste of the sibling's installer.
**Delivers:** `install.ps1`/`uninstall.ps1` writing HKCU-only registry keys (COM class + Office discovery key `LoadBehavior=3` + Resiliency `DoNotDisableAddinList`), with bitness-aware registry-view handling.
**Avoids:** Bitness mismatch, reaching for `regasm`/admin-requiring registration.

### Phase 5: CI/CD Pipeline & Release Runbook
**Rationale:** Depends on everything above existing and buildable; validates that the vendored-PIA approach from Phase 3 actually builds headless on a CI runner with no Office installed — the concrete proof of the milestone's CI requirement.
**Delivers:** GitHub Actions workflow (`windows-latest`, build/test/package/release on `v*.*.*` tag push), release `.zip` containing the add-in DLL + all four vendored interop DLLs + installer scripts, `gh` CLI runbook for manual/AI-assisted releases, final validation that `archive/vba-legacy` cleanly holds the legacy code.
**Avoids:** Interop assembly resolution failures on `windows-latest`.

### Phase Ordering Rationale

- Dependency-driven: `Domain` has zero Excel/COM dependency and must be correct before anything is built on top of it (the architecture research's own "Suggested Build Order" explicitly recommends this sequencing).
- Testability-driven: phases 1-2 are fully verifiable via `dotnet test` alone; phases 3-5 require a real Windows+Excel environment and live smoke testing — grouping this way lets CI-only validation happen as early and cheaply as possible, deferring the expensive/manual verification surface to the smallest possible slice.
- Pitfall-driven: the most dangerous, VBA-has-no-equivalent pitfalls (silent callback failures, Resiliency auto-disable, bitness mismatch) cluster in phases 3-4, which is exactly where the pitfalls research's own phase mapping places their prevention/verification.

### Research Flags

Needs deeper research during planning:
- **Phase 3 (COM Entry Point + Real Excel Integration):** PIA sourcing strategy is only MEDIUM confidence (GAC-extraction vs. official-looking NuGet package for `Microsoft.Office.Interop.Excel` — unlike Outlook, Excel's PIA does appear to have an official NuGet listing worth a quick spike); `getPressed`/`InvalidateControl` runtime behavior is community-sourced (MEDIUM), not independently reverified against this specific toolchain.
- **Phase 4 (Installation & Registration):** 32-bit Excel bitness handling is explicitly not solved by the sibling reference project (it only warns) — this is new ground requiring either a real implementation or an explicit single-bitness scope decision, and should get a research pass if the team wants to support both bitnesses live.

Phases with standard, well-documented patterns (research-phase likely unnecessary):
- **Phase 1 (Format Engine Core):** Direct, exhaustively-documented port with exact source strings already extracted during feature research; no ambiguity.
- **Phase 2 (Abstractions + Orchestration):** 1:1 structural copy of the sibling project's already-proven `IOutlookGateway`/`FakeOutlookGateway`/`FormatEngine`-equivalent pattern.
- **Phase 5 (CI/CD Pipeline):** Sibling project's GitHub Actions workflow is a working, directly-copyable template; only the vendored DLL set changes.

## Confidence Assessment

| Area | Confidence | Notes |
|------|------------|-------|
| Stack | HIGH | Core runtime/registration facts verified against official Microsoft docs and a working, shipped sibling implementation; a few version-pin choices (PIA sourcing path) are MEDIUM pending a live smoke test |
| Features | HIGH | VBA source read directly (`archive/vba-legacy`), exact format strings and callback signatures extracted firsthand; Ribbon/COM interop facts confirmed via official Microsoft Learn docs; a few runtime-behavior details (`InvalidateControl` necessity) are community-sourced (MEDIUM) |
| Architecture | HIGH | Pattern verified against a working sibling implementation solving the identical problem for Outlook, plus official Microsoft docs for the Excel-specific pieces (PIA versioning, Ribbon schema) |
| Pitfalls | MEDIUM-HIGH | Most pitfalls are either grounded in official Microsoft docs (WOW64 redirection table, Resiliency mechanism) or were live-diagnosed in the sibling project (`ClassInterfaceType.None`, interop resolution failure); the 32-bit bitness pitfall is new ground the sibling project never fully solved (flagged explicitly, not just assumed safe) |

**Overall confidence:** HIGH

### Gaps to Address

- **PIA sourcing strategy (Excel-specific):** Whether to vendor `Microsoft.Office.Interop.Excel.dll` from a real Office GAC install (proven, matches sibling's pattern) or use the official-looking NuGet package (materially easier than the Outlook case since one may actually exist and be legitimate) is unresolved — resolve with a quick build spike at the start of Phase 1/3, not assumed.
- **32-bit Excel support:** Not addressed by the sibling reference project at all (it only warns for non-x64 Outlook). Needs an explicit decision during roadmap/requirements: either implement full bitness-aware registry-view handling in the installer, or make a documented, deliberate single-bitness constraint (mirroring how the sibling handled it, but stated explicitly rather than left as a silent gap).
- **Extensibility.dll/OFFICE.DLL/stdole.dll reuse across Office hosts:** Assumed genuinely host-agnostic (verified as the "Microsoft Add-In Designer" and "Microsoft Office Object Library" TypeLibs, not Outlook-specific), but not independently smoke-tested loading inside Excel specifically — worth a fast confirmation early in Phase 3 rather than treating as settled.
- **`getPressed`/`InvalidateControl` runtime behavior:** The existing VBA add-in apparently works without ever calling `InvalidateControl`, but community sources say it's generally required — the working theory (Excel optimistically flips the state of the very control just clicked) should be validated against a live Excel smoke test in Phase 3, not just assumed to carry over.

## Sources

### Primary (HIGH confidence)
- `~/pessoal/outlook-classic-delay-send` (`BUILD.md`, `Connect.cs`, `AddInHost.cs`, `IOutlookGateway.cs`, `OutlookGateway.cs`, `RibbonController.cs`, test suite, `install.ps1`, `release.yml`) — working, shipped sibling implementation solving the identical problem for Outlook; directly inspected
- `/home/thomaz/pessoal/finance-fmt-tools/archive/vba-legacy` (`modFormatEngine.bas`, `modRibbon.bas`, `modConfig.bas`, `modUtils.bas`, `customUI14.xml`) — direct source of the exact format strings, callback signatures, and Ribbon XML being ported
- Microsoft Learn — VSTO/COM add-ins cannot use .NET 5+ as final runtime (https://learn.microsoft.com/en-us/answers/questions/1282120/long-term-vsto-addins-support-roadmap-for-ms-outlo)
- Registry Keys Affected by WOW64 — Microsoft Learn (https://learn.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys)
- Customize the Office Fluent ribbon by using a managed COM add-in — Microsoft Learn (https://learn.microsoft.com/en-us/office/vba/Library-Reference/Concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in)
- Range.NumberFormat / NumberFormatLocal Property — Microsoft Learn (https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.numberformat?view=excel-pia)
- Support for keeping add-ins enabled (Resiliency) — Microsoft Learn (https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled)

### Secondary (MEDIUM confidence)
- NuGet Gallery — Microsoft.Office.Interop.Excel (https://www.nuget.org/packages/Microsoft.Office.Interop.Excel) — version/publisher facts, provenance caveat noted
- Custom Ribbon getPressed Issues — Microsoft Q&A (https://learn.microsoft.com/en-us/answers/questions/823104/custom-ribbon-getpressed-issues) and MrExcel Message Board thread on getPressed/getEnabled — `InvalidateControl` requirement discussion
- Add-in Express — How to properly release Excel COM objects (https://www.add-in-express.com/creating-addins-blog/release-excel-com-objects/) — RCW release discipline, "two dots" rule, `foreach` enumerator leak

### Tertiary (LOW confidence)
- WebSearch on Excel Object Library TypeLib GUID (`{00020813-0000-0000-C000-000000000046}`) — multiple independent Q&A threads agree, not a primary Microsoft doc page
- GitHub Community discussions on `windows-latest`/`windows-2022` runner image .NET Framework targeting-pack availability changes over time — motivates pinning a specific image tag as a stopgap

---
*Research completed: 2026-07-10*
*Ready for roadmap: yes*
