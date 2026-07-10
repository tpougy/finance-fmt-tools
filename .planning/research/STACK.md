# Stack Research

**Domain:** C# COM add-in for Excel (IDTExtensibility2 + Ribbon XML), migrating an existing VBA `.xlam` add-in
**Researched:** 2026-07-10
**Confidence:** HIGH (core runtime/registration facts verified against official Microsoft docs + a proven, already-working sibling implementation for Outlook; a few version-pin choices are MEDIUM pending a live smoke test in Phase work)

## Reference Implementation Used

`~/pessoal/outlook-classic-delay-send` is a **working, shipped** C# 9 / .NET Framework 4.8 COM add-in for Outlook Classic (IDTExtensibility2 + Ribbon XML, HKCU-only registration, `dotnet build`/`dotnet test` from VS Code, 25 xUnit tests, GitHub Actions CI on `windows-latest`). Its `BUILD.md` documents empirically-confirmed build/registration facts (TypeLib GUIDs, PIA versions, registry strings, MSBuild failure modes) for that machine/toolchain. This research verifies which parts of that pattern are Office-host-agnostic (transfer to Excel unchanged) versus Outlook-specific (need re-verification for Excel), and cross-checks everything against current (2025–2026) official Microsoft documentation.

**Bottom line: ~90% of the reference project's stack transfers to Excel unchanged.** Only the Excel object-model PIA is genuinely new; the COM lifecycle, Ribbon plumbing, registration mechanics, build tooling, and test approach are identical patterns.

## Recommended Stack

### Core Technologies

| Technology | Version | Purpose | Why Recommended | Confidence |
|------------|---------|---------|------------------|------------|
| .NET Framework 4.8 (`net48`) | 4.8 (fixed) | Runtime target for the add-in DLL | Office's COM add-in / VSTO Runtime loader (`mscoree.dll`) only hosts the classic CLR. Microsoft has explicitly stated it will **not** update the VSTO/COM add-in platform to .NET Core/.NET 5+, and that .NET Framework 4.8 is the last supported major runtime for this scenario — modern .NET assemblies cannot be loaded in-process by Office's add-in host at all (CLR-in-CLR-host conflicts cause load failures). This is a hard technical ceiling, not a style preference. | HIGH — confirmed via Microsoft Learn / Microsoft Q&A (2025–2026) and matches the reference project's already-proven, shipped result |
| .NET 8 SDK (build tool only) | 8.x (LTS, e.g. `8.0.4xx`) | Compiles/tests/packs the `net48` project via `dotnet build`/`dotnet test`/`dotnet pack` | The SDK is only the *build tool*; it does not change the runtime target. `dotnet build` resolves `net48` correctly as long as the .NET Framework 4.8 targeting pack is present (preinstalled on GitHub Actions `windows-latest` images and on any dev machine with VS Build Tools or Visual Studio previously installed). This exact SDK version is what the reference project's CI (`actions/setup-dotnet@v4`, `dotnet-version: '8.x'`) already validates successfully, and it matches the constraint already recorded in `PROJECT.md`. | HIGH — proven in sibling project's CI; PROJECT.md already commits to this pairing |
| C# 9 (`<LangVersion>9.0</LangVersion>`) | 9.0 | Language version | SDK-style `net48` projects default their C# language version to **7.3**, not "latest," unless `LangVersion` is explicitly pinned (a well-known .NET SDK TFM-to-LangVersion mapping quirk). Pinning to 9 (matching the reference project) unlocks records-adjacent features, target-typed `new`, pattern-matching improvements, etc., while staying comfortably within IL/BCL surface that `net48` supports. There is no technical reason to go higher for this project's scope (no source generators, no `required` members needed). | HIGH |
| `Extensibility` (IDTExtensibility2) | 7.0.3300.0, `PublicKeyToken=b03f5f7f11d50a3a` | COM add-in lifecycle contract (`OnConnection`/`OnDisconnection`/`OnStartupComplete`/`OnBeginShutdown`) | This interface assembly is **host-application-agnostic** — it is the "Microsoft Add-In Designer" type library used by every classic Office COM add-in (Word, Excel, PowerPoint, Outlook alike), not something Outlook-specific. **Reuse the exact `lib/Extensibility.dll` file already committed in `outlook-classic-delay-send`** — no new sourcing needed. | HIGH |
| `office` (Microsoft Office Core) | 15.0.0.0, `PublicKeyToken=71e9bce111e9429c` | `Microsoft.Office.Core` namespace: `IRibbonExtensibility`, `IRibbonUI`, `IRibbonControl` | Also host-agnostic — the Ribbon Fluent UI plumbing (`GetCustomUI`, `OnRibbonLoad`, `onAction`/`getPressed` callback dispatch) is identical infrastructure regardless of which Office app hosts it. **Reuse `lib/OFFICE.DLL` unchanged.** | HIGH |
| `Microsoft.Office.Interop.Excel` | 15.0.0.0, `PublicKeyToken=71e9bce111e9429c` (same versioning convention as the Outlook PIA) | Excel object model: `Application`, `Workbook`, `Worksheet`, `Range`, `XlHAlign`, etc. | This is the one genuinely new dependency vs. the Outlook reference. Excel's Primary Interop Assembly, like Outlook's, has had its PIA major version frozen at **15.0.0.0** since Office 2013 for backward compatibility — it works unmodified across Excel 2016/2019/2021/365 even though the installed Office build number is much higher (the same pattern the reference project's `BUILD.md` empirically confirmed for the Outlook PIA: TypeLib version 9.6 → PIA `Version=15.0.0.0`). The underlying Excel Object Library TypeLib GUID is `{00020813-0000-0000-C000-000000000046}`. | HIGH — TypeLib GUID and PIA versioning convention independently verified via Microsoft Learn/Q&A threads |
| `stdole` | 7.0.3300.0, `PublicKeyToken=b03f5f7f11d50a3a` | OLE Automation primitives (`IPictureDisp`, etc.) — transitive dependency of the other three PIAs | Also host-agnostic. **Reuse `lib/stdole.dll` unchanged.** (Note: for Finance Fmt Tools this dependency is likely dead weight at runtime since the existing Ribbon XML uses only `imageMso` stock icons — see "What NOT to Use" — but it must still be referenced because the Excel/Office PIAs themselves depend on it at the IL level.) | HIGH |
| Ribbon XML `customUI` schema, namespace `http://schemas.microsoft.com/office/2009/07/customui` | 2009/07 (a.k.a. "customUI14") | Declarative Ribbon tab/group/button/checkBox definitions, returned as a string from `GetCustomUI` | This is **exactly the schema already used by the existing VBA add-in's `src/customUI14.xml`** — no migration needed for the XML itself, only its delivery mechanism changes (embedded resource + `GetCustomUI()` instead of a Custom-UI-Editor-injected part inside the `.xlam`). Confirmed as still the current/standard Ribbon schema for Excel 2010 through Microsoft 365 in 2026 — Microsoft has not introduced a newer Ribbon XML version since 2009/07 (the "backstage"/`mso` additions layered on top of it, not a schema replacement). | HIGH |

### Supporting Libraries

| Library | Version | Purpose | When to Use | Confidence |
|---------|---------|---------|-------------|------------|
| `xunit` | 2.9.3 | Unit test framework | Test project (`FinanceFmtTools.Tests`), for the pure format-engine logic and any Ribbon-controller logic that doesn't touch live Excel COM objects | HIGH — current stable 2.x release per nuget.org; same major line the reference project already proved works with `dotnet test` on `net48` |
| `xunit.runner.visualstudio` | 2.8.2 (proven) or 3.1.5 (latest, verified net472/net48-compatible and still runs xUnit v2 assemblies) | VSTest adapter so `dotnet test` can discover/run xUnit tests | Test project only | HIGH for 2.8.2 (exact proven pairing, 25 green tests in reference project); MEDIUM-HIGH for 3.1.5 (independently verified on nuget.org to target `net472`/`net8.0` and to run "xUnit.net 1.9.2 and later," i.e. v2 assemblies, but not smoke-tested in this research pass) |
| `Microsoft.NET.Test.Sdk` | 17.11.1 (proven) or 18.7.0 (latest stable) | MSBuild test SDK/targets required by any `dotnet test` project | Test project only | HIGH for 17.11.1 (proven); MEDIUM for 18.7.0 (current per nuget.org, not independently smoke-tested here) |
| `System.Windows.Forms` (via `<UseWindowsForms>true</UseWindowsForms>`) | ships with `net48` | `MessageBox.Show(...)` for the "Sobre" (About) dialog | Only in the main add-in project, only if you keep a native modal dialog for "Sobre" instead of routing it through Excel's own UI. No NuGet package needed — it's a framework reference toggled by an MSBuild property, exactly as done in the reference project (`UseWindowsForms=true`) for its config/conflict dialogs. | HIGH |
| `System.Runtime.Serialization` / any JSON-config library | — | — | **Do not add.** The Outlook reference needed this for `StorageItemConfigStore`/`ConfigService` (persisted settings). This milestone explicitly removes persistence for "Alinhar à direita" / "Zero contábil" — they're session-only fields, no serialization needed at all. | HIGH (explicit scope decision in `PROJECT.md`) |

### Development Tools

| Tool | Purpose | Notes |
|------|---------|-------|
| VS Code + C# extension (or C# Dev Kit) | Editing/building/debugging | No Visual Studio install required for day-to-day dev once the 4 PIA DLLs are sourced (see below) — matches the explicit constraint in `PROJECT.md` |
| `dotnet` CLI (`build`/`test`/`pack`) | Build, test, package | Confirmed sufficient by the reference project; no MSBuild-from-Build-Tools fallback needed unless a future machine lacks the `net48` targeting pack |
| PowerShell 5.1+ | Installer (`Install-FinanceFmtTools.ps1`), HKCU registration | Already a hard constraint of this repo (existing installer already requires 5.1+); no version change needed |
| GitHub Actions (`windows-latest`) | CI: build, test, package, release on `v*.*.*` tag push | `windows-latest` images ship the `net48` targeting pack preinstalled — the reference project's workflow builds `net48` successfully with just `actions/setup-dotnet@v4` + `dotnet-version: '8.x'`, no extra targeting-pack install step needed |
| `gh` CLI | Manual/AI-assisted release creation per the runbook requirement | Already part of this milestone's scope per `PROJECT.md` |

## Installation

There is **no `dotnet add package` step for the four Office interop assemblies** — they are referenced as local files, not NuGet packages (see "PIA Sourcing Strategy" below). Only the test project pulls real NuGet packages.

```bash
# Scaffold projects (adjust names to match the repo's chosen namespace, e.g. FinanceFmtTools)
dotnet new classlib -n FinanceFmtTools -f net48 -o src/FinanceFmtTools
dotnet new xunit -n FinanceFmtTools.Tests -f net48 -o src/FinanceFmtTools.Tests
dotnet new sln -n FinanceFmtTools -o src
dotnet sln src/FinanceFmtTools.sln add src/FinanceFmtTools/FinanceFmtTools.csproj src/FinanceFmtTools.Tests/FinanceFmtTools.Tests.csproj

# Test-project packages (pin to the proven combo; bump later if desired)
dotnet add src/FinanceFmtTools.Tests package xunit --version 2.9.3
dotnet add src/FinanceFmtTools.Tests package xunit.runner.visualstudio --version 2.8.2
dotnet add src/FinanceFmtTools.Tests package Microsoft.NET.Test.Sdk --version 17.11.1

dotnet build src/FinanceFmtTools.sln -c Release
dotnet test  src/FinanceFmtTools.sln -c Release
```

**Manual `.csproj` edits required after scaffolding** (the `dotnet new classlib` template won't set these — copy the pattern from `outlook-classic-delay-send/src/UndoSend/UndoSend.csproj`):

```xml
<PropertyGroup>
  <TargetFramework>net48</TargetFramework>
  <PlatformTarget>AnyCPU</PlatformTarget>
  <UseWindowsForms>true</UseWindowsForms>
  <EnableComHosting>false</EnableComHosting> <!-- explicitly opt OUT of .NET5+ COM hosting; see "What NOT to Use" -->
  <GenerateAssemblyInfo>false</GenerateAssemblyInfo> <!-- pair with a hand-written AssemblyInfo.cs -->
  <LangVersion>9.0</LangVersion>
  <Deterministic>true</Deterministic>
</PropertyGroup>

<ItemGroup>
  <EmbeddedResource Include="Ribbon\ribbon.xml" />
</ItemGroup>

<ItemGroup>
  <Reference Include="Extensibility, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a">
    <SpecificVersion>false</SpecificVersion>
    <Private>true</Private>
    <HintPath>..\..\lib\Extensibility.dll</HintPath>
  </Reference>
  <Reference Include="office, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c">
    <SpecificVersion>false</SpecificVersion>
    <Private>true</Private>
    <HintPath>..\..\lib\OFFICE.DLL</HintPath>
  </Reference>
  <Reference Include="Microsoft.Office.Interop.Excel, Version=15.0.0.0, Culture=neutral, PublicKeyToken=71e9bce111e9429c">
    <SpecificVersion>false</SpecificVersion>
    <Private>true</Private>
    <HintPath>..\..\lib\Microsoft.Office.Interop.Excel.dll</HintPath>
  </Reference>
  <Reference Include="stdole, Version=7.0.3300.0, Culture=neutral, PublicKeyToken=b03f5f7f11d50a3a">
    <SpecificVersion>false</SpecificVersion>
    <Private>true</Private>
    <HintPath>..\..\lib\stdole.dll</HintPath>
  </Reference>
</ItemGroup>
```

### PIA sourcing strategy (the one genuinely new step vs. the reference project)

1. **Copy unchanged from the sibling repo:** `Extensibility.dll`, `OFFICE.DLL`, `stdole.dll` from `outlook-classic-delay-send/lib/` into this repo's `lib/` folder. These are Office-host-agnostic; there is no Excel-specific variant. (HIGH confidence — verified these are the "Microsoft Add-In Designer" and "Microsoft Office x.0 Object Library" TypeLibs, not Outlook-specific ones.)
2. **Source fresh:** `Microsoft.Office.Interop.Excel.dll`. Two viable paths:
   - **Primary (matches reference project's proven approach):** on a Windows machine with a full/MSI-based Office install (or one where PIAs were registered via the Office "Primary Interop Assemblies" redistributable), copy it from the GAC — typically `C:\Windows\assembly\GAC_MSIL\Microsoft.Office.Interop.Excel\15.0.0.0__71e9bce111e9429c\Microsoft.Office.Interop.Excel.dll` — into `lib/`. Commit it to the repo (exactly as the reference project committed its 4 DLLs), so CI never needs Office installed to build.
   - **Fallback (if no such machine is available):** reference the NuGet package `Microsoft.Office.Interop.Excel` (currently `15.0.4795.1001`-style or newer builds, published by a third party — owner "CamronBute" on nuget.org, description states "generated and signed by Microsoft" but "entirely unsupported and there is no license"). Mechanically this works fine with plain `dotnet restore`/`dotnet build` (unlike `<COMReference WrapperTool="tlbimp">`, a `<PackageReference>` never invokes `TlbImp.exe`/`ResolveComReference`, so it does **not** hit the `MSB4803` failure the reference project's `BUILD.md` documents). The tradeoff is provenance/support, not build mechanics. **Recommendation: try the GAC-sourced path first; fall back to the NuGet package only if no Office-installed Windows machine is available during setup.**
   - Click-to-Run (modern Microsoft 365) installs do **not always** register PIAs in the GAC the way legacy MSI installs did — verify PIA presence before assuming any given dev machine has it; if absent, either enable the Office PIA feature (if using an MSI channel) or fall back to the NuGet path above.

## Alternatives Considered

| Recommended | Alternative | When to Use Alternative |
|-------------|-------------|--------------------------|
| Pure COM add-in (IDTExtensibility2 + Ribbon XML), `net48` | VSTO (Visual Studio Tools for Office) | Only if full Visual Studio + ClickOnce/MSI deployment is acceptable. Explicitly excluded by this project's constraints (VS Code + `dotnet` CLI only, no admin install). |
| Pure COM add-in on `net48` | Office Web Add-ins (JS/TypeScript, `Excel.run` JS API) | Only if you need cross-platform reach (web/Mac/new Outlook) or want to align with Microsoft's long-term push away from COM add-ins. Not relevant here: Excel *desktop* classic client has no announced COM add-in deprecation (unlike Outlook's "New Outlook"), and the whole point of this milestone is 1:1 UX parity with the existing VBA Ribbon, which the JS add-in model can't replicate as directly (different manifest/hosting model, network-connected task-pane paradigm). |
| GAC-sourced PIA DLLs committed to `lib/` | `Microsoft.Office.Interop.Excel` NuGet package (community-repackaged) | If no Windows+Office dev machine is available to extract the PIA from the GAC. Functionally works with `dotnet build`; accept the "unsupported/no license" caveat. |
| xUnit v2 (`2.9.3`) | xUnit v3 (`3.2.2`, Core Framework) | If the team wants to start fresh on the newest xUnit generation. v3 changes the test project shape (own entry point, `xunit.v3` package) — a bigger jump than needed here; the proven v2 pairing already matches the reference project 1:1. |
| `.NET 8 SDK` as build tool | `.NET 9`/`.NET 10 SDK` as build tool | If the team wants to standardize tooling on the newest LTS. Purely a build-tool choice — target framework (`net48`) and runtime behavior are unaffected either way. |

## What NOT to Use

| Avoid | Why | Use Instead |
|-------|-----|--------------|
| VSTO | Requires full Visual Studio + ClickOnce/MSI packaging — contradicts the explicit "VS Code + `dotnet` CLI only, no full Visual Studio" constraint in `PROJECT.md`. Also heavier runtime footprint and a different (managed, not raw COM) lifecycle model than what's needed here. | IDTExtensibility2 + Ribbon XML COM add-in, exactly as prototyped in `outlook-classic-delay-send` |
| .NET 8/9/10 (or any .NET 5+) as the **runtime target** of the add-in itself | Microsoft has explicitly confirmed the VSTO/COM add-in platform will not be updated to modern .NET; Office's `mscoree.dll`-based add-in host cannot load a modern-.NET-hosted CLR in the same process — this is a documented, current (2025–2026) hard limitation, not a temporary gap. | `net48` as the **target**, any modern SDK (8.x/9.x/10.x) purely as the **build tool** |
| `<COMReference Include="..." WrapperTool="tlbimp">` in the `.csproj` | Triggers MSBuild's `ResolveComReference`/`TlbImp.exe`/`AxImp.exe`, which are NETFX SDK Tools not installed by default outside full Visual Studio. Empirically fails with `MSB4803` when building via `dotnet build` (confirmed in the reference project's `BUILD.md`). | Plain `<Reference Include="..." HintPath="...">` pointing at pre-built PIA DLLs committed to `lib/` |
| `regasm.exe` / any HKLM registration | Requires administrator rights — directly violates this project's "no admin" install constraint (same constraint the existing VBA installer already satisfies). | Hand-rolled HKCU registry key creation in the PowerShell installer (`New-Item`/`Set-ItemProperty` under `HKCU:\Software\Classes\...` and `HKCU:\Software\Microsoft\Office\Excel\Addins\...`), mirroring exactly what `regasm` would write but scoped to the current user only |
| `ClassInterfaceType.None` (or the attribute omitted) on the `Connect` class | Office's Ribbon Fluent UI engine invokes `OnRibbonLoad`, every `onAction`/`getPressed`/`getLabel` target, etc. by **name** via late-bound `IDispatch.GetIDsOfNames` — not through a formal typed interface. `ClassInterfaceType.None` (the effective default once you specify `[ClassInterface]` at all, and increasingly the safe default assumed by tooling) hides these public methods from IDispatch, so Excel silently can't find them: buttons render but do nothing, `getPressed` checkboxes never initialize. This is a documented root-cause the reference project hit and fixed. | `[ClassInterface(ClassInterfaceType.AutoDispatch)]` on the `Connect` class, exactly as in the reference project |
| Custom PNG icons + `loadImage`/`OnLoadImage`/`IPictureDisp` plumbing | The existing VBA `src/customUI14.xml` uses **only** `imageMso="..."` stock Office icons (`AccountingNumberFormat`, `PercentStyle`, `ChartLine`, `TableInsertDate`, `DateAndTimeInsert`, `TextBox`, `TablePropertiesDialog`, `Info`) — confirmed by direct inspection. Unlike the Outlook reference (which ships 3 embedded PNGs and a `GetImage`/`OnLoadImage` callback), Finance Fmt Tools needs **no custom icon resources and no `loadImage` callback at all**. | Keep `imageMso` attributes as-is in the ported Ribbon XML; skip `IRibbonController.GetImage`/`Connect.OnLoadImage` entirely — one less moving part than the Outlook example |
| Any config-persistence library (`System.Runtime.Serialization`/JSON, `CustomXMLPart`-equivalent, `StorageItem`-equivalent) | This milestone explicitly and deliberately removes persistence for "Alinhar à direita"/"Zero contábil" (session-only, per `PROJECT.md`). Building a config store here would be scope creep relative to the VBA original *and* the stated goal for this migration. | Two private `bool` fields on the Ribbon controller, defaulted per `PROJECT.md` (`ForceAlign = false`, `ZeroDash = true`), reset every Excel session |
| Reusing the Outlook reference project's exact CLSID/GUID | Each COM add-in needs a globally unique CLSID. Copy-pasting `{E593E0A9-6108-4FAE-A361-797DED79D9D7}` from the Outlook example would collide with that add-in's registration if both are ever installed on the same machine. | Generate a fresh GUID (`[guid]::NewGuid()` in PowerShell, or any GUID generator) and hardcode it consistently across `Connect.cs`'s `[Guid(...)]` attribute and the installer script's `$Guid` constant |

## Stack Patterns by Variant

**If supporting both 32-bit and 64-bit Excel installs (likely needed — Office bitness isn't guaranteed to be 64-bit for every existing user of this add-in):**
- Build stays `AnyCPU` regardless (the PIAs are MSIL, platform-neutral).
- The installer must detect Office's bitness (reuse the reference project's `Test-PeMachine`-style PE-header check on `EXCEL.EXE`, or read `HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration` → `Platform`) and point the `InprocServer32` `(default)` value at the **matching** `mscoree.dll` shim: `%windir%\System32\mscoree.dll` for 64-bit Office, `%windir%\SysWOW64\mscoree.dll` for 32-bit Office running on 64-bit Windows. A bitness mismatch here (the reference installer hardcodes the 64-bit path only, with a comment flagging this as a known simplification) means Excel simply fails to load the add-in with no user-facing error.
- Because HKCU-scoped COM registration is **not** subject to `Wow6432Node` redirection the way HKLM is, the *registry key paths* themselves don't need a 32-bit variant — only the shim DLL path value does.

**If keeping the format-engine core logic COM-free (recommended, and consistent with what's testable without Excel installed):**
- Put format lookup/string-building (`GetFormatDef`, `AccountingFmt`-equivalent) in a plain class referencing **no** interop types at all — just `string`/`decimal`/`bool`. This mirrors how the reference project's `Domain/` layer has zero Outlook references, and it's what your xUnit tests should exercise directly, no fakes needed.
- Only the thin "apply this format string to the live selection" seam needs Excel COM (`Range.NumberFormat`, `Range.HorizontalAlignment`). Wrap that seam behind a small interface (e.g., `ICellFormatTarget`) so a fake/stub can verify "was told to set NumberFormat=X" in xUnit without Excel installed at all — a direct structural match to `IOutlookGateway`/`FakeOutlookGateway` in the reference project.

**If the Ribbon controller itself needs a unit test (recommended, and directly provable without any Excel install):**
- Test `GetCustomUiXml()` returns the embedded resource non-empty and contains every expected button/checkbox `id` — exactly the regression test pattern in `RibbonControllerTests.cs`, which caught a real embedded-resource-name mismatch bug in the reference project. The same resource-name-vs-`RootNamespace` pitfall applies here (`EmbeddedResource` logical name = `RootNamespace` + folder path + filename) — verify it with a test, don't assume it.

## Version Compatibility

| Package A | Compatible With | Notes |
|-----------|------------------|-------|
| `Microsoft.Office.Interop.Excel, Version=15.0.0.0` | Excel 2013, 2016, 2019, 2021, Microsoft 365 (Click-to-Run and MSI) | PIA major version has been frozen at `15.0.0.0` since Office 2013 for backward compatibility; the installed Office build's actual TypeLib version (e.g., `1.8`/`1.9`) is higher but the PIA identity string stays constant — same pattern independently confirmed for the Outlook PIA in the reference project. |
| `Extensibility.dll` / `stdole.dll` (`7.0.3300.0`) | Any Office host application, any Office version 2010+ | Host- and version-agnostic; safe to reuse the exact files already validated in `outlook-classic-delay-send/lib/`. |
| `net48` project | .NET SDK `8.x`/`9.x`/`10.x` (as build tool) | Any of these SDKs builds a `net48` `csproj` correctly as long as the .NET Framework 4.8 targeting pack is present — true by default on GitHub Actions `windows-latest` and on any machine that ever had VS/VS Build Tools installed. |
| `xunit 2.9.x` | `xunit.runner.visualstudio` `2.8.x` (proven) or `3.1.x` (verified compatible) | `xunit.runner.visualstudio 3.1.5` is documented on nuget.org as running "xUnit.net 1.9.2 and later" (i.e., v2 assemblies) and targets `net472`/`net8.0` — safe to run v2 tests on `net48`. |
| Ribbon XML `customUI` (`2009/07` namespace) | Excel 2010 through Microsoft 365 (2026) | No newer Ribbon XML schema version has shipped since 2009/07; it remains current. |

## Sources

- `~/pessoal/outlook-classic-delay-send` (`BUILD.md`, `src/UndoSend/Connect.cs`, `src/UndoSend/UndoSend.csproj`, `src/UndoSend/Composition/AddInHost.cs`, `src/UndoSend/Abstractions/IOutlookGateway.cs`, `src/UndoSend/Abstractions/IRibbonController.cs`, `src/UndoSend.Tests/*`, `scripts/install.ps1`, `.github/workflows/release.yml`) — HIGH confidence, proven working implementation, directly inspected
- `/home/thomaz/pessoal/finance-fmt-tools/src/customUI14.xml` — HIGH confidence, directly inspected; confirms `imageMso`-only icons and custom (non-`idMso`) tab placement
- [Microsoft Learn — VSTO Add-ins can't be created with .NET / .NET Framework 4.8 as final runtime](https://learn.microsoft.com/en-us/answers/questions/1282120/long-term-vsto-addins-support-roadmap-for-ms-outlo) — HIGH, confirms modern .NET cannot host Office COM/VSTO add-ins
- [Microsoft Support — Excel COM add-ins and Automation add-ins](https://support.microsoft.com/en-us/topic/excel-com-add-ins-and-automation-add-ins-91f5ff06-0c9c-b98e-06e9-3657964eec72) — MEDIUM-HIGH, confirms `HKEY_CURRENT_USER\Software\Microsoft\Office\Excel\Addins` registration path pattern
- [Microsoft Q&A — Identify COM add-ins / New Outlook vs. Excel deprecation distinction](https://learn.microsoft.com/en-us/microsoft-365-apps/outlook/get-started/state-of-com-add-ins) — MEDIUM-HIGH, confirms Excel desktop has no announced COM add-in deprecation timeline (unlike Outlook's "New Outlook")
- [NuGet Gallery — Microsoft.Office.Interop.Excel](https://www.nuget.org/packages/Microsoft.Office.Interop.Excel) — HIGH for version/publisher facts (community-owned, "unsupported/no license" disclaimer verified directly on the package page)
- [NuGet Gallery — xunit](https://www.nuget.org/packages/xunit), [xunit.runner.visualstudio](https://www.nuget.org/packages/xunit.runner.visualstudio), [Microsoft.NET.Test.Sdk](https://www.nuget.org/packages/microsoft.net.test.sdk) — HIGH, official package pages, current versions confirmed
- WebSearch: Excel Object Library TypeLib GUID `{00020813-0000-0000-C000-000000000046}` — MEDIUM-HIGH (multiple independent Microsoft Q&A threads agree; not a primary Microsoft doc page)
- WebSearch: `.NET 10` current LTS / SDK version as of mid-2026 — MEDIUM (used only to confirm .NET 8 remains a valid, supported build-tool choice, not to change the recommendation)

---
*Stack research for: C# COM add-in for Excel (VBA → C# migration)*
*Researched: 2026-07-10*
