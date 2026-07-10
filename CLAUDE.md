<!-- GSD:project-start source:PROJECT.md -->
## Project

**Finance Fmt Tools**

Add-in do Excel "Finance Fmt" que adiciona uma aba na Ribbon com atalhos de formatação (contábil, percentual, data, texto) para uso em planilhas financeiras. Hoje é implementado em VBA (`.xlam`), distribuído via GitHub Releases com um instalador PowerShell. Este milestone migra a implementação para C# (COM add-in), preservando a experiência da Ribbon para o usuário final, com um fluxo de desenvolvimento e release moderno inspirado no projeto irmão `outlook-classic-delay-send`.

**Core Value:** Aplicar formatos financeiros/contábeis padronizados a células do Excel com um clique — agora sobre uma base de código C# testável, com dev/build/release 100% via terminal (VS Code + dotnet CLI), sem depender de Visual Studio completo.

### Constraints

- **Plataforma**: Windows + Excel 2016+ — Why: manter compatibilidade com a base de usuários existente do add-in VBA
- **Tooling**: Desenvolvimento via VS Code + dotnet CLI, sem depender de Visual Studio completo — Why: pedido explícito do usuário, replicando o fluxo do `outlook-classic-delay-send`
- **Runtime**: .NET Framework 4.8, buildável com .NET 8 SDK — Why: COM interop com Excel exige .NET Framework clássico; o SDK moderno é só a ferramenta de build
- **Instalação**: Registro do add-in só em HKCU, sem exigir admin — Why: mesma UX do instalador atual e do projeto de inspiração
- **Compatibilidade de UX**: Mesmos botões/atalhos visíveis na Ribbon — Why: usuários já acostumados com o add-in atual não devem perceber regressão
<!-- GSD:project-end -->

<!-- GSD:stack-start source:codebase/STACK.md -->
## Technology Stack

## Languages
- VBA (Visual Basic for Applications) - Excel add-in logic, exported as `.bas` modules in `src/` (`src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`, `src/ThisWorkbook.bas`)
- Ribbon XML (Office Fluent UI schema `http://schemas.microsoft.com/office/2009/07/customui`) - `src/customUI14.xml` defines the "Finance Fmt" ribbon tab
- PowerShell 5.1+ - Installer/updater script, `Install-FinanceFmtTools.ps1`
- Batch (`.bat`) - One-line launcher that shells out to PowerShell, `Install-FinanceFmtTools.bat`
- Markdown - Documentation, `README.md`
## Runtime
- Microsoft Excel for Windows, Office 2016 or later (stated in `README.md:5` and `README.md:176`)
- VBA runs inside the Excel host process; no standalone interpreter or Node/Python runtime involved
- PowerShell 5.1 (`#Requires -Version 5.1` in `Install-FinanceFmtTools.ps1:1`) for the installer only, executed on the end user's Windows machine, not part of the add-in runtime
- None. This is not an npm/pip/cargo project. There is no `package.json`, `requirements.txt`, or lockfile of any kind.
- Dependency delivery is a single compiled binary: `FinanceFmtTools.xlam` (Excel Add-In binary), built manually from the `.bas`/`.xml` sources in `src/` and distributed via GitHub Releases (not committed to this repo — no `.xlam` file is tracked in git; see `git log`, releases `v1.0.0` and `v1.0.1`).
## Frameworks
- Office Fluent Ribbon (`customUI14`) - `src/customUI14.xml` - declarative XML defining tab/groups/buttons/checkboxes, wired to VBA callbacks via `onAction`/`getPressed` attributes
- Excel Object Model (`Range`, `Worksheet`, `Workbook`, `CustomXMLPart`, `IRibbonUI`) - the entire runtime API surface; no external framework/library is referenced
- None detected. No test framework, no test files (`.bas` test modules), no CI test runner found in the repo.
- None automated. There is no build script, bundler, or VBA compiler config in the repo. The `.bas`/`.xml` files in `src/` are source-of-truth text exports; the actual `.xlam` binary must be assembled manually in Excel (Import modules via VBA Editor, paste Ribbon XML via a tool like Custom UI Editor, then Save As `.xlam`) and uploaded as a GitHub Release asset named `FinanceFmtTools.xlam`.
## Key Dependencies
- Excel `CustomXMLPart` API - `src/modUtils.bas:116-238` - used for persisting user preferences (`ForceAlign`, `ZeroDash`) inside the `.xlam` file itself, under a custom namespace `urn:finance-fmt-tools`. No external database or file needed.
- Excel `IRibbonUI` - `src/modRibbon.bas:10,15-19` - reference captured in `OnRibbonLoad` and used to callback-drive checkbox state (`getPressed`)
- GitHub Releases - distribution channel for the compiled `.xlam` binary; `Install-FinanceFmtTools.ps1` downloads `https://github.com/tpougy/finance-fmt-tools/releases/latest/download/FinanceFmtTools.xlam`
- GitHub REST API (`api.github.com/repos/tpougy/finance-fmt-tools/releases/latest`) - used by the installer (`Get-LatestReleaseTag`, `Install-FinanceFmtTools.ps1:97-109`) to determine version tag for display only (not used to gate download URL, which always points at "latest")
## Configuration
- No `.env` files or environment-variable-based config are used by the add-in itself.
- All add-in configuration constants live in `src/modConfig.bas` (identity strings, logging flags, format-key constants) — plain public `Const`/`Public` VBA declarations, not externalized.
- Runtime user preferences (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) are persisted inside the `.xlam` binary via `CustomXMLPart` (see `src/modUtils.bas:116-166`), not via OS registry or `.ini` files.
- The PowerShell installer uses a local hashtable `$CFG` (`Install-FinanceFmtTools.ps1:34-44`) for its own config: GitHub owner/repo, add-in filename, temp path, and the Office add-ins destination folder (`%APPDATA%\Microsoft\AddIns`).
- No build config files present (no `tsconfig.json`, `webpack.config.js`, etc.) — this is expected for a VBA project; there is no compile step beyond Excel's own VBA project compiler.
## Platform Requirements
- Windows with Excel (2016+) and the VBA Editor (Alt+F11) to edit/import `.bas` modules
- A Ribbon XML editor (e.g., "Custom UI Editor for Microsoft Office") to inject `src/customUI14.xml` into the `.xlam` package, since Excel itself cannot edit Ribbon XML directly
- PowerShell 5.1+ available on Windows by default; no additional local tooling required to run/test the installer script
- End-user machine: Windows + Excel 2016 or later
- Deployment target: `.xlam` file copied to `%APPDATA%\Microsoft\AddIns` and registered via `Excel.Application` COM automation (`Install-FinanceFmtTools.ps1:251-328`)
- No server-side component; the add-in is entirely client-side/local to the user's Excel session
<!-- GSD:stack-end -->

<!-- GSD:conventions-start source:CONVENTIONS.md -->
## Conventions

Conventions not yet established. Will populate as patterns emerge during development.
<!-- GSD:conventions-end -->

<!-- GSD:architecture-start source:ARCHITECTURE.md -->
## Architecture

## System Overview
```text
```
## Component Responsibilities
| Component | Responsibility | File |
|-----------|----------------|------|
| Ribbon XML | Declares UI (tab, groups, buttons, checkboxes), maps each control to a VBA callback name via `onAction`/`getPressed`, and the load callback via `onLoad="OnRibbonLoad"` | `src/customUI14.xml` |
| Ribbon Callbacks | Thin (1-line) adapters between Ribbon events and the format engine / config persistence; owns the `IRibbonUI` reference for `getPressed` re-invalidation | `src/modRibbon.bas` |
| Format Engine | Core business logic: format registry (`GetFormatDef`), format application (`ApplyFormat`/`ApplyFormatToSelection`), and accounting-format string generation (`AccountingFmt`) | `src/modFormatEngine.bas` |
| Config/Constants | Add-in identity, logging flags, format-key string constants, and the two persisted boolean preferences (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) | `src/modConfig.bas` |
| Utils | Cross-cutting concerns: logging (`Log`/`LogToSheet`), centralized error handling (`HandleError`), safe selection access (`SafeSelection`), About dialog (`ShowAbout`), docs link (`OpenDocsURL`), and `CustomXMLPart` persistence (`LoadConfig`/`SaveConfig` + private XML node helpers) | `src/modUtils.bas` |
| Workbook Events | Lifecycle hook that force-saves the `.xlam` on close as a persistence fallback | `src/ThisWorkbook.bas` |
| Installer | Out-of-band PowerShell script that downloads the compiled `.xlam` from GitHub Releases and registers it in Excel via COM automation | `Install-FinanceFmtTools.ps1` |
## Pattern Overview
- Single entry-point pattern: all format application flows through `ApplyFormat` / `ApplyFormatToSelection` in `src/modFormatEngine.bas` — no ribbon callback touches `Range.NumberFormat` directly.
- Registry pattern via `Select Case` in `GetFormatDef` (`src/modFormatEngine.bas:81-170`) — adding a format means adding one `Case` block plus one constant in `modConfig.bas`, no changes elsewhere (this is explicitly documented as the extension point in `README.md:237-241`).
- State persisted in the artifact itself (`CustomXMLPart` inside the `.xlam`), not in any external store — the add-in is fully self-contained and portable.
- Defensive wrapper pattern: all `Selection` access goes through `SafeSelection()` (`src/modUtils.bas:74-89`), and all format string generation for the "Fin" family goes through `AccountingFmt()` (`src/modFormatEngine.bas:188-222`), avoiding duplicated derivation logic.
## Layers
- Purpose: Define the "Finance Fmt" tab, its 5 groups (Numérico, Percentual, Data, Texto, Info) and all buttons/checkboxes with tooltips.
- Location: `src/customUI14.xml`
- Contains: XML only — no logic, only `onAction`/`getPressed` bindings and `imageMso` icon references.
- Depends on: nothing (pure declaration); Excel loads it from the `.xlam`'s Custom UI part.
- Used by: Excel Fluent Ribbon renderer; invokes `src/modRibbon.bas` callbacks by name.
- Purpose: Translate Ribbon events into engine/config calls; hold the `IRibbonUI` handle needed for checkbox `getPressed` state.
- Location: `src/modRibbon.bas`
- Contains: One `Public Sub` per button/checkbox, each exactly one line of logic (documented convention, `src/modRibbon.bas:5-7`).
- Depends on: `modFormatEngine` (`ApplyFormatToSelection`), `modConfig` (`FMT_*` constants, `CFG_*` state), `modUtils` (`LoadConfig`, `SaveConfig`, `Log`, `ShowAbout`, `OpenDocsURL`).
- Used by: Ribbon XML (`onAction`/`getPressed`/`onLoad` attributes).
- Purpose: Own the mapping from format key → `FormatDef` (display name, number format string, category, alignment) and apply it to a `Range`.
- Location: `src/modFormatEngine.bas`
- Contains: `FormatDef` user-defined type, `ApplyFormat`, `ApplyFormatToSelection`, `GetFormatDef`, and the private `AccountingFmt` helper that builds the 3-section (`positive;negative;zero`) accounting number format string, respecting `CFG_FORCE_ALIGN` and `CFG_ZERO_DASH`.
- Depends on: `modConfig` (format-key constants, `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`), `modUtils` (`Log`, `HandleError`, `SafeSelection`).
- Used by: `modRibbon.bas` callbacks.
- Purpose: Single source of truth for add-in identity strings, logging toggles, and format-key string constants (avoids magic strings across modules); also holds the two live config values loaded at ribbon startup.
- Location: `src/modConfig.bas`
- Contains: `Public Const` declarations and two `Public` mutable booleans (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) set by `LoadConfig` in `modUtils.bas`.
- Depends on: nothing.
- Used by: every other module.
- Purpose: Logging, centralized error handling, safe selection access, About dialog, docs link, and `CustomXMLPart`-based configuration persistence.
- Location: `src/modUtils.bas`
- Contains: `Log`, `LogToSheet`, `HandleError`, `SafeSelection`, `ShowAbout`, `LoadConfig`, `SaveConfig`, and private XML helpers (`ReadBoolNode`, `WriteBoolNode`, `FindOrCreateXMLPart`, `FindChildNode`), plus `OpenDocsURL`.
- Depends on: `modConfig` (constants), Excel's `CustomXMLParts`/`CustomXMLPart`/`CustomXMLNode` API.
- Used by: `modRibbon.bas`, `modFormatEngine.bas`, `ThisWorkbook.bas`.
- Purpose: Ensure config changes are not lost if `SaveConfig`'s explicit `ThisWorkbook.Save` was skipped or failed; runs a fallback save whenever the `.xlam` closes with unsaved changes.
- Location: `src/ThisWorkbook.bas`
- Depends on: `modUtils` (`Log`, `HandleError`).
- Used by: Excel's `Workbook_BeforeClose` event (automatic, not called by other modules).
## Data Flow
### Primary Request Path (apply a format to selected cells)
### Configuration Change / Persistence Flow
- All mutable add-in state is two module-level `Public` booleans in `modConfig.bas` (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`), loaded once at ribbon startup (`OnRibbonLoad`) and mutated only by the two checkbox callbacks. No other in-memory state is tracked; formatting is always derived fresh from the current selection and current config values.
## Key Abstractions
- Purpose: Represents one applicable format — key, display name, Excel number-format string, category, and preferred alignment.
- Examples: constructed inline in each `Case` branch of `GetFormatDef`, `src/modFormatEngine.bas:81-170`
- Pattern: value type returned by a lookup function (`GetFormatDef`), consumed uniformly by `ApplyFormat` regardless of category (numeric/percent/date/text).
- Purpose: Avoid magic strings across ribbon callbacks and the format registry.
- Examples: `FMT_FIN_8D`, `FMT_FIN_4D`, `FMT_FIN_2D`, `FMT_PCT_4D`, `FMT_PCT_2D`, `FMT_SPREAD_BPS`, `FMT_DATE_ISO`, `FMT_DATE_BR`, `FMT_DATE_BR_LONG`, `FMT_TEXT`, `FMT_INTEGER` — all declared in `src/modConfig.bas:19-29`
- Pattern: string constant per format, referenced both by `modRibbon.bas` (call site) and `modFormatEngine.bas` (`Select Case` registry).
## Entry Points
- Location: `src/modRibbon.bas:15-19`
- Triggers: Excel, automatically, when the add-in's Ribbon customUI is loaded (declared via `customUI onLoad="OnRibbonLoad"` in `src/customUI14.xml:3`)
- Responsibilities: Capture the `IRibbonUI` reference (needed later for `getPressed` invalidation) and call `LoadConfig` to hydrate `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` from the persisted `CustomXMLPart`.
- Location: `src/modRibbon.bas` (one per button: `RibbonInteger`, `RibbonFin2D`, `RibbonFin4D`, `RibbonFin8D`, `RibbonPct4D`, `RibbonPct2D`, `RibbonSpreadBps`, `RibbonDateISO`, `RibbonDateBR`, `RibbonDateBRLong`, `RibbonText`, `RibbonFinInfo`, `RibbonAbout`)
- Triggers: user clicking the corresponding Ribbon button (`onAction` attribute in `src/customUI14.xml`)
- Responsibilities: forward to `ApplyFormatToSelection` (format buttons) or a utility action (`OpenDocsURL`, `ShowAbout`).
- Location: `src/modRibbon.bas:80-98`
- Triggers: user toggling a Ribbon checkbox
- Responsibilities: update the corresponding `CFG_*` global and persist it via `SaveConfig`; paired `RibbonGetForceAlign`/`RibbonGetZeroDash` supply the checkbox's current state back to the Ribbon via `getPressed`.
- Location: `src/ThisWorkbook.bas:1-17`
- Triggers: Excel's workbook close event, automatically, whenever the `.xlam` is being closed
- Responsibilities: fallback save if there are unsaved changes, so persisted config is never silently lost.
- Location: `Install-FinanceFmtTools.ps1:347-376` (bottom of file, top-level script flow)
- Triggers: user running the script directly or via the one-line `irm ... | iex` command in `README.md:174`
- Responsibilities: check Excel isn't running, download latest `.xlam` from GitHub Releases, install/update it via Excel COM automation, and clean up the temp file.
## Architectural Constraints
- **Threading:** Single-threaded, synchronous — VBA runs on Excel's main UI thread; there is no async/await or background worker concept in this codebase. `Application.ScreenUpdating` is toggled off/on around format application purely for UI-flicker suppression, not concurrency control (`src/modFormatEngine.bas:42,54`).
- **Global state:** Two module-level public globals in `src/modConfig.bas` (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) are shared mutable state read by `modFormatEngine.bas` and written by `modRibbon.bas`/`modUtils.bas`. A private `mRibbon As IRibbonUI` singleton in `src/modRibbon.bas:10` holds the only Ribbon handle for the add-in's lifetime.
- **Circular imports:** None — dependency direction is strictly one-way: `modRibbon` → `modFormatEngine`/`modUtils` → `modConfig`; `ThisWorkbook` → `modUtils`. No module depends back on a caller.
- **No automated build/test pipeline:** Because VBA source lives as `.bas`/`.xml` text exports, there is no way to verify the compiled `.xlam` matches `src/` without manually re-importing into Excel; any structural change to modules or the Ribbon XML must be manually re-integrated into the binary before a release.
## Anti-Patterns
### Direct `Selection` access outside `SafeSelection()`
### Silent Save side-effect inside a config setter
## Error Handling
- Every public procedure declares `Const PROC As String = "<procedure name>"` and uses `On Error GoTo ErrHandler`, ending with `HandleError PROC, Err` in the handler (e.g., `src/modFormatEngine.bas:24-60`, `src/modUtils.bas:116-136`, `140-166`, `220-238`, `263-273`).
- `HandleError(source, e)` (`src/modUtils.bas:59-68`) logs the error via `Log` and additionally shows a `MsgBox` only when compiled with the `DEBUG_MODE` conditional compilation constant — meaning end users normally see no error dialog, just silent logging (a UX tradeoff, not a bug).
- `ApplyFormat` uses a `Cleanup:` label pattern to guarantee `Application.ScreenUpdating` is restored even on error (`src/modFormatEngine.bas:53-59`).
- `SaveConfig`/`Workbook_BeforeClose` both explicitly re-enable `Application.EnableEvents = True` in their error handlers to avoid leaving Excel event handling disabled after a failure (`src/modUtils.bas:161-165`, `src/ThisWorkbook.bas:13-16`).
## Cross-Cutting Concerns
<!-- GSD:architecture-end -->

<!-- GSD:skills-start source:skills/ -->
## Project Skills

No project skills found. Add skills to any of: `.claude/skills/`, `.agents/skills/`, `.cursor/skills/`, `.github/skills/`, or `.codex/skills/` with a `SKILL.md` index file.
<!-- GSD:skills-end -->

<!-- GSD:workflow-start source:GSD defaults -->
## GSD Workflow Enforcement

Before using Edit, Write, or other file-changing tools, start work through a GSD command so planning artifacts and execution context stay in sync.

Use these entry points:
- `/gsd-quick` for small fixes, doc updates, and ad-hoc tasks
- `/gsd-debug` for investigation and bug fixing
- `/gsd-execute-phase` for planned phase work

Do not make direct repo edits outside a GSD workflow unless the user explicitly asks to bypass it.
<!-- GSD:workflow-end -->



<!-- GSD:profile-start -->
## Developer Profile

> Profile not yet configured. Run `/gsd-profile-user` to generate your developer profile.
> This section is managed by `generate-claude-profile` -- do not edit manually.
<!-- GSD:profile-end -->
