# Codebase Structure

**Analysis Date:** 2026-07-10

## Directory Layout

```
finance-fmt-tools/
├── src/                          # All add-in source (VBA modules + Ribbon XML)
│   ├── ThisWorkbook.bas          # Workbook-level event handlers (BeforeClose fallback save)
│   ├── modConfig.bas             # Constants: identity, logging flags, format-key strings, persisted config globals
│   ├── modFormatEngine.bas       # Core engine: FormatDef type, ApplyFormat, GetFormatDef, AccountingFmt
│   ├── modRibbon.bas             # Ribbon callbacks (1 line each), OnRibbonLoad
│   ├── modUtils.bas              # Log, HandleError, SafeSelection, ShowAbout, CustomXMLPart persistence
│   └── customUI14.xml            # Ribbon XML — tab/groups/buttons/checkboxes definition
├── Install-FinanceFmtTools.ps1   # PowerShell installer/updater (downloads + registers .xlam via COM)
├── Install-FinanceFmtTools.bat   # One-line .bat wrapper that shells out to the PS1 via irm|iex
├── README.md                     # Full documentation: format catalog, format-string tables, architecture, install instructions
├── LICENSE                       # Project license
└── .planning/
    └── codebase/                 # GSD-generated codebase-map documents (this file's directory)
```

There is no `dist/`, `build/`, `test/`, `node_modules/`, or config-file clutter — this is a minimal 6-source-file VBA project. The compiled `.xlam` binary is never committed to the repo; it is built manually and published as a GitHub Release asset (see `INTEGRATIONS.md`).

## Directory Purposes

**`src/`:**
- Purpose: Canonical, human-readable source of truth for the Excel add-in. Each `.bas` file is a plain-text export of a VBA standard module (or, for `ThisWorkbook.bas`, the special `ThisWorkbook` document module) that must be manually imported into the VBA Editor of an `.xlam` project to produce the shippable binary. `customUI14.xml` must be injected into the `.xlam` package's Custom UI part using an external tool (Custom UI Editor for Microsoft Office), since Excel cannot edit Ribbon XML natively.
- Contains: 5 `.bas` files, 1 `.xml` file — no subfolders.
- Key files: `modFormatEngine.bas` (core logic), `customUI14.xml` (UI contract), `modConfig.bas` (all constants).

**`.planning/`:**
- Purpose: GSD workflow artifacts (codebase maps, phase plans) — not part of the shipped add-in.
- Contains: `codebase/` subdirectory holding STACK.md, INTEGRATIONS.md, ARCHITECTURE.md, STRUCTURE.md (this document set).
- Generated: Yes, by `/gsd-map-codebase`.
- Committed: Yes (per repo convention; verify with the user's `.gitignore` if this changes).

## Key File Locations

**Entry Points:**
- `src/modRibbon.bas` (`OnRibbonLoad`): Ribbon load-time initialization, called automatically by Excel.
- `src/customUI14.xml`: declares `onLoad="OnRibbonLoad"` at the `<customUI>` root and every button/checkbox `onAction`/`getPressed` binding.
- `Install-FinanceFmtTools.ps1` (bottom of file, `Write-Host`/`try` block starting near the end): script entry point for end-user installation.

**Configuration:**
- `src/modConfig.bas`: all add-in constants (identity, logging toggles, `FMT_*` format keys) and the two live persisted preference globals (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`).
- Persisted user config lives inside the `.xlam` binary itself (`CustomXMLPart`, namespace `urn:finance-fmt-tools`), managed by `src/modUtils.bas` (`LoadConfig`/`SaveConfig`) — there is no external config file.
- `Install-FinanceFmtTools.ps1:34-44` (`$CFG` hashtable): installer-only configuration (GitHub owner/repo, filenames, paths).

**Core Logic:**
- `src/modFormatEngine.bas`: format registry (`GetFormatDef`) and application (`ApplyFormat`/`ApplyFormatToSelection`); the accounting format-string builder (`AccountingFmt`).
- `src/modUtils.bas`: cross-cutting utilities (logging, error handling, safe selection, XML persistence).

**Testing:**
- None present. No test files, test framework, or test runner exist in this repository.

## Naming Conventions

**Files:**
- VBA modules: `mod<PascalCaseName>.bas` (e.g., `modConfig.bas`, `modFormatEngine.bas`, `modRibbon.bas`, `modUtils.bas`) — the `mod` prefix signals a standard VBA module (as opposed to a class module or document module).
- Document modules: named after the special Excel object they represent, no prefix (`ThisWorkbook.bas`).
- Ribbon XML: fixed filename `customUI14.xml`, mandated by the Office Custom UI schema (the `14` suffix refers to the Office 2010+ ribbon schema version, still current usage for `.xlam` ribbon customization).
- Installer scripts: `Install-<ProductName>.ps1` / `.bat`, PascalCase with hyphen before the verb, consistent with PowerShell verb-noun naming conventions applied to script filenames.

**Directories:**
- Lowercase, singular, purpose-named: `src/` for all shippable source.

**Code-level (VBA) naming, observed throughout `src/`:**
- Constants: `SCREAMING_SNAKE_CASE` with a category prefix — `CFG_*` for configuration (`CFG_ADDIN_NAME`, `CFG_FORCE_ALIGN`), `FMT_*` for format keys (`FMT_FIN_2D`, `FMT_DATE_ISO`).
- Public Subs/Functions: `PascalCase`, verb-first (`ApplyFormat`, `GetFormatDef`, `SafeSelection`, `LoadConfig`, `SaveConfig`, `HandleError`).
- Ribbon callback Subs: `Ribbon<Action>` prefix consistently (`RibbonFin2D`, `RibbonChkForceAlign`, `RibbonGetForceAlign`, `RibbonFinInfo`).
- Private helpers: `PascalCase`, no special prefix, declared `Private` and scoped to their module (`AccountingFmt` in `modFormatEngine.bas`; `ReadBoolNode`/`WriteBoolNode`/`FindOrCreateXMLPart`/`FindChildNode`/`LogToSheet` in `modUtils.bas`).
- Local error-context constant: every procedure with error handling declares `Const PROC As String = "<ProcName>"` at the top, used consistently in `Log`/`HandleError` calls to identify the source procedure in log output.
- Ribbon XML ids: `<control-type prefix><PascalCaseName>` — `btnFin8D`, `chkForceAlign`, `grpNumeric`, `tabFinanceFmt`.

## Where to Add New Code

**New Format (e.g., a new numeric/date/percent format):**
1. Add a `FMT_<NAME>` constant to `src/modConfig.bas` (follow the existing `FMT_*` block, `src/modConfig.bas:19-29`).
2. Add a corresponding `Case FMT_<NAME>` block inside `GetFormatDef` in `src/modFormatEngine.bas` (`src/modFormatEngine.bas:88-167`), populating `f.key`, `f.DisplayName`, `f.NumberFmt`, `f.Category` (and `f.Alignment` if not general).
3. Add a `<button>` (or `<checkBox>`) element to the appropriate `<group>` in `src/customUI14.xml`, with a unique `onAction` name.
4. Add the matching one-line `Public Sub Ribbon<Name>(control As IRibbonControl)` to `src/modRibbon.bas`, calling `ApplyFormatToSelection FMT_<NAME>`.
- This is the explicitly documented extension pattern in `README.md:237-241` — no other file needs modification.

**New persisted preference (checkbox-driven global setting):**
- Add the `Public` boolean to `src/modConfig.bas` near `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`.
- Add read/write calls in `LoadConfig`/`SaveConfig` in `src/modUtils.bas:116-166`, following the `ReadBoolNode`/`WriteBoolNode` pattern (defaults must be specified for backward compatibility with `.xlam` files saved before the new node existed).
- Add a `<checkBox>` in `src/customUI14.xml` plus paired `RibbonChk<Name>`/`RibbonGet<Name>` callbacks in `src/modRibbon.bas`, mirroring `RibbonChkZeroDash`/`RibbonGetZeroDash`.

**Utilities:**
- Shared, cross-module helpers (logging, error handling, selection safety, XML persistence) belong in `src/modUtils.bas`. Do not duplicate logging or selection-guard logic in other modules — always call `Log`/`SafeSelection`/`HandleError` from there.

**Installer changes:**
- All installer logic lives in the single script `Install-FinanceFmtTools.ps1`; the `$CFG` hashtable at the top (`Install-FinanceFmtTools.ps1:34-44`) is the place to update repo/owner/filename if the project is renamed or moved. `Install-FinanceFmtTools.bat` is a thin wrapper only — update its embedded URL if the script filename or repo path changes (note: as of this scan, `Install-FinanceFmtTools.bat:2` still references the old script name `Install-RBRFinanceTools.ps1`, which appears to be a stale reference — verify before relying on the `.bat` launcher).

## Special Directories

**`.git/`:**
- Purpose: standard git metadata.
- Generated: Yes.
- Committed: N/A (not tracked as a repo file).

**`.planning/`:**
- Purpose: GSD workflow state and generated codebase documentation.
- Generated: Yes (by GSD commands).
- Committed: Yes (per current repo state).

No `node_modules/`, `dist/`, `build/`, `vendor/`, or similar generated/dependency directories exist — consistent with this being a dependency-free VBA project.

---

*Structure analysis: 2026-07-10*
