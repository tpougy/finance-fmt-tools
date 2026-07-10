# Technology Stack

**Analysis Date:** 2026-07-10

## Languages

**Primary:**
- VBA (Visual Basic for Applications) - Excel add-in logic, exported as `.bas` modules in `src/` (`src/modConfig.bas`, `src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modUtils.bas`, `src/ThisWorkbook.bas`)
- Ribbon XML (Office Fluent UI schema `http://schemas.microsoft.com/office/2009/07/customui`) - `src/customUI14.xml` defines the "Finance Fmt" ribbon tab

**Secondary:**
- PowerShell 5.1+ - Installer/updater script, `Install-FinanceFmtTools.ps1`
- Batch (`.bat`) - One-line launcher that shells out to PowerShell, `Install-FinanceFmtTools.bat`
- Markdown - Documentation, `README.md`

## Runtime

**Environment:**
- Microsoft Excel for Windows, Office 2016 or later (stated in `README.md:5` and `README.md:176`)
- VBA runs inside the Excel host process; no standalone interpreter or Node/Python runtime involved
- PowerShell 5.1 (`#Requires -Version 5.1` in `Install-FinanceFmtTools.ps1:1`) for the installer only, executed on the end user's Windows machine, not part of the add-in runtime

**Package Manager:**
- None. This is not an npm/pip/cargo project. There is no `package.json`, `requirements.txt`, or lockfile of any kind.
- Dependency delivery is a single compiled binary: `FinanceFmtTools.xlam` (Excel Add-In binary), built manually from the `.bas`/`.xml` sources in `src/` and distributed via GitHub Releases (not committed to this repo — no `.xlam` file is tracked in git; see `git log`, releases `v1.0.0` and `v1.0.1`).

## Frameworks

**Core:**
- Office Fluent Ribbon (`customUI14`) - `src/customUI14.xml` - declarative XML defining tab/groups/buttons/checkboxes, wired to VBA callbacks via `onAction`/`getPressed` attributes
- Excel Object Model (`Range`, `Worksheet`, `Workbook`, `CustomXMLPart`, `IRibbonUI`) - the entire runtime API surface; no external framework/library is referenced

**Testing:**
- None detected. No test framework, no test files (`.bas` test modules), no CI test runner found in the repo.

**Build/Dev:**
- None automated. There is no build script, bundler, or VBA compiler config in the repo. The `.bas`/`.xml` files in `src/` are source-of-truth text exports; the actual `.xlam` binary must be assembled manually in Excel (Import modules via VBA Editor, paste Ribbon XML via a tool like Custom UI Editor, then Save As `.xlam`) and uploaded as a GitHub Release asset named `FinanceFmtTools.xlam`.

## Key Dependencies

**Critical:**
- Excel `CustomXMLPart` API - `src/modUtils.bas:116-238` - used for persisting user preferences (`ForceAlign`, `ZeroDash`) inside the `.xlam` file itself, under a custom namespace `urn:finance-fmt-tools`. No external database or file needed.
- Excel `IRibbonUI` - `src/modRibbon.bas:10,15-19` - reference captured in `OnRibbonLoad` and used to callback-drive checkbox state (`getPressed`)

**Infrastructure:**
- GitHub Releases - distribution channel for the compiled `.xlam` binary; `Install-FinanceFmtTools.ps1` downloads `https://github.com/tpougy/finance-fmt-tools/releases/latest/download/FinanceFmtTools.xlam`
- GitHub REST API (`api.github.com/repos/tpougy/finance-fmt-tools/releases/latest`) - used by the installer (`Get-LatestReleaseTag`, `Install-FinanceFmtTools.ps1:97-109`) to determine version tag for display only (not used to gate download URL, which always points at "latest")

## Configuration

**Environment:**
- No `.env` files or environment-variable-based config are used by the add-in itself.
- All add-in configuration constants live in `src/modConfig.bas` (identity strings, logging flags, format-key constants) — plain public `Const`/`Public` VBA declarations, not externalized.
- Runtime user preferences (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) are persisted inside the `.xlam` binary via `CustomXMLPart` (see `src/modUtils.bas:116-166`), not via OS registry or `.ini` files.
- The PowerShell installer uses a local hashtable `$CFG` (`Install-FinanceFmtTools.ps1:34-44`) for its own config: GitHub owner/repo, add-in filename, temp path, and the Office add-ins destination folder (`%APPDATA%\Microsoft\AddIns`).

**Build:**
- No build config files present (no `tsconfig.json`, `webpack.config.js`, etc.) — this is expected for a VBA project; there is no compile step beyond Excel's own VBA project compiler.

## Platform Requirements

**Development:**
- Windows with Excel (2016+) and the VBA Editor (Alt+F11) to edit/import `.bas` modules
- A Ribbon XML editor (e.g., "Custom UI Editor for Microsoft Office") to inject `src/customUI14.xml` into the `.xlam` package, since Excel itself cannot edit Ribbon XML directly
- PowerShell 5.1+ available on Windows by default; no additional local tooling required to run/test the installer script

**Production:**
- End-user machine: Windows + Excel 2016 or later
- Deployment target: `.xlam` file copied to `%APPDATA%\Microsoft\AddIns` and registered via `Excel.Application` COM automation (`Install-FinanceFmtTools.ps1:251-328`)
- No server-side component; the add-in is entirely client-side/local to the user's Excel session

---

*Stack analysis: 2026-07-10*
