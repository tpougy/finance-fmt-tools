# External Integrations

**Analysis Date:** 2026-07-10

## APIs & External Services

**Distribution:**
- GitHub Releases - `Install-FinanceFmtTools.ps1:41` - hosts the compiled `FinanceFmtTools.xlam` binary, fetched at `https://github.com/tpougy/finance-fmt-tools/releases/latest/download/FinanceFmtTools.xlam`
  - SDK/Client: none — raw HTTP download via `System.Net.HttpWebRequest` (`Install-FinanceFmtTools.ps1:165-224`, function `Get-FileFromWeb`)
  - Auth: none (public repo, public release assets)

- GitHub REST API - `Install-FinanceFmtTools.ps1:97-109` (`Get-LatestReleaseTag`) - queries `https://api.github.com/repos/tpougy/finance-fmt-tools/releases/latest` to determine and display the current release tag
  - SDK/Client: `Invoke-RestMethod` (built-in PowerShell cmdlet) with a custom `User-Agent: FinanceFmtTools-Install` header
  - Auth: none (unauthenticated public API call; subject to GitHub's unauthenticated rate limits)

**Documentation:**
- GitHub repository page - `src/modConfig.bas:7` (`CFG_DOCS_URL = "https://github.com/tpougy/finance-fmt-tools"`) - opened via `ThisWorkbook.FollowHyperlink` in `src/modUtils.bas:263-273` (`OpenDocsURL`) when the user clicks the ribbon "Documentação" button (`RibbonFinInfo` in `src/modRibbon.bas:103-105`)

## Data Storage

**Databases:**
- None. No SQL/NoSQL database, ORM, or DB client of any kind.

**File Storage:**
- Local filesystem only, scoped to the installer:
  - Temp download location: `%TEMP%\FinanceFmtTools.xlam` (`Install-FinanceFmtTools.ps1:42`)
  - Install destination: `%APPDATA%\Microsoft\AddIns\FinanceFmtTools.xlam` (`Install-FinanceFmtTools.ps1:43-44`)
- In-file storage: user preferences are persisted inside the `.xlam` package itself via an Office `CustomXMLPart` (namespace `urn:finance-fmt-tools`), not as a separate file — see `src/modUtils.bas:102-238`.

**Caching:**
- None.

## Authentication & Identity

**Auth Provider:**
- None. The add-in has no login/auth concept — it operates purely within the user's already-open Excel session. GitHub API/download calls are unauthenticated (public resources).

## Monitoring & Observability

**Error Tracking:**
- None (no Sentry/Bugsnag/Application Insights, etc.)

**Logs:**
- Custom lightweight logger in `src/modUtils.bas:17-28` (`Log` sub):
  - Always writes timestamped entries to the VBA Immediate window via `Debug.Print` when `CFG_LOG_ENABLED = True` (`src/modConfig.bas:10`)
  - Optionally (when `CFG_LOG_TO_SHEET = True`, `src/modConfig.bas:11`, disabled by default) appends rows to a very-hidden worksheet named `_FTLog` (`src/modUtils.bas:30-54`, `CFG_LOG_SHEET_NAME` in `src/modConfig.bas:12`)
  - Centralized error handler `HandleError` (`src/modUtils.bas:59-68`) logs `Err.Number`/`Err.Description` and optionally shows a `MsgBox` when compiled with `#If DEBUG_MODE Then` conditional compilation constant
- PowerShell installer logs to the console only, via `Write-Step`/`Write-Ok`/`Write-Warn`/`Write-Fail` helper functions (`Install-FinanceFmtTools.ps1:56-74`); no file-based install log.

## CI/CD & Deployment

**Hosting:**
- N/A (client-side Office add-in; no server hosting)

**CI Pipeline:**
- None detected in the repository (no `.github/workflows/`, no other CI config). Releases (`v1.0.0`, `v1.0.1` observed via `gh release list`) appear to be published manually with the `.xlam` binary attached as an asset.

## Environment Configuration

**Required env vars:**
- None used by the add-in (VBA) or by the installer script. The installer uses `$env:TEMP` and `$env:APPDATA` (standard Windows environment variables, not custom configuration) to compute file paths (`Install-FinanceFmtTools.ps1:42-43`).

**Secrets location:**
- None present or required — no API keys, tokens, or credentials anywhere in the codebase. `.env`/credential files: not present in this repository.

## Webhooks & Callbacks

**Incoming:**
- None. Excel Ribbon `onAction`/`getPressed`/`onLoad` callbacks (`src/customUI14.xml`, wired to `src/modRibbon.bas`) are in-process Office event callbacks, not network webhooks.

**Outgoing:**
- None (no outbound webhook calls; only the plain HTTP GET requests described under APIs & External Services).

---

*Integration audit: 2026-07-10*
