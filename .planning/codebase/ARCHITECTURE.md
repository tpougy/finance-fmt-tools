<!-- refreshed: 2026-07-10 -->
# Architecture

**Analysis Date:** 2026-07-10

## System Overview

```text
┌─────────────────────────────────────────────────────────────┐
│                     Ribbon UI (declarative)                  │
│                  `src/customUI14.xml`                        │
│   tab "Finance Fmt" → groups (Numérico/Percentual/Data/       │
│   Texto/Info) → buttons + 2 checkboxes, each wired via        │
│   onAction / getPressed to a VBA callback                     │
└───────────────────────────┬───────────────────────────────────┘
                             │ onAction="RibbonXxx"
                             ▼
┌─────────────────────────────────────────────────────────────┐
│              Ribbon Callback Layer (1 line each)              │
│                 `src/modRibbon.bas`                            │
│   RibbonFin2D / RibbonPct4D / RibbonDateISO / ... →            │
│   ApplyFormatToSelection(FMT_KEY)                              │
│   RibbonChkForceAlign / RibbonChkZeroDash → set config + save  │
└───────────────────────────┬───────────────────────────────────┘
                             │ ApplyFormatToSelection(key)
                             ▼
┌─────────────────────────────────────────────────────────────┐
│                  Format Engine (business logic)                │
│               `src/modFormatEngine.bas`                        │
│   ApplyFormat(rng, key) → GetFormatDef(key) → FormatDef        │
│   AccountingFmt(decimals, zeroDash) builds NumberFormat string  │
└──────────┬───────────────────────────────┬─────────────────────┘
           │ reads CFG_FORCE_ALIGN/         │ writes rng.NumberFormat
           │ CFG_ZERO_DASH                  │ + rng.HorizontalAlignment
           ▼                                ▼
┌───────────────────────────┐   ┌───────────────────────────────┐
│   Config / Constants        │   │   Excel Range (the workbook)   │
│   `src/modConfig.bas`        │   │   user's selected cells        │
│   FMT_* keys, CFG_* state    │   └───────────────────────────────┘
└──────────┬───────────────────┘
           │ LoadConfig()/SaveConfig()
           ▼
┌─────────────────────────────────────────────────────────────┐
│         Persistence + Utilities (`src/modUtils.bas`)           │
│  CustomXMLPart (urn:finance-fmt-tools) inside the .xlam file   │
│  Log() → Immediate window / optional _FTLog hidden sheet       │
│  SafeSelection() → guards all Selection access                 │
│  HandleError() → centralized error logging                     │
└─────────────────────────────────────────────────────────────┘
           ▲
           │ Workbook_BeforeClose → fallback Save
`src/ThisWorkbook.bas`
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

**Overall:** Layered callback-driven plugin architecture typical of Office VBA add-ins — a declarative UI layer (Ribbon XML) invokes thin controller callbacks, which delegate to a single-purpose "engine" module, backed by a constants module and a utilities module. There is no MVC/service/repository split beyond this; the whole system is four VBA modules plus one XML file.

**Key Characteristics:**
- Single entry-point pattern: all format application flows through `ApplyFormat` / `ApplyFormatToSelection` in `src/modFormatEngine.bas` — no ribbon callback touches `Range.NumberFormat` directly.
- Registry pattern via `Select Case` in `GetFormatDef` (`src/modFormatEngine.bas:81-170`) — adding a format means adding one `Case` block plus one constant in `modConfig.bas`, no changes elsewhere (this is explicitly documented as the extension point in `README.md:237-241`).
- State persisted in the artifact itself (`CustomXMLPart` inside the `.xlam`), not in any external store — the add-in is fully self-contained and portable.
- Defensive wrapper pattern: all `Selection` access goes through `SafeSelection()` (`src/modUtils.bas:74-89`), and all format string generation for the "Fin" family goes through `AccountingFmt()` (`src/modFormatEngine.bas:188-222`), avoiding duplicated derivation logic.

## Layers

**Ribbon UI (declarative):**
- Purpose: Define the "Finance Fmt" tab, its 5 groups (Numérico, Percentual, Data, Texto, Info) and all buttons/checkboxes with tooltips.
- Location: `src/customUI14.xml`
- Contains: XML only — no logic, only `onAction`/`getPressed` bindings and `imageMso` icon references.
- Depends on: nothing (pure declaration); Excel loads it from the `.xlam`'s Custom UI part.
- Used by: Excel Fluent Ribbon renderer; invokes `src/modRibbon.bas` callbacks by name.

**Ribbon Callbacks:**
- Purpose: Translate Ribbon events into engine/config calls; hold the `IRibbonUI` handle needed for checkbox `getPressed` state.
- Location: `src/modRibbon.bas`
- Contains: One `Public Sub` per button/checkbox, each exactly one line of logic (documented convention, `src/modRibbon.bas:5-7`).
- Depends on: `modFormatEngine` (`ApplyFormatToSelection`), `modConfig` (`FMT_*` constants, `CFG_*` state), `modUtils` (`LoadConfig`, `SaveConfig`, `Log`, `ShowAbout`, `OpenDocsURL`).
- Used by: Ribbon XML (`onAction`/`getPressed`/`onLoad` attributes).

**Format Engine:**
- Purpose: Own the mapping from format key → `FormatDef` (display name, number format string, category, alignment) and apply it to a `Range`.
- Location: `src/modFormatEngine.bas`
- Contains: `FormatDef` user-defined type, `ApplyFormat`, `ApplyFormatToSelection`, `GetFormatDef`, and the private `AccountingFmt` helper that builds the 3-section (`positive;negative;zero`) accounting number format string, respecting `CFG_FORCE_ALIGN` and `CFG_ZERO_DASH`.
- Depends on: `modConfig` (format-key constants, `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`), `modUtils` (`Log`, `HandleError`, `SafeSelection`).
- Used by: `modRibbon.bas` callbacks.

**Config/Constants:**
- Purpose: Single source of truth for add-in identity strings, logging toggles, and format-key string constants (avoids magic strings across modules); also holds the two live config values loaded at ribbon startup.
- Location: `src/modConfig.bas`
- Contains: `Public Const` declarations and two `Public` mutable booleans (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`) set by `LoadConfig` in `modUtils.bas`.
- Depends on: nothing.
- Used by: every other module.

**Utils (cross-cutting):**
- Purpose: Logging, centralized error handling, safe selection access, About dialog, docs link, and `CustomXMLPart`-based configuration persistence.
- Location: `src/modUtils.bas`
- Contains: `Log`, `LogToSheet`, `HandleError`, `SafeSelection`, `ShowAbout`, `LoadConfig`, `SaveConfig`, and private XML helpers (`ReadBoolNode`, `WriteBoolNode`, `FindOrCreateXMLPart`, `FindChildNode`), plus `OpenDocsURL`.
- Depends on: `modConfig` (constants), Excel's `CustomXMLParts`/`CustomXMLPart`/`CustomXMLNode` API.
- Used by: `modRibbon.bas`, `modFormatEngine.bas`, `ThisWorkbook.bas`.

**Workbook Events:**
- Purpose: Ensure config changes are not lost if `SaveConfig`'s explicit `ThisWorkbook.Save` was skipped or failed; runs a fallback save whenever the `.xlam` closes with unsaved changes.
- Location: `src/ThisWorkbook.bas`
- Depends on: `modUtils` (`Log`, `HandleError`).
- Used by: Excel's `Workbook_BeforeClose` event (automatic, not called by other modules).

## Data Flow

### Primary Request Path (apply a format to selected cells)

1. User clicks a Ribbon button, e.g. "Fin 2D" (`src/customUI14.xml:27-33`, `onAction="RibbonFin2D"`)
2. `RibbonFin2D(control)` fires → calls `ApplyFormatToSelection FMT_FIN_2D` (`src/modRibbon.bas:28-30`)
3. `ApplyFormatToSelection` resolves the active selection via `SafeSelection()` (`src/modUtils.bas:74-89`), which validates `TypeName(Selection) = "Range"` and shows a friendly `MsgBox` otherwise (`src/modFormatEngine.bas:64-77`)
4. `ApplyFormat(rng, "FIN_2D")` looks up the format via `GetFormatDef("FIN_2D")` (`src/modFormatEngine.bas:24-60`, `81-170`)
5. `GetFormatDef` builds a `FormatDef` whose `NumberFmt` is computed by `AccountingFmt(2, applyZeroDash:=CFG_ZERO_DASH)` (`src/modFormatEngine.bas:99-103`, `188-222`), reading the live `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` globals from `modConfig.bas`
6. `ApplyFormat` sets `rng.NumberFormat` and, if the format defines a non-general alignment, `rng.HorizontalAlignment`, wrapped in `Application.ScreenUpdating = False/True` (`src/modFormatEngine.bas:42-54`)
7. Result logged via `Log` (`src/modUtils.bas:17-28`) to the Immediate window (and optionally the hidden `_FTLog` sheet)

### Configuration Change / Persistence Flow

1. User toggles a checkbox, e.g. "Zero contábil" (`src/customUI14.xml:52-57`, `onAction="RibbonChkZeroDash"`)
2. `RibbonChkZeroDash(control, pressed)` sets `CFG_ZERO_DASH = pressed` then calls `SaveConfig` (`src/modRibbon.bas:90-94`)
3. `SaveConfig` (`src/modUtils.bas:140-166`) locates or creates the `CustomXMLPart` (namespace `urn:finance-fmt-tools`) via `FindOrCreateXMLPart`, writes both boolean nodes via `WriteBoolNode`, then explicitly calls `ThisWorkbook.Save` (with `Application.EnableEvents` toggled off/on around it) to persist the change to disk immediately
4. On next Excel session, `OnRibbonLoad(ribbon)` (`src/modRibbon.bas:15-19`) calls `LoadConfig` (`src/modUtils.bas:116-136`), which reads both nodes back via `ReadBoolNode`, defaulting to `True`/`False` respectively if the node or part is missing (forward-compatible with files saved before a given preference existed)
5. Fallback safety net: `Workbook_BeforeClose` (`src/ThisWorkbook.bas:1-17`) force-saves the `.xlam` if `ThisWorkbook.Saved = False` when Excel closes, in case `SaveConfig`'s save was skipped due to an error

**State Management:**
- All mutable add-in state is two module-level `Public` booleans in `modConfig.bas` (`CFG_FORCE_ALIGN`, `CFG_ZERO_DASH`), loaded once at ribbon startup (`OnRibbonLoad`) and mutated only by the two checkbox callbacks. No other in-memory state is tracked; formatting is always derived fresh from the current selection and current config values.

## Key Abstractions

**FormatDef (user-defined type):**
- Purpose: Represents one applicable format — key, display name, Excel number-format string, category, and preferred alignment.
- Examples: constructed inline in each `Case` branch of `GetFormatDef`, `src/modFormatEngine.bas:81-170`
- Pattern: value type returned by a lookup function (`GetFormatDef`), consumed uniformly by `ApplyFormat` regardless of category (numeric/percent/date/text).

**Format-key constants (`FMT_*`):**
- Purpose: Avoid magic strings across ribbon callbacks and the format registry.
- Examples: `FMT_FIN_8D`, `FMT_FIN_4D`, `FMT_FIN_2D`, `FMT_PCT_4D`, `FMT_PCT_2D`, `FMT_SPREAD_BPS`, `FMT_DATE_ISO`, `FMT_DATE_BR`, `FMT_DATE_BR_LONG`, `FMT_TEXT`, `FMT_INTEGER` — all declared in `src/modConfig.bas:19-29`
- Pattern: string constant per format, referenced both by `modRibbon.bas` (call site) and `modFormatEngine.bas` (`Select Case` registry).

## Entry Points

**`OnRibbonLoad(ribbon As IRibbonUI)`:**
- Location: `src/modRibbon.bas:15-19`
- Triggers: Excel, automatically, when the add-in's Ribbon customUI is loaded (declared via `customUI onLoad="OnRibbonLoad"` in `src/customUI14.xml:3`)
- Responsibilities: Capture the `IRibbonUI` reference (needed later for `getPressed` invalidation) and call `LoadConfig` to hydrate `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH` from the persisted `CustomXMLPart`.

**`RibbonXxx(control As IRibbonControl)` callbacks:**
- Location: `src/modRibbon.bas` (one per button: `RibbonInteger`, `RibbonFin2D`, `RibbonFin4D`, `RibbonFin8D`, `RibbonPct4D`, `RibbonPct2D`, `RibbonSpreadBps`, `RibbonDateISO`, `RibbonDateBR`, `RibbonDateBRLong`, `RibbonText`, `RibbonFinInfo`, `RibbonAbout`)
- Triggers: user clicking the corresponding Ribbon button (`onAction` attribute in `src/customUI14.xml`)
- Responsibilities: forward to `ApplyFormatToSelection` (format buttons) or a utility action (`OpenDocsURL`, `ShowAbout`).

**`RibbonChkForceAlign` / `RibbonChkZeroDash` (control, pressed):**
- Location: `src/modRibbon.bas:80-98`
- Triggers: user toggling a Ribbon checkbox
- Responsibilities: update the corresponding `CFG_*` global and persist it via `SaveConfig`; paired `RibbonGetForceAlign`/`RibbonGetZeroDash` supply the checkbox's current state back to the Ribbon via `getPressed`.

**`Workbook_BeforeClose(Cancel As Boolean)`:**
- Location: `src/ThisWorkbook.bas:1-17`
- Triggers: Excel's workbook close event, automatically, whenever the `.xlam` is being closed
- Responsibilities: fallback save if there are unsaved changes, so persisted config is never silently lost.

**`Install-FinanceFmtTools.ps1` (script entry point):**
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

**What happens:** None found in `src/modRibbon.bas` or `src/modFormatEngine.bas` — this is enforced correctly today.
**Why it's wrong:** `Selection` can be a non-Range object (e.g., a Shape or Chart) depending on what's focused in Excel, causing a runtime type-mismatch error if accessed without a guard.
**Do this instead:** Continue routing all selection access through `SafeSelection()` (`src/modUtils.bas:74-89`) as documented in `README.md:239` ("Todo acesso a Selection passa por SafeSelection() em modUtils — nenhum outro módulo toca Selection diretamente"). Any new ribbon callback that needs the active range must call `ApplyFormatToSelection`/`SafeSelection`, never `Selection` directly.

### Silent Save side-effect inside a config setter

**What happens:** `SaveConfig` (`src/modUtils.bas:140-166`) both persists config to the `CustomXMLPart` AND calls `ThisWorkbook.Save` — meaning every single checkbox toggle triggers a full workbook save with `Application.EnableEvents` toggled off/on.
**Why it's wrong:** Coupling a state mutation to a disk I/O side effect makes `SaveConfig` slow/risky to call frequently and makes future callers (e.g., if more preferences are added and toggled together) prone to redundant saves; it also means any error during `.Save` (e.g., read-only file, locked by another process) surfaces as a config-save failure even though the XML part itself was written successfully.
**Do this instead:** If adding new persisted settings, keep writes batched (already done — both `ForceAlign` and `ZeroDash` are written before the single `Save` call) and be aware that any new checkbox-driven setting should follow the same set-then-save pattern rather than saving independently per field.

## Error Handling

**Strategy:** Centralized `On Error GoTo ErrHandler` + `HandleError` pattern used consistently across all `Public Sub`/`Function` entry points that touch Excel API or file I/O.

**Patterns:**
- Every public procedure declares `Const PROC As String = "<procedure name>"` and uses `On Error GoTo ErrHandler`, ending with `HandleError PROC, Err` in the handler (e.g., `src/modFormatEngine.bas:24-60`, `src/modUtils.bas:116-136`, `140-166`, `220-238`, `263-273`).
- `HandleError(source, e)` (`src/modUtils.bas:59-68`) logs the error via `Log` and additionally shows a `MsgBox` only when compiled with the `DEBUG_MODE` conditional compilation constant — meaning end users normally see no error dialog, just silent logging (a UX tradeoff, not a bug).
- `ApplyFormat` uses a `Cleanup:` label pattern to guarantee `Application.ScreenUpdating` is restored even on error (`src/modFormatEngine.bas:53-59`).
- `SaveConfig`/`Workbook_BeforeClose` both explicitly re-enable `Application.EnableEvents = True` in their error handlers to avoid leaving Excel event handling disabled after a failure (`src/modUtils.bas:161-165`, `src/ThisWorkbook.bas:13-16`).

## Cross-Cutting Concerns

**Logging:** `Log(msg)` in `src/modUtils.bas:17-28` — timestamped `Debug.Print` always; optional very-hidden `_FTLog` worksheet when `CFG_LOG_TO_SHEET = True` (`src/modConfig.bas:11-12`). Every meaningful state transition (format applied, config loaded/saved, ribbon events, errors) is logged with a `PROC: message` convention.

**Validation:** Minimal, defensive-only — `SafeSelection()` validates the active selection is a `Range` before use (`src/modUtils.bas:74-89`); `ApplyFormat` checks `rng Is Nothing` and unknown `formatKey` (empty `FormatDef.key`) before proceeding (`src/modFormatEngine.bas:28-40`). No input validation beyond these guards (e.g., no validation of cell contents/types before applying a format).

**Authentication:** Not applicable — the add-in runs entirely within the user's local Excel session with no login concept.

---

*Architecture analysis: 2026-07-10*
