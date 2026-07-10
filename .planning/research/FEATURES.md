# Feature Research

**Domain:** Excel Ribbon number-formatting add-in — VBA → C# COM add-in re-implementation (subsequent milestone, not a new product)
**Researched:** 2026-07-10
**Confidence:** HIGH for VBA source-level facts (read directly from `archive/vba-legacy`); HIGH for Microsoft Learn–sourced COM/Ribbon interop facts; MEDIUM for community-sourced Ribbon-engine runtime behavior (flagged inline)

> Adaptation note: this is a **technology port**, not a greenfield feature landscape. "Table stakes" below means *"must reproduce byte-for-byte / behaviorally identical in C#"*, and "Anti-features" means *"things that look like natural improvements while porting but are explicitly out of scope this milestone."* There is no competitor analysis section — it's replaced with a **VBA→C# Parity Risk Matrix**, which is what actually matters for this migration.

## Feature Landscape

### Table Stakes (Must Preserve Exactly for Parity)

| Feature | Why Expected | Complexity | Notes |
|---------|--------------|------------|-------|
| Fin family buttons (Fin 0D/2D/4D/8D) → 3-section accounting `NumberFormat` string | Core value prop of the whole add-in; users click these dozens of times/day | MEDIUM | Must port `AccountingFmt(decimals, applyZeroDash)` (`archive/vba-legacy:src/modFormatEngine.bas`) as a **pure function** with identical string output for every `(decimals ∈ {0,2,4,8}) × (forceAlign ∈ {T,F}) × (zeroDash ∈ {T,F})` combination — 16 cases total. See exact strings below. |
| Pct family buttons (`0.0000%`, `0.00%`) | Simple, low-risk, but still core | LOW | Static literal strings, no computed logic, no checkbox interaction. Copy verbatim. |
| Spread (bps) button (`#,##0.0" bps"`) | Distinct format family used for capital-markets spreads | LOW | VBA source is `"#,##0.0"" bps"""` (doubled quotes = VBA's escape for a literal `"`). In C# this is `"#,##0.0\" bps\""` or `@"#,##0.0"" bps"""` — a literal quoted suffix `" bps"` inside the format code. Transcription must preserve the embedded quote characters exactly. |
| Date ISO / Date BR / Date BR Long buttons | Users depend on these for exports vs. local display | LOW–MEDIUM | Literal format strings: `yyyy-mm-dd;@`, `[$-pt-BR]dd/mm/yyyy;@`, `[$-pt-BR]dd/mmm/yyyy;@`. The `[$-pt-BR]` is a **locale-ID prefix baked into the format code itself** — it forces Portuguese month abbreviations (`jan/fev/mar…`) via `NumberFormat` (the invariant property) regardless of the Excel UI language. Do not "clean this up" thinking it's redundant with ambient locale — it is the mechanism that makes "Data BR Longa" render in Portuguese even on an English-language Excel. |
| Text / Integer buttons (`@`, integer variant of AccountingFmt with `decimals=0`) | CEP/CNPJ/CETIP codes must not be Excel-interpreted as numbers | LOW | `Integer`/`Fin 0D` shares the `AccountingFmt(0, …)` path — VBA has an explicit special case for `decimals = 0` (no trailing `.`), documented inline in `modFormatEngine.bas`. This special case must be ported, not just the general `decimals > 0` branch, or "Fin 0D" will regress to showing a stray decimal point. |
| "Alinhar à direita" checkbox (forces `* ` fill-alignment char into the format string + `HorizontalAlignment`) | Existing, documented user preference for column alignment | MEDIUM | Now session-only (see Key Decisions in PROJECT.md): in-memory field, defaults to **OFF** every Excel session. No load/save needed at all. |
| "Zero contábil" checkbox (zero renders as a centered dash `-`) | Existing, documented accounting convention | MEDIUM | Now session-only: in-memory field, defaults to **ON** every Excel session. Same simplification as above. |
| Selection-type guard before applying any format (`SafeSelection()` equivalent) | Prevents runtime errors when Selection isn't a cell range (e.g. a Chart or Shape has focus) | MEDIUM | VBA's `TypeName(Selection) <> "Range"` never throws. The C#/COM-interop equivalent (`Application.Selection as Excel.Range`, or an `is` pattern) **can** throw `COMException`/`InvalidCastException` in some edge states — must wrap defensively, not just null-check. See Parity Risk Matrix. |
| Ribbon buttons wired via `onAction` (12 buttons) | All format/utility buttons must keep working | MEDIUM | Delivery mechanism changes structurally: VBA's `customUI14.xml` is a physical OOXML part inside the `.xlam`, auto-loaded by Excel with callbacks resolved against the VBA project. A COM add-in instead implements `IRibbonExtensibility.GetCustomUI(string RibbonID)` returning the same XML (as an embedded resource string), with callbacks resolved against **public C# methods** on the add-in's COM-visible class. Same XML content, different loading/binding path — see Parity Risk Matrix. |
| Checkbox `getPressed` wiring (2 checkboxes) | Ribbon must show the correct current in-memory state | MEDIUM–HIGH | **Calling-convention translation risk, not just a copy-paste port.** VBA implements `getPressed` as `Sub RibbonGetZeroDash(control As IRibbonControl, ByRef returnValue As Variant)` (VBA's idiom for all Ribbon "get*" callbacks — a `Sub` with an out-parameter). The equivalent, correct C# signature is a **function that returns the value directly**: `public bool GetPressedZeroDash(Office.IRibbonControl control)`. A literal VBA→C# transliteration (e.g. a `void` method with an `out`/`ref` parameter) is a plausible but incorrect port and needs to be caught in review. |
| "Sobre" (About) dialog | Existing, low-value but user-visible | LOW | VBA `MsgBox` is inherently modal/owned by the Excel window. `System.Windows.Forms.MessageBox.Show(text)` with no owner can appear unparented/behind Excel. Pass Excel's main window handle (`IWin32Window` wrapper around `Application.Hwnd`) as the owner for correct modality/Z-order parity. |
| "Documentação" link button (`OpenDocsURL`) | Existing, opens GitHub README in default browser | LOW | `Process.Start(url) { UseShellExecute = true }` — default is already `true` on .NET Framework 4.8 (unlike .NET Core/5+ where the default flipped to `false`), so this is low-risk given the project's pinned .NET Framework 4.8 target, but still worth setting explicitly rather than relying on the default. |
| Portuguese-accented UI strings (labels, tooltips, About text) | Product is pt-BR-first; existing users expect the same wording | LOW–MEDIUM | Both the embedded Ribbon XML resource and C# string literals must round-trip UTF-8 correctly (`ã`, `ç`, `é`, `à`). Reading an embedded XML resource via a bare `new StreamReader(stream)` without an explicit encoding relies on BOM/auto-detection — verify the embedded `customUI14.xml` resource keeps its `encoding="UTF-8"` declaration and, ideally, a BOM, or accented characters in labels/supertips can get mangled at runtime. |
| Ribbon `onLoad` → capture `IRibbonUI` handle | Needed today only for future `Invalidate` calls, not for `getPressed` itself | LOW | Port `OnRibbonLoad(IRibbonUI ribbon)` faithfully but note: it no longer needs to call any `LoadConfig` — with persistence removed, `onLoad` becomes purely "stash the ribbon reference," and the in-memory booleans just need correct field initializers (`_forceAlign = false; _zeroDash = true;`) set at add-in construction, independent of `onLoad` timing. |

**Exact `AccountingFmt` output to reproduce (source of truth, `decimals=2` shown, same pattern for 0/4/8):**

```
forceAlign=False, zeroDash=False:  _(#,##0.00_)_-;(#,##0.00)_-;_(#,##0.00_)_-
forceAlign=False, zeroDash=True:   _(#,##0.00_)_-;(#,##0.00)_-;_(-_)_-
forceAlign=True,  zeroDash=False:   * _(#,##0.00_)_-; * (#,##0.00)_-; * _(#,##0.00_)_-
forceAlign=True,  zeroDash=True:    * _(#,##0.00_)_-; * (#,##0.00)_-; * _(-_)_-
```

`decimals=0` special case (no trailing `.`, must not become `0.`):

```
forceAlign=False, zeroDash=False:  _(#,##0_)_-;(#,##0)_-;_(#,##0_)_-
forceAlign=False, zeroDash=True:   _(#,##0_)_-;(#,##0)_-;_(-_)_-
forceAlign=True,  zeroDash=False:   * _(#,##0_)_-; * (#,##0)_-; * _(#,##0_)_-
forceAlign=True,  zeroDash=True:    * _(#,##0_)_-; * (#,##0)_-; * _(-_)_-
```

This function has zero Excel/COM dependency — it is a pure `string` transform of two `bool`s and an `int`. It is the single highest-value target for xUnit coverage (16 cases exactly enumerable), and the single highest-risk target for a silent transcription bug (quote-escaping, `_-`/`_)`/`_(` padding tokens, `IIf`-style ternary logic) that would compile fine but apply a subtly wrong format at runtime.

### Differentiators (Nice Architectural Improvements Now "Free")

Not new features — but real, valuable side effects of the rewrite that the roadmap should take credit for.

| Improvement | Value Proposition | Complexity | Notes |
|-------------|--------------------|------------|-------|
| `AccountingFmt`-equivalent as a pure, COM-free C# function | Fully unit-testable with xUnit (16 exhaustive cases) — impossible in VBA without a test framework | LOW (once ported) | Directly satisfies PROJECT.md's "Testes automatizados (xUnit) cobrindo a lógica de negócio" requirement. Keep this function in a project/assembly with **zero** reference to `Microsoft.Office.Interop.Excel`, so tests never touch COM/STA threading concerns. |
| Removal of the "checkbox toggle forces a full workbook save" anti-pattern | VBA's `SaveConfig` called `ThisWorkbook.Save` (with `EnableEvents` toggled off/on) on **every single checkbox click** — a documented anti-pattern in `ARCHITECTURE.md`. Removing persistence removes this side effect entirely, for free. | N/A (deletion) | This was flagged as an anti-pattern in the existing codebase mapping — the migration's decision to drop persistence coincidentally fixes it. |
| No more `CustomXMLPart` / namespace-collision surface | One less integration point (`urn:finance-fmt-tools` custom XML part, private `ReadBoolNode`/`WriteBoolNode`/`FindOrCreateXMLPart` helpers) to port or keep correct | N/A (deletion) | Removes ~50 lines of VBA-only XML-node-navigation code that doesn't need a C# equivalent at all. |
| Format-key registry becomes a proper enum + dictionary/switch in C# instead of `Select Case` on magic strings | Same extension pattern as VBA ("add one case + one constant"), but with compiler-checked exhaustiveness | LOW | Preserve the existing "one entry point, one registry" pattern (`ApplyFormat`/`GetFormatDef`) — it's already a good design, just give it static typing. |
| Clear separation possible now: Abstractions/Domain/Services/Ui layering (per the `outlook-classic-delay-send` sibling project referenced in PROJECT.md) | Makes the Excel-COM-touching code a thin shell around testable logic | MEDIUM | Architectural decision for STACK.md/ARCHITECTURE.md research, but directly enables the testability goal — flagging here because it changes which pieces count as "table stakes to port carefully" (the pure logic) vs. "thin, low-risk glue" (the COM calls). |

### Anti-Features / Explicit Non-Goals (This Milestone)

| Anti-Feature | Why It Looks Tempting While Porting | Why Avoid (This Milestone) | Alternative |
|--------------|--------------------------------------|------------------------------|-------------|
| Re-adding persistence for the two checkboxes (registry, JSON/config file, `CustomXMLPart`-equivalent, `Properties.Settings`) | ".NET makes this so easy (`Properties.Settings.Default`), why not just add it back better?" | Explicitly decided against in PROJECT.md ("Out of Scope" + Key Decisions) — deliberate simplification, not a gap | In-memory fields only: `_forceAlign = false`, `_zeroDash = true`, set once at add-in construction; no read/write path at all |
| New number formats or buttons beyond the existing 12 | "While I'm rewriting the engine anyway, this would be trivial to add" | Explicitly out of scope — "este milestone troca a implementação, não adiciona features" | Track as a future milestone request, not here |
| VSTO / ClickOnce / MSI installer | Visual Studio's Ribbon Designer + VSTO would arguably be *easier* to wire up Ribbon callbacks (strongly-typed, no reflection-by-name) | Explicitly rejected in PROJECT.md — requires full Visual Studio, breaks the VS Code + `dotnet` CLI constraint | Raw `IDTExtensibility2` + `IRibbonExtensibility`, embedded-resource Ribbon XML, HKCU registration, PowerShell installer |
| Using `NumberFormatLocal` instead of `NumberFormat` (e.g., "let's make it locale-aware since we're targeting pt-BR anyway") | Feels more "correct" for a Brazilian-Portuguese product to use the "local" property | `NumberFormatLocal` uses format codes **in the Excel UI language** (per Microsoft Learn: "format code for the object as a string in the language of the user") — mixing it with the invariant literal strings ported from VBA (which already encode the one locale nuance needed via `[$-pt-BR]`) would silently break formatting for any user whose Excel isn't in pt-BR display language, and is a strictly worse, untested code path vs. what's proven in production today | Always use `NumberFormat` (language-invariant) with the exact same literal strings as the VBA version |
| Any "smart"/adaptive alignment or zero-dash logic (e.g., auto-detect based on cell content, remember per-workbook) | Seems like a natural quality improvement | Out of scope — checkboxes are deliberately simple, global, session-only toggles | Keep the exact existing semantics: two global booleans read by `AccountingFmt`, mutated only by their own checkbox |
| Localizing the add-in UI (e.g., adding an English toggle) | ".NET resource (`.resx`) localization is trivial, why not?" | Not requested; existing product is pt-BR only, scope creep | Keep pt-BR-only strings, verbatim from the VBA source |
| Async/background format application, multi-threaded batch operations across a large selection | ".NET makes async easy, and large selections could be slow" | Excel COM interop requires STA + the main UI thread; VBA's synchronous single-call model isn't a limitation to "fix," it's a platform constraint | Keep `ApplyFormat` synchronous, exactly like VBA; only use `ScreenUpdating = false/true` around it as before |
| A generic Ribbon-callback framework / plugin system for "future" controls | Feels like a good abstraction opportunity mid-rewrite | Over-engineering for 12 buttons + 2 checkboxes; adds indirection with no current requirement | Keep the existing 1-line-per-callback convention (`onAction="RibbonFin2D"` → `ApplyFormatToSelection(FMT_FIN_2D)`), just in C# |

## Feature Dependencies

```
Ribbon XML (embedded resource, customUI14 schema)
    └──requires──> IRibbonExtensibility.GetCustomUI(RibbonID)
                       └──requires──> add-in class registered COM-visible + HKCU LoadBehavior=3

OnRibbonLoad(IRibbonUI ribbon)
    └──captures──> IRibbonUI handle (only needed for future Invalidate calls, NOT for getPressed itself)

_forceAlign (bool, default false)  ──┐
_zeroDash (bool, default true)     ──┤
                                      ├──read by──> AccountingFmt(decimals, forceAlign, zeroDash)
                                      │                  └──used by──> Fin 0D/2D/4D/8D format buttons
                                      │
    checkBox onAction (sets field) ──┘
    checkBox getPressed (reads field, returns bool) ──enhances──> visual checkbox state on next Ribbon paint

Pct/Date/Text/Spread buttons ── (no dependency on checkboxes; static literal format strings)

ApplyFormatToSelection(key)
    └──requires──> SafeSelection()-equivalent guard (Selection is Excel.Range, else friendly message)
    └──requires──> GetFormatDef(key) registry lookup
```

### Dependency Notes

- **Fin family requires `AccountingFmt`, which requires both checkbox states:** any change to the in-memory defaults (`false`/`true` per PROJECT.md) must be set before the first format button click is even possible in a session — trivially satisfied by field initializers, no ordering hazard like the old `LoadConfig`-before-`OnRibbonLoad` sequencing had to guarantee.
- **`getPressed` does not require the captured `IRibbonUI` handle** — it just reads the field directly. The handle is only needed if something *other* than the checkbox's own click needs to force a re-render (not the case here — see Parity Risk Matrix on `InvalidateControl`).
- **Pct/Date/Text/Spread buttons have zero coupling to the checkbox subsystem** — safe to port/test/ship independently of the Fin family and the checkbox wiring; useful for phase ordering (these are the lowest-risk, highest-confidence port targets and could reasonably be phase 1).

## MVP Definition — Reframed as "Full Parity Port" (No Staged Feature Rollout)

This milestone has no partial-launch concept — it swaps the implementation of an already-shipped product. Still useful to rank by port risk/cost for phase ordering:

### Must Port Byte-for-Byte (v1 = the whole milestone)

- [ ] `AccountingFmt` equivalent (pure function, 16 cases) — highest transcription risk, highest test value
- [ ] `GetFormatDef` registry (12 format keys → `FormatDef`) and `ApplyFormat`/`ApplyFormatToSelection`
- [ ] All 12 Ribbon buttons wired via `onAction`
- [ ] 2 checkboxes wired via `onAction` + `getPressed`, in-memory only, correct default states (off/on)
- [ ] Selection-guard equivalent to `SafeSelection()`
- [ ] About dialog + docs-link button

### Deliberately Simplified vs. VBA (already decided, not open)

- [ ] Checkbox state: session-only, no persistence, fixed defaults every launch

### Explicitly Deferred / Not This Milestone

- [ ] Any new formats, buttons, or preferences
- [ ] Localization beyond pt-BR
- [ ] VSTO/ClickOnce path

## Feature Prioritization Matrix (For Roadmap Phase Ordering)

| Feature | Port Value | Implementation Cost | Priority |
|---------|------------|----------------------|----------|
| `AccountingFmt` pure function + xUnit tests | HIGH | MEDIUM | P1 |
| Format registry + `ApplyFormat`/`ApplyFormatToSelection` | HIGH | LOW–MEDIUM | P1 |
| Pct/Date/Text/Spread static-format buttons | HIGH | LOW | P1 (low-risk, do early to validate the Ribbon wiring pattern) |
| Ribbon XML delivery via `IRibbonExtensibility.GetCustomUI` + `IDTExtensibility2.OnConnection` | HIGH | MEDIUM | P1 (blocks everything UI-facing) |
| Checkbox `onAction`/`getPressed` (session-only state) | HIGH | MEDIUM–HIGH | P1 (signature-translation risk, see Parity Risk Matrix) |
| Selection-guard equivalent | HIGH | MEDIUM | P1 (must exist before any format button is safe to ship) |
| About dialog / docs link | MEDIUM | LOW | P2 (low risk, can trail the core engine work) |
| HKCU-only registration + PowerShell installer | HIGH (ships the whole thing) | MEDIUM | P1 for the milestone, but architecturally separate from "features" — covered in STACK/ARCHITECTURE research |

**Priority key:**
- P1: Must have for this milestone to be considered "shipped"
- P2: Should have, low risk to sequence later within the same milestone

## VBA→C# Parity Risk Matrix

This is the section that most directly answers "what will silently break during the port."

| Risk | What Goes Wrong | Confidence | Mitigation |
|------|------------------|------------|------------|
| `NumberFormat` vs `NumberFormatLocal` mix-up | `NumberFormatLocal` returns/expects format codes "in the language of the user" (Microsoft Learn). If a developer inspects a cell's format in a pt-BR Format Cells dialog and pastes what they see (comma-decimal, semicolon-adjusted) into `.NumberFormat`, or uses `.NumberFormatLocal` "because the product is Brazilian," the string won't match Excel's expected invariant grammar and formatting breaks or silently differs from the VBA original. | HIGH (Microsoft Learn, `Range.NumberFormatLocal` docs) | Always set `.NumberFormat` (never `.NumberFormatLocal`) with the exact literal strings from `modFormatEngine.bas`. The one locale nuance needed (`[$-pt-BR]` for month abbreviations) is already embedded correctly inside the invariant strings — no additional locale logic needed anywhere else. |
| `getPressed` calling-convention mistranslation | VBA's Ribbon "get*" callbacks are `Sub`s with a `ByRef returnValue As Variant` out-parameter — a VBA-only idiom. The correct C# signature is a **function returning `bool` directly** (`public bool GetPressedZeroDash(Office.IRibbonControl control)`). A naive line-by-line port (`void` + `ref`/`out` param) compiles but Excel's Ribbon engine won't bind it correctly, and the checkbox will either always show unchecked or throw at ribbon load. | HIGH (Microsoft Learn managed-COM-add-in walkthrough + multiple independent tutorials converge on function-return form) | Explicitly re-derive each Ribbon callback signature from the *type* of callback (`onAction`, `getPressed`, etc.), not from mechanically translating the VBA `Sub`/`ByRef` shape. |
| Ribbon XML delivery mechanism change | VBA's `customUI14.xml` is a physical OOXML part read automatically by Excel from the `.xlam`. A COM add-in has no such automatic part — it must implement `IRibbonExtensibility.GetCustomUI(string RibbonID)` and return the XML (typically from an embedded resource), plus implement `IDTExtensibility2.OnConnection` to capture the `Excel.Application` object. This is a structural rewrite of the "how the UI shows up at all" layer, not a content copy. | HIGH (Microsoft Learn walkthrough: "Customize the Office Fluent ribbon by using a managed COM add-in") | Treat Ribbon-loading wiring as its own well-tested phase before porting individual button logic — if `GetCustomUI`/`OnConnection` is wrong, *no* button will work, masking whether the format logic itself is correct. |
| `SafeSelection()` exception semantics differ from VBA's `TypeName()` | VBA's `TypeName(Selection) <> "Range"` never throws, even for odd selection states (Shape/Chart focused, nothing selected). The C#/COM-interop equivalent (`Application.Selection as Excel.Range` or an `is` check against the RCW) can surface a `COMException`/`InvalidCastException` in some edge states rather than just returning a mismatched type. | MEDIUM (general COM-interop knowledge; not independently re-verified against every Excel edge case) | Wrap the selection-type check in `try/catch` around COM-specific exceptions, not just a type check, to faithfully reproduce VBA's "never crashes, just shows a friendly message" guarantee. |
| Checkbox visual state without an explicit `Invalidate` call | Community sources (MrExcel, Microsoft Q&A) state Excel does **not** auto-refresh a checkbox's `getPressed` value — you must call `ribbon.InvalidateControl(control.Id)` after changing backing state. Yet the current VBA code (`RibbonChkForceAlign`/`RibbonChkZeroDash`) never calls `Invalidate`/`InvalidateControl` anywhere, and this has apparently worked in production. | MEDIUM (community sources say Invalidate is required in general; but the existing single-control-per-flag VBA implementation appears to work without it — likely because Excel optimistically flips the visual state of *the very control the user clicked*, and only needs an explicit re-query when a *different* control or a non-click code path changes the same backing state) | Keep the same 1-control-per-flag design (no second control mirrors either boolean) so the existing "no Invalidate needed" behavior transfers unchanged. If the roadmap ever adds a second control reflecting the same state, `InvalidateControl` becomes mandatory at that point — flag this explicitly if scope ever expands. |
| UTF-8 encoding of embedded pt-BR strings | Both the embedded Ribbon XML resource and C# string literals contain accented characters (`ã`, `ç`, `é`, `à`, `õ`). A bare `new StreamReader(resourceStream)` without explicit encoding relies on BOM auto-detection; C# source file encoding must also be UTF-8. | LOW–MEDIUM (general .NET knowledge, not independently re-verified against this specific toolchain/CLI setup) | Verify the embedded `customUI14.xml` resource preserves its UTF-8 declaration/BOM, and that `dotnet build` treats `.cs` files as UTF-8 (default in modern C# tooling, but worth an explicit smoke test against real accented labels/tooltips once built). |
| `imageMso` icon ID availability across Excel 2016+ | The existing XML uses `AccountingNumberFormat`, `PercentStyle`, `ChartLine`, `TableInsertDate`, `DateAndTimeInsert`, `TextBox`, `TablePropertiesDialog`, `Info`. This is a copy-paste risk equally present whether the host is VBA or C# (it's just XML content), not a language-specific parity risk — included here only because it's a real regression risk if a target Excel version drops/renames one of these IDs. | LOW (not independently re-verified per-icon; flagged for completeness, not urgent) | Smoke-test the actual rendered Ribbon in the lowest supported Excel version (2016 per PROJECT.md constraints) before considering the port "done." |
| `Range.NumberFormat` PIA property typed as `object`, not `string` | Reading it back (e.g., for logging) requires an explicit cast (`(string)range.NumberFormat`); for a range spanning cells with mixed formats it can return `DBNull.Value` rather than a string. Writing a `string` to it works fine without a cast. | MEDIUM (standard PIA behavior) | Only relevant if the port adds any read-back/logging of the applied format (VBA's `Log` already does `rng.Address` not `rng.NumberFormat`, so this is a non-issue if the same logging shape is preserved — flagged so it doesn't get "improved" into a fragile read-back). |
| Dialog modality/parenting for the About box | VBA's `MsgBox` is naturally modal to/owned by the Excel window. `System.Windows.Forms.MessageBox.Show(text)` with no owner argument can appear un-parented (behind Excel, or not blocking Excel input as expected). | LOW–MEDIUM (general WinForms/COM-interop knowledge) | Pass an `IWin32Window` wrapping Excel's main window handle (`Application.Hwnd`) as the owner argument to `MessageBox.Show`. |

## Sources

- `archive/vba-legacy:src/modFormatEngine.bas`, `src/modRibbon.bas`, `src/modConfig.bas`, `src/modUtils.bas`, `src/customUI14.xml` (this repository) — direct source of the exact format strings and callback signatures being ported
- `.planning/codebase/ARCHITECTURE.md` — existing VBA data-flow mapping used to identify the persistence anti-pattern and layering
- [IRibbonExtensibility.GetCustomUI method (Office) | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/api/office.iribbonextensibility.getcustomui) — HIGH confidence, official
- [Customize the Office Fluent ribbon by using a managed COM add-in | Microsoft Learn](https://learn.microsoft.com/en-us/office/vba/Library-Reference/Concepts/customize-the-office-fluent-ribbon-by-using-a-managed-com-add-in) — HIGH confidence, official; source for `OnConnection`/`GetCustomUI`/`onAction` signature facts
- [Range.NumberFormatLocal Property (Microsoft.Office.Interop.Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.numberformatlocal?view=excel-pia) — HIGH confidence, official; "language of the user" wording quoted directly
- [Range.NumberFormat Property (Microsoft.Office.Interop.Excel) | Microsoft Learn](https://learn.microsoft.com/en-us/dotnet/api/microsoft.office.interop.excel.range.numberformat?view=excel-pia) — HIGH confidence, official
- [XML Callbacks: Adding Functionality to the Office Ribbon — nolongerset.com](https://nolongerset.com/ribbon-callbacks/) — MEDIUM confidence, community, corroborates VBA callback signature conventions
- [Custom Ribbon getPressed Issues - Microsoft Q&A](https://learn.microsoft.com/en-us/answers/questions/823104/custom-ribbon-getpressed-issues) — MEDIUM confidence, community, source for the `InvalidateControl` requirement discussion
- [Custom Ribbon with getEnabled, getVisible & getPressed | MrExcel Message Board](https://www.mrexcel.com/board/threads/custom-ribbon-with-getenabled-getvisible-getpressed.827397/) — MEDIUM confidence, community
- [How a COM Add-In Is Registered | flylib.com](https://flylib.com/books/en/2.53.1/how_a_com_add_in_is_registered.html) — MEDIUM confidence, corroborates HKCU/HKLM registration and `ProgId` key requirements

---
*Feature research for: Excel Ribbon add-in VBA → C# COM add-in migration (finance-fmt-tools)*
*Researched: 2026-07-10*
