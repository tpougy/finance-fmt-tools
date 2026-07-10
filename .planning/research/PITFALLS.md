# Pitfalls Research

**Domain:** C# COM add-in for Excel (IDTExtensibility2 + Ribbon XML, no VSTO, HKCU-only registration) — migrating from a working VBA `.xlam` add-in
**Researched:** 2026-07-10
**Confidence:** MEDIUM-HIGH (grounded in official Microsoft Learn docs, a working sibling C# COM add-in implementation for Outlook — `~/pessoal/outlook-classic-delay-send` — that already solved most of these problems empirically, and verified community sources; flagged individually where confidence is lower)

This document intentionally does **not** repeat VBA-specific pitfalls already documented in `.planning/codebase/ARCHITECTURE.md`'s "Anti-Patterns" section (`SafeSelection()` guarding, silent-save side effects). Everything below is a **new** risk class introduced specifically by moving from an in-process VBA `.xlam` to an out-of-VBA, registry-registered C# COM server.

---

## Critical Pitfalls

### Pitfall 1: COM registration bitness mismatch (HKCU registry view redirection)

**What goes wrong:**
The installer writes `HKCU\Software\Classes\CLSID\{guid}\InprocServer32` (and the `ProgId`/`CLSID` pairing keys) using whichever PowerShell process bitness happens to run the script (64-bit PowerShell is the Windows default). If the user's installed Excel is **32-bit** (still common in enterprises — Microsoft recommended 32-bit Office for legacy add-in compatibility for years, and many corporate images still deploy it even in 2026), Excel never finds the CLSID: the add-in shows as present in the Add-ins manager (or not at all), `LoadBehavior` may look correct, but Excel silently fails to activate it, or after repeated failed loads Office marks it disabled.

**Why it happens:**
Per Microsoft's official WOW64 registry redirection table, `HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID` **is a redirected key** on Windows 7/Server 2008 R2 and newer (this is not limited to `HKLM`). A 32-bit process (32-bit Excel) reading `HKCU\Software\Classes\CLSID\{guid}` is transparently redirected by the OS to a separate, WOW64-specific physical location than what a 64-bit process (64-bit PowerShell, the default host) writes to. Writing once, from the "wrong" bitness process, produces a registration that is invisible to the other bitness of Excel. This is a real risk distinct from Outlook: `outlook-classic-delay-send`'s `install.ps1` only **warns** if Outlook isn't x64 ("baseline validado foi x64") — it never actually handles 32-bit registration, because its validated baseline was x64-only. Finance Fmt Tools' constraint ("Excel 2016+") does not pin a bitness, so this cannot be left as a warning-only gap.

**How to avoid:**
- Detect Excel's bitness at install time (PE header of `EXCEL.EXE`, or `HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration\Platform` — same technique already used in the sibling's `Test-PeMachine` helper).
- Write the three COM registration keys (`ProgId`, `CLSID`, `InprocServer32`) using the **matching registry view**, not just whatever bitness PowerShell happens to run as. The cleanest mechanism is `reg.exe add ... /reg:32` or `/reg:64` (an explicit, well-documented Windows flag that targets the WOW64 or native view regardless of the calling process's own bitness) — safer than hand-crafting an unofficial `Wow6432Node` path under `HKCU` (community reports on whether a literal `Wow6432Node` subkey is inspectable there are inconsistent, even though the OS-level redirection of `HKCU\Software\Classes\CLSID` itself is officially documented).
- Alternative: re-invoke the install logic under the matching-bitness PowerShell host explicitly (`%WINDIR%\SysWOW64\WindowsPowerShell\v1.0\powershell.exe` for 32-bit) so the OS's automatic redirection does the right thing without any manual path juggling.
- Verify post-install by reading back the key **as the target bitness would see it** (e.g., shell out to `reg query ... /reg:32` when target is 32-bit), not just via the PowerShell registry provider running natively.

**Warning signs:**
Add-in installs "successfully" (script reports OK) but the Ribbon tab never appears; Excel's COM Add-ins dialog shows the add-in with no error, or it silently disappears from the list after a few Excel restarts; works on the developer's machine (matching bitness) but fails for a subset of end users.

**Phase to address:**
Install/registration phase — this must be designed in from the start, not discovered after a support ticket. Verify against both a 32-bit and a 64-bit Excel install if both are realistically in the user base (or explicitly document/enforce a single supported bitness in the installer's pre-flight check if only testing one).

---

### Pitfall 2: `ClassInterfaceType.None` hides Ribbon callback methods from Excel

**What goes wrong:**
The Ribbon XML wires buttons/checkboxes to callback method names (`onAction="RibbonFin2D"`, `getPressed="RibbonGetZeroDash"`, `onLoad="OnRibbonLoad"`, etc.). Office resolves these **by name at runtime via `IDispatch.GetIDsOfNames`**, not via a static, compiled COM interface. If the COM entry-point class is decorated with `[ClassInterface(ClassInterfaceType.None)]` (a common "more correct/explicit COM" instinct for .NET developers, and the default many templates encourage for production COM servers), the class exposes **no** auto-generated dispinterface, so none of its public methods are visible via late binding — buttons appear in the Ribbon (the XML still renders) but silently do nothing on click, and checkboxes never reflect state, with **no exception, no log entry, no visible error**.

**Why it happens:**
This is not theoretical: the sibling Outlook C# add-in hit this exact failure mode live, during its Gate 6 smoke test, and the fix — switching to `[ClassInterface(ClassInterfaceType.AutoDispatch)]` — is called out explicitly in its `Connect.cs` with a comment documenting the diagnosis. It generalizes identically to Excel: Ribbon Fluent UI extensibility invokes callbacks the same way (late-bound `IDispatch`) regardless of host application.

**How to avoid:**
Mark the COM entry-point class `[ClassInterface(ClassInterfaceType.AutoDispatch)]` (not `None`, not `AutoDual`). Keep the assembly-level `[assembly: ComVisible(false)]` and only the single entry-point class `[ComVisible(true)]`, so only intended surface area is exposed. Do not rename public Ribbon callback methods casually — a rename without updating the matching `onAction`/`getPressed`/`getEnabled` string in the embedded Ribbon XML fails the exact same way (silent no-op), so add an automated or checklist-based cross-check between method names and XML attribute values before each release.

**Warning signs:**
Ribbon renders correctly (tab, groups, buttons, icons all visible) but clicking any button does nothing; checkboxes never show a pressed/unpressed state; no exception anywhere in logs.

**Phase to address:**
Ribbon/callback wiring phase — verify with a live Excel smoke test (not just unit tests, since this failure is invisible to abstraction-layer unit tests) before considering that phase done.

---

### Pitfall 3: Unhandled exceptions in Ribbon callbacks or `IDTExtensibility2` methods trigger Office's automatic add-in disabling

**What goes wrong:**
If any Ribbon callback (`getPressed`, `getEnabled`, `onAction`, `getLabel`, `loadImage`) or any `IDTExtensibility2` method (`OnConnection`, `OnStartupComplete`, etc.) throws an unhandled exception, Excel's Resiliency subsystem treats this as a crash signal. After enough failures, Excel marks the add-in as disabled and adds a binary entry under `HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems` — the add-in then silently fails to load on next launch, with the only visible symptom being "the add-in is not in the Ribbon anymore," which most users (and even developers who don't know to check Resiliency keys) will misdiagnose as a registration bug rather than a runtime crash.

**Why it happens:**
This is documented, well-known Office add-in behavior (confirmed for Excel specifically, not just Outlook): "Excel disables your ribbon if there are unhandled exceptions thrown from the ribbon-handler routines." `GetCustomUI()` and every handler must be defensively wrapped. The VBA original already does this centrally via `HandleError`/`On Error GoTo` in every public procedure — that discipline must be ported 1:1 (or stricter) to C#, because the *consequence* of a leaked exception is more severe in a COM add-in (auto-disable) than in VBA (an error dialog or a silently failed Sub).

**How to avoid:**
- Wrap every public method reachable from Office (all `IDTExtensibility2` methods, `GetCustomUI`, every Ribbon callback) in `try/catch`, log, and return a safe default (`string.Empty`, `false`, or simply swallow for `void` methods) — mirroring the sibling's `Connect.cs` pattern (`_host.Ribbon?.GetX() ?? default; catch { TryLog(...); }`) exactly.
- Proactively write `HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList\<ProgId> = 1` (DWORD) during install — this doesn't prevent Excel from soft-disabling on a genuine crash, but prevents disabling from *transient* issues and is cheap insurance, matching the sibling project's install script.
- Never let a constructor throw: COM instantiates the class directly; a throwing constructor becomes an opaque `E_FAIL`/class-not-created error with no useful diagnostic surfaced to the user.

**Warning signs:**
Add-in "disappears" after some period of use without the user uninstalling anything; re-installing "fixes" it temporarily (because reinstalling doesn't touch `DisabledItems`, but users often also toggle it back on via Excel's Add-ins manager, resetting the disabled flag).

**Phase to address:**
Ribbon/callback wiring phase (defensive coding pattern) **and** install/registration phase (the `DoNotDisableAddinList` key). Add a "does every callback have try/catch" checklist item to the phase's Definition of Done.

---

### Pitfall 4: RCW / COM object leaks from Range/Application access (memory bloat, stale references, not "true" zombie EXCEL.EXE for the add-in itself — but very real for any automation-based verification scripts)

**What goes wrong:**
Every `Range`, `Worksheet`, `Selection`, or `Application` object obtained through Interop is backed by a Runtime Callable Wrapper (RCW) holding a live COM reference count. Two specific C# habits keep RCWs (and the objects they wrap) alive far longer than intended:
1. Chaining property access ("two dots"), e.g. `Globals.Application.Selection.NumberFormat = "..."` — each intermediate object in the chain (`Application`, `Selection`) is a temporary RCW the runtime never explicitly releases.
2. Using `foreach` over any Excel collection (`Cells`, `Worksheets`, `Range` cast to an enumerable) — `foreach` allocates a hidden COM enumerator object that is never released.
Inside the add-in itself this manifests as gradual memory growth and, in the worst case, "This method or property is not available" / `RPC_E_DISCONNECTED` errors if a stale reference to a since-closed workbook/range is touched again — not a headless zombie process (the add-in is in-process with the single Excel the user is looking at, so there's no separate process to leak). The **actual zombie-EXCEL.EXE** risk shows up in any **out-of-process** verification tooling: a PowerShell or C# smoke-test script that does `New-Object -ComObject Excel.Application` to launch a real Excel and click-test the Ribbon, and then fails to release every intermediate COM object and call `Application.Quit()` correctly, leaves an invisible `EXCEL.EXE` running forever — a very well-known PowerShell+COM automation gotcha, and each repeated CI/local test run accumulates another orphaned process.

**Why it happens:**
The .NET garbage collector has no visibility into COM reference counts; only `Marshal.ReleaseComObject` (or the CLR's finalizer, on its own schedule, far later) decrements them. `foreach`'s hidden enumerator and multi-dot chains create references nothing in application code ever explicitly names or releases.

**How to avoid:**
- Never chain more than one dot deep; assign every intermediate COM object (`Application`, `Selection`/`Range`, `Worksheet`) to a local variable.
- Never `foreach` over a raw Excel COM collection; use an indexed `for` loop (`for (int i = 1; i <= coll.Count; i++)`) and release each non-kept item inside the loop body, mirroring the exact `TryFindChild`-style pattern already used in the sibling's `OutlookGateway.cs` (iterate by index, `Release()` everything not being returned, in a `finally`).
- In the format-application code path, resolve `Application.Selection` **fresh on every button click** (never cache it across calls — see Pitfall 8), do the format/alignment write, then `Marshal.ReleaseComObject` the `Range` in a `finally` block, following the sibling's `Release(object comObject)` helper pattern (`Marshal.IsComObject(x)` guard + swallow-and-log on failure).
- Prefer `Marshal.ReleaseComObject` over `Marshal.FinalReleaseComObject` as the default (decrement-by-one is safer; final-release everywhere risks `InvalidComObjectException` if any other code path still holds a reference to the same RCW).
- **Do not** port the legacy installer's approach of driving Excel via COM automation to register the add-in (per `.planning/codebase/ARCHITECTURE.md`, `Install-FinanceFmtTools.ps1` today "registers it in Excel via COM automation"). The new HKCU-registry-write installer (see Pitfall 1/6) needs **zero** live Excel automation to register the add-in — this eliminates an entire class of zombie-EXCEL.EXE risk that existed in the VBA-era installer.
- If any local/manual smoke-test script does automate a real Excel (e.g., to click-test the Ribbon end to end), it must follow the same release discipline and explicitly call `Application.Quit()` + release the `Application` RCW itself, otherwise every test run leaks a process.

**Warning signs:**
Task Manager shows `EXCEL.EXE` processes lingering after all visible Excel windows are closed; Excel becomes progressively slower over a long session of heavy formatting; `RPC_E_DISCONNECTED` or "COM object that has been separated from its underlying RCW" exceptions in logs.

**Phase to address:**
Format-engine/core-port phase (the release discipline inside `ApplyFormatToSelection`'s C# port) **and** any phase that adds an Excel-automation-based smoke test (CI or manual runbook) — that tooling needs the same discipline applied to itself.

---

### Pitfall 5: `dotnet build`/`dotnet test` on GitHub Actions `windows-latest` fails to resolve Office interop assemblies

**What goes wrong:**
GitHub-hosted runners have **no Office installed** — there's no GAC-registered `Microsoft.Office.Interop.Excel`/`office`/`stdole` PIA to resolve, and `TlbImp.exe`/`AxImp.exe` (the NETFX SDK tools that `<COMReference WrapperTool="tlbimp">` needs to regenerate interop assemblies from a TypeLib) are not installed on `windows-latest` either. A `.csproj` that references Office interop the "textbook VSTO" way (`<COMReference>` blocks, or a NuGet package that merely re-wraps an unofficial PIA build) either fails outright in CI (`MSB4803`/`MSB3086`-class errors) or silently resolves to an unofficial NuGet package whose GUIDs may not match the CLSID/TypeLib strings actually registered on end-user machines.

**Why it happens:**
This is exactly the failure the sibling Outlook project hit and solved (documented in its own `BUILD.md`, item C-2): `dotnet build` with `<COMReference>` failed with `MSB4803` in that project too, on a machine that DID have the tools locally — the problem is strictly worse on a bare CI runner with no Office/NETFX-SDK tools at all. Microsoft has never published first-party, reliably-versioned NuGet packages for the classic Office PIAs (`Interop.Excel`, `office`, `stdole`, `Extensibility`); anything found on nuget.org under those names is unofficial/community-repackaged and risks GUID/version drift from what's actually registered by a real Office install.

**How to avoid:**
- Vendor the exact PIA DLLs (sourced from a real Office installation's GAC, matching the target Office/TypeLib version) directly in the repo, e.g. `lib/Microsoft.Office.Interop.Excel.dll`, `lib/OFFICE.DLL`, `lib/stdole.dll`, `lib/Extensibility.dll` — the same four-DLL pattern the sibling project committed for Outlook, swapping only the Excel interop DLL.
- Reference them via plain `<Reference Include="..."><HintPath>..\..\lib\X.dll</HintPath><SpecificVersion>false</SpecificVersion><Private>true</Private></Reference>` blocks — **not** `<COMReference>` — so `dotnet build` resolves them as ordinary file references, with `Private=true` ensuring they're copied to the output/package directory (needed since end-user machines and CI have no GAC copy to fall back on).
- Confirm the exact PIA GUIDs/versions against a real Office install once (e.g. via `regasm /regfile` run locally, never at install time — see Pitfall 6) and pin that as the vendored version; don't silently let a different Office version's PIA get vendored later without re-verifying.
- Treat `windows-latest` as a moving target: GitHub periodically rotates the underlying OS image (e.g. Windows Server 2022 → 2025), and .NET Framework targeting-pack availability has already changed between image generations for older framework versions (`net40` support was dropped moving from `windows-2019` to `windows-2022`). Pin to a specific image tag (e.g. `windows-2022`) if reproducibility matters more than always getting the newest image, and add a fast "does `dotnet build` for `net48` succeed" canary as the very first CI step so a future image rotation fails loudly instead of breaking releases silently.

**Warning signs:**
CI build fails with `MSB4803`, `MSB3086`, or "reference assemblies for .NETFramework,Version=v4.8 were not found"; build succeeds locally (dev machine has Office installed) but fails only in GitHub Actions.

**Phase to address:**
CI/CD pipeline phase — but the **vendored `lib/` DLLs must be prepared during the core-port/project-scaffolding phase**, before CI is even wired up, since the CI phase's job is to prove the already-working vendoring approach also builds headless.

---

### Pitfall 6: Reaching for `regasm.exe` (or a VSTO/ClickOnce-style install) as the actual install-time mechanism instead of direct registry writes

**What goes wrong:**
`regasm.exe` is genuinely present on any machine with .NET Framework installed (it ships with the Framework, not just the SDK), which makes it tempting to shell out to it from `install.ps1`. But its default behavior targets `HKEY_CLASSES_ROOT`/`HKLM`, which requires admin — directly violating the "HKCU-only, no admin" constraint. Getting `regasm` to *only* write to `HKCU` reliably (without ever touching `HKLM`) is inconsistent across .NET Framework versions and easy to get subtly wrong; and even when it works, it doesn't create the Office-specific "discovery" keys (`...\Excel\Addins\<ProgId>` with `LoadBehavior`) or the Resiliency allow-list key at all — those still need to be hand-written regardless. Reg-free COM (activation-context manifests) is not a viable alternative here either: it lets a *client app* declare its own private CLSID mappings via its own manifest, but Excel (the process that must `CoCreateInstance` a third-party add-in's ProgID) has no mechanism to consume a manifest belonging to an arbitrary add-in DLL it doesn't ship with — Office COM add-in discovery is registry-only, full stop.

**Why it happens:**
Most COM interop tutorials assume a VSTO or admin-installed scenario where `regasm`/MSI conventions apply directly. The HKCU-only constraint is a deliberate, non-default choice (matching the sibling project and this project's stated goals), so following generic COM-registration tutorials literally leads away from the actually-required approach.

**How to avoid:**
- Use `regasm /codebase /regfile:out.reg` **once, locally, as a discovery/documentation tool** (never at actual install time) to capture the exact `Assembly`, `CodeBase` format, `RuntimeVersion`, and `ThreadingModel` string values regasm would use — then translate those into direct `New-Item`/`Set-ItemProperty` calls under `HKCU:\Software\Classes\...` in `install.ps1`, adjusting the key root from `HKEY_CLASSES_ROOT` to `HKEY_CURRENT_USER\Software\Classes` and rewriting `CodeBase` to point at the actual install directory (e.g. `%LocalAppData%\FinanceFmtTools\`), not the build output path.
- Write all three key groups explicitly: (a) COM class (`ProgId`→`CLSID`, `CLSID`→`InprocServer32` with `mscoree.dll` shim, `ThreadingModel=Both`, `Assembly`, `RuntimeVersion`, `CodeBase`), (b) the Office discovery key (`HKCU\Software\Microsoft\Office\Excel\Addins\<ProgId>` — **not** version-numbered, unlike the Resiliency key below — with `LoadBehavior=3` DWORD and `FriendlyName`/`Description`), (c) the Resiliency allow-list (`HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList\<ProgId>=1` — **is** version-numbered).
- Make `install.ps1`/`uninstall.ps1` fully idempotent (safe to re-run, safe to run when never installed) exactly as the sibling's scripts do.

**Warning signs:**
Installer requests or silently requires elevation; installer "succeeds" for the developer (who has local admin) but fails for end users; add-in registers but never appears in Excel because the Office-specific discovery key was never written (only the raw COM class was).

**Phase to address:**
Install/registration phase.

---

### Pitfall 7: Cross-thread COM access (STA apartment violations) from timers, async/await, or background work

**What goes wrong:**
Any Ribbon callback, `IDTExtensibility2` method, or event handler executes on Excel's single main STA (single-threaded apartment) thread. If code spins up `System.Threading.Tasks.Task.Run(...)`, a raw `Thread`, or uses `async`/`await` with `ConfigureAwait(false)` and then touches any Excel COM object (`Application`, `Range`, `Worksheet`) from the continuation, that continuation resumes on a thread-pool thread — a different apartment — and any COM call throws `RPC_E_WRONG_THREAD` (`0x8001010E`, "The application called an interface that was marshalled for a different thread"). `System.Threading.Timer` has the identical problem: its callback fires on a thread-pool thread, not the UI thread.

**Why it happens:**
COM objects exposed by Excel are apartment-threaded; the CLR's async/thread-pool infrastructure has no inherent knowledge of, or respect for, that constraint.

**How to avoid:**
- Keep all Excel COM access synchronous, on the thread Excel invokes the callback on. Finance Fmt Tools' format operations are inherently synchronous (apply a `NumberFormat` string to a `Range`), so there's rarely a good reason to introduce threading here at all — treat this as a hard rule rather than a common need in this codebase.
- If any future feature needs a delay/timer (mirroring the sibling project's send-delay feature), use `System.Windows.Forms.Timer` (fires via the existing Win32 message loop Excel already pumps on its main thread) — not `System.Threading.Timer` — exactly as the sibling's `WinFormsTimerCore` wrapper does. No `Application.Run()` call is needed from the add-in itself; Excel's own message loop already services `WM_TIMER`.
- If background work (e.g., a future GitHub-releases update check) is ever added, marshal any Excel-touching continuation back onto the UI thread (e.g. via a hidden `Control`'s `Invoke`/`BeginInvoke`, or `SynchronizationContext.Current` captured on the UI thread) rather than touching COM objects directly from `await`.

**Warning signs:**
Intermittent `COMException` with HRESULT `0x8001010E`; failures that only reproduce when a delay/timer/background task is involved, never on direct button clicks.

**Phase to address:**
Format-engine/core-port phase (establish "no background threads touch COM objects" as an architectural rule) — flag for deeper research only if/when a future phase actually introduces async work (none is currently in scope per `PROJECT.md`).

---

### Pitfall 8: Caching `Workbook`/`Worksheet`/`Range` references across Ribbon callback invocations

**What goes wrong:**
Unlike Outlook (one MAPI session/mailbox for the whole app), Excel's `Application.Workbooks` can hold many simultaneously open workbooks, and the "active" one changes freely between clicks — including workbooks in entirely separate top-level windows, each potentially a **separate `EXCEL.EXE` process** since Excel 2013 introduced SDI (Single Document Interface): each top-level Excel window typically runs in its own process, each loading its own independent instance of the add-in DLL, its own `Connect`/`IRibbonUI`, and its own in-memory `CFG_FORCE_ALIGN`/`CFG_ZERO_DASH`-equivalent state. If any C# code holds onto a `Workbook`/`Worksheet`/`Range` COM reference across calls (e.g., a field cached at `OnConnection` time) instead of re-resolving `Application.ActiveWorkbook`/`Application.Selection` fresh on every button click, two things can go wrong: (a) the cached reference points at the wrong (non-active) workbook if the user switched windows between clicks, and (b) if the cached workbook was since closed, touching it throws `RPC_E_DISCONNECTED`-class errors.

**Why it happens:**
This is a genuine behavioral difference from the Outlook reference implementation and from a naive C#/WinForms mental model where "the main window" is a stable, long-lived singleton — Excel's multi-workbook, multi-process-per-window model breaks that assumption. It also does not require any new bug in *this* migration if the VBA original already always read `Selection` fresh (it does, via `SafeSelection()`) — the risk is specifically that a well-intentioned C# refactor introduces caching that wasn't present in the VBA version, since "cache the expensive object" is a common (and usually correct) C# performance instinct that is wrong here.

**How to avoid:**
- Resolve `Application.Selection` fresh inside every `onAction` callback — never store a `Range`/`Worksheet`/`Workbook` field for reuse across calls.
- Do not assume a single, process-wide singleton state for the two checkbox preferences; each Excel window/process legitimately has its own copy, and per this milestone's explicit scope decision (no cross-session persistence), that's the accepted behavior — just don't accidentally introduce a mechanism (e.g. a file on disk, a named pipe) that tries to synchronize state across windows, since that's out of scope and adds needless complexity/risk.
- Release the resolved `Selection`/`Range` immediately after use each time (see Pitfall 4), rather than holding it "just in case" for the next click.

**Warning signs:**
Format applies to the wrong workbook/sheet when the user has multiple Excel windows open; intermittent `RPC_E_DISCONNECTED`/"disconnected from clients" exceptions correlated with the user having closed a workbook since the add-in last touched it.

**Phase to address:**
Format-engine/core-port phase — call out explicitly as a design constraint when porting `ApplyFormatToSelection`.

---

## Technical Debt Patterns

| Shortcut | Immediate Benefit | Long-term Cost | When Acceptable |
|----------|-------------------|-----------------|-----------------|
| Skipping the 32-bit Excel bitness handling in the installer (warn-only, like the sibling's Outlook baseline) | Ships faster, matches the sibling's proven pattern exactly | Silent install "success" with a non-functional add-in for any 32-bit Excel user; hard-to-diagnose support burden | Only if the project explicitly commits (in `PROJECT.md`/README) to a single supported bitness and the install script hard-fails with a clear message on the other bitness, rather than silently "succeeding" |
| Using `ClassInterface(ClassInterfaceType.AutoDispatch)` instead of a hand-written explicit dual interface | Zero extra COM plumbing code, matches the proven working sibling pattern | Slightly less "textbook COM" (no static vtable/typelib for external tooling); irrelevant here since Office only calls these methods via late binding by name anyway | Always acceptable for this use case — this is the recommended approach, not a shortcut to later revisit |
| Not writing an integration test that drives a real, live Excel via COM automation | Faster CI, no flaky live-Office dependency, avoids introducing new zombie-process risk in test tooling | Pitfalls 2/3 (silent Ribbon callback failures, resiliency auto-disable) are invisible to abstraction-layer unit tests and can only be caught by an actual Excel smoke test | Acceptable to defer full automation of this smoke test, but a **documented manual smoke-test runbook step** before each release is not optional — mirrors the sibling's Gate 6 |
| Vendoring PIA DLLs in `lib/` rather than solving the "proper" NuGet-based COMReference build | Ships today, matches the sibling's exact validated approach, works headless in CI | `lib/` DLLs must be manually refreshed if targeting a materially different Office/TypeLib version later; binary files committed to git | Always acceptable — this is the recommended, not a temporary, approach for this build model |

## Integration Gotchas

| Integration | Common Mistake | Correct Approach |
|-------------|-----------------|-------------------|
| Office Ribbon Fluent UI (`IRibbonExtensibility.GetCustomUI`) | Marking the embedded `customUI14.xml` as `<Content>`/`<None>` instead of `<EmbeddedResource>` in the `.csproj` | Mark as `<EmbeddedResource>`; resolve by manifest-resource-name suffix at runtime (robust to namespace drift) exactly as the sibling's `RibbonController.LoadEmbeddedXml` does, so a rename of the assembly/root namespace doesn't silently break resource lookup |
| Office Ribbon `imageMso` vs custom icons | Porting the sibling Outlook project's `loadImage`/`IPictureDisp`/`AxHost.GetIPictureDispFromPicture` machinery wholesale, even though the current VBA Ribbon XML only ever references built-in `imageMso` icons | Keep using `imageMso="..."` attributes directly in the Ribbon XML for all existing buttons (Excel resolves these internally with zero callback code needed); only reach for the `loadImage`/`IPictureDisp` pattern if/when a genuinely custom (non-built-in) icon is introduced |
| GitHub Releases distribution | Assuming the release artifact only needs the add-in DLL | The `.zip` must be self-sufficient: add-in DLL + all four vendored interop DLLs + `install.ps1`/`uninstall.ps1`/`verify-environment.ps1`, mirroring the sibling's `releases/*.zip` layout — a DLL-only zip breaks on a machine without Office's PIAs in the GAC |
| `dotnet` CLI build of a classic-style (non-SDK-only) interop reference | Using `<COMReference WrapperTool="tlbimp">`, which requires `TlbImp.exe`/`AxImp.exe` (NETFX SDK tools) not guaranteed present on either dev or CI machines | Reference the vendored `lib/*.dll` PIAs directly via `<Reference><HintPath>...</HintPath><Private>true</Private></Reference>` |

## Performance Traps

| Trap | Symptoms | Prevention | When It Breaks |
|------|----------|------------|-----------------|
| Setting `NumberFormat`/`HorizontalAlignment` cell-by-cell in a loop instead of once on the whole `Range` | Visible UI lag/flicker on large selections; each COM call has fixed marshaling overhead even in-process | Set the format property once on the entire `Range` object (as the VBA original already does) — never iterate cells for a uniform format change | Selections of a few hundred to thousands of cells; a per-cell loop turns an O(1) COM call into an O(n) one |
| Leaving `Application.ScreenUpdating`/`Calculation` untouched during a bulk formatting operation | Excel repaints on every property write, visible flicker on large selections | Port the VBA original's `ScreenUpdating = False/True` wrap around `ApplyFormat` verbatim, restoring it in a `finally`/`Cleanup` path even on error (as the VBA version already does) | Selections spanning full columns/rows or very large ranges |
| Accumulating unreleased RCWs across a long Excel session (see Pitfall 4) | Gradual memory growth over hours of use; eventual `RPC_E_DISCONNECTED` on stale references | Release every `Range`/`Selection` object right after use, per invocation | Long-running sessions with heavy repeated use of the add-in, not a single-click scenario |

## Security Mistakes

| Mistake | Risk | Prevention |
|---------|------|------------|
| Writing an installer that requires or silently escalates to admin "just to be safe" | Contradicts the explicit no-admin constraint; increases the trust boundary a downloaded-and-run script needs, which is a bigger ask for end users pulling a script via `irm ... \| iex` | Keep every registry write under `HKCU` only; if a write to `HKLM`/admin is ever "tempting" (e.g. a shared-machine install), treat that as a distinct, separately-justified feature, not a default |
| Registering the add-in with a `CodeBase` pointing at a temp/download folder instead of a stable per-user install directory | If the temp file is later deleted (e.g. by disk cleanup), Excel fails to load the add-in with a cryptic COM error next launch | Copy binaries to a stable location (e.g. `%LocalAppData%\FinanceFmtTools\`) before registering, and point `CodeBase` at that final path — mirroring the sibling's `install.ps1` exactly |
| Trusting an unofficial/third-party NuGet package for Office PIAs | Supply-chain risk (unverified publisher) plus functional risk (GUID/version mismatch against what's actually registered on the end user's Office install) | Vendor DLLs pulled from a real, known-good local Office installation directly in the repo (`lib/`), as the sibling project does |

## UX Pitfalls

| Pitfall | User Impact | Better Approach |
|---------|-------------|-------------------|
| Checkbox visual state ("Alinhar à direita"/"Zero contábil") drifting from actual applied-format behavior after the add-in is disabled/re-enabled via Excel's Add-ins manager within the same session | User sees a checkbox state that doesn't match what will actually happen on next click, confusing given the intentional removal of persistence | Document this as an accepted, known limitation (matches the milestone's explicit "no persistence" scope decision) rather than attempting a fix; make sure defaults (ForceAlign=off, ZeroDash=on) are applied consistently every time the add-in (re-)connects, not just on first-ever load |
| Installer silently "succeeding" on the wrong Excel bitness (Pitfall 1) | User sees no error but the Ribbon tab never appears; user assumes they did something wrong | Installer must hard-fail with an actionable message if it detects a mismatch it cannot resolve, rather than reporting a green "Instalação concluída com sucesso" |
| Add-in getting soft-disabled by Excel Resiliency after a transient crash, with no user-visible explanation | User loses the Ribbon tab with zero indication why, and "reinstalling" doesn't fix it (only re-enabling via Excel's Add-ins manager, or the `DoNotDisableAddinList` key, does) | Document the Add-ins manager re-enable path in the README/troubleshooting section, and proactively set `DoNotDisableAddinList` at install time |

## "Looks Done But Isn't" Checklist

- [ ] **COM registration:** Verify the discovery key (`HKCU\Software\Microsoft\Office\Excel\Addins\<ProgId>`) has **no** Office-version segment in its path, while the Resiliency key (`HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DoNotDisableAddinList`) **does** — mixing these up is an easy, silent mistake.
- [ ] **Ribbon callbacks:** Confirm every `onAction`/`getPressed`/`getEnabled`/`getLabel` string in the embedded Ribbon XML has an exactly-matching public method name on the `[ComVisible(true)]`/`AutoDispatch` class — a typo fails silently (Pitfall 2), not with an exception.
- [ ] **getPressed wiring:** Verify the `getPressed` callback reads the *exact same* mutable field/service instance that the checkbox's `onAction` handler writes — not a stale copy captured at construction.
- [ ] **CI build:** Confirm `dotnet build`/`dotnet test` succeed on a clean `windows-latest` runner with **no local Office install and no dev-machine GAC state** — a green build on a dev machine that happens to have Office installed proves nothing about CI.
- [ ] **Release artifact completeness:** Confirm the packaged `.zip` contains the add-in DLL **and** all four vendored interop DLLs **and** `install.ps1`/`uninstall.ps1`/`verify-environment.ps1` — a DLL-only artifact silently fails to load on a machine without the Office PIAs already registered.
- [ ] **Bitness:** Confirm install/uninstall have been exercised against every Excel bitness actually supported/claimed, not just the developer's own machine's bitness.
- [ ] **Exception safety:** Confirm every method reachable from Office (all `IDTExtensibility2` methods + every Ribbon callback) has a top-level try/catch that never lets an exception escape to Office.

## Recovery Strategies

| Pitfall | Recovery Cost | Recovery Steps |
|---------|----------------|-----------------|
| Add-in soft-disabled by Excel Resiliency | LOW | Excel → Options → Add-ins → manage "Disabled Items" → re-enable; or delete the specific `DisabledItems` binary entry under `HKCU\Software\Microsoft\Office\16.0\Excel\Resiliency\DisabledItems` for this ProgID; ensure `DoNotDisableAddinList` is set going forward |
| Wrong-bitness registration (Pitfall 1) | LOW-MEDIUM | Re-run `uninstall.ps1`, then re-run `install.ps1` with bitness-aware registry-view writes; no data loss risk since there's no persisted state to migrate |
| Zombie `EXCEL.EXE` from a smoke-test/verification script | LOW | Kill the orphaned process(es) via Task Manager/`Stop-Process`; fix the script's release discipline (Pitfall 4) so it doesn't recur |
| CI break from a `windows-latest` image rotation dropping/moving the .NET Framework 4.8 targeting pack | LOW-MEDIUM | Pin the workflow to a specific runner image tag (e.g. `windows-2022`) as a stopgap; add the `Microsoft.NETFramework.ReferenceAssemblies` NuGet package as a build-time fallback so targeting-pack resolution doesn't depend on the image at all |

## Pitfall-to-Phase Mapping

| Pitfall | Prevention Phase | Verification |
|---------|-------------------|---------------|
| COM registration bitness mismatch | Install/registration phase | Install and load the add-in against both a 32-bit and a 64-bit Excel (or hard-fail explicitly on the unsupported one); read back registration keys "as the target bitness would see them" |
| `ClassInterfaceType.None` hiding Ribbon callbacks | Ribbon/callback wiring phase | Live Excel smoke test: click every button, toggle every checkbox — a passing unit-test suite alone does not catch this |
| Unhandled exceptions triggering Resiliency auto-disable | Ribbon/callback wiring phase + install/registration phase | Code review checklist: every Office-reachable method wrapped in try/catch; `DoNotDisableAddinList` key present after install |
| RCW leaks / stale COM references | Format-engine/core-port phase | Manual long-session soak test (repeatedly click format buttons across many minutes); confirm no `EXCEL.EXE` memory growth or `RPC_E_DISCONNECTED` errors |
| `dotnet build`/`test` failing on `windows-latest` due to interop resolution | CI/CD pipeline phase (but `lib/` vendoring must exist by end of core-port/scaffolding phase) | First green CI run on a freshly pushed tag, with **no** manual pre-warming of the runner |
| Reaching for `regasm`/admin-requiring registration | Install/registration phase | `install.ps1` never invokes `regasm.exe`; runs successfully as a non-admin user with UAC untouched |
| Cross-thread COM access | Format-engine/core-port phase | Architectural rule documented + code review; deeper research needed only if a later phase introduces async/background work (none currently in scope) |
| Caching Workbook/Worksheet/Range across calls | Format-engine/core-port phase | Code review: `ApplyFormatToSelection`'s C# port resolves `Application.Selection` fresh every invocation, never from a cached field |

## Sources

- [Registry Keys Affected by WOW64 — Microsoft Learn](https://learn.microsoft.com/en-us/windows/win32/winprog64/shared-registry-keys) — official table confirming `HKEY_CURRENT_USER\SOFTWARE\Classes\CLSID` is a WOW64-redirected key on current Windows (HIGH confidence)
- `~/pessoal/outlook-classic-delay-send/scripts/install.ps1`, `uninstall.ps1`, `verify-environment.ps1` — working, empirically-validated HKCU-only, no-admin, no-regasm install/uninstall implementation for a sibling C# COM Office add-in
- `~/pessoal/outlook-classic-delay-send/.github/workflows/release.yml` — working GitHub Actions `windows-latest` CI pipeline building/testing/packaging a `net48` COM add-in with vendored Office interop DLLs
- `~/pessoal/outlook-classic-delay-send/BUILD.md` — documents the `dotnet build` + `<COMReference>` failure (`MSB4803`) and the direct-PIA-reference fix; documents `regasm /regfile`-as-discovery-tool pattern; documents the exact registry value set captured
- `~/pessoal/outlook-classic-delay-send/src/UndoSend/Connect.cs` — documents the live `ClassInterfaceType.None` → `AutoDispatch` fix found during its Gate 6 smoke test
- `~/pessoal/outlook-classic-delay-send/src/UndoSend/Services/OutlookGateway.cs` — reference RCW release discipline (`Release()` helper, index-based iteration instead of `foreach`, resolving `Session`/folders fresh per call)
- `~/pessoal/outlook-classic-delay-send/src/UndoSend/Services/RibbonController.cs`, `Composition/AddInHost.cs` — reference `IRibbonUI` caching/invalidation pattern and defensive null-safe wiring
- [How to properly release Excel COM objects: C# code examples — Add-in Express](https://www.add-in-express.com/creating-addins-blog/release-excel-com-objects/) — "two dots" rule, `foreach` enumerator leak, `ReleaseComObject` vs `FinalReleaseComObject`, `GC.Collect`/`WaitForPendingFinalizers` guidance (MEDIUM-HIGH confidence, community source consistent with official `Marshal.ReleaseComObject` docs)
- [Support for keeping add-ins enabled — Microsoft Learn (Outlook, applies identically to Excel's Resiliency mechanism)](https://learn.microsoft.com/en-us/office/vba/outlook/concepts/getting-started/support-for-keeping-add-ins-enabled) — confirms unhandled Ribbon-callback exceptions trigger auto-disable, and the `DisabledItems`/`DoNotDisableAddinList` registry mechanics (HIGH confidence for the general mechanism; verified applicable to Excel via community Q&A sources referencing the identical `...\Office\16.0\Excel\Resiliency\...` path)
- GitHub Community discussions on building legacy `.NET Framework` targets on `windows-latest`/`windows-2022` runners (MEDIUM confidence — confirms general viability but flags that targeting-pack availability has changed across runner image generations for some Framework versions, motivating the "pin the image, add a canary build step" recommendation)
- `.planning/codebase/ARCHITECTURE.md` (this repo) — baseline of what the current VBA/legacy installer does today (COM-automation-based registration), used to identify what changes/improves with the new approach

---
*Pitfalls research for: C# COM add-in migration (Excel, IDTExtensibility2 + Ribbon XML, HKCU-only)*
*Researched: 2026-07-10*
