---
status: issues
files_reviewed: 7
depth: standard
findings:
  critical: 0
  warning: 0
  info: 3
  total: 3
fixed:
  - "CR-1: AddInHost.Teardown() now releases the cached Excel.Application and IRibbonUI via Marshal.ReleaseComObject, then forces two GC.Collect()/WaitForPendingFinalizers() cycles — the standard fix for ghost EXCEL.EXE processes"
  - "WR-1: RealRangeHandle now implements IDisposable (without touching Phase 2's IRangeHandle contract); AddInHost.ApplyFormat disposes the range in a finally block after FormatEngine.Apply completes"
  - "WR-2: RibbonGetForceAlign/RibbonGetZeroDash in Connect.cs now wrapped in try/catch like all 15 other callbacks, with safe fallback return values matching each checkbox's documented default"
  - "WR-3: RealExcelGateway.TryGetSelectedRange now catches COMException from _app.Selection (no workbook open) and returns false instead of letting the exception propagate"
  - "WR-4: RealRangeHandle's NumberFormat/HorizontalAlignment getters now guard against DBNull mixed-value results from a multi-cell selection, returning a safe default instead of throwing InvalidCastException"
skipped:
  - "IN-1: IDTExtensibility2's InterfaceIsIDispatch vs InterfaceIsDual marshaling attribute could not be independently verified further in this environment — deferred to the human_needed live-Excel smoke test"
  - "IN-2: No TraceListener configured for TraceLog — deferred, matches VBA's own silent-by-default logging; not blocking"
  - "IN-3: AddinVersion hardcoded separately from csproj <Version> — minor drift risk, deferred as low-value churn"
---

## Summary

Reviewed all 7 Phase 3 files (`AddInHost.cs`, `Connect.cs`, `Extensibility.cs`, `RealExcelGateway.cs`, `RealRangeHandle.cs`, `TraceLog.cs`, `FinanceFmtTools.ComAddin.csproj`) against `src/customUI14.xml`, `src/modUtils.bas`/`modRibbon.bas` (source of truth for callback names and behavior), and the unmodified Phase 1/2 engine contracts (`IExcelGateway`, `IRangeHandle`, `ILog`, `FormatEngine`).

**Confirmed correct:**
- Every `onAction`/`getPressed`/`onLoad` name in `customUI14.xml` (17 callbacks total) has a matching method in `Connect.cs`, with the right signature (`void(IRibbonControl)` for buttons, `void(IRibbonControl, bool)` for checkbox `onAction`, `bool(IRibbonControl)` for `getPressed`, `void(IRibbonUI)` for `onLoad`).
- `[ClassInterface(ClassInterfaceType.AutoDispatch)]` is used on `Connect` (not `None`), matching 03-RESEARCH.md's documented pitfall.
- `[ComVisible(false)]` at the assembly level (csproj) + explicit `[ComVisible(true)]` only on `Connect` is the correct hide-by-default/opt-in pattern.
- Every `IDTExtensibility2`/`IRibbonExtensibility` method Excel calls into (`OnConnection`, `OnDisconnection`, `GetCustomUI`, `OnRibbonLoad`, and all 15 button/checkbox `onAction` handlers) wraps its body in try/catch + log, so a thrown exception can't cross the COM boundary and crash Excel or trigger Resiliency — except the two gaps noted in WR-2 below.
- No cross-thread/reentrancy issue with the cached `Office.IRibbonUI` field: it's set once in `OnRibbonLoad` and only read from `InvalidateControl`, both of which only ever run on Excel's single STA UI thread (no async/Task/timer code anywhere in this phase).

**Issues found:** one COM-lifecycle gap severe enough to risk the classic "Excel won't fully let go on exit" symptom, plus four narrower correctness/robustness gaps and three lower-priority notes — detailed below.

### CR-1

**file:** `src/FinanceFmtTools.ComAddin/AddInHost.cs:60-64` (also `src/FinanceFmtTools.ComAddin/RealExcelGateway.cs:10-17`)

`Teardown()` — called from `Connect.OnDisconnection` — only nulls the managed `_gateway`/`_ribbonUi` fields:

```csharp
public void Teardown()
{
    _gateway = null;
    _ribbonUi = null;
}
```

It never releases the underlying COM objects those fields point to. `RealExcelGateway` holds `_app` (the live `Excel.Application` RCW) for the entire add-in lifetime and exposes no `Dispose`/`Release` method at all, so there is no code path anywhere in this phase that ever calls `Marshal.ReleaseComObject` on `Application`, and `_ribbonUi` (an `Office.IRibbonUI` COM object) is likewise just dropped. Across all 7 files, the *only* `Marshal.ReleaseComObject` call site targets a throwaway non-`Range` `Selection` object (`RealExcelGateway.cs:29`) — nothing releases `Application`, `Worksheet`, or `IRibbonUI`, and `Teardown()` never forces `GC.Collect()`/`GC.WaitForPendingFinalizers()` to flush pending RCWs before Excel's shutdown sequence proceeds.

**Failure scenario:** User disables the add-in or closes Excel; `OnDisconnection` fires, `Teardown()` nulls references, but the native COM reference count on `Application` (and any not-yet-finalized `Range` RCWs from earlier ribbon clicks) is still elevated because nothing explicitly released them and the CLR GC's timing isn't guaranteed to run before Excel checks its own shutdown state. This is the textbook root cause of an add-in leaving `EXCEL.EXE` lingering in Task Manager after all visible windows close, or of Excel hanging/stuttering on exit.

**Suggested fix:** Give `RealExcelGateway` an internal `Release()`/`IDisposable.Dispose()` that does `if (Marshal.IsComObject(_app)) Marshal.ReleaseComObject(_app);`. Have `AddInHost.Teardown()` call it, do the same for `_ribbonUi`, then force `GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();` before returning, so any pending `Range` RCWs (see WR-1) are also flushed at disconnect time.

### WR-1

**file:** `src/FinanceFmtTools.ComAddin/RealExcelGateway.cs:19-31`

```csharp
public bool TryGetSelectedRange(out IRangeHandle range)
{
    object sel = _app.Selection;
    if (sel is Excel.Range r)
    {
        range = new RealRangeHandle(r);
        return true;
    }

    range = null;
    if (Marshal.IsComObject(sel)) Marshal.ReleaseComObject(sel); // release the non-Range object we don't need
    return false;
}
```

The success branch (`sel is Excel.Range r`) wraps and returns the `Range` RCW but never releases it — inconsistent with the very next lines, which explicitly release `sel` in the failure branch. Since `ApplyFormat` runs on every ribbon-button click, every click creates a fresh `Range` RCW that is reclaimed only whenever the .NET GC happens to finalize it, with no deterministic release anywhere (`RealRangeHandle`/`IRangeHandle` have no `Dispose`).

**Failure scenario:** Heavy Ribbon use over a long session (hundreds of format clicks) accumulates unreleased `Range` RCWs between GC generations; combined with CR-1's missing GC flush on disconnect, this compounds the "Excel won't cleanly free memory / lingers before exit" symptom described above.

**Suggested fix:** Have `RealRangeHandle` implement `IDisposable` wrapping `Marshal.ReleaseComObject(_range)`, and have `AddInHost.ApplyFormat` dispose the handle (`(range as IDisposable)?.Dispose();` in a `finally`) after `FormatEngine.Apply` returns — this can be done entirely within Phase 3's own files without touching Phase 1/2's unmodified `IRangeHandle` interface.

### WR-2

**file:** `src/FinanceFmtTools.ComAddin/Connect.cs:143,151`

```csharp
public bool RibbonGetForceAlign(Office.IRibbonControl control) => _host.GetForceAlign();
...
public bool RibbonGetZeroDash(Office.IRibbonControl control) => _host.GetZeroDash();
```

`RibbonGetForceAlign` and `RibbonGetZeroDash` — the two `getPressed` callbacks Office invokes to refresh checkbox state — are the only two Ribbon-invoked members in `Connect` with no try/catch, unlike all 15 other callbacks in the same class (including their paired `onAction` handlers `RibbonChkForceAlign`/`RibbonChkZeroDash`, which do catch). If `_host.GetForceAlign()`/`GetZeroDash()` ever throws, the exception propagates unhandled across the COM boundary from inside a `getPressed` callback — exactly what the file's own comment (`Connect.cs:167-168`) warns can crash Excel or get the add-in auto-disabled by Resiliency. Currently low-probability (both getters are simple non-throwing property reads today), but it's an inconsistency with the file's own stated exception-safety convention and one refactor away from being a real gap.

**Suggested fix:** Wrap both bodies in the same try/catch-and-log pattern as every other callback, returning a safe default (`false`) from the catch block, since `IRibbonExtensibility` `getPressed` callbacks must return a `bool`.

### WR-3

**file:** `src/FinanceFmtTools.ComAddin/RealExcelGateway.cs:21` (surfaces via `AddInHost.cs:87`)

`object sel = _app.Selection;` is read with no guard. In real Excel, reading `Application.Selection` throws a `COMException` when there is no active workbook/window — e.g. the user closed the last open workbook but left Excel running (the custom Ribbon tab typically stays visible in that state). Because neither `TryGetSelectedRange` nor `AddInHost.ApplyFormat` catches this, the exception skips the FMT-06 friendly "Selecione um intervalo de células..." `MessageBox` that `ApplyFormat`'s own header comment says is its job, and is instead only caught — silently, from the user's point of view — by `Connect.cs`'s generic per-callback try/catch. The net effect: user clicks a format button with no workbook open, and nothing visibly happens (just a swallowed log entry).

**Suggested fix:** Wrap the `_app.Selection` read in `RealExcelGateway.TryGetSelectedRange` in a try/catch for `COMException`, returning `false` (no valid range) so `AddInHost.ApplyFormat`'s existing friendly-message path handles it the same way it already handles a Chart/Shape selection.

### WR-4

**file:** `src/FinanceFmtTools.ComAddin/RealRangeHandle.cs:21,29`

```csharp
public string NumberFormat
{
    get => (string)_range.NumberFormat;
    ...
}
...
var v = (Excel.XlHAlign)_range.HorizontalAlignment;
```

Both getters cast a raw Excel COM Variant directly. `Range.NumberFormat`/`Range.HorizontalAlignment` return `DBNull`/Null (not a string/enum) whenever the underlying multi-cell range has mixed number formats or mixed alignment across cells — a common real-world selection (e.g. a column where some cells are already formatted differently). Casting `DBNull.Value` to `string` or `XlHAlign` throws `InvalidCastException`.

Neither getter is currently invoked anywhere in the reviewed call paths — `FormatEngine.Apply` (Phase 1/2, unmodified) only ever *writes* these properties, never reads them — so this is latent rather than actively triggered today. But it's a real defect in Phase 3's implementation of Phase 2's `IRangeHandle` contract (`string NumberFormat { get; set; }`) and will throw the moment any current or future caller reads either property against a mixed-selection.

**Suggested fix:** Guard both getters against the mixed-value case, e.g. `_range.NumberFormat as string ?? string.Empty`, and check `_range.HorizontalAlignment is Excel.XlHAlign h` before casting (falling back to `CellAlignment.General` otherwise) — mirroring how VBA callers must already check `TypeName(...)` before trusting these same properties.

### IN-1

**file:** `src/FinanceFmtTools.ComAddin/Extensibility.cs:15`

`IDTExtensibility2` is declared with `[InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]`. This matches common hand-rolled "COM add-in without VSTO" reference implementations, but 03-RESEARCH.md's stated verification method ("confirmed byte-identical... via `strings`") only proves the GUID's *text* bytes match a real `Extensibility.dll` — `strings` cannot recover a type library's `InterfaceIsDual` vs. `InterfaceIsIDispatch` marshaling metadata (that's a metadata table entry, not printable text), so this specific attribute choice hasn't actually been independently verified in this environment. Worth a quick double-check against a decompiled `Extensibility.dll` (ildasm/dotPeek/ILSpy) or a real Office install before the first live-Excel smoke test, since a mismatch here would only surface as a runtime COM activation/dispatch failure in real Excel — exactly the class of bug this environment can't catch by building/testing alone.

### IN-2

**file:** `src/FinanceFmtTools.ComAddin/TraceLog.cs` (whole file)

No `App.config` and no `Trace.Listeners.Add(...)` exists anywhere in the repo, so `Trace.TraceWarning/Information/Error` only ever reach the built-in `DefaultTraceListener`, which surfaces solely through an attached debugger or a tool like Sysinternals DebugView — never a file or anything an end user could send back to the developer. This roughly mirrors VBA's `Log()` being silent-by-default (gated by `CFG_LOG_ENABLED`), but VBA additionally had an optional durable `LogToSheet` persistence path (a hidden worksheet baked into the `.xlam` itself, `src/modUtils.bas:30-54`) that `TraceLog` has no C# equivalent for. Not a blocker for this phase — proving the `ILog` contract against a real dependency was the explicit scope — but worth deciding on a durable sink (e.g. a `TextWriterTraceListener` writing to `%APPDATA%\FinanceFmtTools\log.txt`) before relying on these logs to diagnose a real user-reported bug, since right now a support conversation with an end user has literally nothing to retrieve.

### IN-3

**file:** `src/FinanceFmtTools.ComAddin/AddInHost.cs:19`

`AddinVersion = "1.0.0"` is a separate hardcoded literal from `FinanceFmtTools.ComAddin.csproj`'s `<Version>1.0.0.0</Version>` (`FinanceFmtTools.ComAddin.csproj:10`) — two independent places to remember to bump on every release, and their formats already differ (`1.0.0` vs `1.0.0.0`). Purely cosmetic (shows up in the "Sobre" dialog) but a minor drift risk.

**Suggested fix:** Derive the displayed version from `Assembly.GetExecutingAssembly().GetName().Version` at runtime instead of maintaining a second literal.
