// Hand-rolled IDTExtensibility2 shim — no official lightweight NuGet package exists for the
// classic "Extensibility" assembly (see 03-RESEARCH.md Pitfall 3). COM resolves interfaces by
// GUID, not by which assembly declares the .NET type, so this local declaration is functionally
// identical to referencing the real Extensibility.dll. The GUID below is cross-checked against
// Microsoft Learn's published attribute on Extensibility.IDTExtensibility2, and independently
// confirmed byte-identical inside the official Microsoft.VisualStudio.Interop NuGet package via
// `strings` (03-RESEARCH.md session).
//
// InterfaceType must be Dual (the default — do not set InterfaceIsIDispatch here): reflecting
// the real Extensibility.dll (see ~/pessoal/outlook-classic-delay-send/lib/Extensibility.dll)
// shows TypeLibType = 0x1040 (FDual | FDispatchable). Excel's native COM Add-in loader calls
// OnConnection through the vtable, not late-bound IDispatch.Invoke; declaring this interface as
// IDispatch-only silently breaks that vtable layout — the add-in still activates (CreateObject
// succeeds) but Excel's LoadBehavior gets auto-downgraded 3 -> 2 on the very first load attempt,
// with no managed exception and no Windows Event Log entry (confirmed by reproducing the failure
// via COM automation + registry inspection before applying this fix).
//
// The [DispId]/[In]/[MarshalAs] attributes below are not decoration — they were reverse-engineered
// via reflection against the real Extensibility.dll (same sibling-project copy referenced above)
// and matched byte-for-byte: DispId(1..5) per method in declaration order, object params marshaled
// as UnmanagedType.IDispatch (26), and every `ref Array custom` marshaled as UnmanagedType.SafeArray
// (29). Without these, the vtable this interface produces is ABI-incompatible with what Excel's
// native add-in loader expects: QueryInterface for the IID still succeeds (proven independently),
// but the OnConnection call itself never reaches managed code — Excel silently treats the load as
// failed and downgrades LoadBehavior, with zero managed exception and zero Event Log trace.
using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    public interface IDTExtensibility2
    {
        [DispId(1)]
        void OnConnection(
            [In, MarshalAs(UnmanagedType.IDispatch)] object Application,
            [In] ext_ConnectMode ConnectMode,
            [In, MarshalAs(UnmanagedType.IDispatch)] object AddInInst,
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(2)]
        void OnDisconnection(
            [In] ext_DisconnectMode RemoveMode,
            [In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(3)]
        void OnAddInsUpdate([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(4)]
        void OnStartupComplete([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);

        [DispId(5)]
        void OnBeginShutdown([In, MarshalAs(UnmanagedType.SafeArray)] ref Array custom);
    }

    public enum ext_ConnectMode
    {
        ext_cm_AfterStartup = 0,
        ext_cm_Startup = 1,
        ext_cm_External = 2,
        ext_cm_CommandLine = 3
    }

    public enum ext_DisconnectMode
    {
        ext_dm_HostShutdown = 0,
        ext_dm_UserClosed = 1
    }
}
