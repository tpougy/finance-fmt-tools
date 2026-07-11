// Hand-rolled IDTExtensibility2 shim — no official lightweight NuGet package exists for the
// classic "Extensibility" assembly (see 03-RESEARCH.md Pitfall 3). COM resolves interfaces by
// GUID, not by which assembly declares the .NET type, so this local declaration is functionally
// identical to referencing the real Extensibility.dll. The GUID below is cross-checked against
// Microsoft Learn's published attribute on Extensibility.IDTExtensibility2, and independently
// confirmed byte-identical inside the official Microsoft.VisualStudio.Interop NuGet package via
// `strings` (03-RESEARCH.md session).
using System;
using System.Runtime.InteropServices;

namespace Extensibility
{
    [ComImport]
    [Guid("B65AD801-ABAF-11D0-BB8B-00A0C90F2744")]
    [InterfaceType(ComInterfaceType.InterfaceIsIDispatch)]
    public interface IDTExtensibility2
    {
        void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom);
        void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom);
        void OnAddInsUpdate(ref Array custom);
        void OnStartupComplete(ref Array custom);
        void OnBeginShutdown(ref Array custom);
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
