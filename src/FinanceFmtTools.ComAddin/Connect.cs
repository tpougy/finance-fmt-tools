using System;
using System.Runtime.InteropServices;
using Extensibility;
using Office = Microsoft.Office.Core;

namespace FinanceFmtTools.ComAddin
{
    // ================================================================================================
    // Fixed values Phase 4's installer (not yet planned) MUST reuse verbatim — Phase 3 only declares
    // them via attributes; no registry/regasm code is written in this phase (03-RESEARCH.md Pattern 5):
    //   GUID (CLSID):        881EFDF3-424C-4240-BCA0-714DAC2B9CD7
    //   ProgId:              FinanceFmtTools.Connect
    //   AssemblyName:        FinanceFmtTools.ComAddin  (see FinanceFmtTools.ComAddin.csproj <AssemblyName>)
    //   Version:             1.0.0.0                   (see FinanceFmtTools.ComAddin.csproj <Version>)
    //   Excel discovery key: HKCU\Software\Microsoft\Office\Excel\Addins\FinanceFmtTools.Connect
    // ================================================================================================
    [ComVisible(true)]
    [Guid("881EFDF3-424C-4240-BCA0-714DAC2B9CD7")]
    [ProgId("FinanceFmtTools.Connect")]
    // AutoDispatch — never the "no default dispinterface" alternative. Ribbon callbacks below
    // (RibbonFin2D, RibbonChkForceAlign, ...) are not members of IDTExtensibility2/IRibbonExtensibility;
    // Office invokes them by NAME via late-bound IDispatch.GetIDsOfNames, which that alternative does not
    // expose — a first-hand production bug diagnosed in the sibling outlook-classic-delay-send project
    // (03-RESEARCH.md Pitfall 2).
    [ClassInterface(ClassInterfaceType.AutoDispatch)]
    public sealed class Connect : IDTExtensibility2, Office.IRibbonExtensibility
    {
        // Thin COM entry point — zero business logic, every method a 1-3 line delegation to the
        // composition root (03-RESEARCH.md Pattern 1), matching src/modRibbon.bas's own documented
        // convention: "cada callback tem exatamente 1 linha de lógica".
        private readonly AddInHost _host = new AddInHost();

        public void OnConnection(object Application, ext_ConnectMode ConnectMode, object AddInInst, ref Array custom)
        {
            try { _host.Wire(Application); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void OnDisconnection(ext_DisconnectMode RemoveMode, ref Array custom)
        {
            try { _host.Teardown(); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void OnAddInsUpdate(ref Array custom) { }
        public void OnStartupComplete(ref Array custom) { }
        public void OnBeginShutdown(ref Array custom) { }

        public string GetCustomUI(string RibbonID)
        {
            try { return _host.Ribbon?.GetCustomUiXml() ?? string.Empty; }
            catch (Exception ex) { TryLog(ex); return string.Empty; }
        }

        public void OnRibbonLoad(Office.IRibbonUI ribbonUi)
        {
            try { _host.CacheRibbonUi(ribbonUi); }
            catch (Exception ex) { TryLog(ex); }
        }

        // -- Família Fin / Numérico ------------------------------------------------------------------

        public void RibbonInteger(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Integer); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonFin2D(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Fin2D); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonFin4D(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Fin4D); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonFin8D(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Fin8D); }
            catch (Exception ex) { TryLog(ex); }
        }

        // -- Percentual -------------------------------------------------------------------------------

        public void RibbonPct4D(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Pct4D); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonPct2D(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Pct2D); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonSpreadBps(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.SpreadBps); }
            catch (Exception ex) { TryLog(ex); }
        }

        // -- Datas ------------------------------------------------------------------------------------

        public void RibbonDateISO(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.DateIso); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonDateBR(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.DateBr); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonDateBRLong(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.DateBrLong); }
            catch (Exception ex) { TryLog(ex); }
        }

        // -- Texto ------------------------------------------------------------------------------------

        public void RibbonText(Office.IRibbonControl control)
        {
            try { _host.ApplyFormat(FinanceFmtTools.Engine.FormatKeys.Text); }
            catch (Exception ex) { TryLog(ex); }
        }

        // -- Checkboxes de configuração (sessão apenas — sem persistência, RIB-02/RIB-03) --------------

        public void RibbonChkForceAlign(Office.IRibbonControl control, bool pressed)
        {
            try { _host.SetForceAlign(pressed); }
            catch (Exception ex) { TryLog(ex); }
        }

        public bool RibbonGetForceAlign(Office.IRibbonControl control) => _host.GetForceAlign();

        public void RibbonChkZeroDash(Office.IRibbonControl control, bool pressed)
        {
            try { _host.SetZeroDash(pressed); }
            catch (Exception ex) { TryLog(ex); }
        }

        public bool RibbonGetZeroDash(Office.IRibbonControl control) => _host.GetZeroDash();

        // -- Info -------------------------------------------------------------------------------------

        public void RibbonFinInfo(Office.IRibbonControl control)
        {
            try { _host.OpenDocs(); }
            catch (Exception ex) { TryLog(ex); }
        }

        public void RibbonAbout(Office.IRibbonControl control)
        {
            try { _host.ShowAbout(); }
            catch (Exception ex) { TryLog(ex); }
        }

        // Last line of defense — an unhandled exception escaping a COM entry-point method can cause
        // Excel's Resiliency system to silently disable the add-in, so logging itself must never throw.
        private void TryLog(Exception ex)
        {
            try { _host.Log?.Error(ex.ToString()); }
            catch { /* nothing further we can safely do here */ }
        }
    }
}
