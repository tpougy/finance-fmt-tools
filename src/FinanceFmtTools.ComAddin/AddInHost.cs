using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;

namespace FinanceFmtTools.ComAddin
{
    // Composition root — wires the real RealExcelGateway/TraceLog together with Phase 2's unmodified
    // RibbonController, and is the only class that knows about both live COM types
    // (Excel.Application/Office.IRibbonUI) and Phase 1/2's pure-C# engine. Connect.cs (the actual COM
    // entry point) contains zero business logic and delegates every method here (03-RESEARCH.md Pattern 1).
    public sealed class AddInHost
    {
        // Mirrors src/modConfig.bas's CFG_ADDIN_NAME/CFG_ADDIN_VERSION/CFG_DOCS_URL verbatim.
        private const string AddinName = "Finance Fmt Tools";
        private const string AddinVersion = "2.1.1";
        private const string DocsUrl = "https://github.com/tpougy/finance-fmt-tools";

        private readonly TraceLog _log = new TraceLog();

        // Constructed immediately in the field initializer (not deferred to Wire) so GetCustomUI works
        // even before Excel finishes calling OnConnection.
        private readonly RibbonController _ribbon = new RibbonController();

        private Excel.Application _app;
        private RealExcelGateway _gateway;
        private Office.IRibbonUI _ribbonUi;

        public RibbonController Ribbon => _ribbon;
        public ILog Log => _log;

        public void Wire(object application)
        {
            try
            {
                if (_gateway != null)
                {
                    _log.Info("AddInHost.Wire: já conectado — ignorando nova chamada.");
                    return;
                }

                var app = application as Excel.Application;
                if (app == null)
                {
                    _log.Error("AddInHost.Wire: Application recebido não é um Excel.Application.");
                    return;
                }

                _app = app;
                _gateway = new RealExcelGateway(app);
                _log.Info("AddInHost.Wire: conectado ao Excel.Application.");
            }
            catch (Exception ex)
            {
                _log.Error("AddInHost.Wire: " + ex);
            }
        }

        // Releases every cached COM reference (Excel.Application, IRibbonUI) and forces the RCWs to be
        // reclaimed immediately rather than waiting on GC finalization — the standard fix for the classic
        // "EXCEL.EXE lingers as a ghost process after the workbook closes" Office-interop leak.
        public void Teardown()
        {
            try
            {
                if (_ribbonUi != null && Marshal.IsComObject(_ribbonUi))
                {
                    Marshal.ReleaseComObject(_ribbonUi);
                }
            }
            catch (Exception ex)
            {
                _log.Error("AddInHost.Teardown (ribbonUi release): " + ex);
            }
            finally
            {
                _ribbonUi = null;
            }

            try
            {
                if (_app != null && Marshal.IsComObject(_app))
                {
                    Marshal.ReleaseComObject(_app);
                }
            }
            catch (Exception ex)
            {
                _log.Error("AddInHost.Teardown (app release): " + ex);
            }
            finally
            {
                _app = null;
                _gateway = null;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }

        public void CacheRibbonUi(object ribbonUi)
        {
            _ribbonUi = ribbonUi as Office.IRibbonUI;
            if (_ribbonUi == null)
            {
                _log.Warn("AddInHost.CacheRibbonUi: objeto recebido não é um Office.IRibbonUI.");
            }
        }

        // Pattern 3 (03-RESEARCH.md): never re-wraps Phase 2's higher-level FormatEngine.ApplyToSelection —
        // guards the selection itself (showing the FMT-06 friendly message on failure) and calls the
        // lower-level FormatEngine.Apply directly, keeping ApplyToSelection's dotnet-test-proven contract
        // (log warning, never throw, never show UI) completely untouched.
        public void ApplyFormat(string formatKey)
        {
            if (_gateway == null)
            {
                _log.Warn("AddInHost.ApplyFormat: add-in ainda não conectado (not wired) — abortando '" + formatKey + "'.");
                return;
            }

            if (!_gateway.TryGetSelectedRange(out IRangeHandle range))
            {
                MessageBox.Show(
                    "Selecione um intervalo de células antes de aplicar a formatação.",
                    AddinName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
                _log.Warn("AddInHost.ApplyFormat: seleção inválida para '" + formatKey + "'.");
                return;
            }

            try
            {
                FormatEngine.Apply(range, _log, formatKey, _ribbon.Config.ForceAlign, _ribbon.Config.ZeroDash);
            }
            finally
            {
                // Release the wrapped Range RCW promptly rather than waiting on GC finalization —
                // IDisposable is on RealRangeHandle itself, not on Phase 2's IRangeHandle contract.
                (range as IDisposable)?.Dispose();
            }
        }

        public bool GetForceAlign() => _ribbon.Config.ForceAlign;

        public void SetForceAlign(bool pressed)
        {
            _ribbon.Config.ForceAlign = pressed;
            _log.Info("ForceAlign alterado para " + pressed);
            InvalidateControl("chkForceAlign");
        }

        public bool GetZeroDash() => _ribbon.Config.ZeroDash;

        public void SetZeroDash(bool pressed)
        {
            _ribbon.Config.ZeroDash = pressed;
            _log.Info("ZeroDash alterado para " + pressed);
            InvalidateControl("chkZeroDash");
        }

        // Defensive (Pattern 4, 03-RESEARCH.md) — src/modRibbon.bas's VBA never calls InvalidateControl
        // after a checkbox toggle, relying on Excel's native checkbox widget to reflect its own click
        // state. Whether that "just works" without an explicit invalidate could not be resolved to HIGH
        // confidence without a live Excel session, so this call costs nothing if it turns out unnecessary.
        private void InvalidateControl(string controlId)
        {
            try
            {
                _ribbonUi?.InvalidateControl(controlId);
            }
            catch (Exception ex)
            {
                _log.Error("AddInHost.InvalidateControl(" + controlId + "): " + ex);
            }
        }

        public void ShowAbout()
        {
            MessageBox.Show(
                AddinName + " v" + AddinVersion + "\n\n" +
                "Formatação financeira padronizada para mercado de capitais." + "\n\n" +
                "Autor: Thomaz Pougy",
                "Sobre",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information);
        }

        public void OpenDocs()
        {
            try
            {
                // Explicit ProcessStartInfo form always (03-RESEARCH.md Pitfall 4/Anti-Patterns) — never
                // the bare single-string Process.Start(url) overload.
                Process.Start(new ProcessStartInfo(DocsUrl) { UseShellExecute = true });
                _log.Info("AddInHost.OpenDocs: abriu " + DocsUrl);
            }
            catch (Exception ex)
            {
                _log.Error("AddInHost.OpenDocs: " + ex);
            }
        }
    }
}
