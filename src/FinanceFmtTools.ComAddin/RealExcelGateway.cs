using System.Runtime.InteropServices;
using FinanceFmtTools.Engine.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinanceFmtTools.ComAddin
{
    // Real IExcelGateway implementation over a live Microsoft.Office.Interop.Excel.Application,
    // proving Phase 2's unmodified interface (src/FinanceFmtTools.Engine/Abstractions/IExcelGateway.cs)
    // against the actual Excel object model.
    public sealed class RealExcelGateway : IExcelGateway
    {
        private readonly Excel.Application _app;

        public RealExcelGateway(Excel.Application app)
        {
            _app = app;
        }

        public bool TryGetSelectedRange(out IRangeHandle range)
        {
            object sel;
            try
            {
                // Selection throws COMException when no workbook is open (or nothing is selectable) —
                // that is functionally an invalid selection for FMT-06's purposes, not a crash.
                sel = _app.Selection; // typed `object` in the real PIA — Selection can be Range/Chart/Shape/etc.
            }
            catch (COMException)
            {
                range = null;
                return false;
            }

            if (sel is Excel.Range r)
            {
                range = new RealRangeHandle(r);
                return true;
            }

            range = null;
            if (Marshal.IsComObject(sel)) Marshal.ReleaseComObject(sel); // release the non-Range object we don't need
            return false;
        }
    }
}
