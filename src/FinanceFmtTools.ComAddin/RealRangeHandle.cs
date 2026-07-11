using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinanceFmtTools.ComAddin
{
    // Real IRangeHandle implementation over a live Microsoft.Office.Interop.Excel.Range,
    // proving Phase 2's unmodified interface (src/FinanceFmtTools.Engine/Abstractions/IRangeHandle.cs)
    // against the actual Excel object model.
    public sealed class RealRangeHandle : IRangeHandle
    {
        private readonly Excel.Range _range;

        public RealRangeHandle(Excel.Range range)
        {
            _range = range;
        }

        public string NumberFormat
        {
            get => (string)_range.NumberFormat;
            set => _range.NumberFormat = value;
        }

        public CellAlignment HorizontalAlignment
        {
            get
            {
                var v = (Excel.XlHAlign)_range.HorizontalAlignment;
                if (v == Excel.XlHAlign.xlHAlignRight) return CellAlignment.Right;
                if (v == Excel.XlHAlign.xlHAlignLeft) return CellAlignment.Left;
                return CellAlignment.General;
            }
            set
            {
                _range.HorizontalAlignment =
                    value == CellAlignment.Right ? Excel.XlHAlign.xlHAlignRight :
                    value == CellAlignment.Left ? Excel.XlHAlign.xlHAlignLeft :
                    Excel.XlHAlign.xlHAlignGeneral;
            }
        }

        // External:=True mirrors VBA's rng.Address(External:=True) exactly (src/modFormatEngine.bas logging).
        public string Address => _range.Address[External: true];
    }
}
