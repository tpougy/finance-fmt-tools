using System;
using System.Runtime.InteropServices;
using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;
using Excel = Microsoft.Office.Interop.Excel;

namespace FinanceFmtTools.ComAddin
{
    // Real IRangeHandle implementation over a live Microsoft.Office.Interop.Excel.Range,
    // proving Phase 2's unmodified interface (src/FinanceFmtTools.Engine/Abstractions/IRangeHandle.cs)
    // against the actual Excel object model. Implements IDisposable (not part of IRangeHandle's own
    // contract, so Phase 2's interface stays untouched) so AddInHost can release the wrapped Range RCW
    // promptly after use instead of waiting on GC finalization.
    public sealed class RealRangeHandle : IRangeHandle, IDisposable
    {
        private readonly Excel.Range _range;

        public RealRangeHandle(Excel.Range range)
        {
            _range = range;
        }

        public string NumberFormat
        {
            // A multi-cell selection with mixed number formats returns DBNull, not a string — mirror
            // VBA's own reads (which return Null in that case) instead of throwing InvalidCastException.
            get => _range.NumberFormat as string ?? string.Empty;
            set => _range.NumberFormat = value;
        }

        public CellAlignment HorizontalAlignment
        {
            get
            {
                // A multi-cell selection with mixed alignment returns DBNull, not an XlHAlign — fall back
                // to General rather than throwing InvalidCastException on the direct cast.
                if (!(_range.HorizontalAlignment is int || _range.HorizontalAlignment is Excel.XlHAlign))
                {
                    return CellAlignment.General;
                }

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

        public void Dispose()
        {
            if (Marshal.IsComObject(_range))
            {
                Marshal.ReleaseComObject(_range);
            }
        }
    }
}
