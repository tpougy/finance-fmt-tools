using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FakeExcelGateway : IExcelGateway
    {
        // Test control switch: simulates a Chart/Shape being selected instead of a Range.
        public bool SelectionIsRange { get; set; } = true;
        public FakeRangeHandle SelectedRange { get; set; } = new FakeRangeHandle();

        public bool TryGetSelectedRange(out IRangeHandle range)
        {
            if (!SelectionIsRange)
            {
                range = null;
                return false;
            }

            range = SelectedRange;
            return true;
        }
    }
}
