using FinanceFmtTools.Engine;
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FakeRangeHandle : IRangeHandle
    {
        public string NumberFormat { get; set; } = "General";
        public CellAlignment HorizontalAlignment { get; set; } = CellAlignment.General;
        public string Address { get; set; } = "$A$1";
    }
}
