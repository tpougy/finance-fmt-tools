// Needs "using FinanceFmtTools.Engine;" for CellAlignment — child namespaces do not
// automatically see the parent namespace's types in C#.
using FinanceFmtTools.Engine;

namespace FinanceFmtTools.Engine.Abstractions
{
    public interface IRangeHandle
    {
        string NumberFormat { get; set; }
        CellAlignment HorizontalAlignment { get; set; }   // maps to XlHAlign in the Phase 3 real impl
        string Address { get; }                            // parity with VBA's rng.Address(External:=True) logging
    }
}
