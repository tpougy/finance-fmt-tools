// Kept explicit for readability/clarity, even though C#'s namespace lookup already
// resolves CellAlignment (declared in the enclosing FinanceFmtTools.Engine namespace)
// without this using directive.
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
