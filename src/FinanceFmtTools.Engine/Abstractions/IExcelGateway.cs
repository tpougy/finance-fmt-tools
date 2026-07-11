namespace FinanceFmtTools.Engine.Abstractions
{
    public interface IExcelGateway
    {
        // false => current selection is not a Range (e.g. Chart/Shape selected, or nothing selected).
        // Mirrors VBA's SafeSelection()'s TypeName(Selection) <> "Range" check (src/modUtils.bas:74-89),
        // collapsed into one no-throw boolean query instead of a Nothing-returning function.
        bool TryGetSelectedRange(out IRangeHandle range);
    }
}
