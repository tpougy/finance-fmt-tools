namespace FinanceFmtTools.Engine
{
    // COM-free stand-in for VBA's XlHAlign (xlHAlignRight, xlHAlignLeft, xlHAlignGeneral).
    // Phase 1 has zero Excel/COM references; Phase 2/3 will map this to the real
    // XlHAlign enum when wiring real Excel COM objects.
    public enum CellAlignment
    {
        General,
        Left,
        Right
    }
}
