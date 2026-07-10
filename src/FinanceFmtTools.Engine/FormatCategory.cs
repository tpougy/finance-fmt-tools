namespace FinanceFmtTools.Engine
{
    // Type-safe enum mirroring the 4 distinct string values VBA's FormatDef.Category
    // ever holds ("numeric", "percent", "date", "text") — see src/modFormatEngine.bas:14.
    public enum FormatCategory
    {
        Numeric,
        Percent,
        Date,
        Text
    }
}
