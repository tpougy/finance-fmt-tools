namespace FinanceFmtTools.Engine
{
    // Immutable value type describing a resolved format — mirrors VBA's
    // Public Type FormatDef (src/modFormatEngine.bas:10-16), typed idiomatically.
    // Deliberately a plain class with a constructor and get-only properties —
    // C# 9's alternate immutable-type syntax fails to compile on net48 with CS0518
    // (see 01-RESEARCH.md Anti-Patterns).
    public sealed class FormatDef
    {
        public string Key { get; }
        public string DisplayName { get; }
        public string NumberFormat { get; }
        public FormatCategory Category { get; }
        public CellAlignment Alignment { get; }

        public FormatDef(string key, string displayName, string numberFormat, FormatCategory category, CellAlignment alignment)
        {
            Key = key;
            DisplayName = displayName;
            NumberFormat = numberFormat;
            Category = category;
            Alignment = alignment;
        }
    }
}
