namespace FinanceFmtTools.Engine
{
    // Named constants for every combinable piece of the 3-section (positive;negative;zero)
    // Excel accounting number-format string produced by AccountingFormatBuilder.Build.
    // Extracted so a future tweak to a single token (e.g. swapping the hyphen-padding
    // token for a parenthesis-padding token) touches one obvious constant instead of
    // requiring surgery on string-concatenation expressions. Plain `const` fields only —
    // no records, no init-only properties — compiles under net48 same as FormatDef.cs.
    public static class FormatTokens
    {
        // The fill-character token that forces right alignment when CFG_FORCE_ALIGN
        // (forceAlign) is on.
        public const string FillAlignPrefix = " * ";

        // Padding the width of an opening parenthesis without printing one; used to open
        // the positive/zero sections so their digits line up under the negative section's
        // real "(".
        public const string OpenParenPad = "_(";

        // Padding the width of a closing parenthesis without printing one; closes the
        // positive/zero sections to match the negative section's real ")".
        public const string CloseParenPad = "_)";

        // Literal opening parenthesis wrapping the negative section.
        public const string OpenParen = "(";

        // Literal closing parenthesis wrapping the negative section.
        public const string CloseParen = ")";

        // Padding the width of a hyphen without printing one; appended at the end of
        // every section.
        public const string PadHyphen = "_-";

        // The thousands-grouped digit pattern shared by the decimals==0 and decimals>0
        // cases (the latter appends DecimalPoint + a run of zeros).
        public const string DigitsBase = "#,##0";

        // Decimal point separator, only used when decimals > 0.
        public const string DecimalPoint = ".";

        // The literal character shown in the zero section when CFG_ZERO_DASH (zeroDash)
        // is on.
        public const string ZeroDashLiteral = "-";

        // The fractional-digit character repeated `decimals` times via
        // new string(FormatTokens.ZeroDigit, decimals).
        public const char ZeroDigit = '0';

        // Joins the positive/negative/zero sections into the final 3-section format string.
        public const string SectionSeparator = ";";
    }
}
