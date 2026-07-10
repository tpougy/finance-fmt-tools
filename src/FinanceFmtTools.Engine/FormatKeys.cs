namespace FinanceFmtTools.Engine
{
    // Format-key string constants — mirrors src/modConfig.bas:19-29 (FMT_* Public Const values).
    // Values are copied verbatim from the VBA source; do not change without updating the VBA source too.
    public static class FormatKeys
    {
        public const string Integer = "INTEGER";
        public const string Fin2D = "FIN_2D";
        public const string Fin4D = "FIN_4D";
        public const string Fin8D = "FIN_8D";
        public const string Pct4D = "PCT_4D";
        public const string Pct2D = "PCT_2D";
        public const string SpreadBps = "SPREAD_BPS";
        public const string DateIso = "DATE_ISO";
        public const string DateBr = "DATE_BR";
        public const string DateBrLong = "DATE_BR_LONG";
        public const string Text = "TEXT";
    }
}
