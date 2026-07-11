namespace FinanceFmtTools.Engine
{
    // Port of VBA's GetFormatDef registry (src/modFormatEngine.bas:81-170) — the
    // key -> FormatDef lookup that every Ribbon button click resolves through.
    // The VBA source never assigns f.Alignment in any Case branch, so every entry
    // (including the Fin family, wired in a later task) carries the General
    // alignment value below, never the right-aligned one. The Fin family's visual
    // right-alignment comes entirely from the " * " fill-character token already
    // embedded by AccountingFormatBuilder.Build, not from a HorizontalAlignment
    // COM write.
    //
    // Interim state: only the 7 literal (non-Fin) entries are wired so far.
    // FormatKeys.Integer/Fin2D/Fin4D/Fin8D currently fall through to `default`
    // and return false — the next task adds explicit cases for them.
    public static class FormatRegistry
    {
        public static bool TryGetFormatDef(string key, bool forceAlign, bool zeroDash, out FormatDef def)
        {
            switch (key)
            {
                case FormatKeys.Pct4D:
                    def = new FormatDef(key, "% 4 casas", "0.0000%", FormatCategory.Percent, CellAlignment.General);
                    return true;

                case FormatKeys.Pct2D:
                    def = new FormatDef(key, "% 2 casas", "0.00%", FormatCategory.Percent, CellAlignment.General);
                    return true;

                case FormatKeys.SpreadBps:
                    def = new FormatDef(key, "Spread (bps)", "#,##0.0\" bps\"", FormatCategory.Numeric, CellAlignment.General);
                    return true;

                case FormatKeys.DateIso:
                    def = new FormatDef(key, "Data ISO", "yyyy-mm-dd;@", FormatCategory.Date, CellAlignment.General);
                    return true;

                case FormatKeys.DateBr:
                    def = new FormatDef(key, "Data BR", "[$-pt-BR]dd/mm/yyyy;@", FormatCategory.Date, CellAlignment.General);
                    return true;

                case FormatKeys.DateBrLong:
                    def = new FormatDef(key, "Data BR Longa", "[$-pt-BR]dd/mmm/yyyy;@", FormatCategory.Date, CellAlignment.General);
                    return true;

                case FormatKeys.Text:
                    def = new FormatDef(key, "Texto", "@", FormatCategory.Text, CellAlignment.General);
                    return true;

                default:
                    def = null;
                    return false;
            }
        }
    }
}
