using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine
{
    // Port of VBA's ApplyFormat/ApplyFormatToSelection (src/modFormatEngine.bas:24-77).
    // Static + parameterized, consistent with FormatRegistry.cs/AccountingFormatBuilder.cs.
    public static class FormatEngine
    {
        public static void Apply(IRangeHandle range, ILog log, string formatKey, bool forceAlign, bool zeroDash)
        {
            if (!FormatRegistry.TryGetFormatDef(formatKey, forceAlign, zeroDash, out FormatDef def))
            {
                log.Warn("FormatEngine.Apply: chave de formato desconhecida '" + formatKey + "'.");
                return;
            }

            range.NumberFormat = def.NumberFormat;
            if (def.Alignment != CellAlignment.General)
            {
                range.HorizontalAlignment = def.Alignment;
            }

            log.Info("FormatEngine.Apply: aplicado '" + def.DisplayName + "' em " + range.Address + ".");
        }
    }
}
