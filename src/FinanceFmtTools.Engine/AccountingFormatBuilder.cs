using System;

namespace FinanceFmtTools.Engine
{
    // Pure port of VBA's private Function AccountingFmt (src/modFormatEngine.bas:188-222).
    // Builds the 3-section (positive;negative;zero) Excel accounting number-format
    // string. Deliberately keeps VBA's two-branch structure (general N-decimals case,
    // then an explicit decimals == 0 override) instead of unifying into one formula —
    // see 01-RESEARCH.md Common Pitfalls #4. The VBA global CFG_FORCE_ALIGN becomes the
    // forceAlign parameter here; no static/mutable state is held by this type.
    public static class AccountingFormatBuilder
    {
        public static string Build(int decimals, bool forceAlign, bool zeroDash)
        {
            // new string('0', decimals) throws ArgumentOutOfRangeException natively for
            // negative decimals — mirrors VBA's own String(-1, "0") runtime error rather
            // than silently producing a malformed format string.
            string dec = new string('0', decimals);

            string pos, neg, zer;

            if (forceAlign)
            {
                pos = " * _(#,##0." + dec + "_)_-";
                neg = " * (#,##0." + dec + ")_-";
                zer = zeroDash ? " * _(-_)_-" : pos;
            }
            else
            {
                pos = "_(#,##0." + dec + "_)_-";
                neg = "(#,##0." + dec + ")_-";
                zer = zeroDash ? "_(-_)_-" : pos;
            }

            // Nota (mirrors VBA): decimais = 0 produz "_(#,##0_)_-" (sem ponto decimal).
            // String('0', 0) retorna "" e a concatenacao "0." & "" resultaria em "0.",
            // que nao e o desejado — por isso o caso zero e tratado explicitamente.
            if (decimals == 0)
            {
                if (forceAlign)
                {
                    pos = " * _(#,##0_)_-";
                    neg = " * (#,##0)_-";
                    zer = zeroDash ? " * _(-_)_-" : pos;
                }
                else
                {
                    pos = "_(#,##0_)_-";
                    neg = "(#,##0)_-";
                    zer = zeroDash ? "_(-_)_-" : pos;
                }
            }

            return pos + ";" + neg + ";" + zer;
        }
    }
}
