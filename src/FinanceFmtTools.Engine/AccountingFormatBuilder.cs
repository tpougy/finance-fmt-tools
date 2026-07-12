using System;

namespace FinanceFmtTools.Engine
{
    // Pure port of VBA's private Function AccountingFmt (src/modFormatEngine.bas:188-222).
    // Builds the 3-section (positive;negative;zero) Excel accounting number-format
    // string. Deliberately keeps VBA's two-branch structure (general N-decimals case,
    // then an explicit decimals == 0 override) instead of unifying into one formula —
    // see 01-RESEARCH.md Common Pitfalls #4. The VBA global CFG_FORCE_ALIGN becomes the
    // forceAlign parameter here; no static/mutable state is held by this type.
    // Every literal token composed below (fill-align prefix, parenthesis padding,
    // literal parentheses, hyphen padding, digit pattern, decimal point, zero-dash
    // literal, section separator) is sourced from FormatTokens — see that file for the
    // Excel number-format mini-language meaning of each piece.
    public static class AccountingFormatBuilder
    {
        public static string Build(int decimals, bool forceAlign, bool zeroDash)
        {
            // new string(FormatTokens.ZeroDigit, decimals) throws ArgumentOutOfRangeException
            // natively for negative decimals — mirrors VBA's own String(-1, "0") runtime
            // error rather than silently producing a malformed format string.
            string dec = new string(FormatTokens.ZeroDigit, decimals);

            string pos, neg, zer;

            if (forceAlign)
            {
                pos = FormatTokens.FillAlignPrefix + FormatTokens.OpenParenPad + FormatTokens.DigitsBase + FormatTokens.DecimalPoint + dec + FormatTokens.CloseParenPad + FormatTokens.PadHyphen;
                neg = FormatTokens.FillAlignPrefix + FormatTokens.OpenParen + FormatTokens.DigitsBase + FormatTokens.DecimalPoint + dec + FormatTokens.CloseParen + FormatTokens.PadHyphen;
                zer = zeroDash ? FormatTokens.FillAlignPrefix + FormatTokens.OpenParenPad + FormatTokens.ZeroDashLiteral + FormatTokens.CloseParenPad + FormatTokens.PadHyphen : pos;
            }
            else
            {
                pos = FormatTokens.OpenParenPad + FormatTokens.DigitsBase + FormatTokens.DecimalPoint + dec + FormatTokens.CloseParenPad + FormatTokens.PadHyphen;
                neg = FormatTokens.OpenParen + FormatTokens.DigitsBase + FormatTokens.DecimalPoint + dec + FormatTokens.CloseParen + FormatTokens.PadHyphen;
                zer = zeroDash ? FormatTokens.OpenParenPad + FormatTokens.ZeroDashLiteral + FormatTokens.CloseParenPad + FormatTokens.PadHyphen : pos;
            }

            // Nota (mirrors VBA): decimais = 0 produz "_(#,##0_)_-" (sem ponto decimal).
            // String('0', 0) retorna "" e a concatenacao "0." & "" resultaria em "0.",
            // que nao e o desejado — por isso o caso zero e tratado explicitamente,
            // using FormatTokens.DigitsBase alone (no DecimalPoint/dec) as the digit pattern.
            if (decimals == 0)
            {
                if (forceAlign)
                {
                    pos = FormatTokens.FillAlignPrefix + FormatTokens.OpenParenPad + FormatTokens.DigitsBase + FormatTokens.CloseParenPad + FormatTokens.PadHyphen;
                    neg = FormatTokens.FillAlignPrefix + FormatTokens.OpenParen + FormatTokens.DigitsBase + FormatTokens.CloseParen + FormatTokens.PadHyphen;
                    zer = zeroDash ? FormatTokens.FillAlignPrefix + FormatTokens.OpenParenPad + FormatTokens.ZeroDashLiteral + FormatTokens.CloseParenPad + FormatTokens.PadHyphen : pos;
                }
                else
                {
                    pos = FormatTokens.OpenParenPad + FormatTokens.DigitsBase + FormatTokens.CloseParenPad + FormatTokens.PadHyphen;
                    neg = FormatTokens.OpenParen + FormatTokens.DigitsBase + FormatTokens.CloseParen + FormatTokens.PadHyphen;
                    zer = zeroDash ? FormatTokens.OpenParenPad + FormatTokens.ZeroDashLiteral + FormatTokens.CloseParenPad + FormatTokens.PadHyphen : pos;
                }
            }

            return pos + FormatTokens.SectionSeparator + neg + FormatTokens.SectionSeparator + zer;
        }
    }
}
