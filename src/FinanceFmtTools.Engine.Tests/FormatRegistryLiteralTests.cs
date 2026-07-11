using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    // Proves FormatRegistry.TryGetFormatDef resolves the 7 literal (non-Fin) registry
    // entries to their exact VBA-sourced NumberFormat/DisplayName/Category values
    // (src/modFormatEngine.bas:118-161), plus the no-throw contract for an unrecognized
    // key (FMT-02, FMT-03, FMT-04, FMT-05).
    public class FormatRegistryLiteralTests
    {
        [Fact]
        public void Pct4D_ResolvesToExactVbaFormat()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Pct4D, false, false, out var def);

            Assert.True(found);
            Assert.Equal("0.0000%", def.NumberFormat);
            Assert.Equal("% 4 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Percent, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void Pct2D_ResolvesToExactVbaFormat()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Pct2D, false, false, out var def);

            Assert.True(found);
            Assert.Equal("0.00%", def.NumberFormat);
            Assert.Equal("% 2 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Percent, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void SpreadBps_ResolvesToExactVbaFormat_WithDecodedQuoteEscape()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.SpreadBps, false, false, out var def);

            Assert.True(found);
            Assert.Equal("#,##0.0\" bps\"", def.NumberFormat);
            Assert.Equal("Spread (bps)", def.DisplayName);
            Assert.Equal(FormatCategory.Numeric, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void DateIso_ResolvesToExactVbaFormat()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.DateIso, false, false, out var def);

            Assert.True(found);
            Assert.Equal("yyyy-mm-dd;@", def.NumberFormat);
            Assert.Equal("Data ISO", def.DisplayName);
            Assert.Equal(FormatCategory.Date, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void DateBr_ResolvesToExactVbaFormat()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.DateBr, false, false, out var def);

            Assert.True(found);
            Assert.Equal("[$-pt-BR]dd/mm/yyyy;@", def.NumberFormat);
            Assert.Equal("Data BR", def.DisplayName);
            Assert.Equal(FormatCategory.Date, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void DateBrLong_ResolvesToAbbreviatedMonthFormat_NotSpelledOutTooltipVersion()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.DateBrLong, false, false, out var def);

            Assert.True(found);
            Assert.Equal("[$-pt-BR]dd/mmm/yyyy;@", def.NumberFormat);
            Assert.Equal("Data BR Longa", def.DisplayName);
            Assert.Equal(FormatCategory.Date, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void Text_ResolvesToExactVbaFormat()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Text, false, false, out var def);

            Assert.True(found);
            Assert.Equal("@", def.NumberFormat);
            Assert.Equal("Texto", def.DisplayName);
            Assert.Equal(FormatCategory.Text, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void UnknownKey_ReturnsFalse_AndDoesNotThrow()
        {
            bool found = FormatRegistry.TryGetFormatDef("BOGUS_KEY", false, false, out var def);

            Assert.False(found);
            Assert.Null(def);
        }
    }
}
