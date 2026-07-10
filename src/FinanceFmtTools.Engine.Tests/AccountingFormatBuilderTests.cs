using System;
using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    // Proves AccountingFormatBuilder.Build ports VBA's private AccountingFmt helper
    // (src/modFormatEngine.bas:188-222) byte-for-byte across all 16 combinations of
    // decimals x forceAlign x zeroDash (FMT-01, FMT-07). Expected values copied
    // verbatim from 01-RESEARCH.md's pre-computed table — not re-derived by hand.
    public class AccountingFormatBuilderTests
    {
        [Theory]
        [InlineData(0, false, false, "_(#,##0_)_-;(#,##0)_-;_(#,##0_)_-")]
        [InlineData(0, false, true, "_(#,##0_)_-;(#,##0)_-;_(-_)_-")]
        [InlineData(0, true, false, " * _(#,##0_)_-; * (#,##0)_-; * _(#,##0_)_-")]
        [InlineData(0, true, true, " * _(#,##0_)_-; * (#,##0)_-; * _(-_)_-")]
        [InlineData(2, false, false, "_(#,##0.00_)_-;(#,##0.00)_-;_(#,##0.00_)_-")]
        [InlineData(2, false, true, "_(#,##0.00_)_-;(#,##0.00)_-;_(-_)_-")]
        [InlineData(2, true, false, " * _(#,##0.00_)_-; * (#,##0.00)_-; * _(#,##0.00_)_-")]
        [InlineData(2, true, true, " * _(#,##0.00_)_-; * (#,##0.00)_-; * _(-_)_-")]
        [InlineData(4, false, false, "_(#,##0.0000_)_-;(#,##0.0000)_-;_(#,##0.0000_)_-")]
        [InlineData(4, false, true, "_(#,##0.0000_)_-;(#,##0.0000)_-;_(-_)_-")]
        [InlineData(4, true, false, " * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(#,##0.0000_)_-")]
        [InlineData(4, true, true, " * _(#,##0.0000_)_-; * (#,##0.0000)_-; * _(-_)_-")]
        [InlineData(8, false, false, "_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(#,##0.00000000_)_-")]
        [InlineData(8, false, true, "_(#,##0.00000000_)_-;(#,##0.00000000)_-;_(-_)_-")]
        [InlineData(8, true, false, " * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(#,##0.00000000_)_-")]
        [InlineData(8, true, true, " * _(#,##0.00000000_)_-; * (#,##0.00000000)_-; * _(-_)_-")]
        public void Build_MatchesVbaAlgorithm(int decimals, bool forceAlign, bool zeroDash, string expected)
        {
            Assert.Equal(expected, AccountingFormatBuilder.Build(decimals, forceAlign, zeroDash));
        }

        [Fact]
        public void Build_NegativeDecimals_Throws()
        {
            Assert.Throws<ArgumentOutOfRangeException>(() => AccountingFormatBuilder.Build(-1, false, false));
        }
    }
}
