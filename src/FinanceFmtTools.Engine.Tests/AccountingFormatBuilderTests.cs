using System;
using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    // Proves AccountingFormatBuilder.Build ports VBA's private AccountingFmt helper
    // (src/modFormatEngine.bas:188-222) byte-for-byte across all 16 combinations of
    // decimals x forceAlign x zeroDash (FMT-01, FMT-07). Expected values were derived
    // by removing the deleted trailing hyphen-padding token's "_-" substrings from the
    // prior verified expected strings (originally copied from 01-RESEARCH.md's
    // pre-computed table).
    public class AccountingFormatBuilderTests
    {
        [Theory]
        [InlineData(0, false, false, "_(#,##0_);(#,##0);_(#,##0_)")]
        [InlineData(0, false, true, "_(#,##0_);(#,##0);_(-_)")]
        [InlineData(0, true, false, " * _(#,##0_); * (#,##0); * _(#,##0_)")]
        [InlineData(0, true, true, " * _(#,##0_); * (#,##0); * _(-_)")]
        [InlineData(2, false, false, "_(#,##0.00_);(#,##0.00);_(#,##0.00_)")]
        [InlineData(2, false, true, "_(#,##0.00_);(#,##0.00);_(-_)")]
        [InlineData(2, true, false, " * _(#,##0.00_); * (#,##0.00); * _(#,##0.00_)")]
        [InlineData(2, true, true, " * _(#,##0.00_); * (#,##0.00); * _(-_)")]
        [InlineData(4, false, false, "_(#,##0.0000_);(#,##0.0000);_(#,##0.0000_)")]
        [InlineData(4, false, true, "_(#,##0.0000_);(#,##0.0000);_(-_)")]
        [InlineData(4, true, false, " * _(#,##0.0000_); * (#,##0.0000); * _(#,##0.0000_)")]
        [InlineData(4, true, true, " * _(#,##0.0000_); * (#,##0.0000); * _(-_)")]
        [InlineData(8, false, false, "_(#,##0.00000000_);(#,##0.00000000);_(#,##0.00000000_)")]
        [InlineData(8, false, true, "_(#,##0.00000000_);(#,##0.00000000);_(-_)")]
        [InlineData(8, true, false, " * _(#,##0.00000000_); * (#,##0.00000000); * _(#,##0.00000000_)")]
        [InlineData(8, true, true, " * _(#,##0.00000000_); * (#,##0.00000000); * _(-_)")]
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
