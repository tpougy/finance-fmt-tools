using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    // Proves FormatRegistry.TryGetFormatDef wires the 4 Fin/Integer entries by
    // delegating to AccountingFormatBuilder.Build (src/modFormatEngine.bas:93-116),
    // rather than hardcoding a duplicate format string, and that Alignment is always
    // CellAlignment.General — never Right — for every forceAlign/zeroDash
    // combination (FMT-05; the Integer/Fin-2D/4D/8D members of FMT-01/FMT-07's
    // accounting family). Also re-verifies the unknown-key guard once the Fin/
    // Integer cases are wired, to confirm `default` still only catches genuinely
    // unrecognized keys.
    public class FormatRegistryFinFamilyTests
    {
        [Fact]
        public void Integer_DelegatesToAccountingFormatBuilder()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Integer, false, true, out var def);

            Assert.True(found);
            Assert.Equal(AccountingFormatBuilder.Build(0, false, true), def.NumberFormat);
            Assert.Equal("Financeiro 0 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Numeric, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void Fin2D_DelegatesToAccountingFormatBuilder()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Fin2D, true, false, out var def);

            Assert.True(found);
            Assert.Equal(AccountingFormatBuilder.Build(2, true, false), def.NumberFormat);
            Assert.Equal("Financeiro 2 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Numeric, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void Fin4D_DelegatesToAccountingFormatBuilder()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Fin4D, false, false, out var def);

            Assert.True(found);
            Assert.Equal(AccountingFormatBuilder.Build(4, false, false), def.NumberFormat);
            Assert.Equal("Financeiro 4 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Numeric, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void Fin8D_DelegatesToAccountingFormatBuilder()
        {
            bool found = FormatRegistry.TryGetFormatDef(FormatKeys.Fin8D, true, true, out var def);

            Assert.True(found);
            Assert.Equal(AccountingFormatBuilder.Build(8, true, true), def.NumberFormat);
            Assert.Equal("Financeiro 8 casas", def.DisplayName);
            Assert.Equal(FormatCategory.Numeric, def.Category);
            Assert.Equal(CellAlignment.General, def.Alignment);
        }

        [Fact]
        public void FinFamily_NeverUsesRightAlignment_AcrossEveryForceAlignZeroDashCombination()
        {
            var keys = new[] { FormatKeys.Integer, FormatKeys.Fin2D, FormatKeys.Fin4D, FormatKeys.Fin8D };
            var boolValues = new[] { false, true };

            foreach (var key in keys)
            {
                foreach (var forceAlign in boolValues)
                {
                    foreach (var zeroDash in boolValues)
                    {
                        bool found = FormatRegistry.TryGetFormatDef(key, forceAlign, zeroDash, out var def);

                        Assert.True(found);
                        Assert.Equal(FormatCategory.Numeric, def.Category);
                        Assert.Equal(CellAlignment.General, def.Alignment);
                    }
                }
            }
        }

        [Fact]
        public void UnknownKey_StillReturnsFalse_AfterFinFamilyIsWired()
        {
            bool found = FormatRegistry.TryGetFormatDef("NOT_A_REAL_KEY", true, true, out var def);

            Assert.False(found);
            Assert.Null(def);
        }
    }
}
