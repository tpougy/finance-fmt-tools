using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FormatEngineTests
    {
        [Fact]
        public void Apply_ChaveValida_ResolveViaRegistryEAplicaNoRange()
        {
            var range = new FakeRangeHandle();
            var log = new SpyLog();

            FormatEngine.Apply(range, log, FormatKeys.Fin2D, forceAlign: false, zeroDash: true);

            Assert.Equal(AccountingFormatBuilder.Build(2, false, true), range.NumberFormat);
            Assert.Single(log.Infos);
            Assert.Empty(log.Warnings);
        }

        [Fact]
        public void Apply_ChaveDesconhecida_LogaAvisoENaoLanca()
        {
            var range = new FakeRangeHandle();
            var originalNumberFormat = range.NumberFormat;
            var log = new SpyLog();

            var ex = Record.Exception(() =>
                FormatEngine.Apply(range, log, "CHAVE_INEXISTENTE", forceAlign: false, zeroDash: false));

            Assert.Null(ex);
            Assert.Equal(originalNumberFormat, range.NumberFormat);
            Assert.Single(log.Warnings);
            Assert.Empty(log.Infos);
        }
    }
}
