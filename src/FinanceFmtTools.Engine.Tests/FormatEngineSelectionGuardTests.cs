using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class FormatEngineSelectionGuardTests
    {
        [Fact]
        public void ApplyToSelection_SelecaoNaoEhRange_LogaAvisoENaoLanca()
        {
            var gateway = new FakeExcelGateway { SelectionIsRange = false };
            var log = new SpyLog();

            var ex = Record.Exception(() =>
                FormatEngine.ApplyToSelection(gateway, log, FormatKeys.Fin2D, forceAlign: false, zeroDash: true));

            Assert.Null(ex);
            Assert.Single(log.Warnings);
            Assert.Empty(log.Infos);
        }

        [Fact]
        public void ApplyToSelection_SelecaoValida_DelegaParaApplyEAplicaFormato()
        {
            var gateway = new FakeExcelGateway { SelectionIsRange = true };
            var log = new SpyLog();

            FormatEngine.ApplyToSelection(gateway, log, FormatKeys.DateIso, forceAlign: false, zeroDash: false);

            Assert.Equal("yyyy-mm-dd;@", gateway.SelectedRange.NumberFormat);
            Assert.Single(log.Infos);
        }
    }
}
