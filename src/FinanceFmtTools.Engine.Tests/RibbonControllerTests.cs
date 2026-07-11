using FinanceFmtTools.Engine;
using Xunit;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class RibbonControllerTests
    {
        [Fact]
        public void Config_ValoresPadrao_ForceAlignFalseEZeroDashTrue()
        {
            var controller = new RibbonController();

            Assert.False(controller.Config.ForceAlign);
            Assert.True(controller.Config.ZeroDash);
        }

        [Fact]
        public void Config_Mutavel_RefleteAlteracoesDeCheckbox()
        {
            var controller = new RibbonController();

            controller.Config.ForceAlign = true;
            controller.Config.ZeroDash = false;

            Assert.True(controller.Config.ForceAlign);
            Assert.False(controller.Config.ZeroDash);
        }

        [Fact]
        public void ConstrutorComConfigInjetada_UsaValoresFornecidos()
        {
            var config = new RibbonSessionConfig { ForceAlign = true, ZeroDash = false };

            var controller = new RibbonController(config);

            Assert.True(controller.Config.ForceAlign);
            Assert.False(controller.Config.ZeroDash);
        }
    }
}
