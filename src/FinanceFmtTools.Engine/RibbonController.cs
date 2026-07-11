using System;

namespace FinanceFmtTools.Engine
{
    public sealed class RibbonController
    {
        public RibbonSessionConfig Config { get; }

        public RibbonController() : this(new RibbonSessionConfig()) { }

        public RibbonController(RibbonSessionConfig config)
        {
            if (config == null) throw new ArgumentNullException("config");
            Config = config;
        }
    }
}
