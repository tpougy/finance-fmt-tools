using System;
using System.IO;
using System.Reflection;

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

        // Resolves the embedded src/customUI14.xml resource by suffix match rather than a
        // hardcoded full logical name — SDK-computed resource names drift silently with
        // RootNamespace/folder changes (02-RESEARCH.md Pitfall 3).
        public string GetCustomUiXml()
        {
            Assembly asm = typeof(RibbonController).Assembly;
            string resourceName = null;
            foreach (string name in asm.GetManifestResourceNames())
            {
                if (name.EndsWith("customUI14.xml", StringComparison.OrdinalIgnoreCase))
                {
                    resourceName = name;
                    break;
                }
            }
            if (resourceName == null) return string.Empty;

            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
