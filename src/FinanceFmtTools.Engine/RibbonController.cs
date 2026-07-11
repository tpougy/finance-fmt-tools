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
            if (config == null) throw new ArgumentNullException(nameof(config));
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
            if (resourceName == null)
            {
                // Unlike FMT-06's invalid-selection guard, a missing embedded ribbon resource is an
                // unrecoverable build/packaging defect (e.g. csproj EmbeddedResource entry removed or
                // renamed) — fail loudly rather than silently returning an empty ribbon XML string
                // that would render no Finance Fmt tab with zero error trail.
                throw new InvalidOperationException("Embedded resource 'customUI14.xml' not found in assembly.");
            }

            using (Stream stream = asm.GetManifestResourceStream(resourceName))
            using (StreamReader reader = new StreamReader(stream))
            {
                return reader.ReadToEnd();
            }
        }
    }
}
