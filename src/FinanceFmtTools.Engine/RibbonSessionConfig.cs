namespace FinanceFmtTools.Engine
{
    // Authoritative defaults per REQUIREMENTS.md RIB-02/RIB-03 and 02-CONTEXT.md — deliberately
    // NOT copied from either of src/modConfig.bas's or src/modUtils.bas's contradictory VBA
    // defaults (see 02-RESEARCH.md Pitfall 1). No persistence anywhere in this class — that is
    // explicitly out of scope for this migration.
    public sealed class RibbonSessionConfig
    {
        public bool ForceAlign { get; set; } = false;  // "Alinhar à direita" starts OFF (RIB-02)
        public bool ZeroDash { get; set; } = true;      // "Zero contábil" starts ON (RIB-03)
    }
}
