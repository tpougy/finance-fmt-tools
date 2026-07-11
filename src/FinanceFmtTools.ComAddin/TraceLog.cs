using System.Diagnostics;
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.ComAddin
{
    // Real ILog implementation via System.Diagnostics.Trace — zero extra dependency,
    // proving Phase 2's unmodified interface (src/FinanceFmtTools.Engine/Abstractions/ILog.cs).
    public sealed class TraceLog : ILog
    {
        public void Warn(string message) => Trace.TraceWarning(message);

        public void Info(string message) => Trace.TraceInformation(message);

        public void Error(string message) => Trace.TraceError(message);
    }
}
