using System.Collections.Generic;
using FinanceFmtTools.Engine.Abstractions;

namespace FinanceFmtTools.Engine.Tests
{
    public sealed class SpyLog : ILog
    {
        public List<string> Warnings { get; } = new List<string>();
        public List<string> Infos { get; } = new List<string>();
        public List<string> Errors { get; } = new List<string>();

        public void Warn(string message) => Warnings.Add(message);
        public void Info(string message) => Infos.Add(message);
        public void Error(string message) => Errors.Add(message);
    }
}
