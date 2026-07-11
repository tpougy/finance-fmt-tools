namespace FinanceFmtTools.Engine.Abstractions
{
    public interface ILog
    {
        void Warn(string message);
        void Info(string message);
        void Error(string message);
    }
}
