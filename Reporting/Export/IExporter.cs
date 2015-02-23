using Reporting.Implementations.Entities;

namespace Reporting.Export
{
    internal interface IExporter
    {
        void Export(string outputFileName, DiffCollection diffs);
    }
}
