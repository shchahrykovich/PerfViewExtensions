using Reporting.Implementations;

namespace Reporting.Export
{
    internal interface IExporter
    {
        void Export(Statistics stat);
    }
}
