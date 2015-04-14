using System.IO;
using Reporting.Implementations;

namespace Reporting.Export
{
    internal class PerfViewLogFileExporter : IExporter
    {
        private readonly TextWriter _logFile;

        public PerfViewLogFileExporter(TextWriter logFile)
        {
            _logFile = logFile;
        }

        public void Export(Statistics stat)
        {
            _logFile.WriteLine("[MaxValue {0}]", stat.MaxValue);
            _logFile.WriteLine("[MinValue {0}]", stat.MinValue);
            _logFile.WriteLine("[StdDeviation {0}]", stat.StdDeviation);
            
            _logFile.WriteLine("[90% {0}]", stat.Percentile90);
            _logFile.WriteLine("[95% {0}]", stat.Percentile95);
            _logFile.WriteLine("[99% {0}]", stat.Percentile99);

            _logFile.WriteLine("[TotalCount {0}]", stat.TotalCount);
        }
    }
}
