using System.IO;
using Reporting.Implementations;

namespace Reporting.Viewers
{
    internal class TextWriterViewer : IViewer
    {
        private readonly TextWriter _logFile;

        public TextWriterViewer(TextWriter logFile)
        {
            _logFile = logFile;
        }

        public void Show(Statistics stat)
        {
            _logFile.WriteLine();
            _logFile.WriteLine("###Stat Report###");
            
            _logFile.WriteLine("[MaxValue {0}]", stat.MaxValue);
            _logFile.WriteLine("[MinValue {0}]", stat.MinValue);
            _logFile.WriteLine("[StdDeviation {0}]", stat.StdDeviation);
            _logFile.WriteLine("[TotalCount {0}]", stat.TotalCount);

            _logFile.WriteLine("{0,12} {1,21} {2,10}", "Value", "Percentile", "TotalCount");
            foreach(var p in stat.Percentiles)
            {
                _logFile.WriteLine("{0,12:F5}  {1,20:F12} {2,10}", p.Value, p.Percentile, p.TotalCount);
            }
        }
    }
}
