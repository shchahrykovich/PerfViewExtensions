using System.Collections.Generic;

namespace Reporting.Implementations
{
    internal class Statistics
    {
        public Point[] Points { get; set; }
        public long MinValue { get; set; }
        public long MaxValue { get; set; }
        public double StdDeviation { get; set; }
        public long TotalCount { get; set; }
        public List<PercentileRecord> Percentiles { get; set; }

        public Statistics()
        {
            Percentiles = new List<PercentileRecord>();
        }
    }
}
