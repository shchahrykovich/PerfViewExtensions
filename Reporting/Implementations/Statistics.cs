using System.Collections.Generic;

namespace Reporting.Implementations
{
    internal class Statistics
    {
        public Point[] Points { get; set; }
        public double MinValue { get; set; }
        public double MaxValue { get; set; }
        public double StdDeviation { get; set; }
        public long TotalCount { get; set; }
        public List<PercentileRecord> Percentiles { get; set; }
        public List<DiffRecord> Diffs { get; set; }
        public double Median { get; set; }
        public double Sum { get; set; }
        public double Average { get; set; }
        public List<Frequency> Frequencies { get; set; }

        public Statistics()
        {
            Percentiles = new List<PercentileRecord>();
            Diffs = new List<DiffRecord>();
            Frequencies = new List<Frequency>();
        }
    }
}
