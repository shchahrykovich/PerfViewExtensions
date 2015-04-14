namespace Reporting.Implementations
{
    internal class Statistics
    {
        public Point[] Points { get; set; }
        public long MinValue { get; set; }
        public long MaxValue { get; set; }
        public double StdDeviation { get; set; }
        public long TotalCount { get; set; }
        public long Percentile90 { get; set; }
        public long Percentile95 { get; set; }
        public long Percentile99 { get; set; }
    }
}
