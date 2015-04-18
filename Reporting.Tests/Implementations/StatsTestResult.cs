using System;

namespace Reporting.Tests.Implementations
{
    public class StatsTestResult
    {
        public long TotalCount { get; set; }
        public long MaxValue { get; set; }
        public long MinVale { get; set; }
        public double StdDeviation { get; set; }

        private bool Equals(StatsTestResult other)
        {
            return TotalCount == other.TotalCount &&
                   MaxValue == other.MaxValue &&
                   MinVale == other.MinVale &&
                   IsDoubleEqual(other);
        }

        private bool IsDoubleEqual(StatsTestResult other)
        {
            return StdDeviation == other.StdDeviation || (Double.IsNaN(StdDeviation) && Double.IsNaN(other.StdDeviation));
        }

        public override bool Equals(object obj)
        {
            var other = obj as StatsTestResult;
            if (null != other)
            {
                return Equals(other);
            }
            return false;
        }
    }
}