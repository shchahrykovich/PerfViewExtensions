using System.Diagnostics;

namespace Reporting.Implementations
{
    [DebuggerDisplay("{Type} - {TimeStamp}")]
    internal class Point
    {
        public double TimeStamp { get; set; }
        public PointType Type { get; set; }
    }
}
