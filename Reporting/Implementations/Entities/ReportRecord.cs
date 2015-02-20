using System;

namespace Reporting.Implementations.Entities
{
    internal class ReportRecord
    {
        public readonly DateTime TimeStamp;
        public int Counter;

        public ReportRecord(DateTime timeStamp)
        {
            TimeStamp = timeStamp;
        }
    }
}
