using System;

namespace Reporting.DiffExtractor.Entities
{
    internal class DataRecord
    {
        public readonly DateTime TimeStamp;
        public readonly Guid ActivityId;

        public DataRecord(DateTime timeStamp, Guid activityId)
        {
            TimeStamp = timeStamp;
            ActivityId = activityId;
        }
    }
}