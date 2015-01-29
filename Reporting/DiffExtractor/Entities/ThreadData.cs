using System.Collections.Generic;

namespace Reporting.DiffExtractor.Entities
{
    internal class ThreadData
    {
        public List<DataRecord> Records { get; set; }

        public ThreadData()
        {
            Records = new List<DataRecord>();
        }
    }
}