using System.Collections.Generic;

namespace Reporting.DiffExtractor.Entities
{
    internal class ProcessData
    {
        public Dictionary<int, ThreadData> Threads { get; private set; }

        public ProcessData()
        {
            Threads = new Dictionary<int, ThreadData>();
        }
    }
}