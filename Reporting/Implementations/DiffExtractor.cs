using System.Collections.Generic;
using System.Linq;
using Microsoft.Diagnostics.Tracing;
using Reporting.Export;
using Reporting.Implementations.Entities;

namespace Reporting.Implementations
{
    internal class DiffExtractor
    {
        private readonly IExporter _exporter;
        private readonly Dictionary<int, ProcessData> _data = new Dictionary<int, ProcessData>();

        public DiffExtractor(IExporter exporter)
        {
            _exporter = exporter;
        }

        public void ExtractEvents(IEnumerable<TraceEvent> events, DiffExtractorArguments arguments)
        {
            FindEvents(events, arguments.ProviderName, arguments.StartEvent, arguments.StopEvent);
            var diffs = FindDiffs();

            _exporter.Export(arguments.OutputReport, diffs);
        }

        private void FindEvents(IEnumerable<TraceEvent> events, string providerName, string startEvent, string stopEvent)
        {
            foreach (TraceEvent traceEvent in events)
            {
                if (traceEvent.ProviderName == providerName)
                {
                    if (traceEvent.EventName == startEvent || traceEvent.EventName == stopEvent)
                    {
                        AddEvent(traceEvent);
                    }
                }
            }
        }

        private void AddEvent(TraceEvent data)
        {
            if (!_data.ContainsKey(data.ProcessID))
            {
                _data[data.ProcessID] = new ProcessData();
            }

            var process = _data[data.ProcessID];
            if (!process.Threads.ContainsKey(data.ThreadID))
            {
                process.Threads[data.ThreadID] = new ThreadData();
            }

            var thread = process.Threads[data.ThreadID];
            thread.Records.Add(new DataRecord(data.TimeStamp, data.ActivityID));
        }

        private DiffCollection FindDiffs()
        {
            DiffCollection result = new DiffCollection();
            foreach (var process in _data.Values)
            {
                foreach (var thread in process.Threads.Values)
                {
                    var records = thread.Records.OrderBy(r => r.ActivityId).ThenBy(r => r.TimeStamp).ToArray();
                    var length = records.Length;
                    length = length % 2 == 0 ? length : length - 1;
                    for (int i = 0; i < length; i += 2)
                    {
                        var start = records[i];
                        var finish = records[i + 1];

                        result.Diffs.Add(new Diff(start, finish));
                    }
                }
            }
            return result;
        }
    }
}
