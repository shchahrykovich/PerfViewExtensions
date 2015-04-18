using System.Collections.Generic;
using System.Linq;
using HdrHistogram.NET;
using Microsoft.Diagnostics.Tracing;
using Reporting.Viewers;

namespace Reporting.Implementations
{
    internal class Parser
    {
        private readonly IViewer _exporter;
        private readonly List<Point> _points;

        public Parser(IViewer exporter)
        {
            _exporter = exporter;
            _points = new List<Point>();
        }

        public void Parse(IEnumerable<TraceEvent> events, ParserArguments arguments)
        {
            FindEvents(events, arguments.ProviderName, arguments.StartEvent, arguments.StopEvent);
            Statistics s = CalculateStatistics();

            _exporter.Show(s);
        }

        private Statistics CalculateStatistics()
        {
            Statistics result = new Statistics();
            result.Diffs = GetDiffs(result);

            HistogramData histogramData = CreateHistogram(result.Diffs);
            result.MaxValue = histogramData.getMaxValue();
            result.MinValue = histogramData.getMinValue();
            result.StdDeviation = histogramData.getStdDeviation();
            result.TotalCount = histogramData.getTotalCount();

            foreach (var percentile in histogramData.percentiles(5))
            {
                PercentileRecord p = new PercentileRecord
                {
                    Value = percentile.getValueIteratedTo(),
                    Percentile = percentile.getPercentileLevelIteratedTo(),
                    TotalCount = percentile.getTotalCountToThisValue()
                };

                result.Percentiles.Add(p);
            }

            return result;
        }

        private List<double> GetDiffs(Statistics result)
        {
            List<double> diffs = new List<double>();

            result.Points = _points.OrderBy(p => p.TimeStamp).ToArray();
            int firstStatrtIndex = GetStartEventIndex(result.Points);
            for (int i = firstStatrtIndex; i + 1 < result.Points.Length; )
            {
                var start = result.Points[i];
                var stop = result.Points[i + 1];

                if (start.Type == PointType.Start && stop.Type == PointType.Stop)
                {
                    double diff = stop.TimeStamp - start.TimeStamp;
                    diffs.Add(diff);
                    i += 2;
                }
                else
                {
                    i += 1;
                }
            }

            return diffs;
        }

        private HistogramData CreateHistogram(List<double> diffs)
        {
            Histogram h = new Histogram(3600000000000L, 3);

            foreach (var diff in diffs)
            {
                h.recordValue((long)diff);
            }

            HistogramData histogramData = h.getHistogramData();
            return histogramData;
        }

        private int GetStartEventIndex(Point[] points)
        {
            int i = 0;
            while (i < points.Length)
            {
                var point = points[i];
                if (point.Type == PointType.Start)
                {
                    break;
                }

                i++;
            }

            return i;
        }

        private void FindEvents(IEnumerable<TraceEvent> events, string providerName, string startEvent, string stopEvent)
        {
            foreach (TraceEvent traceEvent in events)
            {
                if (traceEvent.ProviderName == providerName)
                {
                    if (traceEvent.EventName == startEvent)
                    {
                        AddEvent(traceEvent, PointType.Start);
                    }
                    else if (traceEvent.EventName == stopEvent)
                    {
                        AddEvent(traceEvent, PointType.Stop);
                    }
                }
            }
        }

        private void AddEvent(TraceEvent data, PointType type)
        {
            Point point = new Point
            {
                TimeStamp = data.TimeStampRelativeMSec,
                Type = type
            };
            _points.Add(point);
        }
    }
}
