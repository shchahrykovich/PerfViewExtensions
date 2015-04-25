using System;
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
            if (result.Diffs.Any())
            {                
                result.MaxValue = result.Diffs.Max(d => d.Value);
                result.MinValue = result.Diffs.Min(d => d.Value);
                result.Median = CalculateMedian(result.Diffs.Select(d => d.Value));
                result.StdDeviation = CalculateStdDeviation(result.Diffs.Select(d => d.Value));
                result.TotalCount = result.Diffs.Count;
                result.Sum = result.Diffs.Sum(d => d.Value);
                result.Average = result.Diffs.Average(d => d.Value);
                result.Frequencies = GetFrequencies(result.Diffs);

                HistogramData histogramData = CreateHistogram(result.Diffs);
                foreach (var percentile in histogramData.percentiles(5))
                {
                    PercentileRecord p = new PercentileRecord
                    {
                        Value = percentile.getValueIteratedTo(),
                        Percentile = percentile.getPercentileLevelIteratedTo(),
                        TotalCount = percentile.getTotalCountToThisValue(),
                        Count = percentile.getCountAddedInThisIterationStep()
                    };

                    result.Percentiles.Add(p);
                }
            }

            return result;
        }

        private List<Frequency> GetFrequencies(List<DiffRecord> diffs)
        {
            List<Frequency> result = new List<Frequency>();

            double sum = diffs.Sum(d => d.Value);
            int totalCount = diffs.Count;

            double totalValueCumulative = 0;
            int totalCountCumulative = 0;
            foreach (var diff in diffs.GroupBy(d => d.Value).OrderBy(d => d.Key))
            {
                var value = diff.Key;
                var count = diff.Count();

                totalValueCumulative += value * count;
                totalCountCumulative += count;
                
                result.Add(new Frequency
                {
                    Value = value,
                    Count = count,
                    TotalValue = totalValueCumulative,
                    TotalValuePercent = 100 * (totalValueCumulative / sum),
                    TotalCount = totalCountCumulative,
                    TotalCountPercent = 100 * ((double)totalCountCumulative / totalCount),
                });
            }
            return result;
        }

        //http://en.wikipedia.org/wiki/Standard_deviation
        private double CalculateStdDeviation(IEnumerable<double> values)
        {
            double[] sorted = values.OrderBy(v => v).ToArray();
            
            double average = sorted.Average();
            double sum = sorted.Sum(v => Math.Pow(v - average, 2));
            
            double result = Math.Sqrt(sum / sorted.Length);
            return result;
        }

        private double CalculateMedian(IEnumerable<double> diffs)
        {
            double result = 0;

            var sorted = diffs.OrderBy(d => d).ToArray();

            var isEven = sorted.Length % 2 == 1;
            var halfIndex = sorted.Length/2;
            if (isEven)
            {
                result = sorted[halfIndex];
            }
            else
            {
                result = (sorted[halfIndex - 1] + sorted[halfIndex]) / 2;
            }

            return result;
        }

        private List<DiffRecord> GetDiffs(Statistics result)
        {
            List<DiffRecord> diffs = new List<DiffRecord>();

            result.Points = _points.OrderBy(p => p.TimeStamp).ToArray();
            int firstStatrtIndex = GetStartEventIndex(result.Points);
            for (int i = firstStatrtIndex; i + 1 < result.Points.Length; )
            {
                var start = result.Points[i];
                var stop = result.Points[i + 1];

                if (start.Type == PointType.Start && stop.Type == PointType.Stop)
                {
                    double diff = stop.TimeStamp - start.TimeStamp;
                    diffs.Add(new DiffRecord
                    {
                        Value = diff,
                        TimeStamp = start.TimeStamp
                    });
                    i += 2;
                }
                else
                {
                    i += 1;
                }
            }

            return diffs;
        }

        private HistogramData CreateHistogram(List<DiffRecord> diffs)
        {
            Histogram h = new Histogram(3600000000000L, 3);

            foreach (var diff in diffs)
            {
                h.recordValue((long)diff.Value);
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
