﻿using System.Collections.Generic;
using System.Linq;
using HdrHistogram.NET;
using Microsoft.Diagnostics.Tracing;
using PEFile;
using Reporting.Export;

namespace Reporting.Implementations
{
    internal class Parser
    {
        private readonly IExporter _exporter;
        private readonly List<Point> _points;

        public Parser(IExporter exporter)
        {
            _exporter = exporter;
            _points = new List<Point>();
        }

        public void ExtractEvents(IEnumerable<TraceEvent> events, ParserArguments arguments)
        {
            FindEvents(events, arguments.ProviderName, arguments.StartEvent, arguments.StopEvent);
            Statistics s = CalculateStatistics();

            _exporter.Export(s);
        }

        private Statistics CalculateStatistics()
        {
            Statistics result = new Statistics();

            Histogram h = new Histogram(3600000000000L, 3);

            result.Points = _points.OrderBy(p => p.TimeStamp).ToArray();
            int firstStatrtIndex = GetStartEventIndex(result.Points);
            for (int i = firstStatrtIndex; i + 1 < result.Points.Length; )
            {
                var start = result.Points[i];
                var stop = result.Points[i + 1];

                if (start.Type == PointType.Start && stop.Type == PointType.Stop)
                {
                    var diff = stop.TimeStamp - start.TimeStamp;
                    h.recordValue((long)diff);
                    i += 2;
                }
                else
                {
                    i += 1;
                }
            }

            HistogramData histogramData = h.getHistogramData();
            result.MaxValue = histogramData.getMaxValue();
            result.MinValue = histogramData.getMinValue();
            result.StdDeviation = histogramData.getStdDeviation();
            result.TotalCount = histogramData.getTotalCount();
            result.Percentile90 = histogramData.getValueAtPercentile(90);
            result.Percentile95 = histogramData.getValueAtPercentile(95);
            result.Percentile99 = histogramData.getValueAtPercentile(99);

            return result;
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
