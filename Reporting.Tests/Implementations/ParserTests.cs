using System;
using System.Collections;
using Microsoft.Diagnostics.Tracing.Etlx;
using NUnit.Framework;
using Reporting.Implementations;
using Reporting.Tests.Viewers;

namespace Reporting.Tests.Implementations
{
    [TestFixture]
    public class ParserTests
    {
        [Test]
        [TestCaseSource("GetAllProvidersAndCheckStats")]
        public StatsTestResult Should_Return_Correct_TotatlCount(String provider, String start, String finish)
        {
            ParserArguments arguments = new ParserArguments
            {
                ProviderName = provider,
                StartEvent = start,
                StopEvent = finish
            };

            var result = Execute(arguments);

            return new StatsTestResult
            {
                TotalCount = result.TotalCount,
                MaxValue = result.MaxValue,
                MinVale = result.MinValue,
                StdDeviation = result.StdDeviation
            };
        }

        private Statistics Execute(ParserArguments parserArguments)
        {
            string fileName = @"Resources\PerfViewData.etl";
            using (var file = TraceLog.OpenOrConvert(fileName))
            {              
                FakeViewer viewer = new FakeViewer();
                var parser = new Parser(viewer);
                parser.Parse(file.Events, parserArguments);
                return viewer.Stat;
            }
        }

        private static IEnumerable GetAllProvidersAndCheckStats()
        {
            yield return new TestCaseData("Company-First-Logger", "StartAction", "StopAction").Returns(new StatsTestResult
            {
                TotalCount = 100,
                MaxValue = 172,
                MinVale = 61,
                StdDeviation = 29.511428294814873d
            });
            yield return new TestCaseData("Company-First-Logger", "Begin", "End").Returns(new StatsTestResult
            {
                TotalCount = 1,
                MaxValue = 11376,
                MinVale = 11376,
                StdDeviation = 0
            });
            yield return new TestCaseData("Company-Second-Logger", "Begin", "End").Returns(new StatsTestResult
            {
                TotalCount = 60,
                MaxValue = 610,
                MinVale = 7,
                StdDeviation = 173.59138723514545d
            });
            yield return new TestCaseData("Company-Second-Logger", "Start", "Finish").Returns(new StatsTestResult
            {
                TotalCount = 1,
                MaxValue = 405,
                MinVale = 405,
                StdDeviation = 0
            });
            yield return new TestCaseData("Company-Second-Logger", "Start", "Fake").Returns(new StatsTestResult
            {
                TotalCount = 0,
                MaxValue = 0,
                MinVale = 0,
                StdDeviation = Double.NaN
            });
            yield return new TestCaseData("Company-Second-Logger", "Fake", "Finish").Returns(new StatsTestResult
            {
                TotalCount = 0,
                MaxValue = 0,
                MinVale = 0,
                StdDeviation = Double.NaN
            });
            yield return new TestCaseData("Fake", "Start", "Finish").Returns(new StatsTestResult
            {
                TotalCount = 0,
                MaxValue = 0,
                MinVale = 0,
                StdDeviation = Double.NaN
            });
        }
    }
}
