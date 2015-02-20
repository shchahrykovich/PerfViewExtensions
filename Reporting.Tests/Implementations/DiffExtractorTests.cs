using System;
using System.Collections.Generic;
using Microsoft.Diagnostics.Tracing;
using Microsoft.Diagnostics.Tracing.Etlx;
using NUnit.Framework;
using PerfViewExtensibility;
using Reporting.Implementations;
using Reporting.Tests.Fakes;

namespace Reporting.Tests.Implementations
{
    [TestFixture]
    public class DiffExtractorTests
    {
        private DiffExtractor _extractor;
        private FakeExporter _exporter;
        private ETLDataFile _file;

        [SetUp]
        public void Setup()
        {
            _exporter = new FakeExporter();
            _extractor = new DiffExtractor(_exporter);
            _file = new ETLDataFile("PerfViewData.etl.zip");
        }

        [TearDown]
        public void TearDown()
        {
            if (null != _file)
            {
                _file.Dispose();
            }
        }

        [Test]
        public void Should_Not_Extract_Events_When_Provider_Is_Not_Equal()
        {
            // Arrange
            var arguments = CreateArguments();
            var events = CreateEvents("provider2");

            // Act
            _extractor.ExtractEvents(events, arguments);

            // Assert
            Assert.That(_exporter.Diffs, Is.Empty);
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
        }

        [Test]
        public void Should_Not_Extract_Events_When_Events_Are_Not_Equal()
        {
            // Arrange
            var arguments = CreateArguments();
            var events = CreateEvents(startEventName: "start2", stopEventName: "stop2");

            // Act
            _extractor.ExtractEvents(events, arguments);

            // Assert
            Assert.That(_exporter.Diffs, Is.Empty);
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
        }

        [Test]
        public void Should_Extract_Events()
        {
            // Arrange
            var arguments = CreateArguments();
            var events = CreateEvents();

            // Act
            _extractor.ExtractEvents(events, arguments);

            // Assert
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
            Assert.That(_exporter.Diffs, Is.Not.Empty);
        }

        private IEnumerable<TraceEvent> CreateEvents(String providerName = "provider", String startEventName = "start", String stopEventName = "stop")
        {
            TimeSpan now = DateTime.UtcNow.TimeOfDay;
            Guid activityId = Guid.NewGuid();

            for (int i = 0; i < 10; i++)
            {
                yield return CreateEvent(providerName, startEventName, now, activityId);
                yield return CreateEvent(providerName, stopEventName, now, activityId);
            }
        }

        private static FakeEvent CreateEvent(string providerName, string stopEventName, TimeSpan now, Guid activityId)
        {
            return new FakeEvent(stopEventName, providerName, now, activityId);
        }

        private static DiffExtractorArguments CreateArguments()
        {
            DiffExtractorArguments arguments = new DiffExtractorArguments
            {
                ProviderName = "provider",
                StartEvent = "start",
                StopEvent = "end"
            };
            return arguments;
        }
    }
}