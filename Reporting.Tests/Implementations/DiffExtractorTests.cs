using System;
using System.IO;
using Microsoft.Diagnostics.Tracing.Etlx;
using NUnit.Framework;
using Reporting.Implementations;
using Reporting.Tests.Fakes;
using Reporting.Tests.Resources;

namespace Reporting.Tests.Implementations
{
    [TestFixture]
    public class DiffExtractorTests
    {
        private DiffExtractor _extractor;
        private FakeExporter _exporter;
        private string _etlFilePath;

        [SetUp]
        public void Setup()
        {
            _exporter = new FakeExporter();
            _extractor = new DiffExtractor(_exporter);

            _etlFilePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".etl");
            File.WriteAllBytes(_etlFilePath, Files.PerfViewData);
        }

        [TearDown]
        public void TearDown()
        {
            if (!String.IsNullOrEmpty(_etlFilePath))
            {
                if (File.Exists(_etlFilePath))
                {
                    File.Delete(_etlFilePath);
                }
            }
        }

        [Test]
        public void Should_Not_Extract_Events_When_Provider_Is_Not_Equal()
        {
            // Arrange
            var arguments = CreateArguments(providerName: "test");

            // Act
            Extract(arguments);

            // Assert
            Assert.That(_exporter.Diffs, Is.Empty);
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
        }

        [Test]
        public void Should_Not_Extract_Events_When_Events_Are_Not_Equal()
        {
            // Arrange
            var arguments = CreateArguments(startEventName: "start", stopEventName: "stop");

            // Act
            Extract(arguments);

            // Assert
            Assert.That(_exporter.Diffs, Is.Empty);
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
        }

        [Test]
        public void Should_Extract_Events()
        {
            // Arrange
            var arguments = CreateArguments();

            // Act
            Extract(arguments);

            // Assert
            Assert.That(_exporter.CallCounter, Is.EqualTo(1));
            Assert.That(_exporter.Diffs, Is.Not.Empty);
        }

        private void Extract(DiffExtractorArguments arguments)
        {
            using (var log = TraceLog.OpenOrConvert(_etlFilePath))
            {
                _extractor.ExtractEvents(log.Events, arguments);
            }
        }

        private static DiffExtractorArguments CreateArguments(string providerName = "Company-First-Logger",
                                                              string stopEventName = "StopAction",
                                                              string startEventName = "StartAction")
        {
            DiffExtractorArguments arguments = new DiffExtractorArguments
            {
                ProviderName = providerName,
                StartEvent = startEventName,
                StopEvent = stopEventName
            };
            return arguments;
        }
    }
}