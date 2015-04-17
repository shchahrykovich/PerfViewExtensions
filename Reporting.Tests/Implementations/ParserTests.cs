using System;
using System.IO;
using NUnit.Framework;
using PerfView;
using PerfViewExtensibility;
using Reporting.Implementations;
using Reporting.Tests.Viewers;

namespace Reporting.Tests.Implementations
{
    [TestFixture]
    public class ParserTests
    {
        [Test]
        public void TestMethod()
        {
            // Arrange           
            ParserArguments arguments = new ParserArguments
            {
                ProviderName = "Company-First-Logger",
                StartEvent = "StartAction",
                StopEvent = "StopAction"
            };

            // Act
            var result = Execute(arguments);

            // Assert
            Assert.That(result.TotalCount, Is.EqualTo(100));
        }

        private Statistics Execute(ParserArguments parserArguments)
        {
            string fileName = Path.Combine(TestContext.CurrentContext.WorkDirectory, @"Resources\PerfViewData.etl");

            App.CommandProcessor = new CommandProcessor();
            using(ETLDataFile file = new ETLDataFile(fileName))
            {                
                FakeViewer viewer = new FakeViewer();
                var parser = new Parser(viewer);
                parser.Parse(file.TraceLog.Events, parserArguments);
                return viewer.Stat;
            }
        }
    }
}
