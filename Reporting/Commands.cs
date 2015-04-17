using System;
using Microsoft.Diagnostics.Tracing.Etlx;
using PerfViewExtensibility;
using Reporting.Implementations;
using Reporting.Viewers;

public class Commands : CommandEnvironment
{
    public void CalculateStatistics(string etlFileName, string providerName, string startEvent, string stopEvent, string outputReport = "report.xlsx")
    {
        Parser parser = new Parser(GetViewer(outputReport));
        using (ETLDataFile etlFile = OpenETLFile(etlFileName))
        {
            ParserArguments arguments = new ParserArguments
            {
                ProviderName = providerName,
                StartEvent = startEvent,
                StopEvent = stopEvent,
                OutputReport = outputReport
            };

            TraceEvents events = GetTraceEventsWithProcessFilter(etlFile);
            parser.Parse(events, arguments);
        }
    }

    private static IViewer GetViewer(string reportFileName)
    {
        if (String.IsNullOrEmpty(reportFileName))
        {
            return new TextWriterViewer(LogFile);
        }
        else
        {
            return new XlsxViewer(reportFileName);
        }
    }

    /// <summary>
    /// Gets the TraceEvents list of events from etlFile, applying a process filter if the /process argument is given. 
    /// </summary>
    private TraceEvents GetTraceEventsWithProcessFilter(ETLDataFile etlFile)
    {
        // If the user asked to focus on one process, do so.  
        TraceEvents events;
        if (CommandLineArgs.Process != null)
        {
            var process = etlFile.Processes.LastProcessWithName(CommandLineArgs.Process);
            if (process == null)
            {
                throw new ApplicationException("Could not find process named " + CommandLineArgs.Process);
            }
            events = process.EventsInProcess;
        }
        else
        {
            events = etlFile.TraceLog.Events;
        }
        return events;
    }
}
