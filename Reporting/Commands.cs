using System;
using Microsoft.Diagnostics.Tracing.Etlx;
using PerfViewExtensibility;
using Reporting.DiffExtractor;
using Address = System.UInt64;

public class Commands : CommandEnvironment
{
    /// <summary>
    /// Extracts difference between two events.
    /// </summary>
    /// <param name="etlFileName"></param>
    /// <param name="providerName"></param>
    /// <param name="startEvent"></param>
    /// <param name="stopEvent"></param>
    /// <param name="outputReport"></param>
    public void ExtractDiffs(string etlFileName, string providerName, string startEvent, string stopEvent, string outputReport = "report.xlsx")
    {
        DiffExtractor extractor = new DiffExtractor();
        using (var etlFile = OpenETLFile(etlFileName))
        {
            DiffExtractorArguments arguments = new DiffExtractorArguments
            {
                ProviderName = providerName,
                StartEvent = startEvent,
                StopEvent = stopEvent,
                OutputReport = outputReport
            };

            TraceEvents events = GetTraceEventsWithProcessFilter(etlFile);
            extractor.ExtractEvents(events, arguments);
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
