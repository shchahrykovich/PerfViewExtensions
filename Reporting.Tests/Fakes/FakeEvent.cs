using System;
using Microsoft.Diagnostics.Tracing;
using Microsoft.Diagnostics.Tracing.Etlx;
using Microsoft.Diagnostics.Tracing.Parsers.Symbol;

namespace Reporting.Tests.Fakes
{
    public class FakeEvent : TraceEvent
    {
        public FakeEvent(string taskName, string providerName, TimeSpan timeSpan, Guid activityId) :
            base(1, 1, taskName, Guid.NewGuid(), 1, String.Empty, Guid.NewGuid(), providerName)
        {
            
        }

        public override object PayloadValue(int index)
        {
            throw new NotImplementedException();
        }

        public override string[] PayloadNames
        {
            get { throw new NotImplementedException(); }
        }

        protected override Delegate Target { get; set; }

        public override int ProcessID
        {
            get { return 1; }
        }
    }
}