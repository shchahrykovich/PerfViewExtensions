using System.Collections.Generic;
using Reporting.Export;
using Reporting.Implementations.Entities;

namespace Reporting.Tests.Fakes
{
    internal class FakeExporter: IExporter
    {
        public int CallCounter { get; private set; }

        public List<Diff> Diffs { get; private set; }

        public void Export(string outputFileName, List<Diff> diffs)
        {
            CallCounter++;

            Diffs = diffs;
        }
    }
}
