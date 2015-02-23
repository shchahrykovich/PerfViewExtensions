using System.Collections.Generic;
using Reporting.Export;
using Reporting.Implementations.Entities;

namespace Reporting.Tests.Fakes
{
    internal class FakeExporter: IExporter
    {
        public int CallCounter { get; private set; }

        public DiffCollection Result { get; private set; }

        public void Export(string outputFileName, DiffCollection diffs)
        {
            CallCounter++;

            Result = diffs;
        }
    }
}
