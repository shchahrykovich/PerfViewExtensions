using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Reporting.Implementations.Entities;

namespace Reporting.Export
{
    internal interface IExporter
    {
        void Export(string outputFileName, List<Diff> diffs);
    }
}
