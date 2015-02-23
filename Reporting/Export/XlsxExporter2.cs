﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Reporting.Implementations.Entities;

namespace Reporting.Export
{
    internal class XlsxExporter2 : IExporter
    {
        public void Export(string outputFileName, DiffCollection diffs)
        {
            XlsxExporter.WriteReport(outputFileName, diffs.Diffs);
        }
    }
}
