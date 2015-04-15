using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal class StatSheet : BaseSheet
    {
        internal override void AddData(Statistics stat)
        {
            IEnumerable<OpenXmlElement> header = AddHeader("", "");
            SheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddStatsRecords(stat);
            SheetData.Append(rows);
        }

        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Stats");
        }

        private IEnumerable<OpenXmlElement> AddStatsRecords(Statistics stat)
        {
            List<Row> rows = new List<Row>();

            AddRow(rows, "MaxValue", stat.MaxValue);
            AddRow(rows, "MinValue", stat.MinValue);
            AddRow(rows, "StdDeviation", stat.StdDeviation);
            AddRow(rows, "TotalCount", stat.TotalCount);

            return rows;
        }
    }
}
