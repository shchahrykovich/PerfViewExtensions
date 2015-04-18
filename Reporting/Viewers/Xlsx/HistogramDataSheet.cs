using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal class HistogramDataSheet : DataSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Histogram data");
        }

        internal override void AddData(Statistics stat)
        {
            IEnumerable<OpenXmlElement> header = AddHeader("Value", "Percentile", "TotalCount", "Count");
            SheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddHistogramRecords(stat);
            SheetData.Append(rows);
        }

        private IEnumerable<OpenXmlElement> AddHistogramRecords(Statistics stat)
        {
            List<Row> rows = new List<Row>();

            foreach (var p in stat.Percentiles)
            {
                int rowIndex = rows.Count + OneBasedArray + HeaderRow;

                Row row = new Row
                {
                    RowIndex = (UInt32)rowIndex
                };

                row.Append(CreateCell(rowIndex, 1, p.Value));
                row.Append(CreateCell(rowIndex, 2, p.Percentile));
                row.Append(CreateCell(rowIndex, 3, p.TotalCount));
                row.Append(CreateCell(rowIndex, 4, p.Count));

                rows.Add(row);
            }

            return rows;
        }

    }
}
