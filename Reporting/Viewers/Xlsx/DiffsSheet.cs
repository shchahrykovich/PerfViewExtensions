using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal class DiffsSheet : DataSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Diffs");
        }

        internal override void AddData(Statistics stat)
        {
            IEnumerable<OpenXmlElement> header = AddHeader("Diffs");
            SheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddDiffs(stat);
            SheetData.Append(rows);
        }

        private static IEnumerable<OpenXmlElement> AddDiffs(Statistics stat)
        {
            List<Row> rows = new List<Row>();

            for (int i = 0; i < stat.Diffs.Count; i++)
            {
                double diff = stat.Diffs[i];

                int rowIndex = i + HeaderRow + OneBasedArray;
                Row row = new Row
                {
                    RowIndex = (UInt32)rowIndex
                };

                row.Append(CreateCell(rowIndex, 1, diff));

                rows.Add(row);
            }

            return rows;
        }
    }
}
