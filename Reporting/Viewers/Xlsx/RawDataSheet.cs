using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal class RawDataSheet: DataSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Raw data");
        }

        internal override void AddData(Statistics stat)
        {
            IEnumerable<OpenXmlElement> header = AddHeader("TimeStamp", "Type");
            SheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddPoints(stat);
            SheetData.Append(rows);
        }

        private static IEnumerable<OpenXmlElement> AddPoints(Statistics stat)
        {
            List<Row> rows = new List<Row>();

            for (int i = 0; i < stat.Points.Length; i++)
            {
                Point point = stat.Points[i];
                
                int rowIndex = i + HeaderRow + OneBasedArray;
                Row row = new Row
                {
                    RowIndex = (UInt32)rowIndex
                };

                row.Append(CreateCell(rowIndex, 1, point.TimeStamp));
                row.Append(CreateCell(rowIndex, 2, point.Type.ToString()));

                rows.Add(row);
            }

            return rows;
        }
    }
}
