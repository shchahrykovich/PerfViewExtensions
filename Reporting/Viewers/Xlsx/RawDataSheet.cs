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

        //if (i == 0)
        //{
        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.String,
        //        CellReference = GetColumnName(4) + rowIndex,
        //        CellValue = new CellValue("Count")
        //    });

        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.Number,
        //        CellReference = GetColumnName(5) + rowIndex,
        //        CellValue = new CellValue(stat.Count.ToString())
        //    });
        //}

        //if (i == 1)
        //{
        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.String,
        //        CellReference = GetColumnName(4) + rowIndex,
        //        CellValue = new CellValue("Average")
        //    });

        //    var average = "=AVERAGE(B2:B" + (stat.Count + 1) + ")";
        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.Number,
        //        CellReference = GetColumnName(5) + rowIndex,
        //        CellFormula = new CellFormula(average)
        //    });
        //}

        //if (i == 2)
        //{
        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.String,
        //        CellReference = GetColumnName(4) + rowIndex,
        //        CellValue = new CellValue("Median")
        //    });

        //    var median = "=MEDIAN(B2:B" + (stat.Count + 1) + ")";
        //    row.Append(new Cell
        //    {
        //        DataType = CellValues.Number,
        //        CellReference = GetColumnName(5) + rowIndex,
        //        CellFormula = new CellFormula(median)
        //    });
        //}

    }
}
