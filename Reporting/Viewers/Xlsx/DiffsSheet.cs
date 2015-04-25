using System;
using System.Collections.Generic;
using System.Linq;
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
            IEnumerable<OpenXmlElement> header = AddHeader("TimeStamp, ms", "Diffs, ms", "", "Name", "Value", "", "Diffs, ms", "Frequency", "Total Value", "Total Value %", "Total Count", "Total Count %");
            SheetData.Append(header);

            IEnumerable<Row> rows = AddDiffs(stat);
            AddStats(rows, stat);
            AddFrequencies(rows, stat);

            SheetData.Append(rows);
        }

        private void AddFrequencies(IEnumerable<Row> rows, Statistics stat)
        {
            for (int i = 0; i < stat.Frequencies.Count; i++)
            {
                var diff = stat.Frequencies[i];

                var row = rows.ElementAt(i);
                row.Append(CreateCell((int)row.RowIndex.Value, 7, diff.Value));
                row.Append(CreateCell((int)row.RowIndex.Value, 8, diff.Count));
                row.Append(CreateCell((int)row.RowIndex.Value, 9, diff.TotalValue));
                row.Append(CreateCell((int)row.RowIndex.Value, 10, diff.TotalValuePercent));
                row.Append(CreateCell((int)row.RowIndex.Value, 11, diff.TotalCount));
                row.Append(CreateCell((int)row.RowIndex.Value, 12, diff.TotalCountPercent));
            }
        }

        private void AddStats(IEnumerable<Row> rows, Statistics stat)
        {
            AddFormula(rows.ElementAt(0), "Count", "=" + stat.TotalCount);
            AddFormula(rows.ElementAt(1), "Max", "=" + stat.MaxValue);
            AddFormula(rows.ElementAt(2), "Average", "=" + stat.Average);
            AddFormula(rows.ElementAt(3), "Median", "=" + stat.Median);
            AddFormula(rows.ElementAt(4), "Min", "=" + stat.MinValue);
            AddFormula(rows.ElementAt(5), "StdDeviation", "=" + stat.StdDeviation);
            AddFormula(rows.ElementAt(6), "Sum", "=" + stat.Sum);
        }

        private static void AddFormula(Row row, string name, string formula)
        {
            row.Append(CreateCell((int) row.RowIndex.Value, 4, name));
            row.Append(CreateCellWithFormula((int) row.RowIndex.Value, 5, formula));
        }

        private static IEnumerable<Row> AddDiffs(Statistics stat)
        {
            List<Row> rows = new List<Row>();

            for (int i = 0; i < stat.Diffs.Count; i++)
            {
                var diff = stat.Diffs[i];

                int rowIndex = i + HeaderRow + OneBasedArray;
                Row row = new Row
                {
                    RowIndex = (UInt32)rowIndex
                };

                row.Append(CreateCell(rowIndex, 1, diff.TimeStamp));
                row.Append(CreateCell(rowIndex, 2, diff.Value));

                rows.Add(row);
            }

            return rows;
        }
    }
}
