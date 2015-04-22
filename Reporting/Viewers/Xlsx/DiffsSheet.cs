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
            IEnumerable<OpenXmlElement> header = AddHeader("TimeStamp, ms", "Diffs, ms", "", "Name", "Value", "", "Diffs, ms", "Frequency", "Cumulative, %");
            SheetData.Append(header);

            IEnumerable<Row> rows = AddDiffs(stat);
            AddStats(rows);
            AddFrequencies(rows, stat);

            SheetData.Append(rows);
        }

        private void AddFrequencies(IEnumerable<Row> rows, Statistics stat)
        {
            var diffs = stat.Diffs.GroupBy(d => d.Value).OrderBy(d => d.Key).ToArray();
            for (int i = 0; i < diffs.Length; i++)
            {
                var diff = diffs[i];

                var row = rows.ElementAt(i);
                row.Append(CreateCell((int)row.RowIndex.Value, 7, diff.Key));
                row.Append(CreateCell((int)row.RowIndex.Value, 8, diff.Count()));
                if (0 == i)
                {
                    row.Append(CreateCellWithFormula((int) row.RowIndex.Value, 9, String.Format("=100 *(G{0} * H{0})/E8", row.RowIndex.Value)));
                }
                else
                {
                    row.Append(CreateCellWithFormula((int)row.RowIndex.Value, 9, String.Format("=100 *(G{0} * H{0})/E8 + I{1}", row.RowIndex.Value, row.RowIndex.Value - 1)));
                }
            }
        }

        private void AddStats(IEnumerable<Row> rows)
        {
            var rowsCount = rows.Count() + OneBasedArray;
            AddFormula(rows.ElementAt(0), "Count", "=" + rowsCount);
            AddFormula(rows.ElementAt(1), "Max", "=MAX(B2:B" + rowsCount + ")");
            AddFormula(rows.ElementAt(2), "Average", "=AVERAGE(B2:B" + rowsCount + ")");
            AddFormula(rows.ElementAt(3), "Median", "=MEDIAN(B2:B" + rowsCount + ")");
            AddFormula(rows.ElementAt(4), "Min", "=MIN(B2:B" + rowsCount + ")");
            AddFormula(rows.ElementAt(5), "StdDeviation", "=STDEVPA(B2:B" + rowsCount + ")");
            AddFormula(rows.ElementAt(6), "Sum", "=SUM(B2:B" + rowsCount + ")");
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
