using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Reporting.Viewers.Xlsx
{
    internal abstract class DataSheet : BaseSheet
    {
        protected const int OneBasedArray = 1;
        protected const int HeaderRow = 1;

        protected static IEnumerable<OpenXmlElement> AddHeader(params String[] columnNames)
        {
            Row header = new Row
            {
                RowIndex = (UInt32)OneBasedArray
            };

            for (int i = 0; i < columnNames.Length; i++)
            {
                Cell column = CreateCell(1, i + OneBasedArray, columnNames[i]);
                header.Append(column);
            }

            return new[] { header };
        }

        protected static void AddRow(List<Row> rows, string name, long value)
        {
            int rowIndex = rows.Count + OneBasedArray;

            Row row = new Row
            {
                RowIndex = (UInt32)rowIndex
            };

            row.Append(CreateCell(rowIndex, 1, name));
            row.Append(CreateCell(rowIndex, 2, value));

            rows.Add(row);
        }

        protected static void AddRow(List<Row> rows, string name, double value)
        {
            int rowIndex = rows.Count + OneBasedArray;

            Row row = new Row
            {
                RowIndex = (UInt32)rowIndex
            };

            row.Append(CreateCell(rowIndex, 1, name));
            row.Append(CreateCell(rowIndex, 2, value));

            rows.Add(row);
        }

        protected static Cell CreateCell(int rowIndex, int columnIndex, String value)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.InlineString,
                CellReference = GetColumnName(columnIndex) + rowIndex
            };

            InlineString inlineString = new InlineString();
            inlineString.AppendChild(new Text
            {
                Text = value
            });

            cell.AppendChild(inlineString);

            return cell;
        }

        protected static Cell CreateCell(int rowIndex, int columnIndex, double value)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = GetColumnName(columnIndex) + rowIndex,
                CellValue = new CellValue(value.ToString())
            };

            return cell;
        }

        protected static Cell CreateCellWithFormula(int rowIndex, int columnIndex, String formula)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = GetColumnName(columnIndex) + rowIndex,
                CellFormula = new CellFormula(formula)
            };

            return cell;
        }

        protected static Cell CreateCell(int rowIndex, int columnIndex, int value)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = GetColumnName(columnIndex) + rowIndex,
                CellValue = new CellValue(value.ToString())
            };

            return cell;
        }

        private static string GetColumnName(int columnIndex)
        {
            int dividend = columnIndex;
            string columnName = String.Empty;

            while (dividend > 0)
            {
                int modifier = (dividend - 1) % 26;
                columnName =
                    Convert.ToChar(65 + modifier).ToString() + columnName;
                dividend = (int)((dividend - modifier) / 26);
            }

            return columnName;
        }
    }
}
