using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal abstract class BaseSheet
    {
        protected WorksheetPart Part { get; private set; }
        protected SheetData SheetData { get; private set; }
        protected const int OneBasedArray = 1;
        protected const int HeaderRow = 1;

        internal abstract void AddData(Statistics stat);

        internal abstract void Create(WorkbookPart workBookPart, Sheets sheets);        

        protected void Create(WorkbookPart workBookPart, Sheets sheets, String sheetName)
        {
            Part = CreateSheet(workBookPart, sheets, sheetName);
            SheetData = CreateSheetData();
        }

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

        private SheetData CreateSheetData()
        {
            SheetData sheetData = new SheetData();

            Worksheet worksheet = new Worksheet();
            worksheet.Append(sheetData);
            Part.Worksheet = worksheet;

            return sheetData;
        }

        private static WorksheetPart CreateSheet(WorkbookPart workbookPart, Sheets sheets, string sheetName)
        {
            uint sheetId = (uint) sheets.ChildElements.Count + 1;
            String sheetIdName = "rId" + sheetId;

            Sheet sheet = new Sheet()
            {
                Name = sheetName,
                SheetId = (UInt32Value) sheetId,
                Id = sheetIdName
            };
            sheets.Append(sheet);

            return workbookPart.AddNewPart<WorksheetPart>(sheetIdName);
        }
    }
}
