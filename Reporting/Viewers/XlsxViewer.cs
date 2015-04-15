using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers
{
    internal class XlsxViewer: IViewer
    {
        private readonly string _reportFileName;
        private const int OneBasedArray = 1;
        private const int HeaderRow = 1;

        public XlsxViewer(string reportFileName)
        {
            _reportFileName = reportFileName;
        }

        public void Show(Statistics stat)
        {
            using (var stream = File.Open(_reportFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = package.AddWorkbookPart();
                    Sheets sheets = CreateSheets(workbookPart);

                    AddStats(stat, workbookPart, sheets, 1);
                    AddHistogram(stat, workbookPart, sheets, 2);
                    AddPoints(stat, workbookPart, sheets, 3);
                }
            }
        }

        private static Sheets CreateSheets(WorkbookPart workbookPart)
        {
            Workbook workbook = new Workbook();
            workbook.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            workbookPart.Workbook = workbook;

            Sheets sheets = new Sheets();
            workbook.Append(sheets);
            return sheets;
        }

        private void AddHistogram(Statistics stat, WorkbookPart workBookPart, Sheets sheets, uint sheetId)
        {
            WorksheetPart part = CreateSheet(workBookPart, sheets, "Histogram", sheetId);
            SheetData sheetData = CreateSheetData(part);

            IEnumerable<OpenXmlElement> header = AddHeader("Value", "Percentile", "TotalCount");
            sheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddHistogramRecords(stat);
            sheetData.Append(rows);
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

                rows.Add(row);
            }

            return rows;
        }

        private void AddStats(Statistics stat, WorkbookPart workBookPart, Sheets sheets, uint sheetId)
        {
            WorksheetPart part = CreateSheet(workBookPart, sheets, "Stats", sheetId);
            SheetData sheetData = CreateSheetData(part);

            IEnumerable<OpenXmlElement> header = AddHeader("", "");
            sheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddStatsRecords(stat);
            sheetData.Append(rows);
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

        private static void AddRow(List<Row> rows, string name, long value)
        {
            int rowIndex = rows.Count + OneBasedArray;

            Row row = new Row
            {
                RowIndex = (UInt32) rowIndex
            };

            row.Append(CreateCell(rowIndex, 1, name));
            row.Append(CreateCell(rowIndex, 2, value));

            rows.Add(row);
        }

        private static void AddRow(List<Row> rows, string name, double value)
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

        private static void AddPoints(Statistics stat, WorkbookPart workBookPart, Sheets sheets, uint sheetId)
        {
            WorksheetPart part = CreateSheet(workBookPart, sheets, "Raw data", sheetId);
            SheetData sheetData = CreateSheetData(part);

            IEnumerable<OpenXmlElement> header = AddHeader("TimeStamp", "Type");
            sheetData.Append(header);

            IEnumerable<OpenXmlElement> rows = AddPoints(stat);
            sheetData.Append(rows);
        }

        private static SheetData CreateSheetData(WorksheetPart part)
        {
            SheetData sheetData = new SheetData();

            Worksheet worksheet = new Worksheet();
            worksheet.Append(sheetData);
            part.Worksheet = worksheet;

            return sheetData;
        }

        private static WorksheetPart CreateSheet(WorkbookPart workbookPart, Sheets sheets, string sheetName, uint sheetId)
        {
            String sheetIdName = "rId" + sheetId;

            Sheet sheet = new Sheet()
            {
                Name = sheetName,
                SheetId = (UInt32Value)sheetId,
                Id = sheetIdName
            };
            sheets.Append(sheet);

            return workbookPart.AddNewPart<WorksheetPart>(sheetIdName);
        }

        private static IEnumerable<OpenXmlElement> AddHeader(params String[] columnNames)
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

        private static Cell CreateCell(int rowIndex, int columnIndex, String value)
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

        private static Cell CreateCell(int rowIndex, int columnIndex, double value)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = GetColumnName(columnIndex) + rowIndex,
                CellValue = new CellValue(value.ToString())
            };

            return cell;
        }

        private static Cell CreateCell(int rowIndex, int columnIndex, int value)
        {
            Cell cell = new Cell
            {
                DataType = CellValues.Number,
                CellReference = GetColumnName(columnIndex) + rowIndex,
                CellValue = new CellValue(value.ToString())
            };

            return cell;
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