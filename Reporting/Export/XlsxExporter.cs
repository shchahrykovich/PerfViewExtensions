using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations.Entities;

namespace Reporting.Export
{
    internal static class XlsxExporter
    {
        private static readonly String[] ColumnNames = new[]
        {
            "TimeStamp",
            "Diff",
            "",
            "",
            "",
            "",
            "",
            ""
        };

        private static void CreateParts(SpreadsheetDocument document)
        {
            WorkbookPart workbookPart1 = document.AddWorkbookPart();
            GenerateWorkbookPart1Content(workbookPart1);
            workbookPart1.AddNewPart<WorksheetPart>("rId1");
        }

        private static void GenerateWorkbookPart1Content(WorkbookPart workbookPart1)
        {
            Workbook workbook1 = new Workbook();
            workbook1.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");

            Sheets sheets1 = new Sheets();
            Sheet sheet1 = new Sheet()
            {
                Name = "Raw data",
                SheetId = (UInt32Value)1U,
                Id = "rId1"
            };
            sheets1.Append(sheet1);

            workbook1.Append(sheets1);
            workbookPart1.Workbook = workbook1;
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

        private static IEnumerable<OpenXmlElement> CreateColumns()
        {
            const int oneBasedArray = 1;

            Row header = new Row
            {
                RowIndex = (UInt32)oneBasedArray
            };

            for (int i = 0; i < ColumnNames.Length; i++)
            {
                Cell column = CreateCell(1, i + oneBasedArray, ColumnNames[i]);
                header.Append(column);
            }

            return new[] { header };
        }

        public static void WriteReport(string reportFile, List<Diff> report)
        {
            using (var stream = File.Open(reportFile, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    CreateParts(package);

                    WorksheetPart part = package.WorkbookPart.WorksheetParts.First();

                    Worksheet worksheet1 = new Worksheet();
                    SheetData sheetData1 = new SheetData();

                    IEnumerable<OpenXmlElement> header = CreateColumns();
                    sheetData1.Append(header);

                    IEnumerable<OpenXmlElement> rows = CreateRows(report);
                    sheetData1.Append(rows);

                    worksheet1.Append(sheetData1);
                    part.Worksheet = worksheet1;
                }
            }
        }

        private static IEnumerable<OpenXmlElement> CreateRows(List<Diff> report)
        {
            const int headerRow = 1;
            const int oneBasedArray = 1;

            List<Row> rows = new List<Row>();

            for (int i = 0; i < report.Count; i++)
            {
                var diff = report[i];

                int rowIndex = i + headerRow + oneBasedArray;
                Row row = new Row
                {
                    RowIndex = (UInt32)rowIndex
                };

                TimeSpan timeSpan = diff.Start.TimeStamp.TimeOfDay;
                string time = timeSpan.Hours + ":" + timeSpan.Minutes + ":" + timeSpan.Seconds + ":" + timeSpan.Milliseconds;
                row.Append(CreateCell(rowIndex, 1, time));
                row.Append(CreateCell(rowIndex, 2, (diff.Finish.TimeStamp - diff.Start.TimeStamp).TotalMilliseconds));

                if (i == 0)
                {
                    row.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellReference = GetColumnName(4) + rowIndex,
                        CellValue = new CellValue("Count")
                    });

                    row.Append(new Cell
                    {
                        DataType = CellValues.Number,
                        CellReference = GetColumnName(5) + rowIndex,
                        CellValue = new CellValue(report.Count.ToString())
                    });
                }

                if (i == 1)
                {
                    row.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellReference = GetColumnName(4) + rowIndex,
                        CellValue = new CellValue("Average")
                    });

                    var average = "=AVERAGE(B2:B" + (report.Count + 1) + ")";
                    row.Append(new Cell
                    {
                        DataType = CellValues.Number,
                        CellReference = GetColumnName(5) + rowIndex,
                        CellFormula = new CellFormula(average)
                    });
                }

                if (i == 2)
                {
                    row.Append(new Cell
                    {
                        DataType = CellValues.String,
                        CellReference = GetColumnName(4) + rowIndex,
                        CellValue = new CellValue("Median")
                    });

                    var median = "=MEDIAN(B2:B" + (report.Count + 1) + ")";
                    row.Append(new Cell
                    {
                        DataType = CellValues.Number,
                        CellReference = GetColumnName(5) + rowIndex,
                        CellFormula = new CellFormula(median)
                    });
                }

                rows.Add(row);
            }

            return rows;
        }
    }
}
