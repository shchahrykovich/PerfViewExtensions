using System;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;

namespace Reporting.Viewers.Xlsx
{
    internal abstract class BaseSheet
    {
        protected SheetData SheetData { get; private set; }
        protected Sheet Sheet { get; private set; }
        protected WorkbookPart WorkbookPart { get; private set; }
        protected WorksheetPart WorksheetPart { get; private set; }

        internal abstract void AddData(Statistics stat);

        internal abstract void Create(WorkbookPart workBookPart, Sheets sheets);

        protected virtual void Create(WorkbookPart workBookPart, Sheets sheets, String sheetName)
        {
            WorkbookPart = workBookPart;
            WorksheetPart = CreateSheet(sheets, sheetName);
            SheetData = CreateSheetData();
        }

        private WorksheetPart CreateSheet(Sheets sheets, string sheetName)
        {
            uint sheetId = (uint)sheets.ChildElements.Count + 1;
            String sheetIdName = "rId" + sheetId;

            Sheet = new Sheet()
            {
                Name = sheetName,
                SheetId = (UInt32Value)sheetId,
                Id = sheetIdName
            };
            sheets.Append(Sheet);

            return WorkbookPart.AddNewPart<WorksheetPart>(sheetIdName);
        }

        private SheetData CreateSheetData()
        {
            SheetData sheetData = new SheetData();

            Worksheet worksheet = new Worksheet();
            worksheet.Append(sheetData);
            WorksheetPart.Worksheet = worksheet;

            return sheetData;
        }
    }
}