using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;
using Reporting.Viewers.Xlsx;

namespace Reporting.Viewers
{
    internal class XlsxViewer : IViewer
    {
        private readonly string _reportFileName;
        private readonly BaseSheet[] _customSheets;

        public XlsxViewer(string reportFileName)
        {
            _reportFileName = reportFileName;

            _customSheets = new BaseSheet[]
            {
                new StatSheet(),
                new HistogramSheet(),
                new RawDataSheet(),
            };
        }

        public void Show(Statistics stat)
        {
            using (var stream = File.Open(_reportFileName, FileMode.OpenOrCreate, FileAccess.ReadWrite))
            {
                using (SpreadsheetDocument package = SpreadsheetDocument.Create(stream, SpreadsheetDocumentType.Workbook))
                {
                    WorkbookPart workbookPart = package.AddWorkbookPart();
                    Sheets sheets = CreateSheets(workbookPart);

                    foreach (BaseSheet sheet in _customSheets)
                    {
                        sheet.Create(workbookPart, sheets);
                        sheet.AddData(stat);
                    }
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
    }
}