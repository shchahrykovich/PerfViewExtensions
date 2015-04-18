using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Reporting.Implementations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Run = DocumentFormat.OpenXml.Drawing.Run;
using RunProperties = DocumentFormat.OpenXml.Drawing.RunProperties;
using Text = DocumentFormat.OpenXml.Drawing.Text;
using A = DocumentFormat.OpenXml.Drawing;

namespace Reporting.Viewers.Xlsx
{
    internal abstract class GraphSheet : BaseSheet
    {
        protected Title GenerateTitle(string titleText)
        {
            ParagraphProperties paragraphProperties = new ParagraphProperties();
            paragraphProperties.Append(new DefaultRunProperties());

            Run run = new Run();
            Text text = new Text();
            text.Text = titleText;

            run.Append(new RunProperties() {Language = "en-US"});
            run.Append(text);

            Paragraph paragraph = new Paragraph();
            paragraph.Append(paragraphProperties);
            paragraph.Append(run);

            RichText richText = new RichText();
            richText.Append(new BodyProperties());
            richText.Append(new ListStyle());
            richText.Append(paragraph);

            ChartText chartText = new ChartText();
            chartText.Append(richText);

            Title title = new Title();
            title.Append(chartText);
            title.Append(new Layout());
            title.Append(new Overlay() {Val = false});
            return title;
        }

        protected static ChartPart CreateChartPart(DrawingsPart drawingsPart)
        {
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() { Val = new StringValue("en-US") });
            return chartPart;
        }

        protected static void AppendGraphicFrame(DrawingsPart drawingsPart, ChartPart chartPart)
        {
            TwoCellAnchor twoCellAnchor =
    drawingsPart.WorksheetDrawing.AppendChild<TwoCellAnchor>(new TwoCellAnchor());
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.FromMarker(new ColumnId("2"),
                new ColumnOffset("158233"),
                new RowId("2"),
                new RowOffset("16894")));
            twoCellAnchor.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.ToMarker(new ColumnId("17"),
                new ColumnOffset("276225"),
                new RowId("32"),
                new RowOffset("0")));

            // Append a GraphicFrame to the TwoCellAnchor object.
            DocumentFormat.OpenXml.Drawing.Spreadsheet.GraphicFrame graphicFrame =
                twoCellAnchor.AppendChild<DocumentFormat.OpenXml.
                    Drawing.Spreadsheet.GraphicFrame>(new DocumentFormat.OpenXml.Drawing.
                        Spreadsheet.GraphicFrame());
            graphicFrame.Macro = "";

            graphicFrame.Append(new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameProperties(
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualDrawingProperties()
                {
                    Id = new UInt32Value(2u),
                    Name = "Chart 1"
                },
                new DocumentFormat.OpenXml.Drawing.Spreadsheet.NonVisualGraphicFrameDrawingProperties()));

            graphicFrame.Append(new Transform(new Offset() { X = 0L, Y = 0L },
                new Extents() { Cx = 0L, Cy = 0L }));

            graphicFrame.Append(
                new Graphic(new GraphicData(new ChartReference() { Id = drawingsPart.GetIdOfPart(chartPart) })
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }));

            twoCellAnchor.Append(new ClientData());
        }
    }
}