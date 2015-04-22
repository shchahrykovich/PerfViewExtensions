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

// ReSharper disable PossiblyMistakenUseOfParamsMethod

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
            TwoCellAnchor twoCellAnchor = drawingsPart.WorksheetDrawing.AppendChild(new TwoCellAnchor());
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
                    Id = new UInt32Value(3u),
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

        protected void AppendValueAxis(PlotArea plotArea, uint id, string name, uint crossingAxisId, AxisPositionValues axisPos = AxisPositionValues.Left, TickLabelPositionValues tickPos = TickLabelPositionValues.NextTo, bool showMajorGridlines = true)
        {
            List<OpenXmlElement> elements = GetAxisElements(id, name, crossingAxisId, tickPosition: tickPos);

            elements.Add(new AxisPosition { Val = new EnumValue<AxisPositionValues>(axisPos) });
            if (showMajorGridlines)
            {
                elements.Add(new MajorGridlines());
            }
            elements.Add(new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
            {
                FormatCode = new StringValue("General"),
                SourceLinked = new BooleanValue(true)
            });
            elements.Add(new CrossBetween { Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between) });

            plotArea.AppendChild(new ValueAxis(elements));
        }

        protected void AppendCategoryAxis(PlotArea plotArea, uint id, string name, uint crossingAxisId, bool show = true)
        {
            List<OpenXmlElement> elements = GetAxisElements(id, name, crossingAxisId, show);

            elements.Add(new AxisPosition { Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom) });
            elements.Add(new AutoLabeled { Val = new BooleanValue(true) });
            elements.Add(new LabelAlignment { Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center) });
            elements.Add(new LabelOffset { Val = new UInt16Value((ushort)100) });

            plotArea.AppendChild(new CategoryAxis(elements));
        }

        private List<OpenXmlElement> GetAxisElements(uint id, String name, uint crossingAxisId, bool show = true, TickLabelPositionValues tickPosition = TickLabelPositionValues.NextTo)
        {
            var result = new List<OpenXmlElement>();
            if (show)
            {
                result.Add(new Delete { Val = false });
            }
            result.Add(new AxisId { Val = new UInt32Value(id) });
            result.Add(new Scaling(new Orientation()
            {
                Val = new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
            }));
            if (show)
            {
                result.Add(GenerateTitle(name));
            }
            result.Add(new TickLabelPosition
            {
                Val = new EnumValue<TickLabelPositionValues>(tickPosition)
            });
            result.Add(new CrossingAxis { Val = new UInt32Value(crossingAxisId) });
            result.Add(new Crosses { Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero) });
            return result;
        }
    }
}