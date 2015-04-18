﻿using System;
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
    internal class HistogramSheet : BaseSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Histogram");
        }

        internal override void AddData(Statistics stat)
        {
            List<long> count = new List<long>();
            List<double> values = new List<double>();

            foreach (var percentile in stat.Percentiles)
            {
                count.Add(percentile.Count);
                values.Add(percentile.Value);
            }

            InsertChartInSpreadsheet(count, values);
        }

        //https://github.com/OfficeDev/office-content/blob/master/en-us/OpenXMLCon/articles/281776d0-be75-46eb-8fdc-a1f656291175.md
        private void InsertChartInSpreadsheet(List<long> count, List<double> values)
        {
            // Add a new drawing to the worksheet.
            DrawingsPart drawingsPart = WorksheetPart.AddNewPart<DrawingsPart>();
            WorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
            {
                Id = WorksheetPart.GetIdOfPart(drawingsPart)
            });
            WorksheetPart.Worksheet.Save();

            // Add a new chart and set the chart language to English-US.
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = new ChartSpace();
            chartPart.ChartSpace.Append(new EditingLanguage() {Val = new StringValue("en-US")});
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace
                .AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>(
                    new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            Layout layout = plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart =
                plotArea.AppendChild<BarChart>(
                    new BarChart(
                        new BarDirection() {Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column)},
                        new BarGrouping() {Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered)},
                        new VaryColors() {Val = false}
                        ));

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.

            BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
            {
                Val = new UInt32Value((uint) 0)
            },
                new Order() {Val = new UInt32Value((uint) 0)},
                new SeriesText(new NumericValue() {Text = "Histogram"})));

            StringLiteral strLit =
                barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData())
                    .AppendChild<StringLiteral>(new StringLiteral());
            strLit.Append(new PointCount() {Val = new UInt32Value((uint) count.Count)});

            NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                new DocumentFormat.OpenXml.Drawing.Charts.Values())
                .AppendChild<NumberLiteral>(new NumberLiteral());
            numLit.Append(new FormatCode("General"));
            numLit.Append(new PointCount() {Val = new UInt32Value((uint) count.Count)});

            for (uint i = 0; i < count.Count; i++)
            {
                strLit.AppendChild<StringPoint>(new StringPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue(values[(int) i].ToString()));
                numLit.AppendChild<NumericPoint>(new NumericPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue(count[(int) i].ToString()));
            }

            barChart.Append(new AxisId() {Val = new UInt32Value(48650112u)});
            barChart.Append(new AxisId() {Val = new UInt32Value(48672768u)});

            // Add the Category Axis.
            CategoryAxis catAx =
                plotArea.AppendChild<CategoryAxis>(new CategoryAxis(
                    new AxisId()
                    {
                        Val = new UInt32Value(48650112u)
                    },
                    new Scaling(new Orientation()
                    {
                        Val =
                            new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                    }),
                    new Delete() {Val = false},
                    GenerateTitle("Time, ms"),
                    new AxisPosition() {Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Bottom)},
                    new TickLabelPosition()
                    {
                        Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
                    },
                    new CrossingAxis() {Val = new UInt32Value(48672768U)},
                    new Crosses() {Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero)},
                    new AutoLabeled() {Val = new BooleanValue(true)},
                    new LabelAlignment() {Val = new EnumValue<LabelAlignmentValues>(LabelAlignmentValues.Center)},
                    new LabelOffset() {Val = new UInt16Value((ushort) 100)}));

            // Add the Value Axis.
            ValueAxis valAx =
                plotArea.AppendChild<ValueAxis>(new ValueAxis(new AxisId() {Val = new UInt32Value(48672768u)},
                    new Scaling(new Orientation()
                    {
                        Val =
                            new EnumValue<DocumentFormat.OpenXml.Drawing.Charts.OrientationValues>(
                                DocumentFormat.OpenXml.Drawing.Charts.OrientationValues.MinMax)
                    }),
                    new AxisPosition() {Val = new EnumValue<AxisPositionValues>(AxisPositionValues.Left)},
                    new MajorGridlines(),
                    GenerateTitle("Number of samples"),
                    new Delete() {Val = false},
                    new DocumentFormat.OpenXml.Drawing.Charts.NumberingFormat()
                    {
                        FormatCode = new StringValue("General"),
                        SourceLinked = new BooleanValue(true)
                    }, new TickLabelPosition()
                    {
                        Val = new EnumValue<TickLabelPositionValues>(TickLabelPositionValues.NextTo)
                    },
                    new CrossingAxis() {Val = new UInt32Value(48650112U)},
                    new Crosses() {Val = new EnumValue<CrossesValues>(CrossesValues.AutoZero)},
                    new CrossBetween() {Val = new EnumValue<CrossBetweenValues>(CrossBetweenValues.Between)}));

            // Add the chart Legend.
            Legend legend =
                chart.AppendChild<Legend>(
                    new Legend(
                        new LegendPosition() {Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right)},
                        new Layout()));

            chart.Append(new PlotVisibleOnly() {Val = new BooleanValue(true)});

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
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

            graphicFrame.Append(new Transform(new Offset() {X = 0L, Y = 0L},
                new Extents() {Cx = 0L, Cy = 0L}));

            graphicFrame.Append(
                new Graphic(new GraphicData(new ChartReference() {Id = drawingsPart.GetIdOfPart(chartPart)})
                {
                    Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart"
                }));

            twoCellAnchor.Append(new ClientData());

            // Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save();
        }

        public Title GenerateTitle(string title)
        {
            Title title1 = new Title();

            ChartText chartText1 = new ChartText();

            RichText richText1 = new RichText();
            BodyProperties bodyProperties1 = new BodyProperties();
            ListStyle listStyle1 = new ListStyle();

            Paragraph paragraph1 = new Paragraph();

            ParagraphProperties paragraphProperties1 = new ParagraphProperties();
            DefaultRunProperties defaultRunProperties1 = new DefaultRunProperties();

            paragraphProperties1.Append(defaultRunProperties1);

            Run run1 = new Run();
            RunProperties runProperties1 = new RunProperties() {Language = "en-US"};
            Text text1 = new A.Text();
            text1.Text = title;

            run1.Append(runProperties1);
            run1.Append(text1);

            paragraph1.Append(paragraphProperties1);
            paragraph1.Append(run1);

            richText1.Append(bodyProperties1);
            richText1.Append(listStyle1);
            richText1.Append(paragraph1);

            chartText1.Append(richText1);
            Layout layout1 = new Layout();
            Overlay overlay1 = new Overlay() {Val = false};

            title1.Append(chartText1);
            title1.Append(layout1);
            title1.Append(overlay1);
            return title1;
        }
    }
}