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
    internal class HistogramSheet : GraphSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Histogram");
        }

        //https://github.com/OfficeDev/office-content/blob/master/en-us/OpenXMLCon/articles/281776d0-be75-46eb-8fdc-a1f656291175.md
        //Here be dragons
        internal override void AddData(Statistics stat)
        {
            // Add a new drawing to the worksheet.
            DrawingsPart drawingsPart = WorksheetPart.AddNewPart<DrawingsPart>();
            WorksheetPart.Worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing()
            {
                Id = WorksheetPart.GetIdOfPart(drawingsPart)
            });
            WorksheetPart.Worksheet.Save();

            // Add a new chart and set the chart language to English-US.
            var chartPart = CreateChartPart(drawingsPart);
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace
                .AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild(new PlotArea());
            plotArea.AppendChild(new Layout());

            CreateHistogram(plotArea, stat.Percentiles, 0, 48650112U, 48672768U);
            CreateCumulative(plotArea, stat.Percentiles, 1, 438381208U, 438380816U);

            // Add the chart Legend.
            chart.AppendChild(
                new Legend(
                    new LegendPosition() {Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right)},
                    new Layout()));

            chart.Append(new PlotVisibleOnly() {Val = new BooleanValue(true)});

            // Save the chart part.
            chartPart.ChartSpace.Save();

            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            
            AppendGraphicFrame(drawingsPart, chartPart);

            // Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save();
        }

        private void CreateCumulative(PlotArea plotArea, List<PercentileRecord> percentiles, uint index,
            uint categoryAxisId, uint valueAxisId)
        {
            LineChart lineChart = plotArea.AppendChild<LineChart>(new LineChart(
                new ShowMarker() {Val = true},
                new Smooth() {Val = false},
                new Grouping() {Val = GroupingValues.Standard},
                new DataLabels(new ShowLegendKey() {Val = false},
                    new ShowValue() {Val = false},
                    new ShowCategoryName() {Val = false},
                    new ShowSeriesName() {Val = false},
                    new ShowPercent() {Val = false},
                    new ShowBubbleSize() {Val = false})));

            LineChartSeries lineChartSeries = lineChart.AppendChild(
                new LineChartSeries(new Index()
                {
                    Val = new UInt32Value(index),

                },
                    new Order() {Val = new UInt32Value(index)},
                    new SeriesText(new NumericValue() {Text = "Cumulative %"})));

            StringLiteral strLit = lineChartSeries.AppendChild(new CategoryAxisData()).AppendChild(new StringLiteral());
            strLit.Append(new PointCount() {Val = new UInt32Value((uint) percentiles.Count)});

            NumberLiteral numLit =
                lineChartSeries.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Values())
                    .AppendChild(new NumberLiteral());
            numLit.Append(new FormatCode("0.00%"));
            numLit.Append(new PointCount() {Val = new UInt32Value((uint) percentiles.Count)});

            for (uint i = 0; i < percentiles.Count; i++)
            {
                strLit.AppendChild<StringPoint>(new StringPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue(percentiles[(int) i].Value.ToString()));
                numLit.AppendChild<NumericPoint>(new NumericPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue((percentiles[(int) i].Percentile / 100).ToString()));
            }

            lineChart.Append(new AxisId() {Val = new UInt32Value(categoryAxisId)});
            lineChart.Append(new AxisId() {Val = new UInt32Value(valueAxisId)});

            AppendCategoryAxis(plotArea, categoryAxisId, "Time, ms", valueAxisId, false);
            AppendValueAxis(plotArea, valueAxisId, "", categoryAxisId, AxisPositionValues.Right, TickLabelPositionValues.High, false);
        }

        private void CreateHistogram(PlotArea plotArea, List<PercentileRecord> percentiles, uint index, uint categoryAxisId, uint valueAxisId)
        {
            BarChart barChart =
                plotArea.AppendChild<BarChart>(
                    new BarChart(
                        new BarDirection() {Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column)},
                        new BarGrouping() {Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered)},
                        new VaryColors() {Val = false}
                        ));

            BarChartSeries barChartSeries = barChart.AppendChild(new BarChartSeries(new Index()
            {
                Val = new UInt32Value(index)
            },
                new Order() {Val = new UInt32Value(index)},
                new SeriesText(new NumericValue() {Text = "Histogram"})));

            StringLiteral strLit =
                barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData())
                    .AppendChild<StringLiteral>(new StringLiteral());
            strLit.Append(new PointCount() {Val = new UInt32Value((uint) percentiles.Count)});

            NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                new DocumentFormat.OpenXml.Drawing.Charts.Values())
                .AppendChild<NumberLiteral>(new NumberLiteral());
            numLit.Append(new FormatCode("General"));
            numLit.Append(new PointCount() {Val = new UInt32Value((uint) percentiles.Count)});

            for (uint i = 0; i < percentiles.Count; i++)
            {
                strLit.AppendChild<StringPoint>(new StringPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue(percentiles[(int) i].Value.ToString()));
                numLit.AppendChild<NumericPoint>(new NumericPoint() {Index = new UInt32Value(i)})
                    .Append(new NumericValue(percentiles[(int) i].Count.ToString()));
            }

            barChart.Append(new AxisId() { Val = new UInt32Value(categoryAxisId) });
            barChart.Append(new AxisId() { Val = new UInt32Value(valueAxisId) });

            AppendCategoryAxis(plotArea, categoryAxisId, "Time, ms", valueAxisId);
            AppendValueAxis(plotArea, valueAxisId, "Number of samples", categoryAxisId);
        }
    }
}