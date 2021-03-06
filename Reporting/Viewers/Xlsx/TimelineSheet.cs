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
    internal class TimelineSheet : GraphSheet
    {
        internal override void Create(WorkbookPart workBookPart, Sheets sheets)
        {
            Create(workBookPart, sheets, "Timeline");
        }

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
            DocumentFormat.OpenXml.Drawing.Charts.Chart chart = chartPart.ChartSpace.AppendChild(new DocumentFormat.OpenXml.Drawing.Charts.Chart());

            // Create a new clustered column chart.
            PlotArea plotArea = chart.AppendChild<PlotArea>(new PlotArea());
            plotArea.AppendChild<Layout>(new Layout());
            BarChart barChart =
                plotArea.AppendChild<BarChart>(
                    new BarChart(
                        new BarDirection() { Val = new EnumValue<BarDirectionValues>(BarDirectionValues.Column) },
                        new BarGrouping() { Val = new EnumValue<BarGroupingValues>(BarGroupingValues.Clustered) },
                        new VaryColors() { Val = false }
                        ));

            // Iterate through each key in the Dictionary collection and add the key to the chart Series
            // and add the corresponding value to the chart Values.

            BarChartSeries barChartSeries = barChart.AppendChild<BarChartSeries>(new BarChartSeries(new Index()
            {
                Val = new UInt32Value((uint)0)
            },
                new Order() { Val = new UInt32Value((uint)0) },
                new SeriesText(new NumericValue() { Text = "Timeline" })));

            StringLiteral strLit =
                barChartSeries.AppendChild<CategoryAxisData>(new CategoryAxisData())
                    .AppendChild<StringLiteral>(new StringLiteral());
            strLit.Append(new PointCount() { Val = new UInt32Value((uint)stat.Diffs.Count) });

            NumberLiteral numLit = barChartSeries.AppendChild<DocumentFormat.OpenXml.Drawing.Charts.Values>(
                new DocumentFormat.OpenXml.Drawing.Charts.Values())
                .AppendChild<NumberLiteral>(new NumberLiteral());
            numLit.Append(new FormatCode("General"));
            numLit.Append(new PointCount() { Val = new UInt32Value((uint)stat.Diffs.Count) });

            uint i = 0;
            foreach (var diff in stat.Diffs)
            {
                strLit.AppendChild<StringPoint>(new StringPoint() { Index = new UInt32Value(i) })
                    .Append(new NumericValue(diff.TimeStamp.ToString()));
                numLit.AppendChild<NumericPoint>(new NumericPoint() { Index = new UInt32Value(i) })
                    .Append(new NumericValue(diff.Value.ToString()));
                i++;
            }

            barChart.Append(new AxisId() { Val = new UInt32Value(48650112u) });
            barChart.Append(new AxisId() { Val = new UInt32Value(48672768u) });

            AppendCategoryAxis(plotArea, 48650112u, "Time, ms", 48672768U);
            AppendValueAxis(plotArea, 48672768u, "Duration, ms", 48650112U);

            // Add the chart Legend.
            Legend legend =
                chart.AppendChild<Legend>(
                    new Legend(
                        new LegendPosition() { Val = new EnumValue<LegendPositionValues>(LegendPositionValues.Right) },
                        new Layout()));

            chart.Append(new PlotVisibleOnly() { Val = new BooleanValue(true) });

            // Position the chart on the worksheet using a TwoCellAnchor object.
            drawingsPart.WorksheetDrawing = new WorksheetDrawing();

            AppendGraphicFrame(drawingsPart, chartPart);

            // Save the WorksheetDrawing object.
            drawingsPart.WorksheetDrawing.Save();
        }
    }
}
