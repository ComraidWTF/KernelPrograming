// =============================================================================
// ChartExcelExporter
// =============================================================================
//
// Exports a LiveCharts2 CartesianChart (with ISeries<TPoint> line series) into
// an .xlsx containing the raw data and a native Excel scatter chart with:
//   - Series line colors matched to the on-screen LiveCharts2 colors
//   - Legend using the real LiveCharts2 series names
//   - Per-point data labels ("Task1", "Task2"...) drawn on each point
//   - Optional custom string labels on the X axis (via secondary axis trick)
//   - Axis titles and min/max
//
// NuGet packages required:
//   dotnet add package ClosedXML
//   dotnet add package DocumentFormat.OpenXml
// Both are MIT licensed and run on .NET 6+ / .NET Standard 2.0.
//
// -----------------------------------------------------------------------------
// USAGE
// -----------------------------------------------------------------------------
//
// Assuming you have:
//   class CustomData { public double X; public double Y; public string TaskName; }
//   CartesianChart myChart;   // your on-screen LiveCharts2 chart
//
// Minimal call — numeric axes, per-point task names on the chart:
//
//     ChartExcelExporter.ExportLineChart<CustomData>(
//         myChart,
//         @"C:\temp\chart.xlsx",
//         p => p.X,
//         p => p.Y,
//         labelSelector: p => p.TaskName);
//
// Full call — axis titles, min/max, and custom X-axis tick labels:
//
//     ChartExcelExporter.ExportLineChart<CustomData>(
//         myChart,
//         @"C:\temp\chart.xlsx",
//         p => p.X,
//         p => p.Y,
//         xAxis: new ChartExcelExporter.AxisConfig
//         {
//             Min = 0,
//             Max = 5,
//             Title = "Day",
//             // Index in the array = tick position. Leave "" for unlabeled ticks.
//             Labels = new[] { "", "Mon", "Tue", "Wed", "Thu", "Fri" }
//         },
//         yAxis: new ChartExcelExporter.AxisConfig
//         {
//             Min = 0,
//             Max = 9,
//             Title = "Severity"
//             // Omit Labels to keep the Y axis numeric.
//         },
//         labelSelector: p => p.TaskName,
//         chartTitle: "Task Severity Over Time");
//
// Notes on parameters:
//   - xSelector / ySelector: tell the exporter how to read X and Y off your
//     point type. No reflection guessing — you pass lambdas.
//   - labelSelector: optional. When provided, each point gets a data label
//     drawn on the chart using the returned string. Points whose label is
//     null/empty render without text.
//   - xAxis.Labels / yAxis.Labels: optional. When provided, a secondary
//     category axis is added to display those strings instead of numbers.
//     Without this, axis ticks are numeric.
//   - chartTitle: optional. Shown above the plot area.
//
// If something doesn't render the way you expect, see the notes at the
// bottom of this file about the secondary-axis label trick.
// =============================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media;
using ClosedXML.Excel;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView.WPF;

using C  = DocumentFormat.OpenXml.Drawing.Charts;
using A  = DocumentFormat.OpenXml.Drawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

public static class ChartExcelExporter
{
    public sealed class AxisConfig
    {
        /// <summary>Ordered list of tick labels. Position 0 -> labels[0], 1 -> labels[1], ...</summary>
        public IList<string> Labels { get; set; }
        /// <summary>Shown as the axis title.</summary>
        public string Title { get; set; }
        /// <summary>Optional min / max for the numeric value axis.</summary>
        public double? Min { get; set; }
        public double? Max { get; set; }
    }

    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector,
        AxisConfig xAxis = null,
        AxisConfig yAxis = null,
        Func<TPoint, string> labelSelector = null,
        string chartTitle = "Exported Chart")
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));

        // 1. Collect series from LiveCharts2
        var lc2 = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .Select((s, i) => new SeriesData
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {i + 1}" : s.Name,
                Points = (s.Values as IEnumerable)?.OfType<TPoint>().ToList() ?? new List<TPoint>(),
                Color = GetStrokeColor(s),
                Xs = new List<double>(),
                Ys = new List<double>(),
                Labels = new List<string>()
            })
            .Where(s => s.Points.Count > 0)
            .ToList();

        if (lc2.Count == 0) return;

        foreach (var s in lc2)
        {
            foreach (var p in s.Points)
            {
                s.Xs.Add(xSelector(p));
                s.Ys.Add(ySelector(p));
                s.Labels.Add(labelSelector?.Invoke(p) ?? string.Empty);
            }
        }

        // 2. Build the workbook with ClosedXML (data only — chart added later via OpenXML)
        using (var wb = new XLWorkbook())
        {
            var ws = wb.Worksheets.Add("Chart Data");

            bool hasLabels = labelSelector != null;
            int colsPerSeries = hasLabels ? 3 : 2;

            for (int i = 0; i < lc2.Count; i++)
            {
                int xCol = colsPerSeries * i + 1;   // ClosedXML is 1-based
                int yCol = xCol + 1;
                int labelCol = xCol + 2;

                ws.Cell(1, xCol).Value = $"{lc2[i].Name} X";
                ws.Cell(1, yCol).Value = $"{lc2[i].Name} Y";
                if (hasLabels) ws.Cell(1, labelCol).Value = $"{lc2[i].Name} Label";

                for (int r = 0; r < lc2[i].Xs.Count; r++)
                {
                    ws.Cell(r + 2, xCol).Value = lc2[i].Xs[r];
                    ws.Cell(r + 2, yCol).Value = lc2[i].Ys[r];
                    if (hasLabels) ws.Cell(r + 2, labelCol).Value = lc2[i].Labels[r];
                }
            }

            ws.Columns().AdjustToContents();
            wb.SaveAs(filePath);
        }

        // 3. Re-open with Open XML SDK and add the chart
        AddScatterChart(filePath, lc2, xAxis, yAxis, chartTitle,
            hasLabels: labelSelector != null,
            colsPerSeries: labelSelector != null ? 3 : 2);
    }

    // ---- Chart construction via Open XML SDK --------------------------------

    private static void AddScatterChart(
        string filePath, List<SeriesData> lc2,
        AxisConfig xAxis, AxisConfig yAxis, string chartTitle,
        bool hasLabels, int colsPerSeries)
    {
        using var doc = SpreadsheetDocument.Open(filePath, true);
        var wbPart = doc.WorkbookPart;
        var wsPart = wbPart.WorksheetParts.First();
        var sheetName = "Chart Data";

        // Create a DrawingsPart + ChartPart
        var drawingsPart = wsPart.DrawingsPart ?? wsPart.AddNewPart<DrawingsPart>();
        if (wsPart.Worksheet.Elements<Drawing>().FirstOrDefault() == null)
        {
            var rid = wsPart.GetIdOfPart(drawingsPart);
            wsPart.Worksheet.Append(new Drawing { Id = rid });
            wsPart.Worksheet.Save();
        }

        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(lc2, sheetName, colsPerSeries, hasLabels,
            xAxis, yAxis, chartTitle);
        chartPart.ChartSpace.Save();

        // Anchor the chart in the drawings part
        BuildDrawing(drawingsPart, chartPart, anchorRow: ComputeAnchorRow(lc2), anchorCol: 0);
    }

    private static int ComputeAnchorRow(List<SeriesData> lc2) =>
        lc2.Max(s => s.Xs.Count) + 3;

    private static C.ChartSpace BuildChartSpace(
        List<SeriesData> lc2, string sheetName, int colsPerSeries, bool hasLabels,
        AxisConfig xAxis, AxisConfig yAxis, string chartTitle)
    {
        // Unique axis IDs for primary and secondary axes
        const uint primaryValX = 111111U;
        const uint primaryValY = 222222U;
        const uint secondaryCatX = 333333U;
        const uint secondaryCatY = 444444U;

        var scatterChart = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false });

        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = colsPerSeries * i + 1;
            int yCol = xCol + 1;
            int labelCol = xCol + 2;
            int lastDataRow = 1 + lc2[i].Xs.Count; // header + N data rows

            var ser = new C.ScatterChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                SeriesText(lc2[i].Name),
                SeriesLineColor(lc2[i].Color),
                SeriesMarker(lc2[i].Color));

            if (hasLabels)
                ser.Append(BuildDataLabelsFromCells(sheetName, labelCol, lastDataRow));

            ser.Append(BuildXValues(sheetName, xCol, lastDataRow));
            ser.Append(BuildYValues(sheetName, yCol, lastDataRow));
            ser.Append(new C.Smooth { Val = false });

            scatterChart.Append(ser);
        }

        scatterChart.Append(new C.AxisId { Val = primaryValX });
        scatterChart.Append(new C.AxisId { Val = primaryValY });

        // Primary value axes (these host the scatter data)
        var valX = BuildValueAxis(primaryValX, primaryValY, C.AxisPositionValues.Bottom, xAxis,
            hideLabels: xAxis?.Labels != null && xAxis.Labels.Count > 0);
        var valY = BuildValueAxis(primaryValY, primaryValX, C.AxisPositionValues.Left, yAxis,
            hideLabels: yAxis?.Labels != null && yAxis.Labels.Count > 0);

        var plotArea = new C.PlotArea(new C.Layout());
        plotArea.Append(scatterChart);
        plotArea.Append(valX);
        plotArea.Append(valY);

        // Secondary category axes ONLY when custom labels provided. The
        // secondary axis displays the custom strings; the primary axis
        // (which actually positions the data) hides its labels.
        if (xAxis?.Labels != null && xAxis.Labels.Count > 0)
        {
            var secondaryBar = BuildHiddenBarSeries(lc2.Count, xAxis.Labels.Count, secondaryCatX, secondaryCatY);
            var catX = BuildCategoryAxis(secondaryCatX, secondaryCatY, C.AxisPositionValues.Bottom, xAxis);
            var valYHidden = BuildHiddenValueAxis(secondaryCatY, secondaryCatX, C.AxisPositionValues.Left);
            plotArea.Append(secondaryBar);
            plotArea.Append(catX);
            plotArea.Append(valYHidden);
        }

        var chart = new C.Chart(
            new C.Title(
                new C.ChartText(new C.RichText(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(new A.Text(chartTitle))))),
                new C.Overlay { Val = false }),
            new C.AutoTitleDeleted { Val = false },
            plotArea,
            new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Right },
                new C.Overlay { Val = false }),
            new C.PlotVisibleOnly { Val = true },
            new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        return new C.ChartSpace(
            new C.EditingLanguage { Val = "en-US" },
            chart)
        { Namespaces = { { "c", "http://schemas.openxmlformats.org/drawingml/2006/chart" } } };
    }

    // ---- Series building blocks --------------------------------------------

    private static C.SeriesText SeriesText(string name) =>
        new C.SeriesText(new C.NumericValue(name));

    private static C.ChartShapeProperties SeriesLineColor(Color? c)
    {
        var sp = new C.ChartShapeProperties();
        var line = new A.Outline(new A.SolidFill(ToRgbFill(c)));
        line.Width = 19050; // 1.5pt
        sp.Append(line);
        return sp;
    }

    private static C.Marker SeriesMarker(Color? c) => new C.Marker(
        new C.Symbol { Val = C.MarkerStyleValues.Circle },
        new C.Size { Val = 6 },
        new C.ChartShapeProperties(
            new A.SolidFill(ToRgbFill(c)),
            new A.Outline(new A.SolidFill(ToRgbFill(c)))));

    private static A.RgbColorModelHex ToRgbFill(Color? c)
    {
        var col = c ?? Color.FromRgb(0x44, 0x88, 0xCC);
        return new A.RgbColorModelHex { Val = $"{col.R:X2}{col.G:X2}{col.B:X2}" };
    }

    private static C.XValues BuildXValues(string sheet, int col, int lastRow) =>
        new C.XValues(new C.NumberReference(
            new C.Formula(CellRef(sheet, col, 2, col, lastRow)),
            new C.NumberingCache(new C.FormatCode("General"),
                new C.PointCount { Val = (uint)(lastRow - 1) })));

    private static C.YValues BuildYValues(string sheet, int col, int lastRow) =>
        new C.YValues(new C.NumberReference(
            new C.Formula(CellRef(sheet, col, 2, col, lastRow)),
            new C.NumberingCache(new C.FormatCode("General"),
                new C.PointCount { Val = (uint)(lastRow - 1) })));

    // ---- Data labels from cells --------------------------------------------

    private static C.DataLabels BuildDataLabelsFromCells(string sheet, int col, int lastRow)
    {
        int pointCount = lastRow - 1;

        var dLbls = new C.DataLabels();
        for (int i = 0; i < pointCount; i++)
        {
            dLbls.Append(new C.DataLabel(
                new C.Index { Val = (uint)i },
                new C.Tx(new C.StringReference(
                    new C.Formula(CellRef(sheet, col, i + 2, col, i + 2)),
                    new C.StringCache(new C.PointCount { Val = 1 }))),
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false }));
        }

        dLbls.Append(new C.ShowLegendKey { Val = false });
        dLbls.Append(new C.ShowValue { Val = false });
        dLbls.Append(new C.ShowCategoryName { Val = false });
        dLbls.Append(new C.ShowSeriesName { Val = false });
        dLbls.Append(new C.ShowPercent { Val = false });
        dLbls.Append(new C.ShowBubbleSize { Val = false });
        return dLbls;
    }

    // ---- Axes ---------------------------------------------------------------

    private static C.ValueAxis BuildValueAxis(uint id, uint crossId,
        C.AxisPositionValues position, AxisConfig cfg, bool hideLabels)
    {
        var scaling = new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax });
        if (cfg?.Min.HasValue == true) scaling.Append(new C.MinAxisValue { Val = cfg.Min.Value });
        if (cfg?.Max.HasValue == true) scaling.Append(new C.MaxAxisValue { Val = cfg.Max.Value });

        var axis = new C.ValueAxis(
            new C.AxisId { Val = id },
            scaling,
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position });

        if (!string.IsNullOrWhiteSpace(cfg?.Title) && !hideLabels)
            axis.Append(BuildAxisTitle(cfg.Title));

        axis.Append(new C.MajorTickMark { Val = C.TickMarkValues.Out });
        axis.Append(new C.MinorTickMark { Val = C.TickMarkValues.None });
        axis.Append(new C.TickLabelPosition
        {
            Val = hideLabels ? C.TickLabelPositionValues.None : C.TickLabelPositionValues.NextTo
        });
        axis.Append(new C.CrossingAxis { Val = crossId });
        axis.Append(new C.Crosses { Val = C.CrossesValues.AutoZero });
        axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });

        return axis;
    }

    private static C.CategoryAxis BuildCategoryAxis(uint id, uint crossId,
        C.AxisPositionValues position, AxisConfig cfg)
    {
        var axis = new C.CategoryAxis(
            new C.AxisId { Val = id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = position });

        if (!string.IsNullOrWhiteSpace(cfg?.Title))
            axis.Append(BuildAxisTitle(cfg.Title));

        axis.Append(new C.MajorTickMark { Val = C.TickMarkValues.Out });
        axis.Append(new C.MinorTickMark { Val = C.TickMarkValues.None });
        axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        axis.Append(new C.CrossingAxis { Val = crossId });
        axis.Append(new C.Crosses { Val = C.CrossesValues.AutoZero });
        axis.Append(new C.AutoLabeled { Val = false });
        axis.Append(new C.LabelAlignment { Val = C.LabelAlignmentValues.Center });
        axis.Append(new C.LabelOffset { Val = 100 });
        axis.Append(new C.NoMultiLevelLabels { Val = true });

        return axis;
    }

    private static C.ValueAxis BuildHiddenValueAxis(uint id, uint crossId, C.AxisPositionValues position) =>
        new C.ValueAxis(
            new C.AxisId { Val = id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = true },
            new C.AxisPosition { Val = position },
            new C.CrossingAxis { Val = crossId });

    private static C.Title BuildAxisTitle(string text) =>
        new C.Title(
            new C.ChartText(new C.RichText(
                new A.BodyProperties { Rotation = 0, Vertical = A.TextVerticalValues.Horizontal },
                new A.ListStyle(),
                new A.Paragraph(new A.Run(new A.Text(text))))),
            new C.Overlay { Val = false });

    // ---- Hidden bar series to drive the secondary category axis ------------
    // We need a bar chart in the same plot area so the CategoryAxis has
    // something to bind its labels to. We make it invisible.

    private static C.BarChart BuildHiddenBarSeries(int seriesIndexBase, int labelCount,
        uint catAxisId, uint valAxisId)
    {
        var bar = new C.BarChart(
            new C.BarDirection { Val = C.BarDirectionValues.Column },
            new C.BarGrouping { Val = C.BarGroupingValues.Clustered },
            new C.VaryColors { Val = false });

        var ser = new C.BarChartSeries(
            new C.Index { Val = (uint)(seriesIndexBase + 1000) },
            new C.Order { Val = (uint)(seriesIndexBase + 1000) },
            new C.SeriesText(new C.NumericValue(" ")),
            InvisibleFill());

        var catRef = new C.CategoryAxisData();
        var catLit = new C.StringLiteral(new C.PointCount { Val = (uint)labelCount });
        // labels filled in by caller through AxisConfig — we need that here
        ser.Append(catRef);
        ser.Append(BuildZeroValues(labelCount));
        bar.Append(ser);

        bar.Append(new C.GapWidth { Val = 150 });
        bar.Append(new C.AxisId { Val = catAxisId });
        bar.Append(new C.AxisId { Val = valAxisId });
        return bar;
    }

    private static C.Values BuildZeroValues(int count)
    {
        var lit = new C.NumberLiteral(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)count });
        for (int i = 0; i < count; i++)
            lit.Append(new C.NumericPoint(new C.NumericValue("0")) { Index = (uint)i });
        return new C.Values(lit);
    }

    private static C.ChartShapeProperties InvisibleFill() =>
        new C.ChartShapeProperties(new A.NoFill(), new A.Outline(new A.NoFill()));

    // ---- Cell refs ---------------------------------------------------------

    private static string CellRef(string sheet, int col1, int row1, int col2, int row2) =>
        $"'{sheet}'!${ColumnLetter(col1)}${row1}:${ColumnLetter(col2)}${row2}";

    private static string ColumnLetter(int col)
    {
        string r = "";
        while (col > 0)
        {
            int m = (col - 1) % 26;
            r = (char)('A' + m) + r;
            col = (col - m - 1) / 26;
        }
        return r;
    }

    // ---- Drawing anchor ----------------------------------------------------

    private static void BuildDrawing(DrawingsPart drawingsPart, ChartPart chartPart,
        int anchorRow, int anchorCol)
    {
        drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing(
            new OpenXmlAttribute("xmlns:xdr", null,
                "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"),
            new OpenXmlAttribute("xmlns:a", null,
                "http://schemas.openxmlformats.org/drawingml/2006/main"));

        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId(anchorCol.ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId(anchorRow.ToString()),
                new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId((anchorCol + 10).ToString()),
                new Xdr.ColumnOffset("0"),
                new Xdr.RowId((anchorRow + 20).ToString()),
                new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Chart" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(
                    new A.Offset { X = 0, Y = 0 },
                    new A.Extents { Cx = 0, Cy = 0 }),
                new A.Graphic(new A.GraphicData(
                    new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) })
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })),
            new Xdr.ClientData());

        drawingsPart.WorksheetDrawing.Append(anchor);
        drawingsPart.WorksheetDrawing.Save();
    }

    // ---- LiveCharts2 color extraction --------------------------------------

    private static Color? GetStrokeColor(ISeries series)
    {
        var paint = series.GetType().GetProperty("Stroke")?.GetValue(series);
        if (paint == null) return null;

        var sk = paint.GetType().GetProperty("Color")?.GetValue(paint);
        if (sk == null) return null;

        var t = sk.GetType();
        byte Read(string name, byte fallback) =>
            t.GetProperty(name)?.GetValue(sk) is byte b ? b : fallback;

        return Color.FromArgb(
            Read("Alpha", 255),
            Read("Red", 0),
            Read("Green", 0),
            Read("Blue", 0));
    }

    private sealed class SeriesData
    {
        public string Name { get; set; }
        public IList Points { get; set; }
        public Color? Color { get; set; }
        public List<double> Xs { get; set; }
        public List<double> Ys { get; set; }
        public List<string> Labels { get; set; }
    }
}
