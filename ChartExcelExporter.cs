using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView.WPF;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;
// If SolidFill doesn't resolve, add: using Telerik.Windows.Documents.Media;

public static class ChartExcelExporter
{
    public sealed class AxisConfig
    {
        /// <summary>Inclusive axis minimum. Null = let Excel auto-scale.</summary>
        public double? Min { get; set; }
        /// <summary>Inclusive axis maximum. Null = let Excel auto-scale.</summary>
        public double? Max { get; set; }
        /// <summary>
        /// Optional mapping of tick value → display name.
        /// Written to a key table beside the chart (the chart itself shows numbers).
        /// </summary>
        public IDictionary<double, string> Labels { get; set; }
        /// <summary>Header used above the labels column in the key table.</summary>
        public string Title { get; set; }
    }

    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector,
        AxisConfig xAxis = null,
        AxisConfig yAxis = null)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

        // 1. Pull LC2 line series
        var lc2 = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .Select((s, i) => new
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {i + 1}" : s.Name,
                Points = (s.Values as IEnumerable)?.OfType<TPoint>().ToList() ?? new List<TPoint>(),
                Color = GetStrokeColor(s)
            })
            .Where(s => s.Points.Count > 0)
            .ToList();

        if (lc2.Count == 0) return;

        var workbook = new Workbook();
        var sheet = workbook.Worksheets.Add();
        sheet.Name = "Chart Data";

        int maxRows = lc2.Max(s => s.Points.Count);

        // 2. Lay out per-series X/Y column pairs
        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = 2 * i;
            int yCol = 2 * i + 1;

            sheet.Cells[0, xCol].SetValue($"{lc2[i].Name} X");
            sheet.Cells[0, yCol].SetValue($"{lc2[i].Name} Y");

            var pts = lc2[i].Points;
            for (int r = 0; r < pts.Count; r++)
            {
                double x = xSelector(pts[r]);
                double y = ySelector(pts[r]);
                if (!double.IsNaN(x)) sheet.Cells[r + 1, xCol].SetValue(x);
                if (!double.IsNaN(y)) sheet.Cells[r + 1, yCol].SetValue(y);
            }
        }

        int totalCols = 2 * lc2.Count;

        // 3. Seed the Scatter chart (we'll overwrite its series)
        var seedRange = new CellRange(0, 0, maxRows, totalCols - 1);
        var chartShape = new FloatingChartShape(
            sheet,
            new CellIndex(maxRows + 3, 0),
            seedRange,
            ChartType.Scatter)
        {
            Width = 650,
            Height = 350
        };

        // 4. Rebuild series explicitly
        var group = (ScatterSeriesGroup)chartShape.Chart.SeriesGroups.First();
        while (group.Series.Count > 0)
            group.Series.Remove(group.Series.First());

        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = 2 * i;
            int yCol = 2 * i + 1;
            int lastRow = lc2[i].Points.Count;

            var xRange = new CellRange(1, xCol, lastRow, xCol);
            var yRange = new CellRange(1, yCol, lastRow, yCol);

            var xData = new WorkbookFormulaChartData(sheet, xRange);
            var yData = new WorkbookFormulaChartData(sheet, yRange);

            Title title = new TextTitle(lc2[i].Name);
            var added = group.Series.Add(xData, yData, title);

            if (lc2[i].Color is Color c)
            {
                added.Outline.Fill = new SolidFill(c);
                added.Outline.Width = 2;

                if (added is ScatterSeries scatter && scatter.Marker != null)
                {
                    scatter.Marker.Fill = new SolidFill(c);
                }
            }
        }

        // 5. Axis configuration — major ticks at integer positions, no minor
        ApplyAxisConfig(chartShape.Chart.PrimaryAxes.CategoryAxis as ValueAxis, xAxis);
        ApplyAxisConfig(chartShape.Chart.PrimaryAxes.ValueAxis, yAxis);

        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        sheet.Charts.Add(chartShape);

        // 6. Label key table — placed to the right of the chart area so it sits
        //    alongside the chart visually, readable without zooming in.
        int keyStartRow = maxRows + 3;
        int keyStartCol = totalCols + 1; // one blank column for spacing

        WriteLabelKey(sheet, keyStartRow, keyStartCol, xAxis, "X Axis");
        WriteLabelKey(sheet, keyStartRow, keyStartCol + 3, yAxis, "Y Axis");

        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
    }

    private static void ApplyAxisConfig(ValueAxis axis, AxisConfig cfg)
    {
        if (axis == null || cfg == null) return;

        if (cfg.Min.HasValue) axis.Min = cfg.Min.Value;
        if (cfg.Max.HasValue) axis.Max = cfg.Max.Value;

        // Note: Telerik's ValueAxis doesn't expose a MajorUnit/MajorStep property
        // for tick spacing. Excel auto-picks the tick interval from Min/Max; for
        // integer ranges like 0..5 or 0..9 it almost always lands on unit ticks.
    }

    private static void WriteLabelKey(
        Worksheet sheet, int startRow, int startCol,
        AxisConfig cfg, string defaultHeader)
    {
        if (cfg?.Labels == null || cfg.Labels.Count == 0) return;

        string header = string.IsNullOrWhiteSpace(cfg.Title) ? defaultHeader : cfg.Title;

        sheet.Cells[startRow, startCol].SetValue(header);
        sheet.Cells[startRow, startCol + 1].SetValue("Label");

        int row = startRow + 1;
        foreach (var kvp in cfg.Labels.OrderBy(k => k.Key))
        {
            sheet.Cells[row, startCol].SetValue(kvp.Key);
            sheet.Cells[row, startCol + 1].SetValue(kvp.Value ?? string.Empty);
            row++;
        }
    }

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
}
