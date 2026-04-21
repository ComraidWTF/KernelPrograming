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
        public double? Min { get; set; }
        public double? Max { get; set; }
        /// <summary>Shown as chart axis title AND key-table header.</summary>
        public string Title { get; set; }
        /// <summary>
        /// Optional tick-value -> display-name map, written to a key table
        /// beside the chart. The chart axis itself stays numeric.
        /// </summary>
        public IDictionary<double, string> Labels { get; set; }
    }

    /// <summary>
    /// Exports LiveCharts2 line series into an .xlsx with data + a native
    /// Excel scatter chart. Pure Telerik RadSpreadProcessing.
    ///
    /// labelSelector, if provided, writes a "&lt;Series&gt; Label" column per series
    /// so the text exists in the workbook. Making those appear ON the chart
    /// requires a one-time Excel click per series
    /// (Add Data Labels -> Value From Cells -> pick the Label column), because
    /// Telerik's ScatterSeries in this version does not expose a DataLabels API.
    /// </summary>
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector,
        AxisConfig xAxis = null,
        AxisConfig yAxis = null,
        Func<TPoint, string> labelSelector = null)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

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

        bool hasLabels = labelSelector != null;
        int colsPerSeries = hasLabels ? 3 : 2;
        int maxRows = lc2.Max(s => s.Points.Count);

        // 1. Lay out data: X, Y, (optional Label) per series
        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = colsPerSeries * i;
            int yCol = xCol + 1;
            int labelCol = xCol + 2;

            sheet.Cells[0, xCol].SetValue($"{lc2[i].Name} X");
            sheet.Cells[0, yCol].SetValue($"{lc2[i].Name} Y");
            if (hasLabels) sheet.Cells[0, labelCol].SetValue($"{lc2[i].Name} Label");

            var pts = lc2[i].Points;
            for (int r = 0; r < pts.Count; r++)
            {
                double x = xSelector(pts[r]);
                double y = ySelector(pts[r]);
                if (!double.IsNaN(x)) sheet.Cells[r + 1, xCol].SetValue(x);
                if (!double.IsNaN(y)) sheet.Cells[r + 1, yCol].SetValue(y);
                if (hasLabels)
                    sheet.Cells[r + 1, labelCol].SetValue(labelSelector(pts[r]) ?? string.Empty);
            }
        }

        int totalCols = colsPerSeries * lc2.Count;

        // 2. Scatter chart seeded with a range; we replace its series below
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

        // 3. Rebuild series explicitly so legend + X/Y bindings are correct
        var group = (ScatterSeriesGroup)chartShape.Chart.SeriesGroups.First();
        while (group.Series.Count > 0)
            group.Series.Remove(group.Series.First());

        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = colsPerSeries * i;
            int yCol = xCol + 1;
            int lastRow = lc2[i].Points.Count;

            var xData = new WorkbookFormulaChartData(sheet, new CellRange(1, xCol, lastRow, xCol));
            var yData = new WorkbookFormulaChartData(sheet, new CellRange(1, yCol, lastRow, yCol));

            Title title = new TextTitle(lc2[i].Name);
            var added = group.Series.Add(xData, yData, title);

            if (lc2[i].Color is Color c)
            {
                added.Outline.Fill = new SolidFill(c);
                added.Outline.Width = 2;
                if (added is ScatterSeries scatter && scatter.Marker != null)
                    scatter.Marker.Fill = new SolidFill(c);
            }
        }

        // 4. Axis config
        ApplyAxisConfig(chartShape.Chart.PrimaryAxes.CategoryAxis as ValueAxis, xAxis);
        ApplyAxisConfig(chartShape.Chart.PrimaryAxes.ValueAxis, yAxis);

        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        sheet.Charts.Add(chartShape);

        // 5. Axis label key tables next to the chart
        int keyStartRow = maxRows + 3;
        int keyStartCol = totalCols + 1;
        WriteLabelKey(sheet, keyStartRow, keyStartCol, xAxis, "X Axis");
        WriteLabelKey(sheet, keyStartRow, keyStartCol + 3, yAxis, "Y Axis");

        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        // 6. Save — pure Telerik
        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
    }

    private static void ApplyAxisConfig(ValueAxis axis, AxisConfig cfg)
    {
        if (axis == null || cfg == null) return;

        if (cfg.Min.HasValue) axis.Min = cfg.Min.Value;
        if (cfg.Max.HasValue) axis.Max = cfg.Max.Value;

        if (!string.IsNullOrWhiteSpace(cfg.Title))
            axis.Title = new TextTitle(cfg.Title);
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
