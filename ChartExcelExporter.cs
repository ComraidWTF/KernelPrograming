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
    /// <summary>
    /// Exports LiveCharts2 line series to an .xlsx with a native Excel scatter
    /// chart. Uses a true numeric X axis so fractional X values (e.g. 0.33, 1)
    /// plot proportionally, matching what LiveCharts2 shows on screen.
    /// Each series writes its own X and Y columns so different series may have
    /// different X values and different point counts.
    /// </summary>
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

        // 1. Pull the LC2 line series we care about
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

        // 2. Layout: per-series X/Y column pairs
        //    For series i: col 2i = "<Name> X", col 2i+1 = "<Name> Y"
        int maxRows = lc2.Max(s => s.Points.Count);

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

        // 3. Seed a Scatter chart. We'll overwrite its series.
        int totalCols = 2 * lc2.Count;
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

        // 4. Clear auto-generated series and rebuild each one with explicit
        //    X and Y ranges + a text title for the legend.
        var group = (ScatterSeriesGroup)chartShape.Chart.SeriesGroups.First();
        while (group.Series.Count > 0)
            group.Series.Remove(group.Series.First());

        for (int i = 0; i < lc2.Count; i++)
        {
            int xCol = 2 * i;
            int yCol = 2 * i + 1;
            int lastRow = lc2[i].Points.Count; // inclusive

            var xRange = new CellRange(1, xCol, lastRow, xCol);
            var yRange = new CellRange(1, yCol, lastRow, yCol);

            var xData = new WorkbookFormulaChartData(sheet, xRange);
            var yData = new WorkbookFormulaChartData(sheet, yRange);

            // Add via SeriesCollection.Add — in a ScatterSeriesGroup this
            // creates a ScatterSeries where the first IChartData is the
            // X values and the second is the Y values.
            Title title = new TextTitle(lc2[i].Name);
            var added = group.Series.Add(xData, yData, title);

            // Color the line + marker
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

        // 5. Chart title + legend
        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        sheet.Charts.Add(chartShape);

        // 6. Save
        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
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
