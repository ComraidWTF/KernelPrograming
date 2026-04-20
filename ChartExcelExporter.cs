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
// If SolidFill / ThemableColor don't resolve, add:
// using Telerik.Windows.Documents.Media;

public static class ChartExcelExporter
{
    /// <summary>
    /// Exports every LiveCharts2 line series of type <typeparamref name="TPoint"/>
    /// on <paramref name="chart"/> to an xlsx file, including a native Excel line
    /// chart whose series colors match the on-screen ones.
    ///
    /// Assumption: all series share the same X axis — X values are taken from the
    /// first series. If series have different X values, use ChartType.Scatter and
    /// build per-series X/Y ranges instead (ask and I'll show that variant).
    /// </summary>
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, object> xSelector,
        Func<TPoint, double> ySelector)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

        // 1. Grab the line series whose point type is TPoint
        var lineSeries = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .Select((s, i) => new
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {i + 1}" : s.Name,
                Points = (s.Values as IEnumerable)?.OfType<TPoint>().ToList() ?? new List<TPoint>(),
                Color = GetStrokeColor(s)
            })
            .Where(x => x.Points.Count > 0)
            .ToList();

        if (lineSeries.Count == 0) return;

        // 2. Use the first series' X values as the shared X axis
        var xValues = lineSeries[0].Points.Select(p => xSelector(p)).ToList();
        int rowCount = xValues.Count;

        var workbook = new Workbook();
        var sheet = workbook.Worksheets.Add();
        sheet.Name = "Chart Data";

        // 3. Header row: A1 = "X", B1..N1 = series names
        sheet.Cells[0, 0].SetValue("X");
        for (int i = 0; i < lineSeries.Count; i++)
            sheet.Cells[0, i + 1].SetValue(lineSeries[i].Name);

        // 4. X column + Y columns
        for (int row = 0; row < rowCount; row++)
        {
            WriteX(sheet.Cells[row + 1, 0], xValues[row]);

            for (int col = 0; col < lineSeries.Count; col++)
            {
                var points = lineSeries[col].Points;
                if (row < points.Count)
                {
                    double y = ySelector(points[row]);
                    if (!double.IsNaN(y)) sheet.Cells[row + 1, col + 1].SetValue(y);
                }
            }
        }

        // 5. Build the chart. IMPORTANT: pass SeriesRangesOrientation.Cols so
        //    Telerik knows series are laid out column-wise and column A is the
        //    X axis — otherwise auto-detection can flip and you get series
        //    names appearing as category labels.
        var dataRange = new CellRange(0, 0, rowCount, lineSeries.Count);

        var chartShape = new FloatingChartShape(
            sheet,
            new CellIndex(rowCount + 3, 0),
            dataRange,
            SeriesRangesOrientation.Cols,
            ChartType.Line)
        {
            Width = 600,
            Height = 320
        };

        sheet.Charts.Add(chartShape);

        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        // 6. Match stroke colors to the on-screen series
        var group = chartShape.Chart.SeriesGroups.FirstOrDefault();
        if (group != null)
        {
            var excelSeries = group.Series.ToList();
            for (int i = 0; i < excelSeries.Count && i < lineSeries.Count; i++)
            {
                if (lineSeries[i].Color is Color c)
                {
                    excelSeries[i].Outline.Fill = new SolidFill(c);
                    excelSeries[i].Outline.Width = 2;
                }
            }
        }

        // 7. Auto-fit and save
        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
    }

    private static void WriteX(CellSelection cell, object x)
    {
        switch (x)
        {
            case null:           cell.SetValue(string.Empty); break;
            case double d:       cell.SetValue(d); break;
            case float f:        cell.SetValue(f); break;
            case int i:          cell.SetValue(i); break;
            case long l:         cell.SetValue((double)l); break;
            case decimal m:      cell.SetValue((double)m); break;
            case DateTime dt:    cell.SetValue(dt); break;
            case string s:       cell.SetValue(s); break;
            default:             cell.SetValue(x.ToString()); break;
        }
    }

    /// <summary>
    /// Reads the stroke color off a LiveCharts2 series via reflection so we don't
    /// take a hard dependency on SkiaSharp. Returns null for non-solid paints.
    /// </summary>
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
