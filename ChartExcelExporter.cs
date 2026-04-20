using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows.Media;
using LiveChartsCore;
using LiveChartsCore.Defaults;
using LiveChartsCore.SkiaSharpView.WPF;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;
// If ThemableColor / SolidFill aren't resolved, add:
// using Telerik.Windows.Documents.Media;
// using Telerik.Windows.Documents.Spreadsheet.Theming;

public static class ChartExcelExporter
{
    public static void ExportLineChart(CartesianChart chart, string filePath)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets.Add();
        sheet.Name = "Chart Data";

        // 1. Pull line-like series out of the chart
        var lineSeries = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .ToList();

        if (lineSeries.Count == 0) return;

        // 2. Extract values + stroke color per series
        var columns = lineSeries
            .Select((s, idx) => new SeriesColumn
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {idx + 1}" : s.Name,
                Values = ExtractDoubles(s.Values).ToList(),
                Color = GetStrokeColor(s)
            })
            .ToList();

        int maxLen = columns.Max(c => c.Values.Count);

        // 3. Header row — A = X index, B..N = series names (Telerik reads these for legend titles)
        sheet.Cells[0, 0].SetValue("X");
        for (int i = 0; i < columns.Count; i++)
            sheet.Cells[0, i + 1].SetValue(columns[i].Name);

        // 4. Data rows
        for (int row = 0; row < maxLen; row++)
        {
            sheet.Cells[row + 1, 0].SetValue(row + 1);
            for (int col = 0; col < columns.Count; col++)
            {
                if (row < columns[col].Values.Count)
                    sheet.Cells[row + 1, col + 1].SetValue(columns[col].Values[row]);
            }
        }

        // 5. Insert a native Excel line chart pointing at the block above.
        //    Including row 0 in the range lets Telerik use the headers as series titles,
        //    and including column 0 makes it the category (X) axis.
        var dataRange = new CellRange(0, 0, maxLen, columns.Count);

        var chartShape = new FloatingChartShape(
            sheet,
            new CellIndex(maxLen + 3, 0),
            dataRange,
            ChartType.Line)
        {
            Width = 600,
            Height = 320
        };

        sheet.Charts.Add(chartShape);

        // Correct types: TextTitle and Legend (with LegendPosition enum)
        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        // 6. Match each Excel series color to its LiveCharts2 counterpart.
        //    For line charts, the visible color is Outline.Fill (the stroke), not Fill.
        var group = chartShape.Chart.SeriesGroups.FirstOrDefault();
        if (group != null)
        {
            var excelSeries = group.Series.ToList();
            for (int i = 0; i < excelSeries.Count && i < columns.Count; i++)
            {
                if (columns[i].Color is Color c)
                {
                    excelSeries[i].Outline.Fill = new SolidFill(c);
                    excelSeries[i].Outline.Width = 2; // match a typical LiveCharts stroke thickness
                }
            }
        }

        // 7. Auto-fit the used range in a single call
        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
    }

    /// <summary>
    /// Pulls the stroke color off a LiveCharts2 series via reflection so this
    /// works regardless of which SkiaSharp / LiveCharts2 version is referenced.
    /// Returns null if the series uses a gradient or non-solid paint.
    /// </summary>
    private static Color? GetStrokeColor(ISeries series)
    {
        var paint = series.GetType().GetProperty("Stroke")?.GetValue(series);
        if (paint == null) return null;

        // SolidColorPaint exposes a Color property of type SKColor,
        // which has byte Alpha / Red / Green / Blue properties.
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

    private static IEnumerable<double> ExtractDoubles(IEnumerable values)
    {
        if (values == null) yield break;

        foreach (var v in values)
        {
            switch (v)
            {
                case null: yield return double.NaN; break;
                case double d: yield return d; break;
                case float f: yield return f; break;
                case int i: yield return i; break;
                case long l: yield return l; break;
                case decimal m: yield return (double)m; break;
                case ObservablePoint op: yield return op.Y ?? double.NaN; break;
                case ObservableValue ov: yield return ov.Value ?? double.NaN; break;
                case DateTimePoint dtp: yield return dtp.Value ?? double.NaN; break;
                default:
                    var prop = v.GetType().GetProperty("Value") ?? v.GetType().GetProperty("Y");
                    if (prop?.GetValue(v) is IConvertible c)
                        yield return Convert.ToDouble(c);
                    else
                        yield return double.NaN;
                    break;
            }
        }
    }

    private sealed class SeriesColumn
    {
        public string Name { get; set; } = "";
        public List<double> Values { get; set; } = new();
        public Color? Color { get; set; }
    }
}
