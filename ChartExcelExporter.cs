using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using LiveChartsCore;
using LiveChartsCore.Defaults;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.WPF;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;

public static class ChartExcelExporter
{
    public static void ExportLineChart(CartesianChart chart, string filePath)
    {
        var workbook = new Workbook();
        var sheet = workbook.Worksheets.Add();
        sheet.Name = "Chart Data";

        // 1. Pull the line series out of the chart
        var lineSeries = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .ToList();

        if (lineSeries.Count == 0) return;

        // 2. Extract values per series as doubles (handles double, ObservablePoint, etc.)
        var columns = lineSeries
            .Select(s => new
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {lineSeries.IndexOf(s) + 1}" : s.Name,
                Values = ExtractDoubles(s.Values).ToList()
            })
            .ToList();

        // 3. Header row
        sheet.Cells[0, 0].SetValue("X");
        for (int i = 0; i < columns.Count; i++)
            sheet.Cells[0, i + 1].SetValue(columns[i].Name);

        // 4. Data rows (pad to the longest series)
        int maxLen = columns.Max(c => c.Values.Count);
        for (int row = 0; row < maxLen; row++)
        {
            sheet.Cells[row + 1, 0].SetValue(row + 1);
            for (int col = 0; col < columns.Count; col++)
            {
                if (row < columns[col].Values.Count)
                    sheet.Cells[row + 1, col + 1].SetValue(columns[col].Values[row]);
            }
        }

        // 5. (Optional) Embed a native Excel line chart so users see a graph, not just numbers
        AddExcelLineChart(sheet, columns.Count, maxLen);

        // 6. Auto-fit and save
        for (int c = 0; c <= columns.Count; c++)
            sheet.Columns[c].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, new TimeSpan(0, 0, 30));
    }

    private static IEnumerable<double> ExtractDoubles(IEnumerable values)
    {
        if (values == null) yield break;

        foreach (var v in values)
        {
            switch (v)
            {
                case null:
                    yield return double.NaN; break;
                case double d:
                    yield return d; break;
                case float f:
                    yield return f; break;
                case int i:
                    yield return i; break;
                case long l:
                    yield return l; break;
                case decimal m:
                    yield return (double)m; break;
                case ObservablePoint op:
                    yield return op.Y ?? double.NaN; break;
                case ObservableValue ov:
                    yield return ov.Value ?? double.NaN; break;
                case DateTimePoint dtp:
                    yield return dtp.Value ?? double.NaN; break;
                default:
                    // Fallback: try reflection on a "Value" or "Y" property
                    var prop = v.GetType().GetProperty("Value") ?? v.GetType().GetProperty("Y");
                    if (prop != null && prop.GetValue(v) is IConvertible c)
                        yield return Convert.ToDouble(c);
                    else
                        yield return double.NaN;
                    break;
            }
        }
    }

    private static void AddExcelLineChart(Worksheet sheet, int seriesCount, int rowCount)
    {
        var floatingChart = sheet.Shapes.AddChart(
            new CellRange(rowCount + 3, 0, rowCount + 20, 8),
            new LineChartData());

        var chartData = (LineChartData)floatingChart.ChartGraphicProperties.ChartData;

        // X-axis categories = the X column (A2:A{rowCount+1})
        var categories = new CellRangeChartData(sheet, new CellRange(1, 0, rowCount, 0));

        for (int s = 0; s < seriesCount; s++)
        {
            var titleRef = new CellRangeChartData(sheet, new CellRange(0, s + 1, 0, s + 1));
            var valuesRef = new CellRangeChartData(sheet, new CellRange(1, s + 1, rowCount, s + 1));
            chartData.Series.Add(new LineSeries(titleRef, categories, valuesRef));
        }
    }
}
