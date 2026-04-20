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
    /// Exports LiveCharts2 line series to an .xlsx with a native Excel line chart.
    /// Series share the X axis — X values are taken from the first series.
    /// </summary>
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, object> xSelector,
        Func<TPoint, double> ySelector,
        IList<string> xAxisLabels = null)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

        // 1. Collect the LC2 line series we care about
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

        int rowCount = lc2[0].Points.Count;
        var xValues = lc2[0].Points.Select(p => xSelector(p)).ToList();

        // 2. Create workbook and lay out the data
        var workbook = new Workbook();
        var sheet = workbook.Worksheets.Add();
        sheet.Name = "Chart Data";

        // Header row
        sheet.Cells[0, 0].SetValue("X");
        for (int i = 0; i < lc2.Count; i++)
            sheet.Cells[0, i + 1].SetValue(lc2[i].Name);

        // X column + Y columns
        for (int row = 0; row < rowCount; row++)
        {
            // If labels are provided, column A holds the display label for this
            // step (e.g. "Jan", "Feb"...) — the Excel Line chart will use these
            // directly as X-axis tick labels. Otherwise fall back to the raw X value.
            if (xAxisLabels != null && row < xAxisLabels.Count)
                sheet.Cells[row + 1, 0].SetValue(xAxisLabels[row] ?? string.Empty);
            else
                WriteX(sheet.Cells[row + 1, 0], xValues[row]);

            for (int col = 0; col < lc2.Count; col++)
            {
                var pts = lc2[col].Points;
                if (row < pts.Count)
                {
                    double y = ySelector(pts[row]);
                    if (!double.IsNaN(y)) sheet.Cells[row + 1, col + 1].SetValue(y);
                }
            }
        }

        // 3. Create the FloatingChartShape. The CellRange passed here seeds an
        //    auto-generated chart which we will then completely replace with
        //    our own explicit series, so its exact contents don't matter.
        var seedRange = new CellRange(0, 0, rowCount, lc2.Count);
        var chartShape = new FloatingChartShape(
            sheet,
            new CellIndex(rowCount + 3, 0),
            seedRange,
            ChartType.Line)
        {
            Width = 600,
            Height = 320
        };

        // 4. Get the line series group that was created for us, and CLEAR its
        //    auto-generated series so we can add ours explicitly.
        var group = chartShape.Chart.SeriesGroups.First();
        while (group.Series.Count > 0)
            group.Series.Remove(group.Series.First());

        // 5. Shared X-axis categories = column A, rows 1..rowCount (A2:A{n+1})
        var categoriesRange = new CellRange(1, 0, rowCount, 0);
        var categoriesData = new WorkbookFormulaChartData(sheet, categoriesRange);

        // 6. Add each series explicitly with its own Values range and a
        //    TextTitle so the legend entry is correct.
        for (int i = 0; i < lc2.Count; i++)
        {
            var valuesRange = new CellRange(1, i + 1, rowCount, i + 1);
            var valuesData = new WorkbookFormulaChartData(sheet, valuesRange);

            Title seriesTitle = new TextTitle(lc2[i].Name);
            var added = group.Series.Add(categoriesData, valuesData, seriesTitle);

            // Match LiveCharts2 stroke color
            if (lc2[i].Color is Color c)
            {
                added.Outline.Fill = new SolidFill(c);
                added.Outline.Width = 2;
            }
        }

        // 7. Chart title + legend
        chartShape.Chart.Title = new TextTitle("Exported Chart");
        chartShape.Chart.Legend = new Legend { LegendPosition = LegendPosition.Right };

        sheet.Charts.Add(chartShape);

        // 8. Tidy up and save
        sheet.Columns[sheet.UsedCellRange].AutoFitWidth();

        var provider = new XlsxFormatProvider();
        using var stream = File.Create(filePath);
        provider.Export(workbook, stream, TimeSpan.FromSeconds(30));
    }

    private static void WriteX(CellSelection cell, object x)
    {
        switch (x)
        {
            case null:        cell.SetValue(string.Empty); break;
            case double d:    cell.SetValue(d); break;
            case float f:     cell.SetValue(f); break;
            case int i:       cell.SetValue(i); break;
            case long l:      cell.SetValue((double)l); break;
            case decimal m:   cell.SetValue((double)m); break;
            case DateTime dt: cell.SetValue(dt); break;
            case string s:    cell.SetValue(s); break;
            default:          cell.SetValue(x.ToString()); break;
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
