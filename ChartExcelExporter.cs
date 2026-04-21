// =============================================================================
// ChartExcelExporter — minimal, reliable
// =============================================================================
//
// Writes LiveCharts2 line-series data into an .xlsx with columns per series:
//   <Series> X   |   <Series> Y   |   <Series> Label
// ...repeated for each series.
//
// No chart is generated automatically. After export, open the file in Excel:
//   1. Select your X/Y columns
//   2. Insert -> Scatter (with Straight Lines)
//   3. (Optional) Right-click series -> Add Data Labels -> Value From Cells
//      -> pick the Label column -> uncheck Y Value.
//   Save. Excel preserves all of the above on re-open.
//
// Why not auto-generate the chart? Open XML SDK's chart namespace is large
// and fragile; generating chart XML reliably needs iteration against a
// compiler. Writing data is simple and rock-solid — and the two clicks in
// Excel give you pixel-perfect control over chart style, axis labels, etc.
//
// NuGet: dotnet add package ClosedXML
// Works on .NET 6/7/8+ and .NET Standard 2.0.
//
// Usage:
//
//   ChartExcelExporter.ExportLineChart<CustomData>(
//       myCartesianChart,
//       @"C:\temp\chart.xlsx",
//       p => p.X,
//       p => p.Y,
//       labelSelector: p => p.TaskName);   // optional
// =============================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using ClosedXML.Excel;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView.WPF;

public static class ChartExcelExporter
{
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector,
        Func<TPoint, string> labelSelector = null)
    {
        if (chart is null) throw new ArgumentNullException(nameof(chart));
        if (xSelector is null) throw new ArgumentNullException(nameof(xSelector));
        if (ySelector is null) throw new ArgumentNullException(nameof(ySelector));

        // Gather LiveCharts2 line series and their points
        var seriesList = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .Select((s, i) => new
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {i + 1}" : s.Name,
                Points = (s.Values as IEnumerable)?.OfType<TPoint>().ToList() ?? new List<TPoint>()
            })
            .Where(s => s.Points.Count > 0)
            .ToList();

        if (seriesList.Count == 0) return;

        bool hasLabels = labelSelector != null;
        int colsPerSeries = hasLabels ? 3 : 2;

        using var wb = new XLWorkbook();
        var ws = wb.Worksheets.Add("Chart Data");

        // Header row + data columns for each series
        for (int i = 0; i < seriesList.Count; i++)
        {
            int xCol = colsPerSeries * i + 1;    // ClosedXML is 1-based
            int yCol = xCol + 1;
            int labelCol = xCol + 2;

            ws.Cell(1, xCol).Value = $"{seriesList[i].Name} X";
            ws.Cell(1, yCol).Value = $"{seriesList[i].Name} Y";
            if (hasLabels)
                ws.Cell(1, labelCol).Value = $"{seriesList[i].Name} Label";

            // Bold the header row for this series
            ws.Cell(1, xCol).Style.Font.Bold = true;
            ws.Cell(1, yCol).Style.Font.Bold = true;
            if (hasLabels) ws.Cell(1, labelCol).Style.Font.Bold = true;

            var points = seriesList[i].Points;
            for (int r = 0; r < points.Count; r++)
            {
                ws.Cell(r + 2, xCol).Value = xSelector(points[r]);
                ws.Cell(r + 2, yCol).Value = ySelector(points[r]);
                if (hasLabels)
                    ws.Cell(r + 2, labelCol).Value = labelSelector(points[r]) ?? string.Empty;
            }
        }

        ws.Columns().AdjustToContents();
        wb.SaveAs(filePath);
    }
}
