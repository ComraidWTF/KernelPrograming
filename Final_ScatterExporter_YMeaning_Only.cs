using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using LiveChartsCore.SkiaSharpView;
using SkiaSharp;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;

namespace WpfChartExcelExport
{
    public static class LiveCharts2ScatterWorksheetExporter
    {
        public static void PopulateWorksheet<TRData>(
            Worksheet worksheet,
            IEnumerable<LineSeries<TRData>> series,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xLabelSelector = null,
            string? chartTitle = null,
            string? xAxisTitle = null,
            string? yAxisTitle = null,
            int startRow = 0,
            int startColumn = 0,
            bool includeXLabelColumn = true,
            IComparer<double>? xComparer = null)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var source = series.ToList();
            if (source.Count == 0) throw new InvalidOperationException("No series supplied.");

            var mapped = MapSeries(source, xSelector, ySelector, xLabelSelector);
            var xValues = BuildOrderedX(mapped, xComparer);

            // Only change requested: add X/Y meaning columns to dataset.
            int xValueColumn = startColumn;
            int xMeaningColumn = startColumn + 1;
            int yMeaningColumn = startColumn + 2;
            int firstSeriesColumn = startColumn + 3;

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            WriteHeaderRow(
                worksheet,
                mapped,
                xAxisTitle,
                yAxisTitle,
                xValueColumn,
                xMeaningColumn,
                yMeaningColumn,
                firstSeriesColumn,
                headerRow);

            WriteDataRows(
                worksheet,
                mapped,
                xValues,
                includeXLabelColumn,
                xValueColumn,
                xMeaningColumn,
                yMeaningColumn,
                firstSeriesColumn,
                dataRow);

            int lastRow = dataRow + xValues.Count - 1;
            int chartRow = lastRow + 2;

            CreateScatterChart(
                worksheet,
                mapped,
                dataRow,
                lastRow,
                xValueColumn,
                firstSeriesColumn,
                chartRow,
                startColumn,
                chartTitle,
                xAxisTitle,
                yAxisTitle);
        }

        private static void WriteHeaderRow(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            string? xAxisTitle,
            string? yAxisTitle,
            int xValueColumn,
            int xMeaningColumn,
            int yMeaningColumn,
            int firstSeriesColumn,
            int headerRow)
        {
            worksheet.Cells[headerRow, xValueColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxisTitle) ? "X Value" : xAxisTitle);

            worksheet.Cells[headerRow, xMeaningColumn].SetValue("X Meaning");
            worksheet.Cells[headerRow, yMeaningColumn]
                .SetValue(string.IsNullOrWhiteSpace(yAxisTitle) ? "Y Meaning" : yAxisTitle);

            for (int i = 0; i < mapped.Count; i++)
            {
                worksheet.Cells[headerRow, firstSeriesColumn + i].SetValue(mapped[i].Name);
            }
        }

        private static void WriteDataRows(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            IReadOnlyList<double> xValues,
            bool includeXLabelColumn,
            int xValueColumn,
            int xMeaningColumn,
            int yMeaningColumn,
            int firstSeriesColumn,
            int dataRow)
        {
            var lookups = mapped
                .Select(s => s.Points
                    .GroupBy(p => p.X)
                    .ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int r = 0; r < xValues.Count; r++)
            {
                double x = xValues[r];
                int row = dataRow + r;

                worksheet.Cells[row, xValueColumn].SetValue(x);

                bool isWhole = Math.Abs(x - Math.Round(x)) < 1e-9;

                if (includeXLabelColumn && isWhole)
                {
                    string? label = ResolveLabelForX(mapped, x);
                    if (!string.IsNullOrWhiteSpace(label))
                        worksheet.Cells[row, xMeaningColumn].SetValue(label);
                }

                // Only add Y meaning column; keep chart behavior unchanged.
                if (isWhole)
                {
                    worksheet.Cells[row, yMeaningColumn].SetValue(string.Empty);
                }

                for (int s = 0; s < mapped.Count; s++)
                {
                    int yCol = firstSeriesColumn + s;

                    if (lookups[s].TryGetValue(x, out var point) && point.Y.HasValue)
                        worksheet.Cells[row, yCol].SetValue(point.Y.Value);
                    else
                        worksheet.Cells[row, yCol].SetValueAsFormula("=NA()");
                }
            }
        }

        private static void CreateScatterChart(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            int dataRow,
            int lastRow,
            int xValueColumn,
            int firstSeriesColumn,
            int chartRow,
            int chartColumn,
            string? chartTitle,
            string? xAxisTitle,
            string? yAxisTitle)
        {
            var seedRange = new CellRange(dataRow, xValueColumn, lastRow, firstSeriesColumn);

            var shape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, chartColumn),
                seedRange,
                ChartType.Scatter)
            {
                Width = 1000,
                Height = 600
            };

            var chart = shape.Chart;

            chart.Legend = new Legend { Position = LegendPosition.Right };

            if (!string.IsNullOrWhiteSpace(chartTitle))
                chart.Title = new TextTitle(chartTitle);

            TrySetAxisTitle(chart, xAxisTitle, true);
            TrySetAxisTitle(chart, yAxisTitle, false);

            TryConfigureWholeNumberAxis(chart, true, 1, 1, null);
            TryConfigureWholeNumberAxis(chart, false, 1, 0, null);

            var group = chart.SeriesGroups.First();
            while (group.Series.Count > 0)
                group.Series.Remove(group.Series.First());

            var xVals = new WorkbookFormulaChartData(
                worksheet,
                new CellRange(dataRow, xValueColumn, lastRow, xValueColumn));

            for (int i = 0; i < mapped.Count; i++)
            {
                int yCol = firstSeriesColumn + i;

                var yVals = new WorkbookFormulaChartData(
                    worksheet,
                    new CellRange(dataRow, yCol, lastRow, yCol));

                var added = group.Series.Add(xVals, yVals, new TextTitle(mapped[i].Name));

                // Restore previous scatter look.
                if (added is ScatterSeries s)
                {
                    s.ScatterStyle = ScatterStyle.LineMarker;
                    s.Marker = new Marker { Size = 7, Symbol = MarkerStyle.Circle };
                }
            }

            worksheet.Charts.Add(shape);
        }

        private static void TryConfigureWholeNumberAxis(object chart, bool isX, double? major, double? min, double? max)
        {
            try
            {
                var axes = chart.GetType().GetProperty("PrimaryAxes")?.GetValue(chart);
                var axis = axes?.GetType()
                    .GetProperty(isX ? "CategoryAxis" : "ValueAxis")
                    ?.GetValue(axes);

                var t = axis?.GetType();
                if (t == null) return;

                t.GetProperty("NumberFormat")?.SetValue(axis, "0");

                if (major.HasValue) t.GetProperty("MajorUnit")?.SetValue(axis, major);
                if (min.HasValue) t.GetProperty("Minimum")?.SetValue(axis, min);
                if (max.HasValue) t.GetProperty("Maximum")?.SetValue(axis, max);
            }
            catch { }
        }

        private static void TrySetAxisTitle(object chart, string? title, bool isX)
        {
            if (string.IsNullOrWhiteSpace(title)) return;

            try
            {
                var axes = chart.GetType().GetProperty("PrimaryAxes")?.GetValue(chart);
                var axis = axes?.GetType()
                    .GetProperty(isX ? "CategoryAxis" : "ValueAxis")
                    ?.GetValue(axes);

                axis?.GetType().GetProperty("Title")
                    ?.SetValue(axis, new TextTitle(title));
            }
            catch { }
        }

        private static List<MappedSeries> MapSeries<TRData>(
            List<LineSeries<TRData>> input,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xLabelSelector)
        {
            return input.Select(s => new MappedSeries
            {
                Name = s.Name ?? "Series",
                Points = s.Values?.Select(v => new MappedPoint(
                    xSelector(v),
                    ySelector(v),
                    xLabelSelector?.Invoke(v))).ToList()
                    ?? new List<MappedPoint>()
            }).ToList();
        }

        private static List<double> BuildOrderedX(List<MappedSeries> s, IComparer<double>? c)
        {
            var vals = s.SelectMany(x => x.Points).Select(p => p.X).Distinct();
            return c != null ? vals.OrderBy(x => x, c).ToList() : vals.OrderBy(x => x).ToList();
        }

        private static string? ResolveLabelForX(IReadOnlyList<MappedSeries> m, double x)
        {
            return m.SelectMany(s => s.Points).FirstOrDefault(p => p.X == x).Label;
        }

        private class MappedSeries
        {
            public string Name = "";
            public List<MappedPoint> Points = new();
        }

        private record struct MappedPoint(double X, double? Y, string? Label);
    }
}
