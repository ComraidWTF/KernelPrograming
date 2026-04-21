using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView;
using SkiaSharp;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;

namespace WpfChartExcelExport
{
    public static class LiveCharts2WorksheetExporter
    {
        public static void PopulateWorksheet<TRData, TX>(
            Worksheet worksheet,
            IEnumerable<LineSeries<TRData>> series,
            Axis? xAxis,
            Axis? yAxis,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0)
            where TX : notnull
        {
            var source = series.ToList();
            var mapped = MapSeries(source, xSelector, ySelector, labelSelector);
            var xValues = BuildX(mapped);

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            // HEADER
            worksheet.Cells[headerRow, startColumn].SetValue(xAxis?.Name ?? "X");

            for (int i = 0; i < mapped.Count; i++)
            {
                worksheet.Cells[headerRow, startColumn + 1 + i].SetValue(mapped[i].Name);
            }

            // DATA
            var lookups = mapped
                .Select(s => s.Points.ToDictionary(p => p.X, p => p))
                .ToList();

            for (int r = 0; r < xValues.Count; r++)
            {
                var x = xValues[r];
                worksheet.Cells[dataRow + r, startColumn].SetValue(x?.ToString());

                for (int s = 0; s < mapped.Count; s++)
                {
                    if (lookups[s].TryGetValue(x, out var p) && p.Y.HasValue)
                    {
                        worksheet.Cells[dataRow + r, startColumn + 1 + s].SetValue(p.Y.Value);
                    }
                }
            }

            int lastRow = dataRow + xValues.Count - 1;
            int chartRow = lastRow + 2;

            CreateChart(
                worksheet,
                mapped,
                headerRow,
                dataRow,
                lastRow,
                startColumn,
                chartRow,
                chartTitle);
        }

        private static void CreateChart<TX>(
            Worksheet ws,
            List<MappedSeries<TX>> series,
            int headerRow,
            int dataRow,
            int lastRow,
            int startCol,
            int chartRow,
            string? title)
            where TX : notnull
        {
            var seed = new CellRange(headerRow, startCol, lastRow, startCol + 1);

            var shape = new FloatingChartShape(
                ws,
                new CellIndex(chartRow, startCol),
                seed,
                ChartType.Line)
            {
                Width = 900,
                Height = 500
            };

            var chart = shape.Chart;

            chart.Legend = new Legend { Position = LegendPosition.Right };

            if (!string.IsNullOrWhiteSpace(title))
                chart.Title = new TextTitle(title);

            var group = chart.SeriesGroups.First();

            // remove auto-generated garbage
            while (group.Series.Count > 0)
                group.Series.Remove(group.Series.First());

            var categories = new WorkbookFormulaChartData(
                ws,
                new CellRange(dataRow, startCol, lastRow, startCol));

            for (int i = 0; i < series.Count; i++)
            {
                int col = startCol + 1 + i;

                var values = new WorkbookFormulaChartData(
                    ws,
                    new CellRange(dataRow, col, lastRow, col));

                var added = group.Series.Add(
                    categories,
                    values,
                    new TextTitle(series[i].Name));

                // optional styling
                try
                {
                    if (added is LineSeries ls)
                    {
                        ls.Outline.Fill = new SolidFill(Hex(series[i].Color));
                    }
                }
                catch { }
            }

            ws.Charts.Add(shape);
        }

        private static List<MappedSeries<TX>> MapSeries<TRData, TX>(
            List<LineSeries<TRData>> input,
            Func<TRData, TX> x,
            Func<TRData, double?> y,
            Func<TRData, string?>? label)
            where TX : notnull
        {
            return input.Select(s => new MappedSeries<TX>
            {
                Name = s.Name ?? "Series",
                Color = GetColor(s.Stroke),
                Points = s.Values?
                    .Select(v => new MappedPoint<TX>(x(v), y(v), label?.Invoke(v)))
                    .ToList()
                    ?? new List<MappedPoint<TX>>()
            }).ToList();
        }

        private static List<TX> BuildX<TX>(List<MappedSeries<TX>> s)
            where TX : notnull
        {
            var set = new HashSet<TX>();
            var list = new List<TX>();

            foreach (var series in s)
            foreach (var p in series.Points)
                if (set.Add(p.X)) list.Add(p.X);

            return list;
        }

        private static string GetColor(object? stroke)
        {
            if (stroke is SolidColorPaint p)
            {
                var c = p.Color;
                return $"#{c.Red:X2}{c.Green:X2}{c.Blue:X2}";
            }
            return "#000000";
        }

        private static ThemableColor Hex(string hex)
        {
            hex = hex.TrimStart('#');
            return new ThemableColor(Color.FromRgb(
                byte.Parse(hex[..2], NumberStyles.HexNumber),
                byte.Parse(hex.Substring(2, 2), NumberStyles.HexNumber),
                byte.Parse(hex.Substring(4, 2), NumberStyles.HexNumber)));
        }

        private class MappedSeries<TX> where TX : notnull
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "#000000";
            public List<MappedPoint<TX>> Points { get; set; } = new();
        }

        private record struct MappedPoint<TX>(TX X, double? Y, string? Label);
    }
}
