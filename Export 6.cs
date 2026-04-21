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
        public static void PopulateWorksheet<TRData>(
            Worksheet worksheet,
            IEnumerable<LineSeries<TRData>> series,
            Axis? xAxis,
            Axis? yAxis,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xLabelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            bool includeLabelColumn = true,
            IComparer<double>? xComparer = null)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var source = series.ToList();
            if (source.Count == 0)
                throw new InvalidOperationException("No series supplied.");

            var mapped = MapSeries(source, xSelector, ySelector, xLabelSelector);
            var xValues = BuildX(mapped, xComparer);

            if (xValues.Count == 0)
                throw new InvalidOperationException("No data points found.");

            int xValueColumn = startColumn;
            int xLabelColumn = includeLabelColumn ? startColumn + 1 : -1;
            int firstSeriesColumn = includeLabelColumn ? startColumn + 2 : startColumn + 1;

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            WriteHeaderRow(
                worksheet,
                mapped,
                xAxis,
                includeLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                headerRow);

            WriteDataRows(
                worksheet,
                mapped,
                xValues,
                includeLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                dataRow);

            ApplyHeaderStyles(
                worksheet,
                mapped,
                includeLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                headerRow);

            SetColumnWidths(
                worksheet,
                mapped.Count,
                includeLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn);

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
                chartTitle);
        }

        public static void PopulateWorksheet<TRData>(
            Worksheet worksheet,
            LineSeries<TRData> series,
            Axis? xAxis,
            Axis? yAxis,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xLabelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            bool includeLabelColumn = true,
            IComparer<double>? xComparer = null)
        {
            if (series == null) throw new ArgumentNullException(nameof(series));

            PopulateWorksheet(
                worksheet,
                new[] { series },
                xAxis,
                yAxis,
                xSelector,
                ySelector,
                xLabelSelector,
                chartTitle,
                startRow,
                startColumn,
                includeLabelColumn,
                xComparer);
        }

        private static void WriteHeaderRow(
            Worksheet worksheet,
            List<MappedSeries> mapped,
            Axis? xAxis,
            bool includeLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn,
            int headerRow)
        {
            worksheet.Cells[headerRow, xValueColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Value" : xAxis!.Name!);

            if (includeLabelColumn)
            {
                worksheet.Cells[headerRow, xLabelColumn].SetValue("X Label");
            }

            for (int i = 0; i < mapped.Count; i++)
            {
                worksheet.Cells[headerRow, firstSeriesColumn + i].SetValue(mapped[i].Name);
            }
        }

        private static void WriteDataRows(
            Worksheet worksheet,
            List<MappedSeries> mapped,
            List<double> xValues,
            bool includeLabelColumn,
            int xValueColumn,
            int xLabelColumn,
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

                if (includeLabelColumn)
                {
                    string? label = ResolveLabelForX(mapped, x);
                    worksheet.Cells[row, xLabelColumn].SetValue(string.IsNullOrWhiteSpace(label) ? x.ToString(CultureInfo.InvariantCulture) : label);
                }

                for (int s = 0; s < mapped.Count; s++)
                {
                    int valueColumn = firstSeriesColumn + s;

                    if (lookups[s].TryGetValue(x, out var point) && point.Y.HasValue)
                    {
                        worksheet.Cells[row, valueColumn].SetValue(point.Y.Value);
                    }
                    else
                    {
                        worksheet.Cells[row, valueColumn].SetValueAsFormula("=NA()");
                    }
                }
            }
        }

        private static string? ResolveLabelForX(List<MappedSeries> mapped, double x)
        {
            foreach (var series in mapped)
            {
                var point = series.Points.LastOrDefault(p => p.X.Equals(x));
                if (!string.IsNullOrWhiteSpace(point.Label))
                    return point.Label;
            }

            return null;
        }

        private static void ApplyHeaderStyles(
            Worksheet worksheet,
            List<MappedSeries> mapped,
            bool includeLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn,
            int headerRow)
        {
            string baseHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);

            worksheet.Cells[headerRow, xValueColumn].SetStyleName(baseHeaderStyle);

            if (includeLabelColumn)
            {
                worksheet.Cells[headerRow, xLabelColumn].SetStyleName(baseHeaderStyle);
            }

            for (int i = 0; i < mapped.Count; i++)
            {
                string style = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: mapped[i].Color,
                    fontHex: IdealTextColor(mapped[i].Color));

                worksheet.Cells[headerRow, firstSeriesColumn + i].SetStyleName(style);
            }
        }

        private static void SetColumnWidths(
            Worksheet worksheet,
            int seriesCount,
            bool includeLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn)
        {
            worksheet.Columns[xValueColumn].SetWidth(new ColumnWidth(90, true));

            if (includeLabelColumn)
            {
                worksheet.Columns[xLabelColumn].SetWidth(new ColumnWidth(180, true));
            }

            for (int i = 0; i < seriesCount; i++)
            {
                worksheet.Columns[firstSeriesColumn + i].SetWidth(new ColumnWidth(110, true));
            }
        }

        private static void CreateScatterChart(
            Worksheet worksheet,
            List<MappedSeries> mapped,
            int dataRow,
            int lastRow,
            int xValueColumn,
            int firstSeriesColumn,
            int chartRow,
            int chartColumn,
            string? title)
        {
            var seedRange = new CellRange(dataRow, xValueColumn, lastRow, firstSeriesColumn);

            var shape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, chartColumn),
                seedRange,
                ChartType.Scatter)
            {
                Width = 900,
                Height = 500
            };

            var chart = shape.Chart;

            chart.Legend = new Legend
            {
                Position = LegendPosition.Right
            };

            if (!string.IsNullOrWhiteSpace(title))
            {
                chart.Title = new TextTitle(title);
            }

            var group = chart.SeriesGroups.First();

            while (group.Series.Count > 0)
            {
                group.Series.Remove(group.Series.First());
            }

            var xValues = new WorkbookFormulaChartData(
                worksheet,
                new CellRange(dataRow, xValueColumn, lastRow, xValueColumn));

            for (int i = 0; i < mapped.Count; i++)
            {
                int yColumn = firstSeriesColumn + i;

                var yValues = new WorkbookFormulaChartData(
                    worksheet,
                    new CellRange(dataRow, yColumn, lastRow, yColumn));

                var added = group.Series.Add(
                    xValues,
                    yValues,
                    new TextTitle(mapped[i].Name));

                try
                {
                    if (added is ScatterSeries scatterSeries)
                    {
                        scatterSeries.Outline.Fill = new SolidFill(HexToThemableColor(mapped[i].Color));
                    }
                    else if (added is Telerik.Windows.Documents.Spreadsheet.Model.Charts.LineSeries lineSeries)
                    {
                        lineSeries.Outline.Fill = new SolidFill(HexToThemableColor(mapped[i].Color));
                    }
                    else if (added is CategorySeriesBase categorySeries)
                    {
                        categorySeries.Outline.Fill = new SolidFill(HexToThemableColor(mapped[i].Color));
                    }
                }
                catch
                {
                }
            }

            worksheet.Charts.Add(shape);
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
                Color = GetColor(s.Stroke),
                Points = s.Values?
                    .Select(v => new MappedPoint(
                        xSelector(v),
                        ySelector(v),
                        xLabelSelector?.Invoke(v)))
                    .ToList()
                    ?? new List<MappedPoint>()
            }).ToList();
        }

        private static List<double> BuildX(
            List<MappedSeries> series,
            IComparer<double>? comparer = null)
        {
            var values = series
                .SelectMany(s => s.Points)
                .Select(p => p.X)
                .Distinct();

            return comparer != null
                ? values.OrderBy(x => x, comparer).ToList()
                : values.OrderBy(x => x).ToList();
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

        private static ThemableColor HexToThemableColor(string hex)
        {
            hex = hex.TrimStart('#');

            if (hex.Length != 6)
            {
                return new ThemableColor(Colors.Black);
            }

            return new ThemableColor(Color.FromRgb(
                byte.Parse(hex.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(hex.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture),
                byte.Parse(hex.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture)));
        }

        private static string CreateStyle(
            Workbook workbook,
            bool isBold = false,
            string? fillHex = null,
            string? fontHex = null)
        {
            string styleName = "style_" + Guid.NewGuid().ToString("N");
            var style = workbook.Styles.Add(styleName);

            if (isBold)
            {
                style.FontProperties = new FontProperties(style.FontProperties)
                {
                    IsBold = true
                };
            }

            if (!string.IsNullOrWhiteSpace(fillHex))
            {
                style.Fill = PatternFill.CreateSolidFill(HexToThemableColor(fillHex));
            }

            if (!string.IsNullOrWhiteSpace(fontHex))
            {
                style.ForeColor = HexToThemableColor(fontHex);
            }

            return styleName;
        }

        private static string IdealTextColor(string backgroundHex)
        {
            string clean = backgroundHex.TrimStart('#');
            if (clean.Length != 6) return "#000000";

            int r = int.Parse(clean.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int g = int.Parse(clean.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            int b = int.Parse(clean.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            double luminance = (0.299 * r) + (0.587 * g) + (0.114 * b);
            return luminance < 140 ? "#FFFFFF" : "#000000";
        }

        private class MappedSeries
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "#000000";
            public List<MappedPoint> Points { get; set; } = new();
        }

        private readonly record struct MappedPoint(double X, double? Y, string? Label);
    }
}
