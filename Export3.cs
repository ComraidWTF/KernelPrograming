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
            int startColumn = 0,
            IComparer<TX>? xComparer = null)
            where TX : notnull
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var source = series.ToList();
            if (source.Count == 0)
                throw new InvalidOperationException("No series supplied.");

            var mapped = MapSeries(source, xSelector, ySelector, labelSelector);
            var xValues = BuildX(mapped, xComparer);

            if (xValues.Count == 0)
                throw new InvalidOperationException("No data points found.");

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            WriteHeaderRow(worksheet, mapped, xAxis, headerRow, startColumn);
            WriteDataRows(worksheet, mapped, xValues, dataRow, startColumn);
            ApplyHeaderStyles(worksheet, mapped, headerRow, startColumn);
            SetColumnWidths(worksheet, mapped.Count, startColumn);

            int lastRow = dataRow + xValues.Count - 1;
            int chartRow = lastRow + 2;

            CreateChart(
                worksheet,
                mapped,
                dataRow,
                lastRow,
                startColumn,
                chartRow,
                chartTitle);
        }

        public static void PopulateWorksheet<TRData, TX>(
            Worksheet worksheet,
            LineSeries<TRData> series,
            Axis? xAxis,
            Axis? yAxis,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            IComparer<TX>? xComparer = null)
            where TX : notnull
        {
            if (series == null) throw new ArgumentNullException(nameof(series));

            PopulateWorksheet(
                worksheet,
                new[] { series },
                xAxis,
                yAxis,
                xSelector,
                ySelector,
                labelSelector,
                chartTitle,
                startRow,
                startColumn,
                xComparer);
        }

        private static void WriteHeaderRow<TX>(
            Worksheet worksheet,
            List<MappedSeries<TX>> mapped,
            Axis? xAxis,
            int headerRow,
            int startColumn)
            where TX : notnull
        {
            worksheet.Cells[headerRow, startColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            for (int i = 0; i < mapped.Count; i++)
            {
                worksheet.Cells[headerRow, startColumn + 1 + i].SetValue(mapped[i].Name);
            }
        }

        private static void WriteDataRows<TX>(
            Worksheet worksheet,
            List<MappedSeries<TX>> mapped,
            List<TX> xValues,
            int dataRow,
            int startColumn)
            where TX : notnull
        {
            var lookups = mapped
                .Select(s => s.Points
                    .GroupBy(p => p.X)
                    .ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int r = 0; r < xValues.Count; r++)
            {
                var x = xValues[r];
                worksheet.Cells[dataRow + r, startColumn].SetValue(FormatCellValue(x));

                for (int s = 0; s < mapped.Count; s++)
                {
                    int cellColumn = startColumn + 1 + s;

                    if (lookups[s].TryGetValue(x, out var point) && point.Y.HasValue)
                    {
                        worksheet.Cells[dataRow + r, cellColumn].SetValue(point.Y.Value);
                    }
                    else
                    {
                        worksheet.Cells[dataRow + r, cellColumn].SetValueAsFormula("=NA()");
                    }
                }
            }
        }

        private static void ApplyHeaderStyles<TX>(
            Worksheet worksheet,
            List<MappedSeries<TX>> mapped,
            int headerRow,
            int startColumn)
            where TX : notnull
        {
            string xHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[headerRow, startColumn].SetStyleName(xHeaderStyle);

            for (int i = 0; i < mapped.Count; i++)
            {
                string style = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: mapped[i].Color,
                    fontHex: IdealTextColor(mapped[i].Color));

                worksheet.Cells[headerRow, startColumn + 1 + i].SetStyleName(style);
            }
        }

        private static void SetColumnWidths(
            Worksheet worksheet,
            int seriesCount,
            int startColumn)
        {
            worksheet.Columns[startColumn].SetWidth(new ColumnWidth(130, true));

            for (int i = 0; i < seriesCount; i++)
            {
                worksheet.Columns[startColumn + 1 + i].SetWidth(new ColumnWidth(110, true));
            }
        }

        private static void CreateChart<TX>(
            Worksheet worksheet,
            List<MappedSeries<TX>> mapped,
            int dataRow,
            int lastRow,
            int startColumn,
            int chartRow,
            string? title)
            where TX : notnull
        {
            var seedRange = new CellRange(dataRow, startColumn, lastRow, startColumn + 1);

            var shape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, startColumn),
                seedRange,
                ChartType.Line)
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

            var categories = new WorkbookFormulaChartData(
                worksheet,
                new CellRange(dataRow, startColumn, lastRow, startColumn));

            for (int i = 0; i < mapped.Count; i++)
            {
                int valueColumn = startColumn + 1 + i;

                var values = new WorkbookFormulaChartData(
                    worksheet,
                    new CellRange(dataRow, valueColumn, lastRow, valueColumn));

                var added = group.Series.Add(
                    categories,
                    values,
                    new TextTitle(mapped[i].Name));

                try
                {
                    if (added is Telerik.Windows.Documents.Spreadsheet.Model.Charts.LineSeries lineSeries)
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

        private static List<MappedSeries<TX>> MapSeries<TRData, TX>(
            List<LineSeries<TRData>> input,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector)
            where TX : notnull
        {
            return input.Select(s => new MappedSeries<TX>
            {
                Name = s.Name ?? "Series",
                Color = GetColor(s.Stroke),
                Points = s.Values?
                    .Select(v => new MappedPoint<TX>(xSelector(v), ySelector(v), labelSelector?.Invoke(v)))
                    .ToList()
                    ?? new List<MappedPoint<TX>>()
            }).ToList();
        }

        private static List<TX> BuildX<TX>(
            List<MappedSeries<TX>> series,
            IComparer<TX>? comparer = null)
            where TX : notnull
        {
            var values = series
                .SelectMany(s => s.Points)
                .Select(p => p.X)
                .Distinct();

            if (comparer != null)
            {
                return values.OrderBy(x => x, comparer).ToList();
            }

            try
            {
                return values.OrderBy(x => x).ToList();
            }
            catch
            {
                return values.ToList();
            }
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

        private static string FormatCellValue<TX>(TX value)
        {
            if (value == null) return string.Empty;
            if (value is string s) return s;
            if (value is DateTime dt) return dt.ToString("yyyy-MM-dd HH:mm:ss", CultureInfo.InvariantCulture);
            if (value is DateTimeOffset dto) return dto.ToString("yyyy-MM-dd HH:mm:ss zzz", CultureInfo.InvariantCulture);
            if (value is IFormattable formattable) return formattable.ToString(null, CultureInfo.InvariantCulture);

            return value.ToString() ?? string.Empty;
        }

        private class MappedSeries<TX> where TX : notnull
        {
            public string Name { get; set; } = "";
            public string Color { get; set; } = "#000000";
            public List<MappedPoint<TX>> Points { get; set; } = new();
        }

        private readonly record struct MappedPoint<TX>(TX X, double? Y, string? Label);
    }
}
