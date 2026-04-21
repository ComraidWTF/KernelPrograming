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
    /// <summary>
    /// Populates an existing Telerik Worksheet from one or more LiveCharts2 LineSeries&lt;TRData&gt;
    /// and creates an Excel line chart from the populated range.
    ///
    /// Public contract:
    /// - depends on Worksheet
    /// - depends on LineSeries&lt;TRData&gt;
    /// - depends on Axis for X and Y names/labels
    /// - no ViewModel dependency
    ///
    /// Notes:
    /// - X values are taken from xSelector, not from the LiveCharts Axis object.
    /// - xAxis.Name / yAxis.Name are used as worksheet headers when available.
    /// - chart title and legend are added to the Excel chart.
    /// - optional labelSelector can write data-label text into extra worksheet columns.
    /// - this file keeps chart styling compile-safe and does not force per-series Excel line color.
    /// </summary>
    public static class LiveCharts2GenericWorksheetExporter
    {
        public static void PopulateWorksheet<TRData, TX>(
            Worksheet worksheet,
            IEnumerable<LineSeries<TRData>> series,
            LiveChartsCore.Measure.Axis? xAxis,
            LiveChartsCore.Measure.Axis? yAxis,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520,
            bool includeLabelColumns = false)
            where TX : notnull
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var sourceSeries = series.ToList();
            if (sourceSeries.Count == 0)
                throw new InvalidOperationException("No series supplied.");

            var mapped = MapSeries(sourceSeries, xSelector, ySelector, labelSelector);

            var orderedX = BuildOrderedXValues(mapped);
            if (orderedX.Count == 0)
                throw new InvalidOperationException("No data points found in the supplied series.");

            int labelColumnsPerSeries = includeLabelColumns && labelSelector != null ? 1 : 0;

            int headerRow = startRow;
            int dataStartRow = headerRow + 1;

            WriteHeaders(
                worksheet,
                mapped,
                xAxis,
                startRow: headerRow,
                startColumn: startColumn,
                includeLabelColumns: includeLabelColumns,
                hasLabelSelector: labelSelector != null);

            WriteRows(
                worksheet,
                mapped,
                orderedX,
                xSelector,
                dataStartRow,
                startColumn,
                includeLabelColumns,
                labelSelector != null);

            ApplyHeaderStyles(
                worksheet,
                mapped,
                headerRow,
                startColumn,
                includeLabelColumns,
                labelSelector != null);

            SetColumnWidths(
                worksheet,
                startColumn,
                mapped.Count,
                includeLabelColumns,
                labelSelector != null);

            int chartDataLastColumn = startColumn + mapped.Count; // x column + one Y column per series only
            int chartDataLastRow = dataStartRow + orderedX.Count - 1;

            int resolvedChartRow = chartRow >= 0 ? chartRow : startRow + 1;
            int resolvedChartColumn = chartColumn >= 0
                ? chartColumn
                : startColumn + 1 + mapped.Count + (mapped.Count * labelColumnsPerSeries) + 2;

            CreateLineChart(
                worksheet,
                new CellRange(headerRow, startColumn, chartDataLastRow, chartDataLastColumn),
                resolvedChartRow,
                resolvedChartColumn,
                chartWidth,
                chartHeight,
                chartTitle);
        }

        public static void PopulateWorksheet<TRData, TX>(
            Worksheet worksheet,
            LineSeries<TRData> series,
            LiveChartsCore.Measure.Axis? xAxis,
            LiveChartsCore.Measure.Axis? yAxis,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector = null,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520,
            bool includeLabelColumns = false)
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
                chartRow,
                chartColumn,
                chartWidth,
                chartHeight,
                includeLabelColumns);
        }

        private static List<MappedSeries<TRData, TX>> MapSeries<TRData, TX>(
            IReadOnlyList<LineSeries<TRData>> series,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector)
            where TX : notnull
        {
            var result = new List<MappedSeries<TRData, TX>>();

            foreach (var seriesItem in series)
            {
                var values = seriesItem.Values?.ToList() ?? new List<TRData>();
                var points = new List<MappedPoint<TX>>();

                foreach (var item in values)
                {
                    var x = xSelector(item);
                    var y = ySelector(item);
                    var label = labelSelector?.Invoke(item);

                    points.Add(new MappedPoint<TX>(x, y, label));
                }

                result.Add(new MappedSeries<TRData, TX>
                {
                    Name = string.IsNullOrWhiteSpace(seriesItem.Name) ? "Series" : seriesItem.Name!,
                    HexColor = ResolveHexColor(seriesItem.Stroke),
                    Points = points
                });
            }

            return result;
        }

        private static List<TX> BuildOrderedXValues<TRData, TX>(IReadOnlyList<MappedSeries<TRData, TX>> series)
            where TX : notnull
        {
            var ordered = new List<TX>();
            var seen = new HashSet<TX>();

            foreach (var currentSeries in series)
            {
                foreach (var point in currentSeries.Points)
                {
                    if (seen.Add(point.X))
                    {
                        ordered.Add(point.X);
                    }
                }
            }

            return ordered;
        }

        private static void WriteHeaders<TRData, TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TRData, TX>> series,
            LiveChartsCore.Measure.Axis? xAxis,
            int startRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            worksheet.Cells[startRow, startColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            int currentColumn = startColumn + 1;

            foreach (var item in series)
            {
                worksheet.Cells[startRow, currentColumn].SetValue(item.Name);
                currentColumn++;

                if (includeLabelColumns && hasLabelSelector)
                {
                    worksheet.Cells[startRow, currentColumn].SetValue(item.Name + " Label");
                    currentColumn++;
                }
            }
        }

        private static void WriteRows<TRData, TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TRData, TX>> series,
            IReadOnlyList<TX> orderedX,
            Func<TRData, TX> xSelector,
            int dataStartRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            var lookupPerSeries = series
                .Select(s => s.Points
                    .GroupBy(p => p.X)
                    .ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int rowIndex = 0; rowIndex < orderedX.Count; rowIndex++)
            {
                int worksheetRow = dataStartRow + rowIndex;
                var xValue = orderedX[rowIndex];

                worksheet.Cells[worksheetRow, startColumn].SetValue(FormatCellValue(xValue));

                int currentColumn = startColumn + 1;

                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (lookupPerSeries[seriesIndex].TryGetValue(xValue, out var point))
                    {
                        if (point.Y.HasValue)
                        {
                            worksheet.Cells[worksheetRow, currentColumn].SetValue(point.Y.Value);
                        }

                        currentColumn++;

                        if (includeLabelColumns && hasLabelSelector)
                        {
                            if (!string.IsNullOrWhiteSpace(point.Label))
                            {
                                worksheet.Cells[worksheetRow, currentColumn].SetValue(point.Label);
                            }

                            currentColumn++;
                        }
                    }
                    else
                    {
                        currentColumn++;

                        if (includeLabelColumns && hasLabelSelector)
                        {
                            currentColumn++;
                        }
                    }
                }
            }
        }

        private static void ApplyHeaderStyles<TRData, TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TRData, TX>> series,
            int headerRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            string xHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[headerRow, startColumn].SetStyleName(xHeaderStyle);

            int currentColumn = startColumn + 1;

            foreach (var item in series)
            {
                string seriesStyle = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: item.HexColor,
                    fontHex: IdealTextColor(item.HexColor));

                worksheet.Cells[headerRow, currentColumn].SetStyleName(seriesStyle);
                currentColumn++;

                if (includeLabelColumns && hasLabelSelector)
                {
                    worksheet.Cells[headerRow, currentColumn].SetStyleName(seriesStyle);
                    currentColumn++;
                }
            }
        }

        private static void SetColumnWidths(
            Worksheet worksheet,
            int startColumn,
            int seriesCount,
            bool includeLabelColumns,
            bool hasLabelSelector)
        {
            worksheet.Columns[startColumn].SetWidth(new ColumnWidth(130, true));

            int currentColumn = startColumn + 1;

            for (int i = 0; i < seriesCount; i++)
            {
                worksheet.Columns[currentColumn].SetWidth(new ColumnWidth(110, true));
                currentColumn++;

                if (includeLabelColumns && hasLabelSelector)
                {
                    worksheet.Columns[currentColumn].SetWidth(new ColumnWidth(150, true));
                    currentColumn++;
                }
            }
        }

        private static void CreateLineChart(
            Worksheet worksheet,
            CellRange dataRange,
            int chartRow,
            int chartColumn,
            double chartWidth,
            double chartHeight,
            string? chartTitle)
        {
            var chartShape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, chartColumn),
                dataRange,
                ChartType.Line)
            {
                Width = chartWidth,
                Height = chartHeight
            };

            if (!string.IsNullOrWhiteSpace(chartTitle))
            {
                chartShape.Chart.Title = new TextTitle(chartTitle);
            }

            chartShape.Chart.Legend = new Legend
            {
                Position = LegendPosition.Right
            };

            worksheet.Charts.Add(chartShape);
        }

        private static string FormatCellValue<TX>(TX value)
        {
            if (value == null) return string.Empty;

            if (value is string s) return s;
            if (value is DateTime dt) return dt.ToString("O", CultureInfo.InvariantCulture);
            if (value is IFormattable formattable) return formattable.ToString(null, CultureInfo.InvariantCulture);

            return value.ToString() ?? string.Empty;
        }

        private static string ResolveHexColor(object? stroke)
        {
            if (stroke is SolidColorPaint solid)
            {
                SKColor c = solid.Color;
                return $"#{c.Red:X2}{c.Green:X2}{c.Blue:X2}";
            }

            return "#000000";
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

        private static ThemableColor HexToThemableColor(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex))
            {
                return new ThemableColor(Colors.Black);
            }

            string clean = hex.TrimStart('#');
            if (clean.Length != 6)
            {
                return new ThemableColor(Colors.Black);
            }

            byte r = byte.Parse(clean.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte g = byte.Parse(clean.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte b = byte.Parse(clean.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            return new ThemableColor(Color.FromRgb(r, g, b));
        }

        private sealed class MappedSeries<TRData, TX> where TX : notnull
        {
            public string Name { get; set; } = string.Empty;
            public string HexColor { get; set; } = "#000000";
            public List<MappedPoint<TX>> Points { get; set; } = new();
        }

        private readonly record struct MappedPoint<TX>(TX X, double? Y, string? Label) where TX : notnull;
    }
}
