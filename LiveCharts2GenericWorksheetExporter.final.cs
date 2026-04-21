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
    public static class LiveCharts2GenericWorksheetExporter
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
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520,
            bool includeLabelColumns = false,
            bool includeAxisMetadataRows = true)
            where TX : notnull
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var sourceSeries = series.ToList();
            if (sourceSeries.Count == 0)
                throw new InvalidOperationException("No series supplied.");

            var mappedSeries = MapSeries(sourceSeries, xSelector, ySelector, labelSelector);
            var orderedXValues = BuildOrderedXValues(mappedSeries);

            if (orderedXValues.Count == 0)
                throw new InvalidOperationException("No data points found in the supplied series.");

            int currentRow = startRow;

            if (includeAxisMetadataRows)
            {
                WriteAxisMetadataRows(worksheet, xAxis, yAxis, currentRow, startColumn);
                currentRow += 2;
            }

            int headerRow = currentRow;
            int dataStartRow = headerRow + 1;

            WriteTableHeaders(
                worksheet,
                mappedSeries,
                xAxis,
                headerRow,
                startColumn,
                includeLabelColumns,
                labelSelector != null);

            WriteTableRows(
                worksheet,
                mappedSeries,
                orderedXValues,
                dataStartRow,
                startColumn,
                includeLabelColumns,
                labelSelector != null);

            ApplyHeaderStyles(
                worksheet,
                mappedSeries,
                headerRow,
                startColumn,
                includeLabelColumns,
                labelSelector != null);

            SetColumnWidths(
                worksheet,
                startColumn,
                mappedSeries.Count,
                includeLabelColumns,
                labelSelector != null);

            int numericChartLastColumn = startColumn + mappedSeries.Count;
            int numericChartLastRow = dataStartRow + orderedXValues.Count - 1;

            int resolvedChartRow = chartRow >= 0 ? chartRow : numericChartLastRow + 2;
            int resolvedChartColumn = chartColumn >= 0 ? chartColumn : startColumn;

            CreateLineChart(
                worksheet,
                new CellRange(headerRow, startColumn, numericChartLastRow, numericChartLastColumn),
                resolvedChartRow,
                resolvedChartColumn,
                chartWidth,
                chartHeight,
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
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520,
            bool includeLabelColumns = false,
            bool includeAxisMetadataRows = true)
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
                includeLabelColumns,
                includeAxisMetadataRows);
        }

        private static List<MappedSeries<TX>> MapSeries<TRData, TX>(
            IReadOnlyList<LineSeries<TRData>> series,
            Func<TRData, TX> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? labelSelector)
            where TX : notnull
        {
            var result = new List<MappedSeries<TX>>();

            foreach (var lineSeries in series)
            {
                var rawValues = lineSeries.Values?.ToList() ?? new List<TRData>();
                var points = new List<MappedPoint<TX>>(rawValues.Count);

                foreach (var item in rawValues)
                {
                    points.Add(new MappedPoint<TX>(
                        xSelector(item),
                        ySelector(item),
                        labelSelector?.Invoke(item)));
                }

                result.Add(new MappedSeries<TX>
                {
                    Name = string.IsNullOrWhiteSpace(lineSeries.Name) ? "Series" : lineSeries.Name!,
                    HexColor = ResolveHexColor(lineSeries.Stroke),
                    Points = points
                });
            }

            return result;
        }

        private static List<TX> BuildOrderedXValues<TX>(IReadOnlyList<MappedSeries<TX>> series)
            where TX : notnull
        {
            var seen = new HashSet<TX>();
            var ordered = new List<TX>();

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

        private static void WriteAxisMetadataRows(
            Worksheet worksheet,
            Axis? xAxis,
            Axis? yAxis,
            int startRow,
            int startColumn)
        {
            worksheet.Cells[startRow, startColumn].SetValue("X Axis Name");
            worksheet.Cells[startRow, startColumn + 1]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            worksheet.Cells[startRow + 1, startColumn].SetValue("Y Axis Name");
            worksheet.Cells[startRow + 1, startColumn + 1]
                .SetValue(string.IsNullOrWhiteSpace(yAxis?.Name) ? "Y Axis" : yAxis!.Name!);

            string metaLabelStyle = CreateStyle(worksheet.Workbook, true);
            worksheet.Cells[startRow, startColumn].SetStyleName(metaLabelStyle);
            worksheet.Cells[startRow + 1, startColumn].SetStyleName(metaLabelStyle);
        }

        private static void WriteTableHeaders<TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TX>> series,
            Axis? xAxis,
            int headerRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            worksheet.Cells[headerRow, startColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            int currentColumn = startColumn + 1;

            foreach (var seriesItem in series)
            {
                worksheet.Cells[headerRow, currentColumn].SetValue(seriesItem.Name);
                currentColumn++;

                if (includeLabelColumns && hasLabelSelector)
                {
                    worksheet.Cells[headerRow, currentColumn].SetValue(seriesItem.Name + " Label");
                    currentColumn++;
                }
            }
        }

        private static void WriteTableRows<TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TX>> series,
            IReadOnlyList<TX> orderedXValues,
            int dataStartRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            var lookups = series
                .Select(s => s.Points
                    .GroupBy(p => p.X)
                    .ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int rowIndex = 0; rowIndex < orderedXValues.Count; rowIndex++)
            {
                int row = dataStartRow + rowIndex;
                TX xValue = orderedXValues[rowIndex];

                worksheet.Cells[row, startColumn].SetValue(FormatCellValue(xValue));

                int currentColumn = startColumn + 1;

                for (int seriesIndex = 0; seriesIndex < series.Count; seriesIndex++)
                {
                    if (lookups[seriesIndex].TryGetValue(xValue, out var point))
                    {
                        if (point.Y.HasValue)
                        {
                            worksheet.Cells[row, currentColumn].SetValue(point.Y.Value);
                        }

                        currentColumn++;

                        if (includeLabelColumns && hasLabelSelector)
                        {
                            if (!string.IsNullOrWhiteSpace(point.Label))
                            {
                                worksheet.Cells[row, currentColumn].SetValue(point.Label);
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

        private static void ApplyHeaderStyles<TX>(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries<TX>> series,
            int headerRow,
            int startColumn,
            bool includeLabelColumns,
            bool hasLabelSelector)
            where TX : notnull
        {
            string xHeaderStyle = CreateStyle(worksheet.Workbook, true);
            worksheet.Cells[headerRow, startColumn].SetStyleName(xHeaderStyle);

            int currentColumn = startColumn + 1;

            foreach (var seriesItem in series)
            {
                string seriesStyle = CreateStyle(
                    worksheet.Workbook,
                    true,
                    seriesItem.HexColor,
                    IdealTextColor(seriesItem.HexColor));

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
            CellRange chartDataRange,
            int chartRow,
            int chartColumn,
            double chartWidth,
            double chartHeight,
            string? chartTitle)
        {
            var chartShape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, chartColumn),
                chartDataRange,
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

        private static string ResolveHexColor(object? stroke)
        {
            if (stroke is SolidColorPaint solid)
            {
                SKColor color = solid.Color;
                return $"#{color.Red:X2}{color.Green:X2}{color.Blue:X2}";
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
                return new ThemableColor(Colors.Black);

            string clean = hex.TrimStart('#');
            if (clean.Length != 6)
                return new ThemableColor(Colors.Black);

            byte r = byte.Parse(clean.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte g = byte.Parse(clean.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte b = byte.Parse(clean.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            return new ThemableColor(Color.FromRgb(r, g, b));
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

        private sealed class MappedSeries<TX> where TX : notnull
        {
            public string Name { get; set; } = string.Empty;
            public string HexColor { get; set; } = "#000000";
            public List<MappedPoint<TX>> Points { get; set; } = new();
        }

        private readonly record struct MappedPoint<TX>(TX X, double? Y, string? Label) where TX : notnull;
    }
}
