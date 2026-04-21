using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using SkiaSharp;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;

namespace WpfChartExcelExport
{
    /// <summary>
    /// Writes LiveCharts2 line-series data into an existing Telerik worksheet
    /// and creates a line chart from that data.
    ///
    /// Public contract:
    /// - depends on Worksheet
    /// - depends on ISeries / IEnumerable<ISeries>
    /// - depends on X and Y Axis
    /// - no ViewModel dependency
    ///
    /// Supported LiveCharts2 series:
    /// - LineSeries<double>
    /// - LineSeries<ObservablePoint>
    ///
    /// Notes:
    /// - X axis labels come from xAxis.Labels when available.
    /// - If X labels are missing, 1..N are generated.
    /// - Legend is enabled in the Excel chart.
    /// - Series names are used as legend entries.
    /// - Header cells are colored using the series stroke color.
    /// - Telerik chart line colors are theme-driven by default, so this file does not force
    ///   per-series Excel chart colors in order to stay compile-safe across Telerik versions.
    /// </summary>
    public static class LiveCharts2WorksheetExporter
    {
        public static void PopulateWorksheet(
            Worksheet worksheet,
            IEnumerable<ISeries> series,
            Axis? xAxis,
            Axis? yAxis,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));

            var inputSeries = series.ToList();
            if (inputSeries.Count == 0)
                throw new InvalidOperationException("No series supplied.");

            var exportSeries = MapSeries(inputSeries);
            if (exportSeries.Count == 0)
            {
                throw new InvalidOperationException(
                    "No supported series found. Supported: LineSeries<double>, LineSeries<ObservablePoint>.");
            }

            var xLabels = ResolveXLabels(xAxis, exportSeries);
            NormalizeSeriesLengths(exportSeries, xLabels.Count);

            WriteWorksheetData(
                worksheet,
                exportSeries,
                xLabels,
                xAxis,
                yAxis,
                startRow,
                startColumn);

            CreateLineChart(
                worksheet,
                exportSeries.Count,
                xLabels.Count,
                chartTitle,
                startRow,
                startColumn,
                chartRow >= 0 ? chartRow : startRow + 1,
                chartColumn >= 0 ? chartColumn : startColumn + exportSeries.Count + 3,
                chartWidth,
                chartHeight);
        }

        public static void PopulateWorksheet(
            Worksheet worksheet,
            ISeries series,
            Axis? xAxis,
            Axis? yAxis,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            int chartRow = -1,
            int chartColumn = -1,
            double chartWidth = 960,
            double chartHeight = 520)
        {
            if (series == null) throw new ArgumentNullException(nameof(series));

            PopulateWorksheet(
                worksheet,
                new[] { series },
                xAxis,
                yAxis,
                chartTitle,
                startRow,
                startColumn,
                chartRow,
                chartColumn,
                chartWidth,
                chartHeight);
        }

        private static List<ExportSeries> MapSeries(IReadOnlyList<ISeries> input)
        {
            var result = new List<ExportSeries>();

            foreach (var item in input)
            {
                if (item is LineSeries<double> doubleSeries)
                {
                    result.Add(new ExportSeries
                    {
                        Name = string.IsNullOrWhiteSpace(doubleSeries.Name) ? "Series" : doubleSeries.Name!,
                        Values = doubleSeries.Values?.Select(v => (double?)v).ToList() ?? new List<double?>(),
                        HexColor = ResolveHexColor(doubleSeries.Stroke)
                    });
                    continue;
                }

                if (item is LineSeries<ObservablePoint> pointSeries)
                {
                    result.Add(new ExportSeries
                    {
                        Name = string.IsNullOrWhiteSpace(pointSeries.Name) ? "Series" : pointSeries.Name!,
                        Values = pointSeries.Values?.Select(v => (double?)v?.Y).ToList() ?? new List<double?>(),
                        HexColor = ResolveHexColor(pointSeries.Stroke)
                    });
                }
            }

            return result;
        }

        private static List<string> ResolveXLabels(Axis? xAxis, IReadOnlyList<ExportSeries> series)
        {
            var labels = xAxis?.Labels?.ToList() ?? new List<string>();
            int maxCount = series.Max(s => s.Values.Count);

            if (labels.Count == 0)
            {
                return Enumerable.Range(1, maxCount)
                    .Select(i => i.ToString(CultureInfo.InvariantCulture))
                    .ToList();
            }

            if (labels.Count < maxCount)
            {
                for (int i = labels.Count; i < maxCount; i++)
                {
                    labels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }
            else if (labels.Count > maxCount)
            {
                labels = labels.Take(maxCount).ToList();
            }

            return labels;
        }

        private static void NormalizeSeriesLengths(IReadOnlyList<ExportSeries> series, int targetLength)
        {
            foreach (var item in series)
            {
                while (item.Values.Count < targetLength)
                {
                    item.Values.Add(null);
                }

                if (item.Values.Count > targetLength)
                {
                    item.Values = item.Values.Take(targetLength).ToList();
                }
            }
        }

        private static void WriteWorksheetData(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            IReadOnlyList<string> xLabels,
            Axis? xAxis,
            Axis? yAxis,
            int startRow,
            int startColumn)
        {
            int currentRow = startRow;

            if (!string.IsNullOrWhiteSpace(yAxis?.Name))
            {
                worksheet.Cells[currentRow, startColumn].SetValue("Y Axis");
                worksheet.Cells[currentRow, startColumn + 1].SetValue(yAxis!.Name!);
                currentRow++;
            }

            worksheet.Cells[currentRow, startColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            for (int s = 0; s < series.Count; s++)
            {
                worksheet.Cells[currentRow, startColumn + s + 1].SetValue(series[s].Name);
            }

            ApplyHeaderStyles(worksheet, series, currentRow, startColumn);

            for (int row = 0; row < xLabels.Count; row++)
            {
                worksheet.Cells[currentRow + row + 1, startColumn].SetValue(xLabels[row]);

                for (int s = 0; s < series.Count; s++)
                {
                    var value = series[s].Values[row];
                    if (value.HasValue)
                    {
                        worksheet.Cells[currentRow + row + 1, startColumn + s + 1].SetValue(value.Value);
                    }
                }
            }

            for (int c = 0; c < series.Count + 1; c++)
            {
                worksheet.Columns[startColumn + c].SetWidth(new ColumnWidth(130, true));
            }
        }

        private static void CreateLineChart(
            Worksheet worksheet,
            int seriesCount,
            int xLabelCount,
            string? chartTitle,
            int startRow,
            int startColumn,
            int chartRow,
            int chartColumn,
            double chartWidth,
            double chartHeight)
        {
            // Data layout:
            // optional Y-axis row at startRow
            // header row is either startRow or startRow + 1 depending on whether Y axis name exists.
            //
            // To keep the API compile-safe and predictable, the chart always reads from the table whose
            // header row is the first row containing "X Axis" + series names.
            int headerRow = startRow;

            var firstCellText = worksheet.Cells[startRow, startColumn].GetValue().Value.RawValue;
            if (string.Equals(firstCellText, "Y Axis", StringComparison.OrdinalIgnoreCase))
            {
                headerRow = startRow + 1;
            }

            int lastRow = headerRow + xLabelCount;
            int lastColumn = startColumn + seriesCount;

            var dataRange = new CellRange(headerRow, startColumn, lastRow, lastColumn);

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

        private static void ApplyHeaderStyles(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            int headerRow,
            int startColumn)
        {
            string xHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[headerRow, startColumn].SetStyleName(xHeaderStyle);

            for (int i = 0; i < series.Count; i++)
            {
                string styleName = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: series[i].HexColor,
                    fontHex: IdealTextColor(series[i].HexColor));

                worksheet.Cells[headerRow, startColumn + i + 1].SetStyleName(styleName);
            }
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
                return new ThemableColor(Colors.Black);

            string clean = hex.TrimStart('#');
            if (clean.Length != 6)
                return new ThemableColor(Colors.Black);

            byte r = byte.Parse(clean.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte g = byte.Parse(clean.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte b = byte.Parse(clean.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            return new ThemableColor(Color.FromRgb(r, g, b));
        }

        private sealed class ExportSeries
        {
            public string Name { get; set; } = string.Empty;
            public List<double?> Values { get; set; } = new();
            public string HexColor { get; set; } = "#000000";
        }
    }
}
