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
    /// Populates a Telerik worksheet from LiveCharts2 line series and axes.
    ///
    /// Responsibility:
    /// - Takes an existing Worksheet
    /// - Takes one or more LiveCharts2 line series
    /// - Takes X and Y axes
    /// - Writes the raw data table into the worksheet
    /// - Creates an Excel line chart in the worksheet
    ///
    /// Notes:
    /// - Supports LineSeries<double> and LineSeries<ObservablePoint>
    /// - Assumes the series are aligned to a common X axis label set
    /// - If X axis labels are missing, numeric labels 1..N are generated
    /// - Telerik chart styling APIs vary slightly by version, so series styling is best effort
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
            int chartTopRowOffset = 1,
            int chartLeftColumnOffset = 3)
        {
            if (worksheet == null) throw new ArgumentNullException(nameof(worksheet));
            if (series == null) throw new ArgumentNullException(nameof(series));

            var seriesList = series.ToList();
            if (seriesList.Count == 0) throw new InvalidOperationException("No series supplied.");

            var exportSeries = MapSeries(seriesList);
            if (exportSeries.Count == 0)
            {
                throw new InvalidOperationException(
                    "No supported line series found. Supported types: LineSeries<double>, LineSeries<ObservablePoint>.");
            }

            var xLabels = ResolveXLabels(xAxis, exportSeries);
            NormalizeSeriesLengths(exportSeries, xLabels.Count);

            WriteHeaders(worksheet, exportSeries, xAxis, startRow, startColumn);
            WriteDataRows(worksheet, exportSeries, xLabels, startRow, startColumn);
            ApplyHeaderFormatting(worksheet, exportSeries, startRow, startColumn);
            SetColumnWidths(worksheet, exportSeries.Count + 1, startColumn);

            CreateChart(
                worksheet,
                exportSeries,
                xLabels,
                xAxis,
                yAxis,
                chartTitle,
                startRow,
                startColumn,
                chartTopRowOffset,
                chartLeftColumnOffset);
        }

        public static void PopulateWorksheet(
            Worksheet worksheet,
            ISeries singleSeries,
            Axis? xAxis,
            Axis? yAxis,
            string? chartTitle = null,
            int startRow = 0,
            int startColumn = 0,
            int chartTopRowOffset = 1,
            int chartLeftColumnOffset = 3)
        {
            if (singleSeries == null) throw new ArgumentNullException(nameof(singleSeries));

            PopulateWorksheet(
                worksheet,
                new[] { singleSeries },
                xAxis,
                yAxis,
                chartTitle,
                startRow,
                startColumn,
                chartTopRowOffset,
                chartLeftColumnOffset);
        }

        private static List<ExportSeries> MapSeries(IReadOnlyList<ISeries> seriesList)
        {
            var result = new List<ExportSeries>();

            foreach (var s in seriesList)
            {
                if (s is LineSeries<double> doubleSeries)
                {
                    result.Add(new ExportSeries
                    {
                        Name = string.IsNullOrWhiteSpace(doubleSeries.Name) ? "Series" : doubleSeries.Name!,
                        HexColor = ResolveHexColor(doubleSeries.Stroke),
                        Values = doubleSeries.Values?.Select(v => (double?)v).ToList() ?? new List<double?>()
                    });
                    continue;
                }

                if (s is LineSeries<ObservablePoint> pointSeries)
                {
                    result.Add(new ExportSeries
                    {
                        Name = string.IsNullOrWhiteSpace(pointSeries.Name) ? "Series" : pointSeries.Name!,
                        HexColor = ResolveHexColor(pointSeries.Stroke),
                        Values = pointSeries.Values?.Select(v => (double?)v?.Y).ToList() ?? new List<double?>()
                    });
                }
            }

            return result;
        }

        private static List<string> ResolveXLabels(Axis? xAxis, IReadOnlyList<ExportSeries> series)
        {
            var labels = xAxis?.Labels?.ToList() ?? new List<string>();
            int maxPointCount = series.Max(s => s.Values.Count);

            if (labels.Count == 0)
            {
                labels = Enumerable.Range(1, maxPointCount)
                    .Select(i => i.ToString(CultureInfo.InvariantCulture))
                    .ToList();
            }
            else if (labels.Count < maxPointCount)
            {
                for (int i = labels.Count; i < maxPointCount; i++)
                {
                    labels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }
            else if (labels.Count > maxPointCount)
            {
                labels = labels.Take(maxPointCount).ToList();
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

        private static void WriteHeaders(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            Axis? xAxis,
            int startRow,
            int startColumn)
        {
            worksheet.Cells[startRow, startColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxis?.Name) ? "X Axis" : xAxis!.Name!);

            for (int i = 0; i < series.Count; i++)
            {
                worksheet.Cells[startRow, startColumn + i + 1].SetValue(series[i].Name);
            }
        }

        private static void WriteDataRows(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            IReadOnlyList<string> xLabels,
            int startRow,
            int startColumn)
        {
            for (int row = 0; row < xLabels.Count; row++)
            {
                worksheet.Cells[startRow + row + 1, startColumn].SetValue(xLabels[row]);

                for (int s = 0; s < series.Count; s++)
                {
                    var value = series[s].Values[row];
                    if (value.HasValue)
                    {
                        worksheet.Cells[startRow + row + 1, startColumn + s + 1].SetValue(value.Value);
                    }
                }
            }
        }

        private static void ApplyHeaderFormatting(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            int startRow,
            int startColumn)
        {
            string baseHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[startRow, startColumn, startRow, startColumn + series.Count].SetStyleName(baseHeaderStyle);

            for (int i = 0; i < series.Count; i++)
            {
                string styleName = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: series[i].HexColor,
                    fontHex: IdealTextColor(series[i].HexColor));

                worksheet.Cells[startRow, startColumn + i + 1].SetStyleName(styleName);
            }
        }

        private static void SetColumnWidths(Worksheet worksheet, int totalColumns, int startColumn)
        {
            for (int c = 0; c < totalColumns; c++)
            {
                worksheet.Columns[startColumn + c].SetWidth(new ColumnWidth(130, true));
            }
        }

        private static void CreateChart(
            Worksheet worksheet,
            IReadOnlyList<ExportSeries> series,
            IReadOnlyList<string> xLabels,
            Axis? xAxis,
            Axis? yAxis,
            string? chartTitle,
            int startRow,
            int startColumn,
            int chartTopRowOffset,
            int chartLeftColumnOffset)
        {
            int lastRow = startRow + xLabels.Count;
            int lastColumn = startColumn + series.Count;

            CellRange dataRange = worksheet.Cells[startRow, startColumn, lastRow, lastColumn];

            var chart = worksheet.Charts.Add(
                ChartType.Line,
                dataRange,
                1,
                series.Count,
                1,
                0);

            chart.SetPosition(
                new CellIndex(startRow + chartTopRowOffset, lastColumn + chartLeftColumnOffset),
                0,
                0);

            chart.SetSize(new System.Windows.Size(960, 520));

            if (!string.IsNullOrWhiteSpace(chartTitle))
            {
                chart.Title = chartTitle;
            }

            if (chart.Legend != null)
            {
                chart.Legend.Position = LegendPosition.Right;
            }

            if (chart.HorizontalAxis != null && !string.IsNullOrWhiteSpace(xAxis?.Name))
            {
                chart.HorizontalAxis.Title = xAxis!.Name!;
            }

            if (chart.VerticalAxis != null && !string.IsNullOrWhiteSpace(yAxis?.Name))
            {
                chart.VerticalAxis.Title = yAxis!.Name!;
            }

            TryApplySeriesStyles(chart, series);
        }

        private static void TryApplySeriesStyles(Chart chart, IReadOnlyList<ExportSeries> series)
        {
            try
            {
                for (int i = 0; i < series.Count && i < chart.Series.Count; i++)
                {
                    var excelSeries = chart.Series[i];
                    var color = HexToThemableColor(series[i].HexColor);
                    ApplySeriesStyle(excelSeries, color);
                }
            }
            catch
            {
                // Best effort only. Data + chart creation should still work.
            }
        }

        private static void ApplySeriesStyle(IChartSeries excelSeries, ThemableColor color)
        {
            if (excelSeries is CategorySeriesBase categorySeries)
            {
                try
                {
                    categorySeries.Outline.Fill = PatternFill.CreateSolidFill(color);
                }
                catch
                {
                }

                try
                {
                    categorySeries.Fill = PatternFill.CreateSolidFill(color);
                }
                catch
                {
                }
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
            string name = "style_" + Guid.NewGuid().ToString("N");
            var style = workbook.Styles.Add(name);

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

            return name;
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
            public string HexColor { get; set; } = "#000000";
            public List<double?> Values { get; set; } = new();
        }
    }
}
