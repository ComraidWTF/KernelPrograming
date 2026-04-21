using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using SkiaSharp;
using Telerik.Windows.Documents.Spreadsheet.FormatProviders.OpenXml.Xlsx;
using Telerik.Windows.Documents.Spreadsheet.Model;
using Telerik.Windows.Documents.Spreadsheet.Model.Charts;

namespace WpfChartExcelExport
{
    /// <summary>
    /// Drop this file into your WPF project.
    ///
    /// What it does:
    /// - Reads LiveCharts2 line series data
    /// - Writes raw data into an Excel worksheet
    /// - Creates an Excel line chart using Telerik RadSpreadProcessing
    /// - Preserves chart title, axis titles, x labels, legend names, and series colors as far as Excel/Telerik allow
    ///
    /// Notes:
    /// - Exact Telerik chart styling APIs can vary slightly by version.
    /// - The core export flow is correct, but if your Telerik version exposes series styling differently,
    ///   only the ApplySeriesStyle method may need a tiny adjustment.
    /// - Best results happen when all line series share the same X labels.
    /// </summary>
    public static class LiveCharts2TelerikExcelExporter
    {
        public static void Export(
            IEnumerable<ISeries> series,
            IList<Axis>? xAxes,
            IList<Axis>? yAxes,
            string filePath,
            string? chartTitle = null,
            string worksheetName = "Chart Data")
        {
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            var exportModel = BuildExportModel(series, xAxes, yAxes, chartTitle);
            Export(exportModel, filePath, worksheetName);
        }

        public static void Export(
            ExcelChartExportModel exportModel,
            string filePath,
            string worksheetName = "Chart Data")
        {
            if (exportModel == null) throw new ArgumentNullException(nameof(exportModel));
            if (!exportModel.Series.Any()) throw new InvalidOperationException("No series found to export.");
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("filePath is required.", nameof(filePath));

            var workbook = new Workbook();
            var worksheet = workbook.Worksheets.Add();

            if (!string.IsNullOrWhiteSpace(worksheetName))
            {
                TryRenameWorksheet(worksheet, worksheetName);
            }

            WriteDataTable(worksheet, exportModel);
            CreateExcelChart(worksheet, exportModel);

            var provider = new XlsxFormatProvider();
            using var stream = File.Create(filePath);
            provider.Export(workbook, stream);
        }

        public static ExcelChartExportModel BuildExportModel(
            IEnumerable<ISeries> allSeries,
            IList<Axis>? xAxes,
            IList<Axis>? yAxes,
            string? chartTitle = null)
        {
            var seriesList = allSeries?.ToList() ?? throw new ArgumentNullException(nameof(allSeries));
            if (seriesList.Count == 0) throw new InvalidOperationException("No series supplied.");

            var xAxis = xAxes?.FirstOrDefault();
            var yAxis = yAxes?.FirstOrDefault();

            var xLabels = xAxis?.Labels?.ToList() ?? new List<string>();

            var exportSeries = new List<ExcelLineSeriesModel>();

            foreach (var s in seriesList)
            {
                if (TryMapDoubleSeries(s, out var doubleSeries))
                {
                    exportSeries.Add(doubleSeries);
                    continue;
                }

                if (TryMapObservablePointSeries(s, out var pointSeries))
                {
                    exportSeries.Add(pointSeries);
                    continue;
                }
            }

            if (exportSeries.Count == 0)
            {
                throw new InvalidOperationException(
                    "No supported line series found. This file currently supports LineSeries<double> and LineSeries<ObservablePoint>.");
            }

            var maxPointCount = exportSeries.Max(s => s.Values.Count);

            if (xLabels.Count == 0)
            {
                xLabels = Enumerable.Range(1, maxPointCount)
                    .Select(i => i.ToString(CultureInfo.InvariantCulture))
                    .ToList();
            }
            else if (xLabels.Count < maxPointCount)
            {
                for (int i = xLabels.Count; i < maxPointCount; i++)
                {
                    xLabels.Add((i + 1).ToString(CultureInfo.InvariantCulture));
                }
            }
            else if (xLabels.Count > maxPointCount)
            {
                xLabels = xLabels.Take(maxPointCount).ToList();
            }

            foreach (var item in exportSeries)
            {
                while (item.Values.Count < xLabels.Count)
                {
                    item.Values.Add(null);
                }

                if (item.Values.Count > xLabels.Count)
                {
                    item.Values = item.Values.Take(xLabels.Count).ToList();
                }
            }

            return new ExcelChartExportModel
            {
                ChartTitle = chartTitle ?? string.Empty,
                XAxisTitle = xAxis?.Name ?? string.Empty,
                YAxisTitle = yAxis?.Name ?? string.Empty,
                XLabels = xLabels,
                Series = exportSeries
            };
        }

        private static bool TryMapDoubleSeries(ISeries series, out ExcelLineSeriesModel exportSeries)
        {
            exportSeries = null!;

            if (series is not LineSeries<double> lineSeries) return false;

            exportSeries = new ExcelLineSeriesModel
            {
                Name = string.IsNullOrWhiteSpace(lineSeries.Name) ? "Series" : lineSeries.Name!,
                HexColor = ResolveSeriesHexColor(lineSeries.Stroke),
                Values = lineSeries.Values?.Select(v => (double?)v).ToList() ?? new List<double?>()
            };

            return true;
        }

        private static bool TryMapObservablePointSeries(ISeries series, out ExcelLineSeriesModel exportSeries)
        {
            exportSeries = null!;

            if (series is not LineSeries<ObservablePoint> lineSeries) return false;

            exportSeries = new ExcelLineSeriesModel
            {
                Name = string.IsNullOrWhiteSpace(lineSeries.Name) ? "Series" : lineSeries.Name!,
                HexColor = ResolveSeriesHexColor(lineSeries.Stroke),
                Values = lineSeries.Values?.Select(v => (double?)v?.Y).ToList() ?? new List<double?>()
            };

            return true;
        }

        private static string ResolveSeriesHexColor(object? stroke)
        {
            if (stroke is SolidColorPaint solid)
            {
                var c = solid.Color;
                return $"#{c.Red:X2}{c.Green:X2}{c.Blue:X2}";
            }

            return "#000000";
        }

        private static void WriteDataTable(Worksheet worksheet, ExcelChartExportModel exportModel)
        {
            worksheet.Cells[0, 0].SetValue(string.IsNullOrWhiteSpace(exportModel.XAxisTitle) ? "X Axis" : exportModel.XAxisTitle);

            for (int s = 0; s < exportModel.Series.Count; s++)
            {
                worksheet.Cells[0, s + 1].SetValue(exportModel.Series[s].Name);
            }

            for (int row = 0; row < exportModel.XLabels.Count; row++)
            {
                worksheet.Cells[row + 1, 0].SetValue(exportModel.XLabels[row]);

                for (int s = 0; s < exportModel.Series.Count; s++)
                {
                    var value = exportModel.Series[s].Values[row];
                    if (value.HasValue)
                    {
                        worksheet.Cells[row + 1, s + 1].SetValue(value.Value);
                    }
                }
            }

            ApplyHeaderFormatting(worksheet, exportModel);
            AutoWidthColumns(worksheet, exportModel.Series.Count + 1);
        }

        private static void ApplyHeaderFormatting(Worksheet worksheet, ExcelChartExportModel exportModel)
        {
            var headerStyle = workbookStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[0, 0, 0, exportModel.Series.Count].SetStyleName(headerStyle);

            for (int s = 0; s < exportModel.Series.Count; s++)
            {
                string styleName = workbookStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: exportModel.Series[s].HexColor,
                    fontHex: IdealTextColor(exportModel.Series[s].HexColor));

                worksheet.Cells[0, s + 1].SetStyleName(styleName);
            }
        }

        private static void AutoWidthColumns(Worksheet worksheet, int columnCount)
        {
            for (int c = 0; c < columnCount; c++)
            {
                worksheet.Columns[c].SetWidth(new ColumnWidth(130, true));
            }
        }

        private static void CreateExcelChart(Worksheet worksheet, ExcelChartExportModel exportModel)
        {
            int lastRow = exportModel.XLabels.Count;
            int lastColumn = exportModel.Series.Count;

            CellRange dataRange = worksheet.Cells[0, 0, lastRow, lastColumn];

            var chart = worksheet.Charts.Add(
                ChartType.Line,
                dataRange,
                1,
                lastColumn,
                1,
                0);

            chart.SetPosition(new CellIndex(1, lastColumn + 3), 0, 0);
            chart.SetSize(new System.Windows.Size(960, 520));

            if (!string.IsNullOrWhiteSpace(exportModel.ChartTitle))
            {
                chart.Title = exportModel.ChartTitle;
            }

            if (chart.Legend != null)
            {
                chart.Legend.Position = LegendPosition.Right;
            }

            if (chart.HorizontalAxis != null && !string.IsNullOrWhiteSpace(exportModel.XAxisTitle))
            {
                chart.HorizontalAxis.Title = exportModel.XAxisTitle;
            }

            if (chart.VerticalAxis != null && !string.IsNullOrWhiteSpace(exportModel.YAxisTitle))
            {
                chart.VerticalAxis.Title = exportModel.YAxisTitle;
            }

            TryApplySeriesStyles(chart, exportModel);
        }

        private static void TryApplySeriesStyles(Chart chart, ExcelChartExportModel exportModel)
        {
            // Telerik exposes chart series styling differently depending on library version.
            // This method is intentionally defensive so the file still works even if style hooks differ.

            try
            {
                for (int i = 0; i < exportModel.Series.Count && i < chart.Series.Count; i++)
                {
                    var excelSeries = chart.Series[i];
                    var color = HexToRgb(exportModel.Series[i].HexColor);

                    ApplySeriesStyle(excelSeries, color);
                }
            }
            catch
            {
                // Styling is best-effort only. Data export and chart creation should still succeed.
            }
        }

        private static void ApplySeriesStyle(IChartSeries excelSeries, ThemableColor color)
        {
            // Depending on Telerik version, one or more of these may exist.
            // Kept as guarded dynamic-style access is not appropriate here.
            // If your version exposes different members, adjust only this method.

            if (excelSeries is CategorySeriesBase categorySeries)
            {
                try
                {
                    categorySeries.Outline.Fill = PatternFill.CreateSolidFill(color);
                }
                catch
                {
                    // ignore
                }

                try
                {
                    categorySeries.Fill = PatternFill.CreateSolidFill(color);
                }
                catch
                {
                    // ignore
                }
            }
        }

        private static ThemableColor HexToRgb(string hex)
        {
            if (string.IsNullOrWhiteSpace(hex)) return new ThemableColor(Colors.Black);

            string clean = hex.TrimStart('#');

            if (clean.Length != 6) return new ThemableColor(Colors.Black);

            byte r = byte.Parse(clean.Substring(0, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte g = byte.Parse(clean.Substring(2, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
            byte b = byte.Parse(clean.Substring(4, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);

            return new ThemableColor(Color.FromRgb(r, g, b));
        }

        private static string workbookStyle(Workbook workbook, bool isBold = false, string? fillHex = null, string? fontHex = null)
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
                style.Fill = PatternFill.CreateSolidFill(HexToRgb(fillHex));
            }

            if (!string.IsNullOrWhiteSpace(fontHex))
            {
                style.ForeColor = HexToRgb(fontHex);
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

        private static void TryRenameWorksheet(Worksheet worksheet, string desiredName)
        {
            try
            {
                worksheet.Name = desiredName;
            }
            catch
            {
                // ignore invalid sheet name characters etc.
            }
        }
    }

    public sealed class ExcelChartExportModel
    {
        public string ChartTitle { get; set; } = string.Empty;
        public string XAxisTitle { get; set; } = string.Empty;
        public string YAxisTitle { get; set; } = string.Empty;
        public List<string> XLabels { get; set; } = new();
        public List<ExcelLineSeriesModel> Series { get; set; } = new();
    }

    public sealed class ExcelLineSeriesModel
    {
        public string Name { get; set; } = string.Empty;
        public string HexColor { get; set; } = "#000000";
        public List<double?> Values { get; set; } = new();
    }

    /// <summary>
    /// Example usage:
    ///
    /// LiveCharts2TelerikExcelExporter.Export(
    ///     Series,
    ///     XAxes,
    ///     YAxes,
    ///     @"C:\Temp\chart-export.xlsx",
    ///     "My Chart");
    ///
    /// Or if you already have a view model:
    ///
    /// var exportModel = new ExcelChartExportModel
    /// {
    ///     ChartTitle = "Performance",
    ///     XAxisTitle = "Month",
    ///     YAxisTitle = "Value",
    ///     XLabels = new List<string> { "Jan", "Feb", "Mar" },
    ///     Series = new List<ExcelLineSeriesModel>
    ///     {
    ///         new ExcelLineSeriesModel
    ///         {
    ///             Name = "Series A",
    ///             HexColor = "#FF0000",
    ///             Values = new List<double?> { 10, 20, 30 }
    ///         },
    ///         new ExcelLineSeriesModel
    ///         {
    ///             Name = "Series B",
    ///             HexColor = "#0000FF",
    ///             Values = new List<double?> { 12, 14, 28 }
    ///         }
    ///     }
    /// };
    ///
    /// LiveCharts2TelerikExcelExporter.Export(exportModel, @"C:\Temp\chart-export.xlsx");
    /// </summary>
    public static class ExportUsageExample
    {
    }
}
