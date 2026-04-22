using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
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
            bool includeXLookupTable = false,
            bool showDataLabels = false,
            bool labelLastPointOnly = true,
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
            var xValues = BuildOrderedX(mapped, xComparer);

            if (xValues.Count == 0)
                throw new InvalidOperationException("No data points found.");

            int xValueColumn = startColumn;
            int xLabelColumn = includeXLabelColumn ? startColumn + 1 : -1;
            int firstSeriesColumn = includeXLabelColumn ? startColumn + 2 : startColumn + 1;

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            WriteHeaderRow(
                worksheet,
                mapped,
                xAxisTitle,
                includeXLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                headerRow);

            WriteDataRows(
                worksheet,
                mapped,
                xValues,
                includeXLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                dataRow);

            ApplyHeaderStyles(
                worksheet,
                mapped,
                includeXLabelColumn,
                xValueColumn,
                xLabelColumn,
                firstSeriesColumn,
                headerRow);

            SetColumnWidths(
                worksheet,
                mapped.Count,
                includeXLabelColumn,
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
                chartTitle,
                xAxisTitle,
                yAxisTitle,
                showDataLabels,
                labelLastPointOnly);

            if (includeXLookupTable)
            {
                int lookupStartColumn = firstSeriesColumn + mapped.Count + 2;
                WriteXLookupTable(worksheet, mapped, xValues, lookupStartColumn, headerRow);
            }
        }

        public static void PopulateWorksheet<TRData>(
            Worksheet worksheet,
            LineSeries<TRData> series,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xLabelSelector = null,
            string? chartTitle = null,
            string? xAxisTitle = null,
            string? yAxisTitle = null,
            int startRow = 0,
            int startColumn = 0,
            bool includeXLabelColumn = true,
            bool includeXLookupTable = false,
            bool showDataLabels = false,
            bool labelLastPointOnly = true,
            IComparer<double>? xComparer = null)
        {
            if (series == null) throw new ArgumentNullException(nameof(series));

            PopulateWorksheet(
                worksheet,
                new[] { series },
                xSelector,
                ySelector,
                xLabelSelector,
                chartTitle,
                xAxisTitle,
                yAxisTitle,
                startRow,
                startColumn,
                includeXLabelColumn,
                includeXLookupTable,
                showDataLabels,
                labelLastPointOnly,
                xComparer);
        }

        private static void WriteHeaderRow(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            string? xAxisTitle,
            bool includeXLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn,
            int headerRow)
        {
            worksheet.Cells[headerRow, xValueColumn]
                .SetValue(string.IsNullOrWhiteSpace(xAxisTitle) ? "X Value" : xAxisTitle);

            if (includeXLabelColumn)
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
            IReadOnlyList<MappedSeries> mapped,
            IReadOnlyList<double> xValues,
            bool includeXLabelColumn,
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

                if (includeXLabelColumn)
                {
                    string? label = ResolveLabelForX(mapped, x);
                    worksheet.Cells[row, xLabelColumn].SetValue(
                        string.IsNullOrWhiteSpace(label)
                            ? x.ToString(CultureInfo.InvariantCulture)
                            : label);
                }

                for (int s = 0; s < mapped.Count; s++)
                {
                    int yColumn = firstSeriesColumn + s;

                    if (lookups[s].TryGetValue(x, out var point) && point.Y.HasValue)
                    {
                        worksheet.Cells[row, yColumn].SetValue(point.Y.Value);
                    }
                    else
                    {
                        worksheet.Cells[row, yColumn].SetValueAsFormula("=NA()");
                    }
                }
            }
        }

        private static void WriteXLookupTable(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            IReadOnlyList<double> xValues,
            int startColumn,
            int startRow)
        {
            worksheet.Cells[startRow, startColumn].SetValue("X Value");
            worksheet.Cells[startRow, startColumn + 1].SetValue("Meaning");

            string headerStyle = CreateStyle(worksheet.Workbook, isBold: true);
            worksheet.Cells[startRow, startColumn].SetStyleName(headerStyle);
            worksheet.Cells[startRow, startColumn + 1].SetStyleName(headerStyle);

            for (int i = 0; i < xValues.Count; i++)
            {
                int row = startRow + 1 + i;
                double x = xValues[i];
                worksheet.Cells[row, startColumn].SetValue(x);

                string? label = ResolveLabelForX(mapped, x);
                worksheet.Cells[row, startColumn + 1].SetValue(
                    string.IsNullOrWhiteSpace(label)
                        ? x.ToString(CultureInfo.InvariantCulture)
                        : label);
            }

            worksheet.Columns[startColumn].SetWidth(new ColumnWidth(90, true));
            worksheet.Columns[startColumn + 1].SetWidth(new ColumnWidth(180, true));
        }

        private static void ApplyHeaderStyles(
            Worksheet worksheet,
            IReadOnlyList<MappedSeries> mapped,
            bool includeXLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn,
            int headerRow)
        {
            string baseHeaderStyle = CreateStyle(worksheet.Workbook, isBold: true);

            worksheet.Cells[headerRow, xValueColumn].SetStyleName(baseHeaderStyle);

            if (includeXLabelColumn)
            {
                worksheet.Cells[headerRow, xLabelColumn].SetStyleName(baseHeaderStyle);
            }

            for (int i = 0; i < mapped.Count; i++)
            {
                string seriesStyle = CreateStyle(
                    worksheet.Workbook,
                    isBold: true,
                    fillHex: mapped[i].Color,
                    fontHex: IdealTextColor(mapped[i].Color));

                worksheet.Cells[headerRow, firstSeriesColumn + i].SetStyleName(seriesStyle);
            }
        }

        private static void SetColumnWidths(
            Worksheet worksheet,
            int seriesCount,
            bool includeXLabelColumn,
            int xValueColumn,
            int xLabelColumn,
            int firstSeriesColumn)
        {
            worksheet.Columns[xValueColumn].SetWidth(new ColumnWidth(90, true));

            if (includeXLabelColumn)
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
            IReadOnlyList<MappedSeries> mapped,
            int dataRow,
            int lastRow,
            int xValueColumn,
            int firstSeriesColumn,
            int chartRow,
            int chartColumn,
            string? chartTitle,
            string? xAxisTitle,
            string? yAxisTitle,
            bool showDataLabels,
            bool labelLastPointOnly)
        {
            var seedRange = new CellRange(dataRow, xValueColumn, lastRow, firstSeriesColumn);

            var chartShape = new FloatingChartShape(
                worksheet,
                new CellIndex(chartRow, chartColumn),
                seedRange,
                ChartType.Scatter)
            {
                Width = 1100,
                Height = 620
            };

            var chart = chartShape.Chart;

            chart.Legend = new Legend
            {
                Position = LegendPosition.Right
            };

            if (!string.IsNullOrWhiteSpace(chartTitle))
            {
                chart.Title = new TextTitle(chartTitle);
            }

            TrySetAxisTitle(chart, xAxisTitle, isX: true);
            TrySetAxisTitle(chart, yAxisTitle, isX: false);

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

                TryApplySeriesStyle(added, mapped[i].Color);
                TryConfigureScatterSeries(added);
                if (showDataLabels)
                {
                    TryConfigureDataLabels(
                        added,
                        positionName: "Above",
                        showSeriesName: false,
                        showValue: !labelLastPointOnly,
                        showCategoryName: false);
                }
            }

            worksheet.Charts.Add(chartShape);
        }

        private static void TrySetAxisTitle(object chart, string? title, bool isX)
        {
            if (string.IsNullOrWhiteSpace(title) || chart == null) return;

            try
            {
                var chartType = chart.GetType();
                var primaryAxesProp = chartType.GetProperty("PrimaryAxes");
                if (primaryAxesProp == null) return;

                var primaryAxes = primaryAxesProp.GetValue(chart);
                if (primaryAxes == null) return;

                var axesType = primaryAxes.GetType();
                var axisProp = axesType.GetProperty(isX ? "CategoryAxis" : "ValueAxis");
                if (axisProp == null) return;

                var axis = axisProp.GetValue(primaryAxes);
                if (axis == null) return;

                var titleProp = axis.GetType().GetProperty("Title");
                if (titleProp == null || !titleProp.CanWrite) return;

                titleProp.SetValue(axis, new TextTitle(title));
            }
            catch
            {
            }
        }

        private static void TryApplySeriesStyle(object series, string hexColor)
        {
            try
            {
                var fill = new SolidFill(HexToThemableColor(hexColor));

                if (series is ScatterSeries scatterSeries)
                {
                    scatterSeries.Outline.Fill = fill;
                    return;
                }

                if (series is Telerik.Windows.Documents.Spreadsheet.Model.Charts.LineSeries lineSeries)
                {
                    lineSeries.Outline.Fill = fill;
                    return;
                }

                if (series is CategorySeriesBase categorySeries)
                {
                    categorySeries.Outline.Fill = fill;
                }
            }
            catch
            {
            }
        }

        private static void TryConfigureScatterSeries(object series)
        {
            try
            {
                if (series is ScatterSeries scatterSeries)
                {
                    scatterSeries.ScatterStyle = ScatterStyle.LineMarker;
                    scatterSeries.Marker = new Marker
                    {
                        Size = 8,
                        Symbol = MarkerStyle.Circle
                    };
                }
            }
            catch
            {
            }
        }

        private static void TryConfigureDataLabels(
            object series,
            string positionName = "Above",
            bool showSeriesName = false,
            bool showValue = true,
            bool showCategoryName = false)
        {
            if (series == null) return;

            try
            {
                var seriesType = series.GetType();
                var dataLabelsProp = seriesType.GetProperty("DataLabels");
                if (dataLabelsProp == null || !dataLabelsProp.CanRead || !dataLabelsProp.CanWrite)
                    return;

                var dataLabels = dataLabelsProp.GetValue(series);
                if (dataLabels == null)
                {
                    var labelsType = dataLabelsProp.PropertyType;
                    var ctor = labelsType.GetConstructor(Type.EmptyTypes);
                    if (ctor == null) return;

                    dataLabels = ctor.Invoke(null);
                    dataLabelsProp.SetValue(series, dataLabels);
                }

                var labelsType2 = dataLabels!.GetType();

                SetBool(labelsType2, dataLabels, "ShowSeriesName", showSeriesName);
                SetBool(labelsType2, dataLabels, "ShowValue", showValue);
                SetBool(labelsType2, dataLabels, "ShowCategoryName", showCategoryName);
                SetBool(labelsType2, dataLabels, "ShowLegendKey", false);
                SetBool(labelsType2, dataLabels, "ShowPercentage", false);
                SetBool(labelsType2, dataLabels, "ShowBubbleSize", false);

                SetEnumByName(labelsType2, dataLabels, "Position", positionName);
                SetEnumByName(labelsType2, dataLabels, "LabelPosition", positionName);
            }
            catch
            {
            }
        }

        private static void SetBool(Type type, object target, string propertyName, bool value)
        {
            PropertyInfo? p = type.GetProperty(propertyName);
            if (p != null && p.CanWrite && p.PropertyType == typeof(bool))
            {
                p.SetValue(target, value);
            }
        }

        private static void SetEnumByName(Type type, object target, string propertyName, string enumName)
        {
            PropertyInfo? p = type.GetProperty(propertyName);
            if (p == null || !p.CanWrite) return;

            Type pt = p.PropertyType;
            if (!pt.IsEnum) return;

            try
            {
                object enumValue = Enum.Parse(pt, enumName, true);
                p.SetValue(target, enumValue);
            }
            catch
            {
            }
        }

        private static string? ResolveLabelForX(IReadOnlyList<MappedSeries> mapped, double x)
        {
            foreach (var series in mapped)
            {
                var point = series.Points.LastOrDefault(p => p.X.Equals(x));
                if (!string.IsNullOrWhiteSpace(point.Label))
                {
                    return point.Label;
                }
            }

            return null;
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

        private static List<double> BuildOrderedX(
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

        private sealed class MappedSeries
        {
            public string Name { get; set; } = string.Empty;
            public string Color { get; set; } = "#000000";
            public List<MappedPoint> Points { get; set; } = new();
        }

        private readonly record struct MappedPoint(double X, double? Y, string? Label);
    }
}
