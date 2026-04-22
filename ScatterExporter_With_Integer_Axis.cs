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
            var source = series.ToList();
            var mapped = MapSeries(source, xSelector, ySelector, xLabelSelector);
            var xValues = BuildOrderedX(mapped, xComparer);

            int xValueColumn = startColumn;
            int xLabelColumn = includeXLabelColumn ? startColumn + 1 : -1;
            int firstSeriesColumn = includeXLabelColumn ? startColumn + 2 : startColumn + 1;

            int headerRow = startRow;
            int dataRow = headerRow + 1;

            WriteHeaderRow(worksheet, mapped, xAxisTitle, includeXLabelColumn,
                xValueColumn, xLabelColumn, firstSeriesColumn, headerRow);

            WriteDataRows(worksheet, mapped, xValues, includeXLabelColumn,
                xValueColumn, xLabelColumn, firstSeriesColumn, dataRow);

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

            chart.Legend = new Legend { Position = LegendPosition.Right };

            if (!string.IsNullOrWhiteSpace(chartTitle))
                chart.Title = new TextTitle(chartTitle);

            TrySetAxisTitle(chart, xAxisTitle, true);
            TrySetAxisTitle(chart, yAxisTitle, false);

            // 👇 NEW: whole number axis config
            TryConfigureWholeNumberAxis(chart, true, 1, 1, null);
            TryConfigureWholeNumberAxis(chart, false, 1, 0, null);

            var group = chart.SeriesGroups.First();
            while (group.Series.Count > 0)
                group.Series.Remove(group.Series.First());

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

                if (showDataLabels)
                    TryConfigureDataLabels(added, "Above", false, !labelLastPointOnly, false);
            }

            worksheet.Charts.Add(chartShape);
        }

        // ===== AXIS CONFIG =====

        private static void TryConfigureWholeNumberAxis(
            object chart,
            bool isX,
            double? majorUnit,
            double? minimum,
            double? maximum)
        {
            try
            {
                var primaryAxes = chart.GetType().GetProperty("PrimaryAxes")?.GetValue(chart);
                if (primaryAxes == null) return;

                var axis = primaryAxes.GetType()
                    .GetProperty(isX ? "CategoryAxis" : "ValueAxis")
                    ?.GetValue(primaryAxes);

                if (axis == null) return;

                var axisType = axis.GetType();

                SetString(axisType, axis, "NumberFormat", "0");

                if (majorUnit.HasValue)
                    SetDouble(axisType, axis, "MajorUnit", majorUnit.Value);

                if (minimum.HasValue)
                {
                    SetDouble(axisType, axis, "Minimum", minimum.Value);
                    SetDouble(axisType, axis, "Min", minimum.Value);
                }

                if (maximum.HasValue)
                {
                    SetDouble(axisType, axis, "Maximum", maximum.Value);
                    SetDouble(axisType, axis, "Max", maximum.Value);
                }
            }
            catch { }
        }

        // ===== HELPERS =====

        private static void SetDouble(Type type, object target, string name, double value)
        {
            var p = type.GetProperty(name);
            if (p == null || !p.CanWrite) return;

            try { p.SetValue(target, Convert.ChangeType(value, p.PropertyType)); }
            catch { }
        }

        private static void SetString(Type type, object target, string name, string value)
        {
            var p = type.GetProperty(name);
            if (p != null && p.CanWrite)
                p.SetValue(target, value);
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

        private static void TryApplySeriesStyle(object series, string hex)
        {
            try
            {
                var fill = new SolidFill(HexToColor(hex));
                if (series is ScatterSeries s) s.Outline.Fill = fill;
            }
            catch { }
        }

        private static void TryConfigureDataLabels(object series,
            string pos, bool showSeries, bool showValue, bool showCat)
        {
            try
            {
                var prop = series.GetType().GetProperty("DataLabels");
                var labels = prop?.GetValue(series) ?? Activator.CreateInstance(prop.PropertyType);

                SetBool(labels.GetType(), labels, "ShowValue", showValue);
                SetEnum(labels.GetType(), labels, "Position", pos);

                prop.SetValue(series, labels);
            }
            catch { }
        }

        private static void SetBool(Type t, object o, string n, bool v)
        {
            var p = t.GetProperty(n);
            if (p != null && p.CanWrite) p.SetValue(o, v);
        }

        private static void SetEnum(Type t, object o, string n, string val)
        {
            var p = t.GetProperty(n);
            if (p == null) return;
            var ev = Enum.Parse(p.PropertyType, val);
            p.SetValue(o, ev);
        }

        private static ThemableColor HexToColor(string hex)
        {
            hex = hex.TrimStart('#');
            return new ThemableColor(Color.FromRgb(
                byte.Parse(hex.Substring(0, 2), NumberStyles.HexNumber),
                byte.Parse(hex.Substring(2, 2), NumberStyles.HexNumber),
                byte.Parse(hex.Substring(4, 2), NumberStyles.HexNumber)));
        }

        private static List<MappedSeries> MapSeries<TRData>(
            List<LineSeries<TRData>> input,
            Func<TRData, double> x,
            Func<TRData, double?> y,
            Func<TRData, string?>? label)
        {
            return input.Select(s => new MappedSeries
            {
                Name = s.Name ?? "Series",
                Color = "#000000",
                Points = s.Values?.Select(v => new MappedPoint(x(v), y(v), label?.Invoke(v))).ToList()
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
            return m.SelectMany(s => s.Points)
                .FirstOrDefault(p => p.X == x).Label;
        }

        private class MappedSeries
        {
            public string Name = "";
            public string Color = "#000000";
            public List<MappedPoint> Points = new();
        }

        private record struct MappedPoint(double X, double? Y, string? Label);
    }
}
