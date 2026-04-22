using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Reflection;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using LiveChartsCore;

namespace WpfChartExcelExport
{
    public static class OpenXmlLiveChartsExporter
    {
        private static readonly string[] ExcelAccentFallback =
        {
            "5B9BD5",
            "ED7D31",
            "A5A5A5",
            "FFC000",
            "4472C4",
            "70AD47"
        };

        public static void ExportScatterChart(
            string outputPath,
            IEnumerable<ISeries> series,
            Func<object, double> xSelector,
            Func<object, double?> ySelector,
            Func<object, string> xMeaningSelector,
            string chartTitle,
            string xAxisTitle,
            string yAxisTitle,
            string worksheetName,
            IComparer<double> xComparer)
        {
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path is required.", nameof(outputPath));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));
            if (string.IsNullOrWhiteSpace(worksheetName)) worksheetName = "Chart Data";
            if (xComparer == null) xComparer = Comparer<double>.Default;

            List<MappedSeries> mapped = series
                .Select((s, i) => MapSeries(s, i, xSelector, ySelector, xMeaningSelector))
                .Where(s => s != null)
                .Cast<MappedSeries>()
                .ToList();

            if (mapped.Count == 0)
                throw new InvalidOperationException("No readable LiveCharts2 series were found.");

            List<double> orderedX = mapped
                .SelectMany(s => s.Points)
                .Select(p => p.X)
                .Distinct()
                .OrderBy(v => v, xComparer)
                .ToList();

            if (orderedX.Count == 0)
                throw new InvalidOperationException("No data points found.");

            if (System.IO.File.Exists(outputPath))
                System.IO.File.Delete(outputPath);

            using (SpreadsheetDocument document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Spreadsheet.Workbook();

                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                DrawingsPart drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                workbookPart.Workbook.AppendChild(new Spreadsheet.BookViews(
                    new Spreadsheet.WorkbookView()));

                Spreadsheet.Sheets sheets = workbookPart.Workbook.AppendChild(new Spreadsheet.Sheets());
                sheets.Append(new Spreadsheet.Sheet
                {
                    Id = workbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1U,
                    Name = worksheetName
                });

                const uint headerRowIndex = 1U;
                const uint dataStartRowIndex = 2U;

                const int xValueCol = 1;
                const int xMeaningCol = 2;
                const int yMeaningCol = 3;
                const int firstSeriesCol = 4;

                Spreadsheet.SheetData sheetData = new Spreadsheet.SheetData();
                foreach (Spreadsheet.Row row in BuildRows(
                    mapped,
                    orderedX,
                    headerRowIndex,
                    dataStartRowIndex,
                    xValueCol,
                    xMeaningCol,
                    yMeaningCol,
                    firstSeriesCol,
                    xAxisTitle,
                    yAxisTitle))
                {
                    sheetData.Append(row);
                }

                Spreadsheet.SheetViews sheetViews = new Spreadsheet.SheetViews(
                    new Spreadsheet.SheetView { WorkbookViewId = 0U });

                worksheetPart.Worksheet = new Spreadsheet.Worksheet(
                    sheetViews,
                    new Spreadsheet.SheetFormatProperties { DefaultRowHeight = 15D },
                    new Spreadsheet.Columns(
                        CreateColumn(1, 1, 12),
                        CreateColumn(2, 2, 24),
                        CreateColumn(3, 3, 18),
                        CreateColumn((uint)firstSeriesCol, (uint)(firstSeriesCol + mapped.Count - 1), 14)
                    ),
                    sheetData,
                    new Spreadsheet.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

                worksheetPart.Worksheet.Save();

                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                AddScatterChart(
                    drawingsPart,
                    worksheetName,
                    mapped,
                    orderedX.Count,
                    headerRowIndex,
                    dataStartRowIndex,
                    xValueCol,
                    firstSeriesCol,
                    chartTitle,
                    xAxisTitle,
                    yAxisTitle);
                drawingsPart.WorksheetDrawing.Save();

                workbookPart.Workbook.Save();
            }
        }

        public static void ExportScatterChart(
            string outputPath,
            IEnumerable<ISeries> series,
            Func<object, double> xSelector,
            Func<object, double?> ySelector)
        {
            ExportScatterChart(
                outputPath,
                series,
                xSelector,
                ySelector,
                null,
                null,
                null,
                null,
                "Chart Data",
                Comparer<double>.Default);
        }

        private static MappedSeries MapSeries(
            ISeries series,
            int index,
            Func<object, double> xSelector,
            Func<object, double?> ySelector,
            Func<object, string> xMeaningSelector)
        {
            PropertyInfo valuesProp = series.GetType().GetProperty("Values", BindingFlags.Instance | BindingFlags.Public);
            if (valuesProp == null) return null;

            IEnumerable rawValues = valuesProp.GetValue(series, null) as IEnumerable;
            if (rawValues == null) return null;

            PropertyInfo nameProp = series.GetType().GetProperty("Name", BindingFlags.Instance | BindingFlags.Public);
            object nameObj = nameProp != null ? nameProp.GetValue(series, null) : null;
            string seriesName = nameObj != null ? nameObj.ToString() : null;
            if (string.IsNullOrWhiteSpace(seriesName))
            {
                seriesName = "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);
            }

            string colorHex = TryResolveSeriesHexColor(series);
            if (string.IsNullOrWhiteSpace(colorHex))
            {
                colorHex = ExcelAccentFallback[index % ExcelAccentFallback.Length];
            }

            List<MappedPoint> points = new List<MappedPoint>();
            foreach (object item in rawValues)
            {
                if (item == null) continue;
                points.Add(new MappedPoint(
                    xSelector(item),
                    ySelector(item),
                    xMeaningSelector != null ? xMeaningSelector(item) : null));
            }

            return new MappedSeries(seriesName, colorHex, points);
        }

        private static string TryResolveSeriesHexColor(ISeries series)
        {
            try
            {
                PropertyInfo strokeProp = series.GetType().GetProperty("Stroke", BindingFlags.Instance | BindingFlags.Public);
                object stroke = strokeProp != null ? strokeProp.GetValue(series, null) : null;
                if (stroke == null) return null;

                PropertyInfo colorProp = stroke.GetType().GetProperty("Color", BindingFlags.Instance | BindingFlags.Public);
                object colorObj = colorProp != null ? colorProp.GetValue(stroke, null) : null;
                if (colorObj != null)
                {
                    string rgb = TryRgbFromColorObject(colorObj);
                    if (!string.IsNullOrWhiteSpace(rgb)) return rgb;
                }

                string rgbDirect = TryRgbFromColorObject(stroke);
                if (!string.IsNullOrWhiteSpace(rgbDirect)) return rgbDirect;
            }
            catch
            {
            }

            return null;
        }

        private static string TryRgbFromColorObject(object colorObj)
        {
            try
            {
                Type t = colorObj.GetType();
                PropertyInfo red = t.GetProperty("Red", BindingFlags.Instance | BindingFlags.Public);
                PropertyInfo green = t.GetProperty("Green", BindingFlags.Instance | BindingFlags.Public);
                PropertyInfo blue = t.GetProperty("Blue", BindingFlags.Instance | BindingFlags.Public);

                if (red != null && green != null && blue != null)
                {
                    byte r = Convert.ToByte(red.GetValue(colorObj, null), CultureInfo.InvariantCulture);
                    byte g = Convert.ToByte(green.GetValue(colorObj, null), CultureInfo.InvariantCulture);
                    byte b = Convert.ToByte(blue.GetValue(colorObj, null), CultureInfo.InvariantCulture);
                    return r.ToString("X2", CultureInfo.InvariantCulture)
                         + g.ToString("X2", CultureInfo.InvariantCulture)
                         + b.ToString("X2", CultureInfo.InvariantCulture);
                }
            }
            catch
            {
            }

            return null;
        }

        private static IEnumerable<Spreadsheet.Row> BuildRows(
            IReadOnlyList<MappedSeries> mapped,
            IReadOnlyList<double> orderedX,
            uint headerRowIndex,
            uint dataStartRowIndex,
            int xValueCol,
            int xMeaningCol,
            int yMeaningCol,
            int firstSeriesCol,
            string xAxisTitle,
            string yAxisTitle)
        {
            List<Spreadsheet.Row> rows = new List<Spreadsheet.Row>();

            Spreadsheet.Row header = new Spreadsheet.Row { RowIndex = headerRowIndex };
            header.Append(
                CreateTextCell(CellRef(xValueCol, headerRowIndex), string.IsNullOrWhiteSpace(xAxisTitle) ? "X Value" : xAxisTitle),
                CreateTextCell(CellRef(xMeaningCol, headerRowIndex), "X Meaning"),
                CreateTextCell(CellRef(yMeaningCol, headerRowIndex), string.IsNullOrWhiteSpace(yAxisTitle) ? "Y Meaning" : yAxisTitle));

            for (int i = 0; i < mapped.Count; i++)
            {
                header.Append(CreateTextCell(CellRef(firstSeriesCol + i, headerRowIndex), mapped[i].Name));
            }

            rows.Add(header);

            List<Dictionary<double, MappedPoint>> lookups = mapped
                .Select(s => s.Points
                    .GroupBy(p => p.X)
                    .ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int i = 0; i < orderedX.Count; i++)
            {
                uint rowIndex = dataStartRowIndex + (uint)i;
                double x = orderedX[i];

                Spreadsheet.Row row = new Spreadsheet.Row { RowIndex = rowIndex };
                row.Append(CreateNumberCell(CellRef(xValueCol, rowIndex), x));

                bool isWhole = Math.Abs(x - Math.Round(x)) < 1e-9;
                string xMeaning = isWhole ? (ResolveLabelForX(mapped, x) ?? string.Empty) : string.Empty;
                row.Append(CreateTextCell(CellRef(xMeaningCol, rowIndex), xMeaning));
                row.Append(CreateTextCell(CellRef(yMeaningCol, rowIndex), isWhole ? (yAxisTitle ?? string.Empty) : string.Empty));

                for (int s = 0; s < mapped.Count; s++)
                {
                    MappedPoint point;
                    if (lookups[s].TryGetValue(x, out point) && point.Y.HasValue)
                    {
                        row.Append(CreateNumberCell(CellRef(firstSeriesCol + s, rowIndex), point.Y.Value));
                    }
                    else
                    {
                        row.Append(CreateFormulaCell(CellRef(firstSeriesCol + s, rowIndex), "NA()"));
                    }
                }

                rows.Add(row);
            }

            return rows;
        }

        private static void AddScatterChart(
            DrawingsPart drawingsPart,
            string worksheetName,
            IReadOnlyList<MappedSeries> mapped,
            int pointCount,
            uint headerRowIndex,
            uint dataStartRowIndex,
            int xValueCol,
            int firstSeriesCol,
            string chartTitle,
            string xAxisTitle,
            string yAxisTitle)
        {
            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            GenerateChartPartContent(chartPart, worksheetName, mapped, pointCount, headerRowIndex, dataStartRowIndex, xValueCol, firstSeriesCol, chartTitle, xAxisTitle, yAxisTitle);

            string chartRelId = drawingsPart.GetIdOfPart(chartPart);

            Xdr.GraphicFrame graphicFrame = new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Scatter Chart" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()
                ),
                new Xdr.Transform(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartRelId })
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            Xdr.TwoCellAnchor anchor = new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((dataStartRowIndex + (uint)pointCount + 2).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("14"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId((dataStartRowIndex + (uint)pointCount + 22).ToString(CultureInfo.InvariantCulture)),
                    new Xdr.RowOffset("0")),
                graphicFrame,
                new Xdr.ClientData());

            drawingsPart.WorksheetDrawing.Append(anchor);
        }

        private static void GenerateChartPartContent(
            ChartPart chartPart,
            string worksheetName,
            IReadOnlyList<MappedSeries> mapped,
            int pointCount,
            uint headerRowIndex,
            uint dataStartRowIndex,
            int xValueCol,
            int firstSeriesCol,
            string chartTitle,
            string xAxisTitle,
            string yAxisTitle)
        {
            const uint xAxisId = 48650112U;
            const uint yAxisId = 48672768U;

            C.ChartSpace chartSpace = new C.ChartSpace();
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });
            chartSpace.Append(new C.RoundedCorners { Val = false });

            C.Chart chart = new C.Chart();

            if (!string.IsNullOrWhiteSpace(chartTitle))
                chart.Append(CreateChartTitle(chartTitle));

            chart.Append(new C.AutoTitleDeleted { Val = false });

            C.PlotArea plotArea = new C.PlotArea(
                new C.Layout(
                    new C.ManualLayout(
                        new C.LayoutTarget { Val = C.LayoutTargetValues.Inner },
                        new C.LeftMode { Val = C.LayoutModeValues.Edge },
                        new C.TopMode { Val = C.LayoutModeValues.Edge },
                        new C.WidthMode { Val = C.LayoutModeValues.Edge },
                        new C.HeightMode { Val = C.LayoutModeValues.Edge },
                        new C.Left { Val = 0.08D },
                        new C.Top { Val = 0.10D },
                        new C.Width { Val = 0.78D },
                        new C.Height { Val = 0.75D })));

            C.ScatterChart scatterChart = new C.ScatterChart(
                new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                new C.VaryColors { Val = false });

            string xRange = RangeRef(worksheetName, xValueCol, dataStartRowIndex, xValueCol, dataStartRowIndex + (uint)pointCount - 1);

            for (int i = 0; i < mapped.Count; i++)
            {
                int yCol = firstSeriesCol + i;
                string yRange = RangeRef(worksheetName, yCol, dataStartRowIndex, yCol, dataStartRowIndex + (uint)pointCount - 1);
                string titleRef = RangeRef(worksheetName, yCol, headerRowIndex, yCol, headerRowIndex);
                string seriesColor = mapped[i].HexColor;

                C.ScatterChartSeries ser = new C.ScatterChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i },
                    new C.SeriesText(new C.StringReference(new C.Formula(titleRef))),
                    new C.ChartShapeProperties(
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = seriesColor }),
                            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                        { Width = 28575 },
                        new A.EffectList()),
                    new C.Marker(
                        new C.Symbol { Val = C.MarkerStyleValues.Circle },
                        new C.Size { Val = 8 },
                        new C.ChartShapeProperties(
                            new A.SolidFill(new A.RgbColorModelHex { Val = seriesColor }),
                            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = seriesColor }))
                        )),
                    new C.Smooth { Val = false },
                    new C.XValues(new C.NumberReference(new C.Formula(xRange))),
                    new C.YValues(new C.NumberReference(new C.Formula(yRange)))
                );

                scatterChart.Append(ser);
            }

            scatterChart.Append(new C.AxisId { Val = xAxisId });
            scatterChart.Append(new C.AxisId { Val = yAxisId });

            plotArea.Append(scatterChart);
            plotArea.Append(CreateValueAxis(xAxisId, yAxisId, xAxisTitle, C.AxisPositionValues.Bottom, false));
            plotArea.Append(CreateValueAxis(yAxisId, xAxisId, yAxisTitle, C.AxisPositionValues.Left, true));

            chart.Append(plotArea);
            chart.Append(new C.PlotVisibleOnly { Val = true });
            chart.Append(new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Right },
                new C.Layout(),
                new C.Overlay { Val = false }));

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
        }

        private static C.Title CreateChartTitle(string text)
        {
            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                            new A.Run(
                                new A.RunProperties { Language = "en-US", FontSize = 1800 },
                                new A.Text(text))))),
                new C.Layout(),
                new C.Overlay { Val = false });
        }

        private static C.ValueAxis CreateValueAxis(
            uint axisId,
            uint crossesAxisId,
            string title,
            C.AxisPositionValues position,
            bool includeMajorGridlines)
        {
            C.ValueAxis axis = new C.ValueAxis(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = position },
                CreateAxisTitle(title),
                new C.NumberingFormat { FormatCode = "0", SourceLinked = false });

            if (includeMajorGridlines)
            {
                axis.Append(new C.MajorGridlines(
                    new C.ChartShapeProperties(
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "D9D9D9" }),
                            new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                        { Width = 9525 })));
            }

            axis.Append(
                new C.MajorTickMark { Val = C.TickMarkValues.None },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.ChartShapeProperties(
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = "BFBFBF" }))
                    { Width = 9525 }),
                new C.TextProperties(
                    new A.BodyProperties(),
                    new A.ListStyle(),
                    new A.Paragraph(
                        new A.ParagraphProperties(),
                        new A.Run(
                            new A.RunProperties { Language = "en-US", FontSize = 1100 },
                            new A.Text(string.Empty)))),
                new C.CrossingAxis { Val = crossesAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });

            return axis;
        }

        private static OpenXmlElement CreateAxisTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
                return new C.Title(new C.Layout(), new C.Overlay { Val = false });

            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US", FontSize = 1200 },
                                new A.Text(title))))),
                new C.Layout(),
                new C.Overlay { Val = false });
        }

        private static Spreadsheet.Column CreateColumn(uint min, uint max, double width)
        {
            return new Spreadsheet.Column
            {
                Min = min,
                Max = max,
                Width = width,
                CustomWidth = true
            };
        }

        private static Spreadsheet.Cell CreateTextCell(string cellReference, string text)
        {
            return new Spreadsheet.Cell
            {
                CellReference = cellReference,
                DataType = Spreadsheet.CellValues.InlineString,
                InlineString = new Spreadsheet.InlineString(new Spreadsheet.Text(text ?? string.Empty))
            };
        }

        private static Spreadsheet.Cell CreateNumberCell(string cellReference, double value)
        {
            return new Spreadsheet.Cell
            {
                CellReference = cellReference,
                DataType = Spreadsheet.CellValues.Number,
                CellValue = new Spreadsheet.CellValue(value.ToString(CultureInfo.InvariantCulture))
            };
        }

        private static Spreadsheet.Cell CreateFormulaCell(string cellReference, string formula)
        {
            return new Spreadsheet.Cell
            {
                CellReference = cellReference,
                CellFormula = new Spreadsheet.CellFormula(formula)
            };
        }

        private static string CellRef(int colIndex1Based, uint rowIndex1Based)
        {
            return ColumnName(colIndex1Based) + rowIndex1Based.ToString(CultureInfo.InvariantCulture);
        }

        private static string ColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier) + columnName;
                dividend = (dividend - modifier - 1) / 26;
            }

            return columnName;
        }

        private static string RangeRef(string sheetName, int fromCol, uint fromRow, int toCol, uint toRow)
        {
            return "'" + sheetName + "'!$" + ColumnName(fromCol) + "$" + fromRow.ToString(CultureInfo.InvariantCulture)
                 + ":$" + ColumnName(toCol) + "$" + toRow.ToString(CultureInfo.InvariantCulture);
        }

        private static string ResolveLabelForX(IReadOnlyList<MappedSeries> mapped, double x)
        {
            return mapped
                .SelectMany(s => s.Points)
                .Where(p => Math.Abs(p.X - x) < 1e-9)
                .Select(p => p.XMeaning)
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        }

        private sealed class MappedSeries
        {
            public MappedSeries(string name, string hexColor, List<MappedPoint> points)
            {
                Name = name;
                HexColor = hexColor;
                Points = points;
            }

            public string Name { get; private set; }
            public string HexColor { get; private set; }
            public List<MappedPoint> Points { get; private set; }
        }

        private sealed class MappedPoint
        {
            public MappedPoint(double x, double? y, string xMeaning)
            {
                X = x;
                Y = y;
                XMeaning = xMeaning;
            }

            public double X { get; private set; }
            public double? Y { get; private set; }
            public string XMeaning { get; private set; }
        }
    }
}
