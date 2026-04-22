using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Drawing.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace WpfChartExcelExport
{
    public static class OpenXmlScatterChartExporter
    {
        public static void Export<TRData>(
            string outputPath,
            IEnumerable<SeriesInput<TRData>> series,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string?>? xMeaningSelector = null,
            string? chartTitle = null,
            string? xAxisTitle = null,
            string? yAxisTitle = null,
            string worksheetName = "Chart Data",
            IComparer<double>? xComparer = null)
        {
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path is required.", nameof(outputPath));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var mapped = series
                .Select(s => new MappedSeries(
                    s.Name,
                    s.HexColor ?? "#000000",
                    s.Values?.Select(v => new MappedPoint(
                        xSelector(v),
                        ySelector(v),
                        xMeaningSelector?.Invoke(v))).ToList() ?? new List<MappedPoint>()))
                .ToList();

            if (mapped.Count == 0) throw new InvalidOperationException("No series supplied.");

            var orderedX = mapped
                .SelectMany(s => s.Points)
                .Select(p => p.X)
                .Distinct()
                .OrderBy(v => v, xComparer ?? Comparer<double>.Default)
                .ToList();

            if (orderedX.Count == 0) throw new InvalidOperationException("No data points found.");

            if (System.IO.File.Exists(outputPath))
            {
                System.IO.File.Delete(outputPath);
            }

            using var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook);

            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

            var sheets = workbookPart.Workbook.AppendChild(new Sheets());
            sheets.Append(new Sheet
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

            var sheetData = new SheetData();
            foreach (var row in BuildRows(mapped, orderedX, headerRowIndex, dataStartRowIndex, xValueCol, xMeaningCol, yMeaningCol, firstSeriesCol, xAxisTitle, yAxisTitle))
            {
                sheetData.Append(row);
            }

            worksheetPart.Worksheet = new Worksheet(
                new Columns(
                    CreateColumn(1, 1, 12),
                    CreateColumn(2, 2, 24),
                    CreateColumn(3, 3, 18),
                    CreateColumn((uint)firstSeriesCol, (uint)(firstSeriesCol + mapped.Count - 1), 14)),
                sheetData,
                new Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

            worksheetPart.Worksheet.Save();

            drawingsPart.WorksheetDrawing = new WorksheetDrawing();
            AddScatterChart(workbookPart, drawingsPart, worksheetName, mapped, orderedX.Count, headerRowIndex, dataStartRowIndex, xValueCol, firstSeriesCol, chartTitle, xAxisTitle, yAxisTitle);
            drawingsPart.WorksheetDrawing.Save();

            workbookPart.Workbook.Save();
        }

        private static IEnumerable<Row> BuildRows(
            IReadOnlyList<MappedSeries> mapped,
            IReadOnlyList<double> orderedX,
            uint headerRowIndex,
            uint dataStartRowIndex,
            int xValueCol,
            int xMeaningCol,
            int yMeaningCol,
            int firstSeriesCol,
            string? xAxisTitle,
            string? yAxisTitle)
        {
            var rows = new List<Row>();

            var header = new Row { RowIndex = headerRowIndex };
            header.Append(
                CreateTextCell(CellRef(xValueCol, headerRowIndex), string.IsNullOrWhiteSpace(xAxisTitle) ? "X Value" : xAxisTitle!),
                CreateTextCell(CellRef(xMeaningCol, headerRowIndex), "X Meaning"),
                CreateTextCell(CellRef(yMeaningCol, headerRowIndex), string.IsNullOrWhiteSpace(yAxisTitle) ? "Y Meaning" : yAxisTitle!));

            for (int i = 0; i < mapped.Count; i++)
            {
                header.Append(CreateTextCell(CellRef(firstSeriesCol + i, headerRowIndex), mapped[i].Name));
            }

            rows.Add(header);

            var lookups = mapped
                .Select(s => s.Points.GroupBy(p => p.X).ToDictionary(g => g.Key, g => g.Last()))
                .ToList();

            for (int i = 0; i < orderedX.Count; i++)
            {
                var rowIndex = dataStartRowIndex + (uint)i;
                var x = orderedX[i];
                var row = new Row { RowIndex = rowIndex };

                row.Append(CreateNumberCell(CellRef(xValueCol, rowIndex), x));

                bool isWhole = Math.Abs(x - Math.Round(x)) < 1e-9;
                string? xMeaning = isWhole ? ResolveLabelForX(mapped, x) : null;
                row.Append(CreateTextCell(CellRef(xMeaningCol, rowIndex), xMeaning ?? string.Empty));
                row.Append(CreateTextCell(CellRef(yMeaningCol, rowIndex), isWhole ? (yAxisTitle ?? string.Empty) : string.Empty));

                for (int s = 0; s < mapped.Count; s++)
                {
                    if (lookups[s].TryGetValue(x, out var point) && point.Y.HasValue)
                    {
                        row.Append(CreateNumberCell(CellRef(firstSeriesCol + s, rowIndex), point.Y.Value));
                    }
                    else
                    {
                        row.Append(CreateBlankCell(CellRef(firstSeriesCol + s, rowIndex)));
                    }
                }

                rows.Add(row);
            }

            return rows;
        }

        private static void AddScatterChart(
            WorkbookPart workbookPart,
            DrawingsPart drawingsPart,
            string worksheetName,
            IReadOnlyList<MappedSeries> mapped,
            int pointCount,
            uint headerRowIndex,
            uint dataStartRowIndex,
            int xValueCol,
            int firstSeriesCol,
            string? chartTitle,
            string? xAxisTitle,
            string? yAxisTitle)
        {
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            GenerateChartPartContent(chartPart, worksheetName, mapped, pointCount, headerRowIndex, dataStartRowIndex, xValueCol, firstSeriesCol, chartTitle, xAxisTitle, yAxisTitle);

            var chartRelId = drawingsPart.GetIdOfPart(chartPart);

            var graphicFrame = new GraphicFrame(
                new NonVisualGraphicFrameProperties(
                    new NonVisualDrawingProperties { Id = 2U, Name = "Scatter Chart" },
                    new NonVisualGraphicFrameDrawingProperties()),
                new Transform(
                    new Offset { X = 0L, Y = 0L },
                    new Extents { Cx = 0L, Cy = 0L }),
                new Graphic(
                    new GraphicData(
                        new DocumentFormat.OpenXml.Drawing.Charts.ChartReference { Id = chartRelId })
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            var anchor = new TwoCellAnchor(
                new FromMarker(
                    new ColumnId("0"),
                    new ColumnOffset("0"),
                    new RowId((dataStartRowIndex + (uint)pointCount + 2).ToString(CultureInfo.InvariantCulture)),
                    new RowOffset("0")),
                new ToMarker(
                    new ColumnId("14"),
                    new ColumnOffset("0"),
                    new RowId((dataStartRowIndex + (uint)pointCount + 22).ToString(CultureInfo.InvariantCulture)),
                    new RowOffset("0")),
                graphicFrame,
                new ClientData());

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
            string? chartTitle,
            string? xAxisTitle,
            string? yAxisTitle)
        {
            const uint xAxisId = 48650112U;
            const uint yAxisId = 48672768U;

            var chartSpace = new ChartSpace();
            chartSpace.Append(new EditingLanguage { Val = "en-US" });

            var chart = new DocumentFormat.OpenXml.Drawing.Charts.Chart();

            if (!string.IsNullOrWhiteSpace(chartTitle))
            {
                chart.Append(CreateChartTitle(chartTitle!));
            }

            chart.Append(new AutoTitleDeleted { Val = false });

            var plotArea = new PlotArea(new Layout());

            var scatterChart = new ScatterChart(
                new ScatterStyle { Val = ScatterStyleValues.LineMarker },
                new VaryColors { Val = false });

            var xRange = RangeRef(worksheetName, xValueCol, dataStartRowIndex, xValueCol, dataStartRowIndex + (uint)pointCount - 1);

            for (int i = 0; i < mapped.Count; i++)
            {
                int yCol = firstSeriesCol + i;
                string yRange = RangeRef(worksheetName, yCol, dataStartRowIndex, yCol, dataStartRowIndex + (uint)pointCount - 1);
                string titleRef = RangeRef(worksheetName, yCol, headerRowIndex, yCol, headerRowIndex);

                var series = new ScatterChartSeries(
                    new Index { Val = (uint)i },
                    new Order { Val = (uint)i },
                    new SeriesText(new StringReference(new Formula(titleRef))),
                    CreateSeriesShapeProperties(mapped[i].HexColor),
                    new Marker(
                        new Symbol { Val = MarkerStyleValues.Circle },
                        new Size { Val = 7 }),
                    new XValues(new NumberReference(new Formula(xRange))),
                    new YValues(new NumberReference(new Formula(yRange))));

                scatterChart.Append(series);
            }

            scatterChart.Append(
                new AxisId { Val = xAxisId },
                new AxisId { Val = yAxisId });

            plotArea.Append(scatterChart);
            plotArea.Append(CreateCategoryAxis(xAxisId, yAxisId, xAxisTitle));
            plotArea.Append(CreateValueAxis(yAxisId, xAxisId, yAxisTitle));

            chart.Append(plotArea);
            chart.Append(new PlotVisibleOnly { Val = true });
            chart.Append(new Legend(
                new LegendPosition { Val = LegendPositionValues.Right },
                new Layout(),
                new Overlay { Val = false }));

            chartSpace.Append(chart);
            chartPart.ChartSpace = chartSpace;
        }

        private static DocumentFormat.OpenXml.Drawing.Charts.Title CreateChartTitle(string text)
        {
            return new DocumentFormat.OpenXml.Drawing.Charts.Title(
                new ChartText(
                    new RichText(
                        new BodyProperties(),
                        new ListStyle(),
                        new Paragraph(
                            new Run(
                                new RunProperties { Language = "en-US" },
                                new Text(text))))),
                new Layout(),
                new Overlay { Val = false });
        }

        private static ShapeProperties CreateSeriesShapeProperties(string hexColor)
        {
            return new ShapeProperties(
                new Outline(
                    new SolidFill(new RgbColorModelHex { Val = HexToRgb(hexColor) }),
                    new PresetDash { Val = PresetLineDashValues.Solid })
                { Width = 19050 });
        }

        private static CategoryAxis CreateCategoryAxis(uint axisId, uint crossesAxisId, string? title)
        {
            return new CategoryAxis(
                new AxisId { Val = axisId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Bottom },
                CreateAxisTitle(title),
                new NumberingFormat { FormatCode = "0", SourceLinked = false },
                new MajorTickMark { Val = TickMarkValues.Outside },
                new MinorTickMark { Val = TickMarkValues.None },
                new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis { Val = crossesAxisId },
                new Crosses { Val = CrossesValues.AutoZero },
                new AutoLabeled { Val = true },
                new LabelAlignment { Val = LabelAlignmentValues.Center },
                new LabelOffset { Val = 100 });
        }

        private static ValueAxis CreateValueAxis(uint axisId, uint crossesAxisId, string? title)
        {
            return new ValueAxis(
                new AxisId { Val = axisId },
                new Scaling(new Orientation { Val = OrientationValues.MinMax }),
                new Delete { Val = false },
                new AxisPosition { Val = AxisPositionValues.Left },
                CreateAxisTitle(title),
                new NumberingFormat { FormatCode = "0", SourceLinked = false },
                new MajorGridlines(),
                new MajorTickMark { Val = TickMarkValues.Outside },
                new MinorTickMark { Val = TickMarkValues.None },
                new TickLabelPosition { Val = TickLabelPositionValues.NextTo },
                new CrossingAxis { Val = crossesAxisId },
                new Crosses { Val = CrossesValues.AutoZero },
                new CrossBetween { Val = CrossBetweenValues.MidpointCategory });
        }

        private static OpenXmlElement CreateAxisTitle(string? title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return new Title(new Layout(), new Overlay { Val = false });
            }

            return new Title(
                new ChartText(
                    new RichText(
                        new BodyProperties(),
                        new ListStyle(),
                        new Paragraph(
                            new Run(
                                new RunProperties { Language = "en-US" },
                                new Text(title))))),
                new Layout(),
                new Overlay { Val = false });
        }

        private static Column CreateColumn(uint min, uint max, double width)
        {
            return new Column { Min = min, Max = max, Width = width, CustomWidth = true };
        }

        private static Cell CreateTextCell(string cellReference, string text)
        {
            return new Cell
            {
                CellReference = cellReference,
                DataType = CellValues.InlineString,
                InlineString = new InlineString(new Text(text))
            };
        }

        private static Cell CreateNumberCell(string cellReference, double value)
        {
            return new Cell
            {
                CellReference = cellReference,
                CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture)),
                DataType = CellValues.Number
            };
        }

        private static Cell CreateBlankCell(string cellReference)
        {
            return new Cell { CellReference = cellReference };
        }

        private static string CellRef(int colIndex1Based, uint rowIndex1Based)
        {
            return $"{ColumnName(colIndex1Based)}{rowIndex1Based}";
        }

        private static string ColumnName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = string.Empty;

            while (dividend > 0)
            {
                int modifier = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modifier) + columnName;
                dividend = (dividend - modifier) / 26;
            }

            return columnName;
        }

        private static string RangeRef(string sheetName, int fromCol, uint fromRow, int toCol, uint toRow)
        {
            return $"'{sheetName}'!${ColumnName(fromCol)}${fromRow}:${ColumnName(toCol)}${toRow}";
        }

        private static string? ResolveLabelForX(IReadOnlyList<MappedSeries> mapped, double x)
        {
            return mapped
                .SelectMany(s => s.Points)
                .Where(p => Math.Abs(p.X - x) < 1e-9)
                .Select(p => p.XMeaning)
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        }

        private static string HexToRgb(string hex)
        {
            string clean = hex.TrimStart('#');
            return clean.Length == 6 ? clean.ToUpperInvariant() : "000000";
        }

        public sealed class SeriesInput<T>
        {
            public required string Name { get; init; }
            public required IEnumerable<T> Values { get; init; }
            public string? HexColor { get; init; }
        }

        private sealed record MappedSeries(string Name, string HexColor, List<MappedPoint> Points);
        private sealed record MappedPoint(double X, double? Y, string? XMeaning);
    }
}
