using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace WpfChartExcelExport
{
    public static class OpenXmlScatterChartExporter
    {
        public static void Export<TRData>(
            string outputPath,
            IEnumerable<SeriesInput<TRData>> series,
            Func<TRData, double> xSelector,
            Func<TRData, double?> ySelector,
            Func<TRData, string> xMeaningSelector = null,
            string chartTitle = null,
            string xAxisTitle = null,
            string yAxisTitle = null,
            string worksheetName = "Chart Data",
            IComparer<double> xComparer = null)
        {
            if (string.IsNullOrWhiteSpace(outputPath)) throw new ArgumentException("Output path is required.", nameof(outputPath));
            if (series == null) throw new ArgumentNullException(nameof(series));
            if (xSelector == null) throw new ArgumentNullException(nameof(xSelector));
            if (ySelector == null) throw new ArgumentNullException(nameof(ySelector));

            var mapped = series
                .Select(s => new MappedSeries(
                    s.Name,
                    string.IsNullOrWhiteSpace(s.HexColor) ? "#000000" : s.HexColor,
                    s.Values == null
                        ? new List<MappedPoint>()
                        : s.Values.Select(v => new MappedPoint(
                            xSelector(v),
                            ySelector(v),
                            xMeaningSelector != null ? xMeaningSelector(v) : null)).ToList()))
                .ToList();

            if (mapped.Count == 0) throw new InvalidOperationException("No series supplied.");

            var comparer = xComparer ?? Comparer<double>.Default;
            var orderedX = mapped
                .SelectMany(s => s.Points)
                .Select(p => p.X)
                .Distinct()
                .OrderBy(v => v, comparer)
                .ToList();

            if (orderedX.Count == 0) throw new InvalidOperationException("No data points found.");

            if (System.IO.File.Exists(outputPath))
            {
                System.IO.File.Delete(outputPath);
            }

            using (var document = SpreadsheetDocument.Create(outputPath, SpreadsheetDocumentType.Workbook))
            {
                var workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Spreadsheet.Workbook();

                var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();

                var sheets = workbookPart.Workbook.AppendChild(new Spreadsheet.Sheets());
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

                var sheetData = new Spreadsheet.SheetData();
                foreach (var row in BuildRows(mapped, orderedX, headerRowIndex, dataStartRowIndex, xValueCol, xMeaningCol, yMeaningCol, firstSeriesCol, xAxisTitle, yAxisTitle))
                {
                    sheetData.Append(row);
                }

                worksheetPart.Worksheet = new Spreadsheet.Worksheet(
                    new Spreadsheet.Columns(
                        CreateColumn(1, 1, 12),
                        CreateColumn(2, 2, 24),
                        CreateColumn(3, 3, 18),
                        CreateColumn((uint)firstSeriesCol, (uint)(firstSeriesCol + mapped.Count - 1), 14)),
                    sheetData,
                    new Spreadsheet.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });

                worksheetPart.Worksheet.Save();

                drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
                AddScatterChart(drawingsPart, worksheetName, mapped, orderedX.Count, headerRowIndex, dataStartRowIndex, xValueCol, firstSeriesCol, chartTitle, xAxisTitle, yAxisTitle);
                drawingsPart.WorksheetDrawing.Save();

                workbookPart.Workbook.Save();
            }
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
            var rows = new List<Spreadsheet.Row>();

            var header = new Spreadsheet.Row { RowIndex = headerRowIndex };
            header.Append(
                CreateTextCell(CellRef(xValueCol, headerRowIndex), string.IsNullOrWhiteSpace(xAxisTitle) ? "X Value" : xAxisTitle),
                CreateTextCell(CellRef(xMeaningCol, headerRowIndex), "X Meaning"),
                CreateTextCell(CellRef(yMeaningCol, headerRowIndex), string.IsNullOrWhiteSpace(yAxisTitle) ? "Y Meaning" : yAxisTitle));

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
                uint rowIndex = dataStartRowIndex + (uint)i;
                double x = orderedX[i];
                var row = new Spreadsheet.Row { RowIndex = rowIndex };

                row.Append(CreateNumberCell(CellRef(xValueCol, rowIndex), x));

                bool isWhole = Math.Abs(x - Math.Round(x)) < 1e-9;
                string xMeaning = isWhole ? (ResolveLabelForX(mapped, x) ?? string.Empty) : string.Empty;
                row.Append(CreateTextCell(CellRef(xMeaningCol, rowIndex), xMeaning));
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
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            GenerateChartPartContent(chartPart, worksheetName, mapped, pointCount, headerRowIndex, dataStartRowIndex, xValueCol, firstSeriesCol, chartTitle, xAxisTitle, yAxisTitle);

            var chartRelId = drawingsPart.GetIdOfPart(chartPart);

            var graphicFrame = new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Scatter Chart" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(
                    new A.GraphicData(
                        new C.ChartReference { Id = chartRelId })
                    { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" }));

            var anchor = new Xdr.TwoCellAnchor(
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

            var chartSpace = new C.ChartSpace();
            chartSpace.Append(new C.EditingLanguage { Val = "en-US" });

            var chart = new C.Chart();

            if (!string.IsNullOrWhiteSpace(chartTitle))
            {
                chart.Append(CreateChartTitle(chartTitle));
            }

            chart.Append(new C.AutoTitleDeleted { Val = false });

            var plotArea = new C.PlotArea(new C.Layout());

            var scatterChart = new C.ScatterChart(
                new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
                new C.VaryColors { Val = false });

            string xRange = RangeRef(worksheetName, xValueCol, dataStartRowIndex, xValueCol, dataStartRowIndex + (uint)pointCount - 1);

            for (int i = 0; i < mapped.Count; i++)
            {
                int yCol = firstSeriesCol + i;
                string yRange = RangeRef(worksheetName, yCol, dataStartRowIndex, yCol, dataStartRowIndex + (uint)pointCount - 1);
                string titleRef = RangeRef(worksheetName, yCol, headerRowIndex, yCol, headerRowIndex);

                var series = new C.ScatterChartSeries(
                    new C.Index { Val = (uint)i },
                    new C.Order { Val = (uint)i },
                    new C.SeriesText(new C.StringReference(new C.Formula(titleRef))),
                    CreateSeriesShapeProperties(mapped[i].HexColor),
                    new C.Marker(
                        new C.Symbol { Val = C.MarkerStyleValues.Circle },
                        new C.Size { Val = 7 }),
                    new C.XValues(new C.NumberReference(new C.Formula(xRange))),
                    new C.YValues(new C.NumberReference(new C.Formula(yRange))));

                scatterChart.Append(series);
            }

            scatterChart.Append(
                new C.AxisId { Val = xAxisId },
                new C.AxisId { Val = yAxisId });

            plotArea.Append(scatterChart);
            plotArea.Append(CreateCategoryAxis(xAxisId, yAxisId, xAxisTitle));
            plotArea.Append(CreateValueAxis(yAxisId, xAxisId, yAxisTitle));

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
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(text))))),
                new C.Layout(),
                new C.Overlay { Val = false });
        }

        private static A.ShapeProperties CreateSeriesShapeProperties(string hexColor)
        {
            return new A.ShapeProperties(
                new A.Outline(
                    new A.SolidFill(new A.RgbColorModelHex { Val = HexToRgb(hexColor) }),
                    new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                { Width = 19050 });
        }

        private static C.CategoryAxis CreateCategoryAxis(uint axisId, uint crossesAxisId, string title)
        {
            return new C.CategoryAxis(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
                CreateAxisTitle(title),
                new C.NumberingFormat { FormatCode = "0", SourceLinked = false },
                new C.MajorTickMark { Val = C.TickMarkValues.Outside },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossesAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.AutoLabeled { Val = true },
                new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
                new C.LabelOffset { Val = 100 });
        }

        private static C.ValueAxis CreateValueAxis(uint axisId, uint crossesAxisId, string title)
        {
            return new C.ValueAxis(
                new C.AxisId { Val = axisId },
                new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
                new C.Delete { Val = false },
                new C.AxisPosition { Val = C.AxisPositionValues.Left },
                CreateAxisTitle(title),
                new C.NumberingFormat { FormatCode = "0", SourceLinked = false },
                new C.MajorGridlines(),
                new C.MajorTickMark { Val = C.TickMarkValues.Outside },
                new C.MinorTickMark { Val = C.TickMarkValues.None },
                new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
                new C.CrossingAxis { Val = crossesAxisId },
                new C.Crosses { Val = C.CrossesValues.AutoZero },
                new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });
        }

        private static OpenXmlElement CreateAxisTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return new C.Title(new C.Layout(), new C.Overlay { Val = false });
            }

            return new C.Title(
                new C.ChartText(
                    new C.RichText(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.Run(
                                new A.RunProperties { Language = "en-US" },
                                new A.Text(title))))),
                new C.Layout(),
                new C.Overlay { Val = false });
        }

        private static Spreadsheet.Column CreateColumn(uint min, uint max, double width)
        {
            return new Spreadsheet.Column { Min = min, Max = max, Width = width, CustomWidth = true };
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
                CellValue = new Spreadsheet.CellValue(value.ToString(CultureInfo.InvariantCulture)),
                DataType = Spreadsheet.CellValues.Number
            };
        }

        private static Spreadsheet.Cell CreateBlankCell(string cellReference)
        {
            return new Spreadsheet.Cell { CellReference = cellReference };
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
            return "'" + sheetName + "'!$" + ColumnName(fromCol) + "$" + fromRow.ToString(CultureInfo.InvariantCulture) + ":$" + ColumnName(toCol) + "$" + toRow.ToString(CultureInfo.InvariantCulture);
        }

        private static string ResolveLabelForX(IReadOnlyList<MappedSeries> mapped, double x)
        {
            return mapped
                .SelectMany(s => s.Points)
                .Where(p => Math.Abs(p.X - x) < 1e-9)
                .Select(p => p.XMeaning)
                .FirstOrDefault(v => !string.IsNullOrWhiteSpace(v));
        }

        private static string HexToRgb(string hex)
        {
            string clean = (hex ?? string.Empty).TrimStart('#');
            return clean.Length == 6 ? clean.ToUpperInvariant() : "000000";
        }

        public sealed class SeriesInput<T>
        {
            public string Name { get; set; }
            public IEnumerable<T> Values { get; set; }
            public string HexColor { get; set; }
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
