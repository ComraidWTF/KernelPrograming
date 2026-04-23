#nullable disable
using System;
using System.Collections;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using LiveChartsCore;
using Spreadsheet = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace WpfChartExcelExport;

public sealed class YAxisLabel
{
    public double Position { get; set; }
    public string Text { get; set; }
}

public sealed class ChartExportSpec
{
    public IEnumerable<ISeries> Series { get; set; }
    public Func<object, double> XSelector { get; set; }
    public Func<object, double?> YSelector { get; set; }
    public string ChartTitle { get; set; }
    public string XAxisTitle { get; set; }
    public string YAxisTitle { get; set; }
    public IComparer<double> XComparer { get; set; }
    public double? YMin { get; set; }
    public double? YMax { get; set; }
    public double? YMajorUnit { get; set; }
    public double? XMajorUnit { get; set; }
    public string XTickFormat { get; set; }
    public double? XMin { get; set; }
    public double? XMax { get; set; }
    public IReadOnlyList<YAxisLabel> YCustomLabels { get; set; }
}

public sealed class LiveChartsExcelExporter
{
    private const int XValueCol = 1;
    private const int FirstSeriesCol = 2;
    private const int ChartRowSpan = 22;
    private const int BlockRowGap = 3;
    private const uint WorksheetHeaderRow = 1U;
    private const uint FirstBlockRow = 3U;
    private const uint XAxisId = 48650112U;
    private const uint YAxisId = 48672768U;

    private static readonly string[] AccentFallback =
        { "5B9BD5", "ED7D31", "A5A5A5", "FFC000", "4472C4", "70AD47" };

    private readonly string _filePath;
    private readonly Dictionary<string, List<ChartExportSpec>> _sheets = new(StringComparer.Ordinal);
    private readonly List<string> _sheetOrder = new();

    public LiveChartsExcelExporter(string filePath)
    {
        if (string.IsNullOrWhiteSpace(filePath))
            throw new ArgumentException("File path is required.", nameof(filePath));
        _filePath = filePath;
    }

    public void AddWorksheet(string name)
    {
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("Worksheet name is required.", nameof(name));
        if (_sheets.ContainsKey(name))
            throw new InvalidOperationException($"Worksheet '{name}' already exists.");
        _sheets[name] = new List<ChartExportSpec>();
        _sheetOrder.Add(name);
    }

    public void AddChart(string worksheetName, ChartExportSpec spec)
    {
        if (spec == null) throw new ArgumentNullException(nameof(spec));
        if (!_sheets.TryGetValue(worksheetName, out var list))
            throw new InvalidOperationException($"Worksheet '{worksheetName}' not found. Call AddWorksheet first.");
        list.Add(spec);
    }

    public void Save()
    {
        if (_sheetOrder.Count == 0)
            throw new InvalidOperationException("At least one worksheet is required.");

        if (System.IO.File.Exists(_filePath)) System.IO.File.Delete(_filePath);

        using var doc = SpreadsheetDocument.Create(_filePath, SpreadsheetDocumentType.Workbook);
        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Spreadsheet.Workbook();
        wbPart.Workbook.AppendChild(new Spreadsheet.BookViews(new Spreadsheet.WorkbookView()));

        var stylesPart = wbPart.AddNewPart<WorkbookStylesPart>();
        stylesPart.Stylesheet = CreateStylesheet();
        stylesPart.Stylesheet.Save();

        var sheets = wbPart.Workbook.AppendChild(new Spreadsheet.Sheets());

        uint sheetId = 1U;
        foreach (var name in _sheetOrder)
        {
            var prepared = _sheets[name].Select(Prepare).Where(p => p != null).ToList();
            if (prepared.Count == 0) continue;

            var wsPart = wbPart.AddNewPart<WorksheetPart>();
            var drPart = wsPart.AddNewPart<DrawingsPart>();
            sheets.Append(new Spreadsheet.Sheet
            {
                Id = wbPart.GetIdOfPart(wsPart),
                SheetId = sheetId++,
                Name = name
            });
            WriteSheet(wsPart, drPart, name, prepared);
        }

        wbPart.Workbook.Save();
    }

    private static void WriteSheet(WorksheetPart wsPart, DrawingsPart drPart, string name, List<PreparedChart> prepared)
    {
        int maxSeries = prepared.Max(p => p.Mapped.Count);
        var sheetData = new Spreadsheet.SheetData();
        drPart.WorksheetDrawing = new Xdr.WorksheetDrawing();

        var titleRow1 = new Spreadsheet.Row { RowIndex = 1U, Height = 24D, CustomHeight = true };
        titleRow1.Append(StyledTextCell(CellRef(1, 1U), name, 1U));
        sheetData.Append(titleRow1);

        var titleRow2 = new Spreadsheet.Row { RowIndex = 2U, Height = 24D, CustomHeight = true };
        sheetData.Append(titleRow2);

        var titleRow3 = new Spreadsheet.Row { RowIndex = 3U, Height = 24D, CustomHeight = true };
        sheetData.Append(titleRow3);

        uint cursor = 5U;
        for (int i = 0; i < prepared.Count; i++)
        {
            var pc = prepared[i];
            uint headerRow = cursor;
            uint dataStart = cursor + 1U;

            foreach (var row in BuildRows(pc, headerRow, dataStart)) sheetData.Append(row);
            AddChart(drPart, name, pc, headerRow, dataStart, i);
            cursor = dataStart + (uint)pc.OrderedX.Count + ChartRowSpan + BlockRowGap;
        }

        var mergeCells = new Spreadsheet.MergeCells();
        mergeCells.Append(new Spreadsheet.MergeCell { Reference = new StringValue("A1:J3") });

        wsPart.Worksheet = new Spreadsheet.Worksheet(
            new Spreadsheet.SheetViews(new Spreadsheet.SheetView { WorkbookViewId = 0U }),
            new Spreadsheet.SheetFormatProperties { DefaultRowHeight = 15D },
            new Spreadsheet.Columns(
                Column((uint)XValueCol, (uint)XValueCol, 12),
                Column((uint)FirstSeriesCol, (uint)(FirstSeriesCol + maxSeries - 1), 14),
                Column(4U, 10U, 14)),
            sheetData,
            mergeCells,
            new Spreadsheet.Drawing { Id = wsPart.GetIdOfPart(drPart) });

        wsPart.Worksheet.Save();
        drPart.WorksheetDrawing.Save();
    }

    private static PreparedChart Prepare(ChartExportSpec spec)
    {
        if (spec?.Series == null) return null;
        if (spec.XSelector == null) throw new ArgumentException("XSelector is required on every ChartExportSpec.");
        if (spec.YSelector == null) throw new ArgumentException("YSelector is required on every ChartExportSpec.");

        var comparer = spec.XComparer ?? Comparer<double>.Default;
        var mapped = spec.Series
            .Select((s, i) => MapSeries(s, i, spec.XSelector, spec.YSelector))
            .Where(s => s != null).ToList();
        if (mapped.Count == 0) return null;

        var orderedX = mapped.SelectMany(s => s.Points).Select(p => p.X).Distinct()
            .OrderBy(v => v, comparer).ToList();
        return orderedX.Count == 0 ? null : new PreparedChart(spec, mapped, orderedX);
    }

    private static MappedSeries MapSeries(
        ISeries series, int index,
        Func<object, double> xSel,
        Func<object, double?> ySel)
    {
        if (series.GetType().GetProperty("Values")?.GetValue(series) is not IEnumerable raw) return null;

        var name = series.GetType().GetProperty("Name")?.GetValue(series)?.ToString();
        if (string.IsNullOrWhiteSpace(name))
            name = "Series " + (index + 1).ToString(CultureInfo.InvariantCulture);

        var color = TryHexColor(series) ?? AccentFallback[index % AccentFallback.Length];
        var points = new List<MappedPoint>();
        foreach (var item in raw)
            if (item != null) points.Add(new MappedPoint(xSel(item), ySel(item)));
        return new MappedSeries(name, color, points);
    }

    private static string TryHexColor(ISeries series)
    {
        try
        {
            var stroke = series.GetType().GetProperty("Stroke")?.GetValue(series);
            if (stroke == null) return null;
            var colorObj = stroke.GetType().GetProperty("Color")?.GetValue(stroke);
            return (colorObj != null ? RgbHex(colorObj) : null) ?? RgbHex(stroke);
        }
        catch { return null; }
    }

    private static string RgbHex(object c)
    {
        try
        {
            var t = c.GetType();
            var r = t.GetProperty("Red")?.GetValue(c);
            var g = t.GetProperty("Green")?.GetValue(c);
            var b = t.GetProperty("Blue")?.GetValue(c);
            if (r == null || g == null || b == null) return null;
            return $"{Convert.ToByte(r, CultureInfo.InvariantCulture):X2}" +
                   $"{Convert.ToByte(g, CultureInfo.InvariantCulture):X2}" +
                   $"{Convert.ToByte(b, CultureInfo.InvariantCulture):X2}";
        }
        catch { return null; }
    }

    private static IEnumerable<Spreadsheet.Row> BuildRows(PreparedChart pc, uint headerRow, uint dataStart)
    {
        var header = new Spreadsheet.Row { RowIndex = headerRow };
        header.Append(TextCell(CellRef(XValueCol, headerRow),
            string.IsNullOrWhiteSpace(pc.Spec.XAxisTitle) ? "X Value" : pc.Spec.XAxisTitle));
        for (int i = 0; i < pc.Mapped.Count; i++)
            header.Append(TextCell(CellRef(FirstSeriesCol + i, headerRow), pc.Mapped[i].Name));
        yield return header;

        var lookups = pc.Mapped
            .Select(s => s.Points.GroupBy(p => p.X).ToDictionary(g => g.Key, g => g.Last()))
            .ToList();

        for (int i = 0; i < pc.OrderedX.Count; i++)
        {
            uint rowIdx = dataStart + (uint)i;
            double x = pc.OrderedX[i];

            var row = new Spreadsheet.Row { RowIndex = rowIdx };
            row.Append(NumberCell(CellRef(XValueCol, rowIdx), x));
            for (int s = 0; s < pc.Mapped.Count; s++)
                row.Append(lookups[s].TryGetValue(x, out var pt) && pt.Y.HasValue
                    ? NumberCell(CellRef(FirstSeriesCol + s, rowIdx), pt.Y.Value)
                    : FormulaCell(CellRef(FirstSeriesCol + s, rowIdx), "NA()"));
            yield return row;
        }
    }

    private static void AddChart(DrawingsPart drawings, string sheet, PreparedChart pc, uint headerRow, uint dataStart, int chartIndex)
    {
        var chartPart = drawings.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(sheet, pc, headerRow, dataStart);

        uint top = dataStart + (uint)pc.OrderedX.Count + 2;
        uint bottom = dataStart + (uint)pc.OrderedX.Count + 22;

        drawings.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId(top.ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId("14"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId(bottom.ToString(CultureInfo.InvariantCulture)), new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties
                    {
                        Id = (uint)(chartIndex + 2),
                        Name = "Scatter Chart " + (chartIndex + 1).ToString(CultureInfo.InvariantCulture)
                    },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = drawings.GetIdOfPart(chartPart) })
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })),
            new Xdr.ClientData()));
    }

    private static C.ChartSpace BuildChartSpace(string sheet, PreparedChart pc, uint headerRow, uint dataStart)
    {
        bool customY = pc.Spec.YCustomLabels != null && pc.Spec.YCustomLabels.Count > 0;

        var chart = new C.Chart();
        if (!string.IsNullOrWhiteSpace(pc.Spec.ChartTitle))
            chart.Append(BuildTitle(pc.Spec.ChartTitle, 1800, A.TextAlignmentTypeValues.Center));
        chart.Append(new C.AutoTitleDeleted { Val = false });

        var plot = new C.PlotArea(
            new C.Layout(new C.ManualLayout(
                new C.LayoutTarget { Val = C.LayoutTargetValues.Inner },
                new C.LeftMode { Val = C.LayoutModeValues.Edge },
                new C.TopMode { Val = C.LayoutModeValues.Edge },
                new C.WidthMode { Val = C.LayoutModeValues.Edge },
                new C.HeightMode { Val = C.LayoutModeValues.Edge },
                new C.Left { Val = 0.08D }, new C.Top { Val = 0.10D },
                new C.Width { Val = 0.78D }, new C.Height { Val = 0.75D })));

        plot.Append(BuildScatterChart(sheet, pc, headerRow, dataStart, customY));
        plot.Append(BuildValueAxis(
            XAxisId, YAxisId, pc.Spec.XAxisTitle, C.AxisPositionValues.Bottom,
            pc.Spec.XMin, pc.Spec.XMax, pc.Spec.XMajorUnit,
            C.CrossesValues.Minimum,
            string.IsNullOrWhiteSpace(pc.Spec.XTickFormat) ? "0" : pc.Spec.XTickFormat));
        plot.Append(BuildValueAxis(
            YAxisId, XAxisId, pc.Spec.YAxisTitle, C.AxisPositionValues.Left,
            pc.Spec.YMin, pc.Spec.YMax, pc.Spec.YMajorUnit,
            C.CrossesValues.AutoZero,
            customY ? ";;;" : "0"));

        chart.Append(plot);
        chart.Append(new C.PlotVisibleOnly { Val = true });

        var legend = new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Right });
        if (customY)
            legend.Append(new C.LegendEntry(new C.Index { Val = (uint)pc.Mapped.Count }, new C.Delete { Val = true }));
        legend.Append(new C.Layout());
        legend.Append(new C.Overlay { Val = false });
        chart.Append(legend);

        var space = new C.ChartSpace();
        space.Append(new C.EditingLanguage { Val = "en-US" });
        space.Append(new C.RoundedCorners { Val = false });
        space.Append(chart);
        return space;
    }

    private static C.ScatterChart BuildScatterChart(string sheet, PreparedChart pc, uint headerRow, uint dataStart, bool withCustomYLabels)
    {
        var sc = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false });

        int count = pc.OrderedX.Count;
        string xRange = RangeRef(sheet, XValueCol, dataStart, XValueCol, dataStart + (uint)count - 1);

        for (int i = 0; i < pc.Mapped.Count; i++)
        {
            int yCol = FirstSeriesCol + i;
            string yRange = RangeRef(sheet, yCol, dataStart, yCol, dataStart + (uint)count - 1);
            string titleRef = RangeRef(sheet, yCol, headerRow, yCol, headerRow);
            string color = pc.Mapped[i].HexColor;

            sc.Append(new C.ScatterChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.StringReference(new C.Formula(titleRef))),
                new C.ChartShapeProperties(
                    new A.Outline(
                        new A.SolidFill(new A.RgbColorModelHex { Val = color }),
                        new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                    { Width = 28575 },
                    new A.EffectList()),
                new C.Marker(
                    new C.Symbol { Val = C.MarkerStyleValues.Circle },
                    new C.Size { Val = 8 },
                    new C.ChartShapeProperties(
                        new A.SolidFill(new A.RgbColorModelHex { Val = color }),
                        new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = color })))),
                new C.Smooth { Val = false },
                new C.XValues(new C.NumberReference(new C.Formula(xRange))),
                new C.YValues(new C.NumberReference(new C.Formula(yRange)))));
        }

        if (withCustomYLabels)
            sc.Append(BuildYLabelSeries(pc.Mapped.Count, pc.Spec.YCustomLabels, pc.Spec.XMin ?? 0D));

        sc.Append(new C.AxisId { Val = XAxisId });
        sc.Append(new C.AxisId { Val = YAxisId });
        return sc;
    }

    private static C.ScatterChartSeries BuildYLabelSeries(int afterIndex, IReadOnlyList<YAxisLabel> labels, double xAnchor)
    {
        var dLbls = new C.DataLabels();
        for (int i = 0; i < labels.Count; i++)
            dLbls.Append(new C.DataLabel(
                new C.Index { Val = (uint)i },
                new C.ChartText(new C.RichText(
                    new A.BodyProperties { Wrap = A.TextWrappingValues.None },
                    new A.ListStyle(),
                    new A.Paragraph(new A.Run(
                        new A.RunProperties { Language = "en-US", FontSize = 1000 },
                        new A.Text(labels[i].Text ?? ""))))),
                new C.DataLabelPosition { Val = C.DataLabelPositionValues.Left },
                new C.ShowLegendKey { Val = false }, new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false }, new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false }, new C.ShowBubbleSize { Val = false }));

        dLbls.Append(new C.ShowLegendKey { Val = false });
        dLbls.Append(new C.ShowValue { Val = false });
        dLbls.Append(new C.ShowCategoryName { Val = false });
        dLbls.Append(new C.ShowSeriesName { Val = false });
        dLbls.Append(new C.ShowPercent { Val = false });
        dLbls.Append(new C.ShowBubbleSize { Val = false });

        var xLit = new C.NumberLiteral(new C.FormatCode("General"), new C.PointCount { Val = (uint)labels.Count });
        var yLit = new C.NumberLiteral(new C.FormatCode("General"), new C.PointCount { Val = (uint)labels.Count });
        for (int i = 0; i < labels.Count; i++)
        {
            xLit.Append(new C.NumericPoint(new C.NumericValue(xAnchor.ToString("R", CultureInfo.InvariantCulture))) { Index = (uint)i });
            yLit.Append(new C.NumericPoint(new C.NumericValue(labels[i].Position.ToString("R", CultureInfo.InvariantCulture))) { Index = (uint)i });
        }

        return new C.ScatterChartSeries(
            new C.Index { Val = (uint)afterIndex },
            new C.Order { Val = (uint)afterIndex },
            new C.SeriesText(new C.NumericValue("")),
            new C.ChartShapeProperties(new A.Outline(new A.NoFill())),
            new C.Marker(new C.Symbol { Val = C.MarkerStyleValues.None }),
            dLbls,
            new C.XValues(xLit),
            new C.YValues(yLit),
            new C.Smooth { Val = false });
    }

    private static C.Title BuildTitle(string text, int fontSize, A.TextAlignmentTypeValues? alignment = null)
    {
        var para = new A.Paragraph();
        if (alignment.HasValue) para.Append(new A.ParagraphProperties { Alignment = alignment.Value });
        para.Append(new A.Run(
            new A.RunProperties { Language = "en-US", FontSize = fontSize },
            new A.Text(text)));

        return new C.Title(
            new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), para)),
            new C.Layout(),
            new C.Overlay { Val = false });
    }

    private static C.ValueAxis BuildValueAxis(
        uint axisId, uint crossesAxisId, string title,
        C.AxisPositionValues position,
        double? min, double? max, double? majorUnit,
        C.CrossesValues crosses, string numberFormat)
    {
        var scaling = new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax });
        if (max.HasValue) scaling.Append(new C.MaxAxisValue { Val = max.Value });
        if (min.HasValue) scaling.Append(new C.MinAxisValue { Val = min.Value });

        var axis = new C.ValueAxis();
        axis.Append(new C.AxisId { Val = axisId });
        axis.Append(scaling);
        axis.Append(new C.Delete { Val = false });
        axis.Append(new C.AxisPosition { Val = position });
        axis.Append(new C.MajorGridlines(
            new C.ChartShapeProperties(
                new A.Outline(
                    new A.SolidFill(new A.RgbColorModelHex { Val = "D9D9D9" }),
                    new A.PresetDash { Val = A.PresetLineDashValues.Solid })
                { Width = 9525 })));
        if (!string.IsNullOrWhiteSpace(title)) axis.Append(BuildTitle(title, 1200));
        axis.Append(new C.NumberingFormat { FormatCode = numberFormat ?? "0", SourceLinked = false });
        axis.Append(new C.MajorTickMark { Val = C.TickMarkValues.Outside });
        axis.Append(new C.MinorTickMark { Val = C.TickMarkValues.None });
        axis.Append(new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo });
        axis.Append(new C.ChartShapeProperties(
            new A.Outline(new A.SolidFill(new A.RgbColorModelHex { Val = "BFBFBF" }))
            { Width = 9525 }));
        axis.Append(new C.CrossingAxis { Val = crossesAxisId });
        axis.Append(new C.Crosses { Val = crosses });
        axis.Append(new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });
        if (majorUnit.HasValue) axis.Append(new C.MajorUnit { Val = majorUnit.Value });
        return axis;
    }

    private static Spreadsheet.Stylesheet CreateStylesheet()
    {
        var fonts = new Spreadsheet.Fonts(
            new Spreadsheet.Font(),
            new Spreadsheet.Font(
                new Spreadsheet.Bold(),
                new Spreadsheet.FontSize { Val = 14D }));

        var fills = new Spreadsheet.Fills(
            new Spreadsheet.Fill(new Spreadsheet.PatternFill { PatternType = Spreadsheet.PatternValues.None }),
            new Spreadsheet.Fill(new Spreadsheet.PatternFill { PatternType = Spreadsheet.PatternValues.Gray125 }));

        var borders = new Spreadsheet.Borders(
            new Spreadsheet.Border());

        var cellStyleFormats = new Spreadsheet.CellStyleFormats(
            new Spreadsheet.CellFormat());

        var cellFormats = new Spreadsheet.CellFormats(
            new Spreadsheet.CellFormat(),
            new Spreadsheet.CellFormat
            {
                FontId = 1U,
                FillId = 0U,
                BorderId = 0U,
                ApplyFont = true,
                ApplyAlignment = true,
                Alignment = new Spreadsheet.Alignment
                {
                    Horizontal = Spreadsheet.HorizontalAlignmentValues.Center,
                    Vertical = Spreadsheet.VerticalAlignmentValues.Center,
                    WrapText = true
                }
            });

        return new Spreadsheet.Stylesheet(fonts, fills, borders, cellStyleFormats, cellFormats);
    }

    private static Spreadsheet.Column Column(uint min, uint max, double width) =>
        new() { Min = min, Max = max, Width = width, CustomWidth = true };

    private static Spreadsheet.Cell TextCell(string cellRef, string text) => new()
    {
        CellReference = cellRef,
        DataType = Spreadsheet.CellValues.InlineString,
        InlineString = new Spreadsheet.InlineString(new Spreadsheet.Text(text ?? ""))
    };

    private static Spreadsheet.Cell StyledTextCell(string cellRef, string text, uint styleIndex) => new()
    {
        CellReference = cellRef,
        DataType = Spreadsheet.CellValues.InlineString,
        StyleIndex = styleIndex,
        InlineString = new Spreadsheet.InlineString(new Spreadsheet.Text(text ?? ""))
    };

    private static Spreadsheet.Cell NumberCell(string cellRef, double value) => new()
    {
        CellReference = cellRef,
        DataType = Spreadsheet.CellValues.Number,
        CellValue = new Spreadsheet.CellValue(value.ToString(CultureInfo.InvariantCulture))
    };

    private static Spreadsheet.Cell FormulaCell(string cellRef, string formula) => new()
    {
        CellReference = cellRef,
        CellFormula = new Spreadsheet.CellFormula(formula)
    };

    private static string CellRef(int col, uint row) => ColName(col) + row.ToString(CultureInfo.InvariantCulture);

    private static string ColName(int col)
    {
        string n = "";
        while (col > 0)
        {
            int m = (col - 1) % 26;
            n = (char)('A' + m) + n;
            col = (col - m - 1) / 26;
        }
        return n;
    }

    private static string RangeRef(string sheet, int fromCol, uint fromRow, int toCol, uint toRow) =>
        $"'{sheet}'!${ColName(fromCol)}${fromRow}:${ColName(toCol)}${toRow}";

    private sealed class PreparedChart(ChartExportSpec spec, List<MappedSeries> mapped, List<double> orderedX)
    {
        public ChartExportSpec Spec { get; } = spec;
        public List<MappedSeries> Mapped { get; } = mapped;
        public List<double> OrderedX { get; } = orderedX;
    }

    private sealed class MappedSeries(string name, string hex, List<MappedPoint> points)
    {
        public string Name { get; } = name;
        public string HexColor { get; } = hex;
        public List<MappedPoint> Points { get; } = points;
    }

    private sealed class MappedPoint(double x, double? y)
    {
        public double X { get; } = x;
        public double? Y { get; } = y;
    }
}
