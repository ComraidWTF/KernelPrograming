// =============================================================================
// ChartExcelExporter — DocumentFormat.OpenXml only
// =============================================================================
//
// Creates an .xlsx with:
//   * Per-series columns: X, Y, Label
//   * A native Excel scatter chart (lines + markers) plotting the points
//   * Per-point data labels read from the Label column cells
//   * Series names as they appear in LiveCharts2 used for the legend
//
// NuGet: dotnet add package DocumentFormat.OpenXml
// Works on .NET 6+, .NET Standard 2.0.
//
// Usage:
//   ChartExcelExporter.ExportLineChart<CustomData>(
//       myCartesianChart,
//       @"C:\temp\chart.xlsx",
//       p => p.X,
//       p => p.Y,
//       p => p.TaskName);
// =============================================================================

using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView.WPF;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

public static class ChartExcelExporter
{
    public static void ExportLineChart<TPoint>(
        CartesianChart chart,
        string filePath,
        Func<TPoint, double> xSelector,
        Func<TPoint, double> ySelector,
        Func<TPoint, string> labelSelector = null)
    {
        var seriesList = chart.Series
            .OfType<ISeries>()
            .Where(s => s.GetType().Name.StartsWith("LineSeries"))
            .Select((s, i) => new SeriesData
            {
                Name = string.IsNullOrWhiteSpace(s.Name) ? $"Series {i + 1}" : s.Name,
                Xs = new List<double>(),
                Ys = new List<double>(),
                Labels = new List<string>()
            })
            .ToList();

        int idx = 0;
        foreach (var lc2 in chart.Series.OfType<ISeries>().Where(s => s.GetType().Name.StartsWith("LineSeries")))
        {
            var pts = (lc2.Values as IEnumerable)?.OfType<TPoint>().ToList() ?? new List<TPoint>();
            foreach (var p in pts)
            {
                seriesList[idx].Xs.Add(xSelector(p));
                seriesList[idx].Ys.Add(ySelector(p));
                seriesList[idx].Labels.Add(labelSelector?.Invoke(p) ?? string.Empty);
            }
            idx++;
        }

        seriesList = seriesList.Where(s => s.Xs.Count > 0).ToList();
        if (seriesList.Count == 0) return;

        const string sheetName = "Chart Data";
        using var doc = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);

        var wbPart = doc.AddWorkbookPart();
        wbPart.Workbook = new Workbook();

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        wsPart.Worksheet = BuildWorksheet(seriesList, labelSelector != null);

        var sheets = wbPart.Workbook.AppendChild(new Sheets());
        sheets.Append(new Sheet
        {
            Id = wbPart.GetIdOfPart(wsPart),
            SheetId = 1U,
            Name = sheetName
        });
        wbPart.Workbook.Save();

        // Drawing + chart
        var drawingsPart = wsPart.AddNewPart<DrawingsPart>();
        wsPart.Worksheet.Append(new Drawing { Id = wsPart.GetIdOfPart(drawingsPart) });
        wsPart.Worksheet.Save();

        var chartPart = drawingsPart.AddNewPart<ChartPart>();
        chartPart.ChartSpace = BuildChartSpace(seriesList, sheetName, labelSelector != null);
        chartPart.ChartSpace.Save();

        BuildDrawing(drawingsPart, chartPart, anchorRow: seriesList.Max(s => s.Xs.Count) + 3);
        drawingsPart.WorksheetDrawing.Save();
    }

    // ---- Worksheet ----------------------------------------------------------

    private static Worksheet BuildWorksheet(List<SeriesData> list, bool hasLabels)
    {
        int colsPerSeries = hasLabels ? 3 : 2;
        int maxRows = list.Max(s => s.Xs.Count);
        var sheetData = new SheetData();

        // Header row
        var header = new Row { RowIndex = 1U };
        for (int i = 0; i < list.Count; i++)
        {
            int xCol = colsPerSeries * i;
            header.Append(TextCell(xCol, 1, $"{list[i].Name} X"));
            header.Append(TextCell(xCol + 1, 1, $"{list[i].Name} Y"));
            if (hasLabels) header.Append(TextCell(xCol + 2, 1, $"{list[i].Name} Label"));
        }
        sheetData.Append(header);

        // Data rows
        for (int r = 0; r < maxRows; r++)
        {
            var row = new Row { RowIndex = (uint)(r + 2) };
            for (int i = 0; i < list.Count; i++)
            {
                int xCol = colsPerSeries * i;
                if (r < list[i].Xs.Count)
                {
                    row.Append(NumberCell(xCol, r + 2, list[i].Xs[r]));
                    row.Append(NumberCell(xCol + 1, r + 2, list[i].Ys[r]));
                    if (hasLabels)
                        row.Append(TextCell(xCol + 2, r + 2, list[i].Labels[r]));
                }
            }
            sheetData.Append(row);
        }

        return new Worksheet(sheetData);
    }

    private static Cell TextCell(int col, int row, string value) => new Cell
    {
        CellReference = ColLetter(col) + row,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(value ?? string.Empty))
    };

    private static Cell NumberCell(int col, int row, double value) => new Cell
    {
        CellReference = ColLetter(col) + row,
        DataType = CellValues.Number,
        CellValue = new CellValue(value.ToString(System.Globalization.CultureInfo.InvariantCulture))
    };

    private static string ColLetter(int zeroBased)
    {
        int c = zeroBased + 1;
        string s = "";
        while (c > 0)
        {
            int m = (c - 1) % 26;
            s = (char)('A' + m) + s;
            c = (c - m - 1) / 26;
        }
        return s;
    }

    // ---- Chart --------------------------------------------------------------

    private static C.ChartSpace BuildChartSpace(List<SeriesData> list, string sheetName, bool hasLabels)
    {
        const uint axisIdX = 111111U;
        const uint axisIdY = 222222U;

        var scatter = new C.ScatterChart(
            new C.ScatterStyle { Val = C.ScatterStyleValues.LineMarker },
            new C.VaryColors { Val = false });

        int colsPerSeries = hasLabels ? 3 : 2;
        for (int i = 0; i < list.Count; i++)
        {
            int xCol = colsPerSeries * i;
            int yCol = xCol + 1;
            int labelCol = xCol + 2;
            int n = list[i].Xs.Count;
            int lastDataRow = n + 1; // header is row 1, data rows 2..n+1

            var ser = new C.ScatterChartSeries(
                new C.Index { Val = (uint)i },
                new C.Order { Val = (uint)i },
                new C.SeriesText(new C.NumericValue(list[i].Name)));

            if (hasLabels)
                ser.Append(BuildDataLabels(sheetName, labelCol, n));

            ser.Append(BuildNumberRef<C.XValues>(sheetName, xCol, lastDataRow, list[i].Xs));
            ser.Append(BuildNumberRef<C.YValues>(sheetName, yCol, lastDataRow, list[i].Ys));
            ser.Append(new C.Smooth { Val = false });

            scatter.Append(ser);
        }

        scatter.Append(new C.AxisId { Val = axisIdX });
        scatter.Append(new C.AxisId { Val = axisIdY });

        var plotArea = new C.PlotArea(new C.Layout(), scatter);
        plotArea.Append(BuildValueAxis(axisIdX, axisIdY, C.AxisPositionValues.Bottom));
        plotArea.Append(BuildValueAxis(axisIdY, axisIdX, C.AxisPositionValues.Left));

        var chart = new C.Chart(
            new C.AutoTitleDeleted { Val = true },
            plotArea,
            new C.Legend(
                new C.LegendPosition { Val = C.LegendPositionValues.Right },
                new C.Overlay { Val = false }),
            new C.PlotVisibleOnly { Val = true },
            new C.DisplayBlanksAs { Val = C.DisplayBlanksAsValues.Gap });

        var cs = new C.ChartSpace(new C.EditingLanguage { Val = "en-US" }, chart);
        cs.AddNamespaceDeclaration("c", "http://schemas.openxmlformats.org/drawingml/2006/chart");
        cs.AddNamespaceDeclaration("a", "http://schemas.openxmlformats.org/drawingml/2006/main");
        cs.AddNamespaceDeclaration("r", "http://schemas.openxmlformats.org/officeDocument/2006/relationships");
        return cs;
    }

    private static T BuildNumberRef<T>(string sheet, int col, int lastRow, List<double> values)
        where T : OpenXmlCompositeElement, new()
    {
        var cache = new C.NumberingCache(
            new C.FormatCode("General"),
            new C.PointCount { Val = (uint)values.Count });
        for (int i = 0; i < values.Count; i++)
        {
            cache.Append(new C.NumericPoint(
                new C.NumericValue(values[i].ToString(System.Globalization.CultureInfo.InvariantCulture)))
            { Index = (uint)i });
        }

        var numRef = new C.NumberReference(
            new C.Formula($"'{sheet}'!${ColLetter(col)}$2:${ColLetter(col)}${lastRow}"),
            cache);

        var container = new T();
        container.Append(numRef);
        return container;
    }

    private static C.DataLabels BuildDataLabels(string sheet, int labelCol, int count)
    {
        var dLbls = new C.DataLabels();

        // Per-point label, each pointing at its own cell for the text
        for (int i = 0; i < count; i++)
        {
            int row = i + 2;
            var cache = new C.StringCache(new C.PointCount { Val = 1U });
            cache.Append(new C.StringPoint(new C.NumericValue("")) { Index = 0U });

            dLbls.Append(new C.DataLabel(
                new C.Index { Val = (uint)i },
                new C.ChartText(new C.StringReference(
                    new C.Formula($"'{sheet}'!${ColLetter(labelCol)}${row}"),
                    cache)),
                new C.ShowLegendKey { Val = false },
                new C.ShowValue { Val = false },
                new C.ShowCategoryName { Val = false },
                new C.ShowSeriesName { Val = false },
                new C.ShowPercent { Val = false },
                new C.ShowBubbleSize { Val = false }));
        }

        // Series-level defaults
        dLbls.Append(new C.ShowLegendKey { Val = false });
        dLbls.Append(new C.ShowValue { Val = false });
        dLbls.Append(new C.ShowCategoryName { Val = false });
        dLbls.Append(new C.ShowSeriesName { Val = false });
        dLbls.Append(new C.ShowPercent { Val = false });
        dLbls.Append(new C.ShowBubbleSize { Val = false });
        return dLbls;
    }

    private static C.ValueAxis BuildValueAxis(uint id, uint crossId, C.AxisPositionValues pos) =>
        new C.ValueAxis(
            new C.AxisId { Val = id },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = pos },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = "General", SourceLinked = true },
            new C.MajorTickMark { Val = C.TickMarkValues.Out },
            new C.MinorTickMark { Val = C.TickMarkValues.None },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = crossId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.MidpointCategory });

    // ---- Drawing / anchor ---------------------------------------------------

    private static void BuildDrawing(DrawingsPart drawingsPart, ChartPart chartPart, int anchorRow)
    {
        drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
        drawingsPart.WorksheetDrawing.AddNamespaceDeclaration(
            "xdr", "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing");
        drawingsPart.WorksheetDrawing.AddNamespaceDeclaration(
            "a", "http://schemas.openxmlformats.org/drawingml/2006/main");

        var anchor = new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(
                new Xdr.ColumnId("0"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId(anchorRow.ToString()), new Xdr.RowOffset("0")),
            new Xdr.ToMarker(
                new Xdr.ColumnId("10"), new Xdr.ColumnOffset("0"),
                new Xdr.RowId((anchorRow + 20).ToString()), new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Chart 1" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(
                    new A.Offset { X = 0L, Y = 0L },
                    new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(new A.GraphicData(
                    new C.ChartReference { Id = drawingsPart.GetIdOfPart(chartPart) })
                { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })),
            new Xdr.ClientData());

        drawingsPart.WorksheetDrawing.Append(anchor);
    }

    private sealed class SeriesData
    {
        public string Name;
        public List<double> Xs;
        public List<double> Ys;
        public List<string> Labels;
    }
}
