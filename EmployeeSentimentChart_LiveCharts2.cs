// =============================================================================
//  EmployeeSentimentChart — WPF + LiveCharts2 MVVM Line Chart (single file)
//
//  SETUP:
//    1. Install NuGet: LiveChartsCore.SkiaSharpView.WPF
//    2. Split into files as indicated by each region header
//    3. Paste the XAML block (bottom of this file) into your View .xaml
// =============================================================================


// ─────────────────────────────────────────────────────────────────────────────
// REGION 1 — SENTIMENT LABELS (shared lookup)
// File: ViewModels/SentimentLabels.cs
// ─────────────────────────────────────────────────────────────────────────────
using System.Collections.Generic;

namespace EmployeeSentimentChart.ViewModels
{
    public static class SentimentLabels
    {
        private static readonly Dictionary<double, string> Map = new Dictionary<double, string>
        {
            { -2, "Negative 2" },
            { -1, "Negative 1" },
            {  0, "Neutral"    },
            {  1, "Positive 1" },
            {  2, "Positive 2" },
            {  3, "Positive 3" },
            {  4, "Positive 4" },
        };

        public static string GetLabel(double value) =>
            Map.TryGetValue(value, out var label) ? label : value.ToString();
    }
}


// ─────────────────────────────────────────────────────────────────────────────
// REGION 2 — VIEW MODEL
// File: ViewModels/EmployeeChartViewModel.cs
// ─────────────────────────────────────────────────────────────────────────────
using System;
using System.Collections.Generic;
using LiveChartsCore;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.Painting;
using SkiaSharp;

namespace EmployeeSentimentChart.ViewModels
{
    public class EmployeeChartViewModel
    {
        // ── Series — one LineSeries<ObservablePoint> per employee ─────────────
        // Use a plain List (not ObservableCollection) for large datasets —
        // avoids per-item change notification overhead at scale.
        public List<ISeries> Series { get; set; }

        // ── X-Axis — month name labels ────────────────────────────────────────
        public List<Axis> XAxes { get; set; }

        // ── Y-Axis — converts raw numbers to sentiment words ──────────────────
        public List<Axis> YAxes { get; set; }

        public EmployeeChartViewModel()
        {
            // ── Data ──────────────────────────────────────────────────────────
            // ObservablePoint(xIndex, sentimentValue)
            // X is a zero-based index that maps to the XAxes.Labels array below.
            // Add as many employees as needed — no XAML changes required.
            Series = new List<ISeries>
            {
                BuildSeries("Emp1", SKColors.DodgerBlue, new[]
                {
                    (0, 1.0),   // Jan  → Positive 1
                    (1, 4.0),   // Feb  → Positive 4
                    (2, 2.0),   // Mar  → Positive 2
                }),

                BuildSeries("Emp2", SKColors.OrangeRed, new[]
                {
                    (0, 4.0),   // Jan  → Positive 4
                    (1, 0.0),   // Feb  → Neutral
                    (2, 3.0),   // Mar  → Positive 3
                }),
            };

            // ── X-Axis ────────────────────────────────────────────────────────
            XAxes = new List<Axis>
            {
                new Axis
                {
                    Labels = new[] { "Jan", "Feb", "March" },
                    LabelsRotation = 0,
                    TextSize = 13,
                    LabelsPaint = new SolidColorPaint(SKColors.White),
                    SeparatorsPaint = new SolidColorPaint(SKColors.Gray)
                    {
                        StrokeThickness = 1
                    }
                }
            };

            // ── Y-Axis — Labeler lambda replaces IValueConverter entirely ─────
            YAxes = new List<Axis>
            {
                new Axis
                {
                    MinLimit  = -2,
                    MaxLimit  =  4,
                    MinStep   =  1,   // one tick per integer step
                    TextSize  = 13,
                    LabelsPaint = new SolidColorPaint(SKColors.White),
                    SeparatorsPaint = new SolidColorPaint(new SKColor(80, 80, 80))
                    {
                        StrokeThickness = 1,
                        PathEffect = new LiveChartsCore.SkiaSharpView.Painting.Effects.DashEffect(new float[] { 4, 4 })
                    },
                    // ↓ This single lambda replaces the entire IValueConverter from Telerik
                    Labeler = value => SentimentLabels.GetLabel(value)
                }
            };
        }

        // ── Helper — builds a styled LineSeries for one employee ──────────────
        private static LineSeries<LiveChartsCore.Defaults.ObservablePoint> BuildSeries(
            string name, SKColor color, (int x, double y)[] points)
        {
            var data = new LiveChartsCore.Defaults.ObservablePoint[points.Length];
            for (int i = 0; i < points.Length; i++)
                data[i] = new LiveChartsCore.Defaults.ObservablePoint(points[i].x, points[i].y);

            return new LineSeries<LiveChartsCore.Defaults.ObservablePoint>
            {
                Name   = name,
                Values = data,

                // ── Performance settings (safe for 10k+ points) ───────────────
                GeometrySize     = 8,                  // set to 0 to hide markers at 10k scale
                AnimationsSpeed  = TimeSpan.FromMilliseconds(400),  // set to TimeSpan.Zero at 10k scale
                LineSmoothness   = 0,                  // straight lines = faster render

                // ── Styling ───────────────────────────────────────────────────
                Stroke = new SolidColorPaint(color) { StrokeThickness = 2.5f },
                Fill   = null,                         // no area fill under the line
                GeometryStroke = new SolidColorPaint(color) { StrokeThickness = 2 },
                GeometryFill   = new SolidColorPaint(SKColors.White),

                // ── Tooltip — shows human label, not raw number ───────────────
                TooltipLabelFormatter = point =>
                    $"{name}: {SentimentLabels.GetLabel(point.PrimaryValue)}"
            };
        }
    }
}


/*
═══════════════════════════════════════════════════════════════════════════════
 XAML — Paste into EmployeeChartView.xaml
═══════════════════════════════════════════════════════════════════════════════

<UserControl x:Class="EmployeeSentimentChart.Views.EmployeeChartView"
             xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
             xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
             xmlns:lvc="clr-namespace:LiveChartsCore.SkiaSharpView.WPF;assembly=LiveChartsCore.SkiaSharpView.WPF"
             xmlns:vm="clr-namespace:EmployeeSentimentChart.ViewModels">

    <UserControl.DataContext>
        <vm:EmployeeChartViewModel />
    </UserControl.DataContext>

    <Grid Background="#1E1E2E" Margin="20">

        <lvc:CartesianChart Series="{Binding Series}"
                            XAxes="{Binding XAxes}"
                            YAxes="{Binding YAxes}"
                            TooltipPosition="Top"
                            LegendPosition="Bottom"
                            Background="Transparent" />

    </Grid>

</UserControl>

═══════════════════════════════════════════════════════════════════════════════
 ADDING A NEW EMPLOYEE
═══════════════════════════════════════════════════════════════════════════════

 In EmployeeChartViewModel constructor, call BuildSeries and add to Series:

     Series.Add(BuildSeries("Emp3", SKColors.LimeGreen, new[]
     {
         (0, -1.0),   // Jan  → Negative 1
         (1,  2.0),   // Feb  → Positive 2
         (2,  0.0),   // Mar  → Neutral
     }));

 No XAML changes needed.

═══════════════════════════════════════════════════════════════════════════════
 SCALING TO 10k POINTS — flip these two lines in BuildSeries()
═══════════════════════════════════════════════════════════════════════════════

     GeometrySize    = 0,                        // was 8  — hides per-point markers
     AnimationsSpeed = TimeSpan.Zero,             // was 400ms — skips entry animation

═══════════════════════════════════════════════════════════════════════════════
 ADDING A NEW SENTIMENT LEVEL
═══════════════════════════════════════════════════════════════════════════════

 In SentimentLabels.Map, add one entry:

     { 5, "Excellent" }

 Y-axis label and tooltip both pick it up automatically.

*/
