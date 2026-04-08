using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using LiveChartsCore;
using LiveChartsCore.Kernel.Sketches;
using LiveChartsCore.Measure;
using LiveChartsCore.SkiaSharpView;
using LiveChartsCore.SkiaSharpView.WPF;

namespace YourApp.Charts;

/// <summary>
/// A drop-in CartesianChart that tries to keep axes visible inside the control bounds,
/// especially when multiple charts are placed in a tight dashboard layout (for example a 2x2 Grid).
///
/// What it does:
/// - enables WPF clipping on the chart control
/// - applies a safe default DrawMargin
/// - reapplies a responsive DrawMargin on load/resize
/// - reduces extra series padding so the plot stays inside the draw area
/// - keeps clipping enabled on supported series
/// - optionally reduces axis text size when the chart becomes very small
///
/// Usage in XAML:
/// <local:AxisSafeCartesianChart
///     Series="{Binding Series}"
///     XAxes="{Binding XAxes}"
///     YAxes="{Binding YAxes}"/>
///
/// If you do not want automatic font reduction, set AutoReduceAxisTextSize="False".
/// </summary>
public class AxisSafeCartesianChart : CartesianChart
{
    public static readonly DependencyProperty BaseLeftMarginProperty = DependencyProperty.Register(
        nameof(BaseLeftMargin), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(48d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty BaseTopMarginProperty = DependencyProperty.Register(
        nameof(BaseTopMargin), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(12d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty BaseRightMarginProperty = DependencyProperty.Register(
        nameof(BaseRightMargin), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(18d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty BaseBottomMarginProperty = DependencyProperty.Register(
        nameof(BaseBottomMargin), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(34d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty AutoReduceAxisTextSizeProperty = DependencyProperty.Register(
        nameof(AutoReduceAxisTextSize), typeof(bool), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(true, OnLayoutPropertyChanged));

    public static readonly DependencyProperty CompactAxisTextSizeProperty = DependencyProperty.Register(
        nameof(CompactAxisTextSize), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(10d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty NormalAxisTextSizeProperty = DependencyProperty.Register(
        nameof(NormalAxisTextSize), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(12d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty CompactThresholdWidthProperty = DependencyProperty.Register(
        nameof(CompactThresholdWidth), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(320d, OnLayoutPropertyChanged));

    public static readonly DependencyProperty CompactThresholdHeightProperty = DependencyProperty.Register(
        nameof(CompactThresholdHeight), typeof(double), typeof(AxisSafeCartesianChart),
        new PropertyMetadata(220d, OnLayoutPropertyChanged));

    public double BaseLeftMargin
    {
        get => (double)GetValue(BaseLeftMarginProperty);
        set => SetValue(BaseLeftMarginProperty, value);
    }

    public double BaseTopMargin
    {
        get => (double)GetValue(BaseTopMarginProperty);
        set => SetValue(BaseTopMarginProperty, value);
    }

    public double BaseRightMargin
    {
        get => (double)GetValue(BaseRightMarginProperty);
        set => SetValue(BaseRightMarginProperty, value);
    }

    public double BaseBottomMargin
    {
        get => (double)GetValue(BaseBottomMarginProperty);
        set => SetValue(BaseBottomMarginProperty, value);
    }

    public bool AutoReduceAxisTextSize
    {
        get => (bool)GetValue(AutoReduceAxisTextSizeProperty);
        set => SetValue(AutoReduceAxisTextSizeProperty, value);
    }

    public double CompactAxisTextSize
    {
        get => (double)GetValue(CompactAxisTextSizeProperty);
        set => SetValue(CompactAxisTextSizeProperty, value);
    }

    public double NormalAxisTextSize
    {
        get => (double)GetValue(NormalAxisTextSizeProperty);
        set => SetValue(NormalAxisTextSizeProperty, value);
    }

    public double CompactThresholdWidth
    {
        get => (double)GetValue(CompactThresholdWidthProperty);
        set => SetValue(CompactThresholdWidthProperty, value);
    }

    public double CompactThresholdHeight
    {
        get => (double)GetValue(CompactThresholdHeightProperty);
        set => SetValue(CompactThresholdHeightProperty, value);
    }

    public AxisSafeCartesianChart()
    {
        ClipToBounds = true;
        Loaded += OnLoaded;
        SizeChanged += OnSizeChanged;
        DataContextChanged += (_, _) => ApplySafeLayout();
    }

    private static void OnLayoutPropertyChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
    {
        if (d is AxisSafeCartesianChart chart) chart.ApplySafeLayout();
    }

    private void OnLoaded(object sender, RoutedEventArgs e)
    {
        ApplySafeLayout();
    }

    private void OnSizeChanged(object sender, SizeChangedEventArgs e)
    {
        ApplySafeLayout();
    }

    private void ApplySafeLayout()
    {
        // 1) Keep the control itself clipped by WPF.
        ClipToBounds = true;

        // 2) Give the axes enough room using DrawMargin.
        // LiveCharts2 documents DrawMargin as the distance from the axes/edge to the draw margin area,
        // so this is the main lever to stop the axis from getting visually cut off. citeturn979466search1turn979466search8
        DrawMargin = BuildResponsiveMargin();

        // 3) Reduce series overflow tendencies.
        // LiveCharts2 series expose DataPadding and ClippingMode; clipping defaults to XY,
        // and DataPadding controls how far the series is drawn from the chart edge. citeturn979466search6
        ApplySeriesDefaults(Series);

        // 4) When the chart is very small, reduce axis text size so labels remain inside view.
        ApplyAxisDefaults(XAxes);
        ApplyAxisDefaults(YAxes);
    }

    private Margin BuildResponsiveMargin()
    {
        var width = Math.Max(0, ActualWidth);
        var height = Math.Max(0, ActualHeight);

        var left = BaseLeftMargin;
        var top = BaseTopMargin;
        var right = BaseRightMargin;
        var bottom = BaseBottomMargin;

        // Small cells in a 2x2 dashboard usually need more bottom room for the X axis labels,
        // and a minimum left margin for Y axis values.
        if (width < 360)
        {
            left = Math.Max(left, 42);
            right = Math.Max(right, 12);
        }

        if (height < 240)
        {
            bottom = Math.Max(bottom, 38);
            top = Math.Max(top, 8);
        }

        return new Margin(left, top, right, bottom);
    }

    private void ApplySeriesDefaults(IEnumerable<ISeries>? seriesCollection)
    {
        if (seriesCollection is null) return;

        foreach (var series in seriesCollection)
        {
            switch (series)
            {
                case LineSeries<double> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;

                case LineSeries<ObservablePoint> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;

                case StepLineSeries<double> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;

                case ColumnSeries<double> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;

                case RowSeries<double> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;

                case ScatterSeries<ObservablePoint> s:
                    s.DataPadding = new LvcPoint(0, 0);
                    s.ClippingMode = ClipMode.XY;
                    break;
            }
        }
    }

    private void ApplyAxisDefaults(IEnumerable<ICartesianAxis>? axes)
    {
        if (axes is null) return;

        var compact = AutoReduceAxisTextSize &&
                      (ActualWidth <= CompactThresholdWidth || ActualHeight <= CompactThresholdHeight);

        var textSize = compact ? CompactAxisTextSize : NormalAxisTextSize;

        foreach (var axis in axes.OfType<Axis>())
        {
            axis.TextSize = textSize;

            // Make separators/ticks align in a predictable way for dashboards.
            axis.UnitWidth = axis.UnitWidth <= 0 ? 1 : axis.UnitWidth;

            // In tight layouts, rotated labels make clipping worse unless really needed.
            if (compact && Math.Abs(axis.LabelsRotation) > 30)
            {
                axis.LabelsRotation = 0;
            }
        }
    }

    /// <summary>
    /// Optional helper you can call from your view model when building axes.
    /// </summary>
    public static ObservableCollection<Axis> CreateDefaultXAxes(params string[] labels)
    {
        return new ObservableCollection<Axis>
        {
            new Axis
            {
                Labels = labels,
                LabelsRotation = 0,
                MinStep = 1,
                TextSize = 12
            }
        };
    }

    /// <summary>
    /// Optional helper you can call from your view model when building numeric Y axes.
    /// </summary>
    public static ObservableCollection<Axis> CreateDefaultYAxes(string? name = null)
    {
        return new ObservableCollection<Axis>
        {
            new Axis
            {
                Name = name,
                MinStep = 1,
                TextSize = 12
            }
        };
    }
}
