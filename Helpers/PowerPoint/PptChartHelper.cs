using Aspose.Slides;
using Aspose.Slides.Charts;

namespace AsposeMcpServer.Helpers.PowerPoint;

/// <summary>
///     Helper class for PowerPoint chart operations.
/// </summary>
public static class PptChartHelper
{
    /// <summary>
    ///     Gets a chart by index from a slide.
    /// </summary>
    /// <param name="slide">The slide containing the chart.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="slideIndex">The slide index for error messages.</param>
    /// <returns>The chart at the specified index.</returns>
    /// <exception cref="ArgumentException">Thrown when no charts exist or index is out of range.</exception>
    public static IChart GetChartByIndex(ISlide slide, int chartIndex, int slideIndex)
    {
        var charts = slide.Shapes.OfType<IChart>().ToList();

        if (charts.Count == 0)
            throw new ArgumentException($"Slide {slideIndex} contains no charts.");

        if (chartIndex < 0 || chartIndex >= charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} out of range for slide {slideIndex}. " +
                $"Valid range: 0 to {charts.Count - 1} (Total charts: {charts.Count})");

        return charts[chartIndex];
    }

    /// <summary>
    ///     Parses a chart type string to ChartType enum.
    /// </summary>
    /// <param name="chartTypeStr">The chart type string to parse.</param>
    /// <param name="defaultType">The default chart type if parsing fails.</param>
    /// <returns>The parsed ChartType enum value.</returns>
    public static ChartType ParseChartType(string? chartTypeStr, ChartType defaultType = ChartType.ClusteredColumn)
    {
        if (string.IsNullOrEmpty(chartTypeStr))
            return defaultType;

        return chartTypeStr.ToLower() switch
        {
            "column" => ChartType.ClusteredColumn,
            "bar" => ChartType.ClusteredBar,
            "line" => ChartType.Line,
            "pie" => ChartType.Pie,
            "area" => ChartType.Area,
            "scatter" => ChartType.ScatterWithSmoothLines,
            "doughnut" => ChartType.Doughnut,
            "bubble" => ChartType.Bubble,
            "radar" => ChartType.Radar,
            "treemap" => ChartType.Treemap,
            _ => defaultType
        };
    }

    /// <summary>
    ///     Sets the chart title.
    /// </summary>
    /// <param name="chart">The chart to set the title on.</param>
    /// <param name="title">The title text.</param>
    public static void SetChartTitle(IChart chart, string title)
    {
        chart.HasTitle = true;
        var chartTitle = chart.ChartTitle;
        if (chartTitle != null)
        {
            if (chartTitle.TextFrameForOverriding != null)
                chartTitle.TextFrameForOverriding.Text = title;
            else
                chartTitle.AddTextFrameForOverriding(title);
        }
    }
}
