using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Helpers.Excel;

/// <summary>
///     Helper class providing shared utility methods for Excel chart operations.
/// </summary>
public static class ExcelChartHelper
{
    /// <summary>
    ///     Parses chart type string to ChartType enum.
    /// </summary>
    /// <param name="chartTypeStr">The chart type string to parse.</param>
    /// <param name="defaultType">The default chart type if parsing fails.</param>
    /// <returns>The parsed ChartType enum value.</returns>
    public static ChartType ParseChartType(string? chartTypeStr, ChartType defaultType = ChartType.Column)
    {
        if (string.IsNullOrEmpty(chartTypeStr))
            return defaultType;

        return Enum.TryParse<ChartType>(chartTypeStr, true, out var result) ? result : defaultType;
    }

    /// <summary>
    ///     Parses legend position string to LegendPositionType enum.
    /// </summary>
    /// <param name="positionStr">The legend position string to parse.</param>
    /// <param name="defaultPosition">The default legend position if parsing fails.</param>
    /// <returns>The parsed LegendPositionType enum value.</returns>
    public static LegendPositionType ParseLegendPosition(string? positionStr,
        LegendPositionType defaultPosition = LegendPositionType.Bottom)
    {
        if (string.IsNullOrEmpty(positionStr))
            return defaultPosition;

        return positionStr.ToLower() switch
        {
            "bottom" => LegendPositionType.Bottom,
            "top" => LegendPositionType.Top,
            "left" => LegendPositionType.Left,
            "right" => LegendPositionType.Right,
            "topright" => LegendPositionType.Right,
            _ => defaultPosition
        };
    }

    /// <summary>
    ///     Sets category data for chart series.
    /// </summary>
    /// <param name="chart">The chart to set category data for.</param>
    /// <param name="categoryAxisDataRange">The range for category axis data.</param>
    public static void SetCategoryData(Chart chart, string categoryAxisDataRange)
    {
        if (string.IsNullOrEmpty(categoryAxisDataRange) || chart.NSeries.Count == 0)
            return;

        chart.NSeries.CategoryData = categoryAxisDataRange;
    }

    /// <summary>
    ///     Adds data series to chart.
    /// </summary>
    /// <param name="chart">The chart to add data series to.</param>
    /// <param name="dataRange">The data range for the series.</param>
    public static void AddDataSeries(Chart chart, string dataRange)
    {
        chart.NSeries.Clear();
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

        foreach (var range in ranges)
        {
            var seriesIndex = chart.NSeries.Add(range, true);
            chart.NSeries[seriesIndex].Values = range;
        }
    }

    /// <summary>
    ///     Gets a chart from the worksheet by index.
    /// </summary>
    /// <param name="worksheet">The worksheet containing the chart.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <returns>The chart at the specified index.</returns>
    /// <exception cref="ArgumentException">Thrown when chart index is out of range.</exception>
    public static Chart GetChart(Worksheet worksheet, int chartIndex)
    {
        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

        return worksheet.Charts[chartIndex];
    }
}
