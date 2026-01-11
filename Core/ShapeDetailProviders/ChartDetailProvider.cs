using Aspose.Slides;
using Aspose.Slides.Charts;

namespace AsposeMcpServer.Core.ShapeDetailProviders;

/// <summary>
///     Provider for extracting details from Chart elements
/// </summary>
public class ChartDetailProvider : IShapeDetailProvider
{
    /// <inheritdoc />
    public string TypeName => "Chart";

    /// <inheritdoc />
    public bool CanHandle(IShape shape)
    {
        return shape is IChart;
    }

    /// <inheritdoc />
    public object? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IChart chart)
            return null;

        var seriesInfo = chart.ChartData.Series.Select(s => new
        {
            name = s.Name?.ToString(),
            type = s.Type.ToString()
        }).ToArray();

        var categoryCount = chart.ChartData.Categories.Count;

        return new
        {
            chartType = chart.Type.ToString(),
            title = chart.ChartTitle?.TextFrameForOverriding?.Text,
            hasTitle = chart.HasTitle,
            hasLegend = chart.HasLegend,
            seriesCount = chart.ChartData.Series.Count,
            series = seriesInfo.Length > 0 ? seriesInfo : null,
            categoryCount,
            hasDataTable = chart.HasDataTable
        };
    }
}
