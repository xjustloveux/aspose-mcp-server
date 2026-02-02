using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Core.ShapeDetailProviders.Details;

namespace AsposeMcpServer.Core.ShapeDetailProviders.Providers;

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
    public ShapeDetails? GetDetails(IShape shape, IPresentation presentation)
    {
        if (shape is not IChart chart)
            return null;

        var seriesInfo = chart.ChartData.Series.Select(s => new ChartSeriesInfo
        {
            Name = s.Name?.ToString(),
            Type = s.Type.ToString()
        }).ToList();

        var categoryCount = chart.ChartData.Categories.Count;

        return new ChartDetails
        {
            ChartType = chart.Type.ToString(),
            Title = chart.ChartTitle?.TextFrameForOverriding?.Text,
            HasTitle = chart.HasTitle,
            HasLegend = chart.HasLegend,
            LegendPosition = chart.HasLegend ? chart.Legend?.Position.ToString() : null,
            SeriesCount = chart.ChartData.Series.Count,
            Series = seriesInfo.Count > 0 ? seriesInfo : null,
            CategoryCount = categoryCount,
            HasDataTable = chart.HasDataTable
        };
    }
}
