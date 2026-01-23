using Aspose.Slides;
using AsposeMcpServer.Core;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Helpers.PowerPoint;
using AsposeMcpServer.Results.PowerPoint.Chart;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for getting chart data from PowerPoint presentations.
/// </summary>
[ResultType(typeof(GetChartDataPptResult))]
public class GetPptChartDataHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "get_data";

    /// <summary>
    ///     Gets chart data including categories and series.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    /// </param>
    /// <returns>JSON string containing the chart data.</returns>
    public override object Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetChartDataParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, p.ChartIndex, p.SlideIndex);
        var chartData = chart.ChartData;

        List<GetChartCategoryItem> categoriesList = [];
        for (var i = 0; i < chartData.Categories.Count; i++)
        {
            var cat = chartData.Categories[i];
            categoriesList.Add(new GetChartCategoryItem
            {
                Index = i,
                Value = cat.Value?.ToString()
            });
        }

        List<GetChartSeriesItem> seriesList = [];
        for (var i = 0; i < chartData.Series.Count; i++)
        {
            var series = chartData.Series[i];
            List<GetChartDataPoint> dataPointsList = [];
            for (var j = 0; j < series.DataPoints.Count; j++)
            {
                var point = series.DataPoints[j];
                dataPointsList.Add(new GetChartDataPoint
                {
                    Index = j,
                    Value = point.Value?.ToString()
                });
            }

            seriesList.Add(new GetChartSeriesItem
            {
                Index = i,
                Name = series.Name?.ToString(),
                DataPointsCount = series.DataPoints.Count,
                DataPoints = dataPointsList
            });
        }

        var result = new GetChartDataPptResult
        {
            SlideIndex = p.SlideIndex,
            ChartIndex = p.ChartIndex,
            ChartType = chart.Type.ToString(),
            HasTitle = chart.HasTitle,
            Title = chart is { HasTitle: true, ChartTitle: not null }
                ? chart.ChartTitle.TextFrameForOverriding?.Text
                : null,
            Categories = new GetChartCategoriesInfo
            {
                Count = chartData.Categories.Count,
                Items = categoriesList
            },
            Series = new GetChartSeriesInfo
            {
                Count = chartData.Series.Count,
                Items = seriesList
            }
        };

        return result;
    }

    /// <summary>
    ///     Extracts get chart data parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted get chart data parameters.</returns>
    private static GetChartDataParameters ExtractGetChartDataParameters(OperationParameters parameters)
    {
        return new GetChartDataParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"));
    }

    /// <summary>
    ///     Record for holding get chart data parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ChartIndex">The chart shape index.</param>
    private sealed record GetChartDataParameters(int SlideIndex, int ChartIndex);
}
