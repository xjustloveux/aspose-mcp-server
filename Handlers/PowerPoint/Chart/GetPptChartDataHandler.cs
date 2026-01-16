using Aspose.Slides;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for getting chart data from PowerPoint presentations.
/// </summary>
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
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var p = ExtractGetChartDataParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, p.ChartIndex, p.SlideIndex);
        var chartData = chart.ChartData;

        List<object> categoriesList = [];
        for (var i = 0; i < chartData.Categories.Count; i++)
        {
            var cat = chartData.Categories[i];
            categoriesList.Add(new
            {
                index = i,
                value = cat.Value?.ToString()
            });
        }

        List<object> seriesList = [];
        for (var i = 0; i < chartData.Series.Count; i++)
        {
            var series = chartData.Series[i];
            List<object> dataPointsList = [];
            for (var j = 0; j < series.DataPoints.Count; j++)
            {
                var point = series.DataPoints[j];
                dataPointsList.Add(new
                {
                    index = j,
                    value = point.Value?.ToString()
                });
            }

            seriesList.Add(new
            {
                index = i,
                name = series.Name?.ToString(),
                dataPointsCount = series.DataPoints.Count,
                dataPoints = dataPointsList
            });
        }

        var result = new
        {
            slideIndex = p.SlideIndex,
            chartIndex = p.ChartIndex,
            chartType = chart.Type.ToString(),
            hasTitle = chart.HasTitle,
            title = chart is { HasTitle: true, ChartTitle: not null }
                ? chart.ChartTitle.TextFrameForOverriding?.Text
                : null,
            categories = new
            {
                count = chartData.Categories.Count,
                items = categoriesList
            },
            series = new
            {
                count = chartData.Series.Count,
                items = seriesList
            }
        };

        return JsonResult(result);
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
