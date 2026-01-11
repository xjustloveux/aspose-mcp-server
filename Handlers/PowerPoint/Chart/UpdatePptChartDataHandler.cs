using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Core.Handlers;
using AsposeMcpServer.Core.Helpers;

namespace AsposeMcpServer.Handlers.PowerPoint.Chart;

/// <summary>
///     Handler for updating chart data in PowerPoint presentations.
/// </summary>
public class UpdatePptChartDataHandler : OperationHandlerBase<Presentation>
{
    /// <inheritdoc />
    public override string Operation => "update_data";

    /// <summary>
    ///     Updates chart data with new categories and series.
    /// </summary>
    /// <param name="context">The presentation context.</param>
    /// <param name="parameters">
    ///     Required: slideIndex, shapeIndex
    ///     Optional: data, clearExisting
    /// </param>
    /// <returns>Success message with update details.</returns>
    public override string Execute(OperationContext<Presentation> context, OperationParameters parameters)
    {
        var slideIndex = parameters.GetRequired<int>("slideIndex");
        var chartIndex = parameters.GetRequired<int>("shapeIndex");
        var dataObject = parameters.GetOptional<JsonObject?>("data");
        var clearExisting = parameters.GetOptional("clearExisting", false);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, chartIndex, slideIndex);

        var chartData = chart.ChartData;
        var workbook = chartData.ChartDataWorkbook;

        string[]? categories = null;
        List<(string name, double[] values)>? seriesList = null;

        if (dataObject != null)
        {
            var categoriesArray = ValueHelper.GetArray(dataObject, "categories");
            var seriesArray = ValueHelper.GetArray(dataObject, "series");

            if (categoriesArray is { Count: > 0 })
                categories = categoriesArray.Select(c => c?.GetValue<string>() ?? "").ToArray();

            if (seriesArray is { Count: > 0 })
            {
                seriesList = [];
                foreach (var seriesNode in seriesArray)
                    if (seriesNode is JsonObject seriesObj)
                    {
                        var name = ValueHelper.GetString(seriesObj, "name");
                        var valuesArray = ValueHelper.GetArray(seriesObj, "values");
                        if (valuesArray != null)
                        {
                            var values = valuesArray.Select(v =>
                            {
                                if (v == null) return 0.0;
                                if (v.GetValueKind() == JsonValueKind.Number)
                                    return v.GetValue<double>();
                                if (double.TryParse(v.GetValue<string>(), out var d))
                                    return d;
                                return 0.0;
                            }).ToArray();
                            seriesList.Add((name, values));
                        }
                    }
            }
        }

        if (categories == null && seriesList == null && !clearExisting)
            return Success($"No changes made to chart {chartIndex} on slide {slideIndex}.");

        if (clearExisting || categories != null || seriesList != null)
        {
            chartData.Series.Clear();
            chartData.Categories.Clear();
        }

        if (categories is { Length: > 0 })
            for (var i = 0; i < categories.Length; i++)
            {
                var cell = workbook.GetCell(0, i + 1, 0, categories[i]);
                chartData.Categories.Add(cell);
            }

        if (seriesList is { Count: > 0 })
        {
            var baseColumnOffset = categories != null ? 1 : 0;
            var yColumnStart = baseColumnOffset + 1;

            for (var seriesIdx = 0; seriesIdx < seriesList.Count; seriesIdx++)
            {
                var seriesName = seriesList[seriesIdx].name;
                var yColumnIndex = yColumnStart + seriesIdx;

                var headerCell = workbook.GetCell(0, 0, yColumnIndex, seriesName);
                var series = chartData.Series.Add(headerCell, chart.Type);
                var values = seriesList[seriesIdx].values;

                if (chart.Type == ChartType.Bubble || chart.Type == ChartType.BubbleWith3D)
                {
                    var sizeColumnIndex = yColumnStart + seriesList.Count + seriesIdx;

                    for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                    {
                        var xValue = (double)(dataIdx + 1);
                        var xCell = workbook.GetCell(0, dataIdx + 1, 0, xValue);
                        var yValue = values[dataIdx];
                        var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, yValue);
                        var sizeValue = Math.Abs(yValue) > 0 ? Math.Abs(yValue) * 0.5 : 10.0;
                        var sCell = workbook.GetCell(0, dataIdx + 1, sizeColumnIndex, sizeValue);
                        series.DataPoints.AddDataPointForBubbleSeries(xCell, yCell, sCell);
                    }
                }
                else if (chart.Type == ChartType.ScatterWithSmoothLines ||
                         chart.Type == ChartType.ScatterWithStraightLines ||
                         chart.Type == ChartType.ScatterWithStraightLinesAndMarkers ||
                         chart.Type == ChartType.ScatterWithSmoothLinesAndMarkers)
                {
                    for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                    {
                        var xValue = (double)(dataIdx + 1);
                        var xCell = workbook.GetCell(0, dataIdx + 1, 0, xValue);
                        var yValue = values[dataIdx];
                        var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, yValue);
                        series.DataPoints.AddDataPointForScatterSeries(xCell, yCell);
                    }
                }
                else if (chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut)
                {
                    for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                    {
                        var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
                        series.DataPoints.AddDataPointForPieSeries(dataCell);
                    }
                }
                else
                {
                    for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                    {
                        var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
                        series.DataPoints.AddDataPointForBarSeries(dataCell);
                    }
                }
            }
        }

        var maxRow = Math.Max(
            categories?.Length ?? 0,
            seriesList?.Max(s => s.values.Length) ?? 0
        );

        var maxCol = (seriesList?.Count ?? 0) + (categories != null ? 1 : 0);
        if (chart.Type == ChartType.Bubble || chart.Type == ChartType.BubbleWith3D)
            maxCol += seriesList?.Count ?? 0;

        if (maxRow > 0 && maxCol > 0)
        {
            var lastCol = (char)('A' + maxCol - 1);
            var range = $"Sheet1!$A$1:${lastCol}${maxRow + 1}";
            chartData.SetRange(range);
        }

        MarkModified(context);

        return Success($"Chart {chartIndex} data updated on slide {slideIndex}.");
    }
}
