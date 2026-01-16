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

        var (categories, seriesList) = ParseChartData(dataObject);

        if (categories == null && seriesList == null && !clearExisting)
            return Success($"No changes made to chart {chartIndex} on slide {slideIndex}.");

        var chartData = chart.ChartData;
        var workbook = chartData.ChartDataWorkbook;

        if (clearExisting || categories != null || seriesList != null)
        {
            chartData.Series.Clear();
            chartData.Categories.Clear();
        }

        AddCategories(chartData, workbook, categories);
        AddSeriesData(chart, chartData, workbook, categories, seriesList);
        SetChartRange(chart, chartData, categories, seriesList);

        MarkModified(context);

        return Success($"Chart {chartIndex} data updated on slide {slideIndex}.");
    }

    private static (string[]? categories, List<(string name, double[] values)>? seriesList) ParseChartData(
        JsonObject? dataObject)
    {
        if (dataObject == null)
            return (null, null);

        string[]? categories = null;
        List<(string name, double[] values)>? seriesList = null;

        var categoriesArray = ValueHelper.GetArray(dataObject, "categories");
        var seriesArray = ValueHelper.GetArray(dataObject, "series");

        if (categoriesArray is { Count: > 0 })
            categories = categoriesArray.Select(c => c?.GetValue<string>() ?? "").ToArray();

        if (seriesArray is { Count: > 0 })
            seriesList = ParseSeriesArray(seriesArray);

        return (categories, seriesList);
    }

    private static List<(string name, double[] values)> ParseSeriesArray(JsonArray seriesArray)
    {
        List<(string name, double[] values)> seriesList = [];

        foreach (var seriesNode in seriesArray)
        {
            if (seriesNode is not JsonObject seriesObj) continue;

            var name = ValueHelper.GetString(seriesObj, "name");
            var valuesArray = ValueHelper.GetArray(seriesObj, "values");
            if (valuesArray == null) continue;

            var values = valuesArray.Select(ParseDoubleValue).ToArray();
            seriesList.Add((name, values));
        }

        return seriesList;
    }

    private static double ParseDoubleValue(JsonNode? v)
    {
        if (v == null) return 0.0;
        if (v.GetValueKind() == JsonValueKind.Number)
            return v.GetValue<double>();
        if (double.TryParse(v.GetValue<string>(), out var d))
            return d;
        return 0.0;
    }

    private static void AddCategories(IChartData chartData, IChartDataWorkbook workbook, string[]? categories)
    {
        if (categories is not { Length: > 0 }) return;

        for (var i = 0; i < categories.Length; i++)
        {
            var cell = workbook.GetCell(0, i + 1, 0, categories[i]);
            chartData.Categories.Add(cell);
        }
    }

    private static void AddSeriesData(IChart chart, IChartData chartData, IChartDataWorkbook workbook,
        string[]? categories, List<(string name, double[] values)>? seriesList)
    {
        if (seriesList is not { Count: > 0 }) return;

        var baseColumnOffset = categories != null ? 1 : 0;
        var yColumnStart = baseColumnOffset + 1;

        for (var seriesIdx = 0; seriesIdx < seriesList.Count; seriesIdx++)
        {
            var (seriesName, values) = seriesList[seriesIdx];
            var yColumnIndex = yColumnStart + seriesIdx;

            var headerCell = workbook.GetCell(0, 0, yColumnIndex, seriesName);
            var series = chartData.Series.Add(headerCell, chart.Type);

            AddDataPointsForChartType(chart, series, workbook, values, yColumnIndex, yColumnStart, seriesList.Count,
                seriesIdx);
        }
    }

    private static void AddDataPointsForChartType(IChart chart, IChartSeries series, IChartDataWorkbook workbook,
        double[] values, int yColumnIndex, int yColumnStart, int seriesCount, int seriesIdx)
    {
        if (IsBubbleChart(chart.Type))
            AddBubbleDataPoints(series, workbook, values, yColumnIndex, yColumnStart + seriesCount + seriesIdx);
        else if (IsScatterChart(chart.Type))
            AddScatterDataPoints(series, workbook, values, yColumnIndex);
        else if (IsPieChart(chart.Type))
            AddPieDataPoints(series, workbook, values, yColumnIndex);
        else
            AddBarDataPoints(series, workbook, values, yColumnIndex);
    }

    private static bool IsBubbleChart(ChartType chartType)
    {
        return chartType == ChartType.Bubble || chartType == ChartType.BubbleWith3D;
    }

    private static bool IsScatterChart(ChartType chartType)
    {
        return chartType == ChartType.ScatterWithSmoothLines ||
               chartType == ChartType.ScatterWithStraightLines ||
               chartType == ChartType.ScatterWithStraightLinesAndMarkers ||
               chartType == ChartType.ScatterWithSmoothLinesAndMarkers;
    }

    private static bool IsPieChart(ChartType chartType)
    {
        return chartType == ChartType.Pie || chartType == ChartType.Doughnut;
    }

    private static void AddBubbleDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex, int sizeColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var xCell = workbook.GetCell(0, dataIdx + 1, 0, (double)(dataIdx + 1));
            var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            var sizeValue = Math.Abs(values[dataIdx]) > 0 ? Math.Abs(values[dataIdx]) * 0.5 : 10.0;
            var sCell = workbook.GetCell(0, dataIdx + 1, sizeColumnIndex, sizeValue);
            series.DataPoints.AddDataPointForBubbleSeries(xCell, yCell, sCell);
        }
    }

    private static void AddScatterDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var xCell = workbook.GetCell(0, dataIdx + 1, 0, (double)(dataIdx + 1));
            var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            series.DataPoints.AddDataPointForScatterSeries(xCell, yCell);
        }
    }

    private static void AddPieDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            series.DataPoints.AddDataPointForPieSeries(dataCell);
        }
    }

    private static void AddBarDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            series.DataPoints.AddDataPointForBarSeries(dataCell);
        }
    }

    private static void SetChartRange(IChart chart, IChartData chartData, string[]? categories,
        List<(string name, double[] values)>? seriesList)
    {
        var maxRow = Math.Max(
            categories?.Length ?? 0,
            seriesList?.Max(s => s.values.Length) ?? 0
        );

        var maxCol = (seriesList?.Count ?? 0) + (categories != null ? 1 : 0);
        if (IsBubbleChart(chart.Type))
            maxCol += seriesList?.Count ?? 0;

        if (maxRow > 0 && maxCol > 0)
        {
            var lastCol = (char)('A' + maxCol - 1);
            var range = $"Sheet1!$A$1:${lastCol}${maxRow + 1}";
            chartData.SetRange(range);
        }
    }
}
