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
        var p = ExtractUpdateChartDataParameters(parameters);

        var presentation = context.Document;
        var slide = PowerPointHelper.GetSlide(presentation, p.SlideIndex);
        var chart = PptChartHelper.GetChartByIndex(slide, p.ChartIndex, p.SlideIndex);

        var (categories, seriesList) = ParseChartData(p.Data);

        if (categories == null && seriesList == null && !p.ClearExisting)
            return Success($"No changes made to chart {p.ChartIndex} on slide {p.SlideIndex}.");

        var chartData = chart.ChartData;
        var workbook = chartData.ChartDataWorkbook;

        if (p.ClearExisting || categories != null || seriesList != null)
        {
            chartData.Series.Clear();
            chartData.Categories.Clear();
        }

        AddCategories(chartData, workbook, categories);
        AddSeriesData(chart, chartData, workbook, categories, seriesList);
        SetChartRange(chart, chartData, categories, seriesList);

        MarkModified(context);

        return Success($"Chart {p.ChartIndex} data updated on slide {p.SlideIndex}.");
    }

    /// <summary>
    ///     Parses chart data from a JSON object.
    /// </summary>
    /// <param name="dataObject">The JSON object containing chart data.</param>
    /// <returns>A tuple containing categories and series list.</returns>
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

    /// <summary>
    ///     Parses series data from a JSON array.
    /// </summary>
    /// <param name="seriesArray">The JSON array containing series data.</param>
    /// <returns>A list of series with name and values.</returns>
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

    /// <summary>
    ///     Parses a JSON node to a double value.
    /// </summary>
    /// <param name="v">The JSON node to parse.</param>
    /// <returns>The parsed double value, or 0.0 if parsing fails.</returns>
    private static double ParseDoubleValue(JsonNode? v)
    {
        if (v == null) return 0.0;
        if (v.GetValueKind() == JsonValueKind.Number)
            return v.GetValue<double>();
        if (double.TryParse(v.GetValue<string>(), out var d))
            return d;
        return 0.0;
    }

    /// <summary>
    ///     Adds categories to the chart data.
    /// </summary>
    /// <param name="chartData">The chart data.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="categories">The categories to add.</param>
    private static void AddCategories(IChartData chartData, IChartDataWorkbook workbook, string[]? categories)
    {
        if (categories is not { Length: > 0 }) return;

        for (var i = 0; i < categories.Length; i++)
        {
            var cell = workbook.GetCell(0, i + 1, 0, categories[i]);
            chartData.Categories.Add(cell);
        }
    }

    /// <summary>
    ///     Adds series data to the chart.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <param name="chartData">The chart data.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="categories">The categories.</param>
    /// <param name="seriesList">The series list to add.</param>
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

    /// <summary>
    ///     Adds data points based on chart type.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <param name="series">The chart series.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="values">The data values.</param>
    /// <param name="yColumnIndex">The Y column index.</param>
    /// <param name="yColumnStart">The Y column start index.</param>
    /// <param name="seriesCount">The total series count.</param>
    /// <param name="seriesIdx">The current series index.</param>
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

    /// <summary>
    ///     Determines whether the chart type is a bubble chart.
    /// </summary>
    /// <param name="chartType">The chart type.</param>
    /// <returns>True if the chart is a bubble chart, false otherwise.</returns>
    private static bool IsBubbleChart(ChartType chartType)
    {
        return chartType == ChartType.Bubble || chartType == ChartType.BubbleWith3D;
    }

    /// <summary>
    ///     Determines whether the chart type is a scatter chart.
    /// </summary>
    /// <param name="chartType">The chart type.</param>
    /// <returns>True if the chart is a scatter chart, false otherwise.</returns>
    private static bool IsScatterChart(ChartType chartType)
    {
        return chartType == ChartType.ScatterWithSmoothLines ||
               chartType == ChartType.ScatterWithStraightLines ||
               chartType == ChartType.ScatterWithStraightLinesAndMarkers ||
               chartType == ChartType.ScatterWithSmoothLinesAndMarkers;
    }

    /// <summary>
    ///     Determines whether the chart type is a pie chart.
    /// </summary>
    /// <param name="chartType">The chart type.</param>
    /// <returns>True if the chart is a pie chart, false otherwise.</returns>
    private static bool IsPieChart(ChartType chartType)
    {
        return chartType == ChartType.Pie || chartType == ChartType.Doughnut;
    }

    /// <summary>
    ///     Adds data points for bubble chart series.
    /// </summary>
    /// <param name="series">The chart series.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="values">The data values.</param>
    /// <param name="yColumnIndex">The Y column index.</param>
    /// <param name="sizeColumnIndex">The size column index.</param>
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

    /// <summary>
    ///     Adds data points for scatter chart series.
    /// </summary>
    /// <param name="series">The chart series.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="values">The data values.</param>
    /// <param name="yColumnIndex">The Y column index.</param>
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

    /// <summary>
    ///     Adds data points for pie chart series.
    /// </summary>
    /// <param name="series">The chart series.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="values">The data values.</param>
    /// <param name="yColumnIndex">The Y column index.</param>
    private static void AddPieDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            series.DataPoints.AddDataPointForPieSeries(dataCell);
        }
    }

    /// <summary>
    ///     Adds data points for bar chart series.
    /// </summary>
    /// <param name="series">The chart series.</param>
    /// <param name="workbook">The chart data workbook.</param>
    /// <param name="values">The data values.</param>
    /// <param name="yColumnIndex">The Y column index.</param>
    private static void AddBarDataPoints(IChartSeries series, IChartDataWorkbook workbook, double[] values,
        int yColumnIndex)
    {
        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
        {
            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
            series.DataPoints.AddDataPointForBarSeries(dataCell);
        }
    }

    /// <summary>
    ///     Sets the chart data range.
    /// </summary>
    /// <param name="chart">The chart.</param>
    /// <param name="chartData">The chart data.</param>
    /// <param name="categories">The categories.</param>
    /// <param name="seriesList">The series list.</param>
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

    /// <summary>
    ///     Extracts update chart data parameters from operation parameters.
    /// </summary>
    /// <param name="parameters">The operation parameters.</param>
    /// <returns>The extracted update chart data parameters.</returns>
    private static UpdateChartDataParameters ExtractUpdateChartDataParameters(OperationParameters parameters)
    {
        return new UpdateChartDataParameters(
            parameters.GetRequired<int>("slideIndex"),
            parameters.GetRequired<int>("shapeIndex"),
            parameters.GetOptional<JsonObject?>("data"),
            parameters.GetOptional("clearExisting", false));
    }

    /// <summary>
    ///     Record for holding update chart data parameters.
    /// </summary>
    /// <param name="SlideIndex">The slide index.</param>
    /// <param name="ChartIndex">The chart shape index.</param>
    /// <param name="Data">The optional chart data object.</param>
    /// <param name="ClearExisting">Whether to clear existing data.</param>
    private record UpdateChartDataParameters(int SlideIndex, int ChartIndex, JsonObject? Data, bool ClearExisting);
}
