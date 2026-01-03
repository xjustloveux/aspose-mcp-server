using System.ComponentModel;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint charts (add, edit, delete, get data, update data)
///     Merges: PptAddChartTool, PptEditChartTool, PptDeleteChartTool, PptGetChartDataTool, PptUpdateChartDataTool
/// </summary>
[McpServerToolType]
public class PptChartTool
{
    /// <summary>
    ///     Session manager for document lifecycle management.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="PptChartTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public PptChartTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "ppt_chart")]
    [Description(@"Manage PowerPoint charts. Supports 5 operations: add, edit, delete, get_data, update_data.

Usage examples:
- Add chart: ppt_chart(operation='add', path='presentation.pptx', slideIndex=0, chartType='Column', x=100, y=100, width=400, height=300)
- Edit chart: ppt_chart(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, title='New Title')
- Delete chart: ppt_chart(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get data: ppt_chart(operation='get_data', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Update data: ppt_chart(operation='update_data', path='presentation.pptx', slideIndex=0, shapeIndex=0, data={categories:['A','B'],series:[{name:'Sales',values:[1,2]}]})

Note: shapeIndex refers to the chart index (0-based) among all charts on the slide, not the absolute shape index.")]
    public string Execute(
        [Description("Operation: add, edit, delete, get_data, update_data")]
        string operation,
        [Description("Slide index (0-based)")] int slideIndex,
        [Description("Presentation file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Chart index (0-based, required for edit/delete/get_data/update_data)")]
        int? shapeIndex = null,
        [Description("Chart type (Column, Bar, Line, Pie, etc., required for add, optional for edit)")]
        string? chartType = null,
        [Description("Chart title (optional)")]
        string? title = null,
        [Description("Chart X position in points (optional for add, default: 50)")]
        float x = 50,
        [Description("Chart Y position in points (optional for add, default: 50)")]
        float y = 50,
        [Description("Chart width in points (optional for add, default: 500)")]
        float width = 500,
        [Description("Chart height in points (optional for add, default: 400)")]
        float height = 400,
        [Description("Chart data object with categories and series (optional, for edit/update_data)")]
        JsonObject? data = null,
        [Description("Clear existing data before adding new (optional, for update_data, default: false)")]
        bool clearExisting = false)
    {
        using var ctx = DocumentContext<Presentation>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddChart(ctx, outputPath, slideIndex, chartType, title, x, y, width, height),
            "edit" => EditChart(ctx, outputPath, slideIndex, shapeIndex, title, chartType),
            "delete" => DeleteChart(ctx, outputPath, slideIndex, shapeIndex),
            "get_data" => GetChartData(ctx, slideIndex, shapeIndex),
            "update_data" => UpdateChartData(ctx, outputPath, slideIndex, shapeIndex, data, clearExisting),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a chart to a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="chartTypeStr">The chart type string.</param>
    /// <param name="title">The chart title.</param>
    /// <param name="x">The X position in points.</param>
    /// <param name="y">The Y position in points.</param>
    /// <param name="width">The chart width in points.</param>
    /// <param name="height">The chart height in points.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when chartType is not provided.</exception>
    private static string AddChart(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        string? chartTypeStr, string? title, float x, float y, float width, float height)
    {
        if (string.IsNullOrEmpty(chartTypeStr))
            throw new ArgumentException("chartType is required for add operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

        var chartType = ParseChartType(chartTypeStr);
        var chart = slide.Shapes.AddChart(chartType, x, y, width, height);

        if (!string.IsNullOrEmpty(title))
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

        ctx.Save(outputPath);

        var result = $"Chart '{chartTypeStr}' added to slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Edits chart properties.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="title">The new chart title.</param>
    /// <param name="chartTypeStr">The new chart type string.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided.</exception>
    /// <exception cref="InvalidOperationException">Thrown when chart editing fails.</exception>
    private static string EditChart(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? chartIndex, string? title, string? chartTypeStr)
    {
        if (!chartIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for edit operation");

        try
        {
            var presentation = ctx.Document;
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
            var chart = GetChartByIndex(slide, chartIndex.Value, slideIndex);

            if (!string.IsNullOrEmpty(title))
                try
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
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to set chart title: {ex.Message}", ex);
                }

            if (!string.IsNullOrEmpty(chartTypeStr))
                try
                {
                    chart.Type = ParseChartType(chartTypeStr, chart.Type);
                }
                catch (Exception ex)
                {
                    throw new InvalidOperationException($"Failed to change chart type: {ex.Message}", ex);
                }

            ctx.Save(outputPath);

            var result = $"Chart {chartIndex} updated on slide {slideIndex}.\n";
            result += ctx.GetOutputMessage(outputPath);
            return result;
        }
        catch (ArgumentException)
        {
            throw;
        }
        catch (Exception ex)
        {
            throw new InvalidOperationException($"Error editing chart: {ex.Message}", ex);
        }
    }

    /// <summary>
    ///     Deletes a chart from a slide.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided.</exception>
    private static string DeleteChart(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? chartIndex)
    {
        if (!chartIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for delete operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var chart = GetChartByIndex(slide, chartIndex.Value, slideIndex);

        slide.Shapes.Remove(chart);

        ctx.Save(outputPath);

        var result = $"Chart {chartIndex} deleted from slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets chart data.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <returns>A JSON string containing the chart data.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided.</exception>
    private static string GetChartData(DocumentContext<Presentation> ctx, int slideIndex, int? chartIndex)
    {
        if (!chartIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for get_data operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var chart = GetChartByIndex(slide, chartIndex.Value, slideIndex);
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
            slideIndex,
            chartIndex,
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

        return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
    }

    /// <summary>
    ///     Updates chart data.
    /// </summary>
    /// <param name="ctx">The document context.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="slideIndex">The slide index (0-based).</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="dataObject">The chart data object containing categories and series.</param>
    /// <param name="clearExisting">Whether to clear existing data before adding new data.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when shapeIndex is not provided.</exception>
    private static string UpdateChartData(DocumentContext<Presentation> ctx, string? outputPath, int slideIndex,
        int? chartIndex, JsonObject? dataObject, bool clearExisting)
    {
        if (!chartIndex.HasValue)
            throw new ArgumentException("shapeIndex is required for update_data operation");

        var presentation = ctx.Document;
        var slide = PowerPointHelper.GetSlide(presentation, slideIndex);
        var chart = GetChartByIndex(slide, chartIndex.Value, slideIndex);

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
        {
            ctx.Save(outputPath);
            var noChangeResult = $"No changes made to chart {chartIndex} on slide {slideIndex}.\n";
            noChangeResult += ctx.GetOutputMessage(outputPath);
            return noChangeResult;
        }

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

        ctx.Save(outputPath);

        var result = $"Chart {chartIndex} data updated on slide {slideIndex}.\n";
        result += ctx.GetOutputMessage(outputPath);
        return result;
    }

    /// <summary>
    ///     Gets a chart by index from a slide.
    /// </summary>
    /// <param name="slide">The slide containing the chart.</param>
    /// <param name="chartIndex">The chart index (0-based).</param>
    /// <param name="slideIndex">The slide index for error messages.</param>
    /// <returns>The chart at the specified index.</returns>
    /// <exception cref="ArgumentException">Thrown when no charts exist or index is out of range.</exception>
    private static IChart GetChartByIndex(ISlide slide, int chartIndex, int slideIndex)
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
    private static ChartType ParseChartType(string? chartTypeStr, ChartType defaultType = ChartType.ClusteredColumn)
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
}