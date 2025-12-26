using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Slides;
using Aspose.Slides.Charts;
using Aspose.Slides.Export;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.PowerPoint;

/// <summary>
///     Unified tool for managing PowerPoint charts (add, edit, delete, get data, update data)
///     Merges: PptAddChartTool, PptEditChartTool, PptDeleteChartTool, PptGetChartDataTool, PptUpdateChartDataTool
/// </summary>
public class PptChartTool : IAsposeTool
{
    public string Description =>
        @"Manage PowerPoint charts. Supports 5 operations: add, edit, delete, get_data, update_data.

Usage examples:
- Add chart: ppt_chart(operation='add', path='presentation.pptx', slideIndex=0, chartType='Column', x=100, y=100, width=400, height=300)
- Edit chart: ppt_chart(operation='edit', path='presentation.pptx', slideIndex=0, shapeIndex=0, title='New Title')
- Delete chart: ppt_chart(operation='delete', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Get data: ppt_chart(operation='get_data', path='presentation.pptx', slideIndex=0, shapeIndex=0)
- Update data: ppt_chart(operation='update_data', path='presentation.pptx', slideIndex=0, shapeIndex=0, data=[['A','B'],['1','2']])

Note: shapeIndex refers to the chart index (0-based) among all charts on the slide, not the absolute shape index.";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a chart (required params: path, slideIndex, chartType)
- 'edit': Edit chart properties (required params: path, slideIndex, shapeIndex)
- 'delete': Delete a chart (required params: path, slideIndex, shapeIndex)
- 'get_data': Get chart data (required params: path, slideIndex, shapeIndex)
- 'update_data': Update chart data (required params: path, slideIndex, shapeIndex, data)",
                @enum = new[] { "add", "edit", "delete", "get_data", "update_data" }
            },
            path = new
            {
                type = "string",
                description = "Presentation file path (required for all operations)"
            },
            slideIndex = new
            {
                type = "number",
                description = "Slide index (0-based)"
            },
            shapeIndex = new
            {
                type = "number",
                description =
                    "Chart index (0-based, required for edit/delete/get_data/update_data). This refers to the N-th chart on the slide, not the absolute shape index."
            },
            chartType = new
            {
                type = "string",
                description = "Chart type (Column, Bar, Line, Pie, etc., required for add, optional for edit)"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            },
            data = new
            {
                type = "object",
                description = "Chart data object with series and categories (optional, for edit/update_data)",
                properties = new
                {
                    categories = new
                    {
                        type = "array",
                        items = new { type = "string" },
                        description = "Category labels"
                    },
                    series = new
                    {
                        type = "array",
                        items = new
                        {
                            type = "object",
                            properties = new
                            {
                                name = new { type = "string" },
                                values = new
                                {
                                    type = "array",
                                    items = new { type = "number" }
                                }
                            }
                        },
                        description = "Series data"
                    }
                }
            },
            clearExisting = new
            {
                type = "boolean",
                description = "Clear existing data before adding new (optional, for update_data, default: false)"
            },
            outputPath = new
            {
                type = "string",
                description =
                    "Output file path (optional, for add/edit/delete/update_data operations, defaults to input path)"
            }
        },
        required = new[] { "operation", "path", "slideIndex" }
    };

    /// <summary>
    ///     Executes the tool operation with the provided JSON arguments
    /// </summary>
    /// <param name="arguments">JSON arguments object containing operation parameters</param>
    /// <returns>Result message as a string</returns>
    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var operation = ArgumentHelper.GetString(arguments, "operation");
        var path = ArgumentHelper.GetAndValidatePath(arguments);
        var slideIndex = ArgumentHelper.GetInt(arguments, "slideIndex");

        return operation.ToLower() switch
        {
            "add" => await AddChartAsync(arguments, path, slideIndex),
            "edit" => await EditChartAsync(arguments, path, slideIndex),
            "delete" => await DeleteChartAsync(arguments, path, slideIndex),
            "get_data" => await GetChartDataAsync(arguments, path, slideIndex),
            "update_data" => await UpdateChartDataAsync(arguments, path, slideIndex),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Adds a chart to a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartType, optional title, data, x, y, width, height, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message with chart index</returns>
    private Task<string> AddChartAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var chartTypeStr = ArgumentHelper.GetString(arguments, "chartType");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");

            using var presentation = new Presentation(path);
            if (slideIndex < 0 || slideIndex >= presentation.Slides.Count)
                throw new ArgumentException($"slideIndex must be between 0 and {presentation.Slides.Count - 1}");

            var slide = presentation.Slides[slideIndex];

            var chartType = chartTypeStr.ToLower() switch
            {
                "column" => ChartType.ClusteredColumn,
                "bar" => ChartType.ClusteredBar,
                "line" => ChartType.Line,
                "pie" => ChartType.Pie,
                "area" => ChartType.Area,
                "scatter" => ChartType.ScatterWithSmoothLines,
                _ => ChartType.ClusteredColumn
            };

            var chart = slide.Shapes.AddChart(chartType, 50, 50, 500, 400);

            if (!string.IsNullOrEmpty(title))
            {
                chart.HasTitle = true;
                var chartTitle = chart.ChartTitle;
                if (chartTitle != null)
                {
                    // Check if TextFrameForOverriding exists, if not create it
                    if (chartTitle.TextFrameForOverriding != null)
                        chartTitle.TextFrameForOverriding.Text = title;
                    else
                        chartTitle.AddTextFrameForOverriding(title);
                }
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);

            return $"Chart added to slide {slideIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Edits chart properties
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex, optional title, outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> EditChartAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            try
            {
                var chartIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
                var title = ArgumentHelper.GetStringNullable(arguments, "title");
                var chartTypeStr = ArgumentHelper.GetStringNullable(arguments, "chartType");

                using var presentation = new Presentation(path);
                var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

                var charts = slide.Shapes.OfType<IChart>().ToList();

                if (chartIndex < 0 || chartIndex >= charts.Count)
                {
                    var totalShapes = slide.Shapes.Count;
                    var totalCharts = charts.Count;
                    throw new ArgumentException(
                        $"Slide {slideIndex} does not contain a chart at index {chartIndex}. " +
                        $"(Total charts found: {totalCharts}, Total shapes: {totalShapes})");
                }

                var chart = charts[chartIndex];

                if (!string.IsNullOrEmpty(title))
                    try
                    {
                        chart.HasTitle = true;
                        var chartTitle = chart.ChartTitle;
                        if (chartTitle != null)
                        {
                            // Check if TextFrameForOverriding exists, if not create it
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
                        var chartType = chartTypeStr.ToLower() switch
                        {
                            "column" => ChartType.ClusteredColumn,
                            "bar" => ChartType.ClusteredBar,
                            "line" => ChartType.Line,
                            "pie" => ChartType.Pie,
                            "area" => ChartType.Area,
                            "scatter" => ChartType.ScatterWithSmoothLines,
                            "doughnut" => ChartType.Doughnut,
                            "bubble" => ChartType.Bubble,
                            _ => chart.Type
                        };
                        chart.Type = chartType;
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"Failed to change chart type: {ex.Message}", ex);
                    }

                var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                presentation.Save(outputPath, SaveFormat.Pptx);
                return $"Chart {chartIndex} updated on slide {slideIndex}: {outputPath}";
            }
            catch (ArgumentException)
            {
                throw; // Re-throw ArgumentException as-is
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Error editing chart: {ex.Message}", ex);
            }
        });
    }

    /// <summary>
    ///     Deletes a chart from a slide
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> DeleteChartAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var charts = slide.Shapes.OfType<IChart>().ToList();

            if (chartIndex < 0 || chartIndex >= charts.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalCharts = charts.Count;
                throw new ArgumentException(
                    $"Chart index {chartIndex} out of range for slide {slideIndex}. " +
                    $"(Total charts found: {totalCharts}, Total shapes: {totalShapes})");
            }

            var chart = charts[chartIndex];
            slide.Shapes.Remove(chart);

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Chart {chartIndex} deleted from slide {slideIndex}: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets chart data
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Formatted string with chart data</returns>
    private Task<string> GetChartDataAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var charts = slide.Shapes.OfType<IChart>().ToList();

            if (chartIndex < 0 || chartIndex >= charts.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalCharts = charts.Count;
                throw new ArgumentException(
                    $"Chart index {chartIndex} out of range for slide {slideIndex}. " +
                    $"(Total charts found: {totalCharts}, Total shapes: {totalShapes})");
            }

            var chart = charts[chartIndex];

            var sb = new StringBuilder();
            sb.AppendLine($"Chart Type: {chart.Type}");
            sb.AppendLine($"Has Title: {chart.HasTitle}");
            if (chart is { HasTitle: true, ChartTitle: not null })
                sb.AppendLine($"Title: {chart.ChartTitle.TextFrameForOverriding?.Text ?? ""}");
            sb.AppendLine();

            var chartData = chart.ChartData;
            sb.AppendLine($"Categories ({chartData.Categories.Count}):");
            for (var i = 0; i < chartData.Categories.Count; i++)
            {
                var cat = chartData.Categories[i];
                sb.AppendLine($"  [{i}] {cat.Value}");
            }

            sb.AppendLine();

            sb.AppendLine($"Series ({chartData.Series.Count}):");
            for (var i = 0; i < chartData.Series.Count; i++)
            {
                var series = chartData.Series[i];
                sb.AppendLine($"  [{i}] {series.Name}");
                sb.AppendLine($"      Data Points: {series.DataPoints.Count}");
                for (var j = 0; j < Math.Min(series.DataPoints.Count, 10); j++)
                {
                    var point = series.DataPoints[j];
                    sb.AppendLine($"        [{j}] Value: {point.Value}");
                }

                if (series.DataPoints.Count > 10) sb.AppendLine($"        ... ({series.DataPoints.Count - 10} more)");
            }

            return sb.ToString();
        });
    }

    /// <summary>
    ///     Updates chart data
    /// </summary>
    /// <param name="arguments">JSON arguments containing chartIndex, data, optional outputPath</param>
    /// <param name="path">PowerPoint file path</param>
    /// <param name="slideIndex">Slide index (0-based)</param>
    /// <returns>Success message</returns>
    private Task<string> UpdateChartDataAsync(JsonObject? arguments, string path, int slideIndex)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "shapeIndex");
            var dataObject = ArgumentHelper.GetObject(arguments, "data", false);
            var clearExisting = ArgumentHelper.GetBool(arguments, "clearExisting", false);

            using var presentation = new Presentation(path);
            var slide = PowerPointHelper.GetSlide(presentation, slideIndex);

            var charts = slide.Shapes.OfType<IChart>().ToList();

            if (chartIndex < 0 || chartIndex >= charts.Count)
            {
                var totalShapes = slide.Shapes.Count;
                var totalCharts = charts.Count;
                throw new ArgumentException(
                    $"Chart index {chartIndex} out of range for slide {slideIndex}. " +
                    $"(Total charts found: {totalCharts}, Total shapes: {totalShapes})");
            }

            var chart = charts[chartIndex];

            var chartData = chart.ChartData;
            var workbook = chartData.ChartDataWorkbook;

            // Parse data object if provided
            string[]? categories = null;
            List<(string name, double[] values)>? seriesList = null;

            if (dataObject != null)
            {
                var categoriesArray = ArgumentHelper.GetArray(dataObject, "categories", false);
                var seriesArray = ArgumentHelper.GetArray(dataObject, "series", false);

                if (categoriesArray is { Count: > 0 })
                    categories = categoriesArray.Select(c => c?.GetValue<string>() ?? "").ToArray();

                if (seriesArray is { Count: > 0 })
                {
                    seriesList = new List<(string name, double[] values)>();
                    foreach (var seriesNode in seriesArray)
                        if (seriesNode is JsonObject seriesObj)
                        {
                            var name = ArgumentHelper.GetString(seriesObj, "name", "series name", false, "");
                            var valuesArray = ArgumentHelper.GetArray(seriesObj, "values", false);
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

            // If no data provided and clearExisting is false, return early
            if (categories == null && seriesList == null && !clearExisting)
            {
                var earlyOutputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
                presentation.Save(earlyOutputPath, SaveFormat.Pptx);
                return $"No changes made to chart {chartIndex} on slide {slideIndex}";
            }

            // Clear existing data if requested or if we have new data to write
            if (clearExisting || categories != null || seriesList != null)
            {
                chartData.Series.Clear();
                chartData.Categories.Clear();
            }

            // Write categories to column A (starting from row 1, which is A2 in Excel)
            if (categories is { Length: > 0 })
                for (var i = 0; i < categories.Length; i++)
                {
                    var cell = workbook.GetCell(0, i + 1, 0, categories[i]);
                    chartData.Categories.Add(cell);
                }

            // Write series data
            if (seriesList is { Count: > 0 })
            {
                // Write series names to header row (row 0, starting from column B)
                // For Bubble charts, we need extra columns: Y values and Size values
                var baseColumnOffset = categories != null ? 1 : 0; // Column A is for categories/X
                var yColumnStart = baseColumnOffset + 1; // Start Y columns after categories/X

                for (var seriesIdx = 0; seriesIdx < seriesList.Count; seriesIdx++)
                {
                    var seriesName = seriesList[seriesIdx].name;
                    var yColumnIndex = yColumnStart + seriesIdx; // Column B, C, D... for Y values

                    // Write series name to header (B1, C1, etc.)
                    var headerCell = workbook.GetCell(0, 0, yColumnIndex, seriesName);

                    // Create series
                    var series = chartData.Series.Add(headerCell, chart.Type);

                    // Write data points
                    var values = seriesList[seriesIdx].values;

                    // Handle different chart types that require different data point structures
                    if (chart.Type == ChartType.Bubble || chart.Type == ChartType.BubbleWith3D)
                    {
                        // Bubble charts require X, Y, and Size values
                        // Column layout: A = X (or categories), B/C/D... = Y values, after Y columns = Size values
                        var sizeColumnIndex =
                            yColumnStart + seriesList.Count + seriesIdx; // Size columns after Y columns

                        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                        {
                            // X value: use index (1, 2, 3, ...)
                            var xValue = (double)(dataIdx + 1);
                            var xCell = workbook.GetCell(0, dataIdx + 1, 0, xValue); // Column A for X

                            // Y value: from values array
                            var yValue = values[dataIdx];
                            var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, yValue); // Y column

                            // Size value: use a default size based on Y value
                            // In a proper implementation, this should come from user input
                            var sizeValue = Math.Abs(yValue) > 0 ? Math.Abs(yValue) * 0.5 : 10.0;
                            var sCell = workbook.GetCell(0, dataIdx + 1, sizeColumnIndex, sizeValue); // Size column

                            // Add bubble data point with X, Y, and Size
                            series.DataPoints.AddDataPointForBubbleSeries(xCell, yCell, sCell);
                        }
                    }
                    else if (chart.Type == ChartType.ScatterWithSmoothLines ||
                             chart.Type == ChartType.ScatterWithStraightLines ||
                             chart.Type == ChartType.ScatterWithStraightLinesAndMarkers ||
                             chart.Type == ChartType.ScatterWithSmoothLinesAndMarkers)
                    {
                        // Scatter charts require X and Y values
                        // If only Y values are provided, use index as X
                        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                        {
                            // X value: use index (1, 2, 3, ...)
                            var xValue = (double)(dataIdx + 1);
                            var xCell = workbook.GetCell(0, dataIdx + 1, 0, xValue); // Column A for X

                            // Y value: from values array
                            var yValue = values[dataIdx];
                            var yCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, yValue); // Y column

                            // Add scatter data point with X and Y
                            series.DataPoints.AddDataPointForScatterSeries(xCell, yCell);
                        }
                    }
                    else if (chart.Type == ChartType.Pie || chart.Type == ChartType.Doughnut)
                    {
                        // Pie and Doughnut charts use single value
                        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                        {
                            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
                            series.DataPoints.AddDataPointForPieSeries(dataCell);
                        }
                    }
                    else
                    {
                        // Most chart types (Bar, Column, Line, Area, etc.) use single value
                        for (var dataIdx = 0; dataIdx < values.Length; dataIdx++)
                        {
                            var dataCell = workbook.GetCell(0, dataIdx + 1, yColumnIndex, values[dataIdx]);
                            series.DataPoints.AddDataPointForBarSeries(dataCell);
                        }
                    }
                }
            }

            // Update chart data range to ensure proper refresh
            // Calculate the range based on data written
            var maxRow = Math.Max(
                categories?.Length ?? 0,
                seriesList?.Max(s => s.values.Length) ?? 0
            );

            // Calculate max column based on chart type
            // Bubble charts need extra columns for Size values
            var maxCol = (seriesList?.Count ?? 0) + (categories != null ? 1 : 0);
            if (chart.Type == ChartType.Bubble || chart.Type == ChartType.BubbleWith3D)
                // Bubble charts use additional columns for Size (one per series)
                maxCol += seriesList?.Count ?? 0;

            if (maxRow > 0 && maxCol > 0)
            {
                // Set range: Sheet1!$A$1:$[lastCol]$[lastRow]
                var lastCol = (char)('A' + maxCol - 1);
                var range = $"Sheet1!$A$1:${lastCol}${maxRow + 1}";
                chartData.SetRange(range);
            }

            var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
            presentation.Save(outputPath, SaveFormat.Pptx);
            return $"Chart {chartIndex} data updated on slide {slideIndex}: {outputPath}";
        });
    }
}