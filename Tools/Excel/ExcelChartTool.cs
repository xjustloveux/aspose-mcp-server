using System.Text.Json;
using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Core;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel charts (add, edit, delete, get, update data, set properties)
/// </summary>
public class ExcelChartTool : IAsposeTool
{
    public string Description =>
        @"Manage Excel charts. Supports 6 operations: add, edit, delete, get, update_data, set_properties.

Usage examples:
- Add chart: excel_chart(operation='add', path='book.xlsx', chartType='Column', dataRange='A1:B10', position='A12')
- Edit chart: excel_chart(operation='edit', path='book.xlsx', chartIndex=0, chartType='Line')
- Delete chart: excel_chart(operation='delete', path='book.xlsx', chartIndex=0)
- Get charts: excel_chart(operation='get', path='book.xlsx')
- Update data: excel_chart(operation='update_data', path='book.xlsx', chartIndex=0, dataRange='A1:C10')
- Set properties: excel_chart(operation='set_properties', path='book.xlsx', chartIndex=0, title='Chart Title')";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            operation = new
            {
                type = "string",
                description = @"Operation to perform.
- 'add': Add a chart (required params: path, chartType, dataRange)
- 'edit': Edit chart properties (required params: path, chartIndex)
- 'delete': Delete a chart (required params: path, chartIndex)
- 'get': Get all charts (required params: path)
- 'update_data': Update chart data (required params: path, chartIndex, dataRange)
- 'set_properties': Set chart properties (required params: path, chartIndex)",
                @enum = new[] { "add", "edit", "delete", "get", "update_data", "set_properties" }
            },
            path = new
            {
                type = "string",
                description = "Excel file path (required for all operations)"
            },
            outputPath = new
            {
                type = "string",
                description = "Output file path (optional, defaults to input path)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            chartIndex = new
            {
                type = "number",
                description = "Chart index (0-based, required for edit/delete/update_data/set_properties)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type (required for add, optional for edit)",
                @enum = new[]
                {
                    "Column", "Bar", "Line", "Pie", "Area", "Scatter", "Doughnut", "Radar", "Bubble", "Cylinder",
                    "Cone", "Pyramid"
                }
            },
            dataRange = new
            {
                type = "string",
                description = "Data range for chart values (e.g., 'B1:B10' or 'B1:C10' for multiple series)"
            },
            categoryAxisDataRange = new
            {
                type = "string",
                description = "Category axis (X-axis) data range (optional, e.g., 'A1:A10')"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            },
            topRow = new
            {
                type = "number",
                description = "Top row index for chart position (0-based, optional, default: auto-detect)"
            },
            leftColumn = new
            {
                type = "number",
                description = "Left column index for chart position (0-based, optional, default: 0)"
            },
            width = new
            {
                type = "number",
                description = "Chart width in columns (optional, default: 10)"
            },
            height = new
            {
                type = "number",
                description = "Chart height in rows (optional, default: 15)"
            },
            showLegend = new
            {
                type = "boolean",
                description = "Show legend (optional, for edit/set_properties)"
            },
            legendPosition = new
            {
                type = "string",
                description = "Legend position: Bottom, Top, Left, Right (optional)"
            },
            removeTitle = new
            {
                type = "boolean",
                description = "Remove title (optional, for set_properties)"
            },
            legendVisible = new
            {
                type = "boolean",
                description = "Legend visibility (optional, for set_properties)"
            }
        },
        required = new[] { "operation", "path" }
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
        var outputPath = ArgumentHelper.GetAndValidateOutputPath(arguments, path);
        var sheetIndex = ArgumentHelper.GetInt(arguments, "sheetIndex", 0);

        return operation.ToLower() switch
        {
            "add" => await AddChartAsync(path, outputPath, sheetIndex, arguments),
            "edit" => await EditChartAsync(path, outputPath, sheetIndex, arguments),
            "delete" => await DeleteChartAsync(path, outputPath, sheetIndex, arguments),
            "get" => await GetChartsAsync(path, sheetIndex),
            "update_data" => await UpdateChartDataAsync(path, outputPath, sheetIndex, arguments),
            "set_properties" => await SetChartPropertiesAsync(path, outputPath, sheetIndex, arguments),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Parses chart type string to ChartType enum
    /// </summary>
    /// <param name="chartTypeStr">Chart type string</param>
    /// <param name="defaultType">Default type if parsing fails</param>
    /// <returns>Parsed ChartType</returns>
    private static ChartType ParseChartType(string? chartTypeStr, ChartType defaultType = ChartType.Column)
    {
        if (string.IsNullOrEmpty(chartTypeStr))
            return defaultType;

        return Enum.TryParse<ChartType>(chartTypeStr, true, out var result) ? result : defaultType;
    }

    /// <summary>
    ///     Parses legend position string to LegendPositionType enum
    /// </summary>
    /// <param name="positionStr">Position string</param>
    /// <param name="defaultPosition">Default position if parsing fails</param>
    /// <returns>Parsed LegendPositionType</returns>
    private static LegendPositionType ParseLegendPosition(string? positionStr,
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
    ///     Sets category data for chart series
    /// </summary>
    /// <param name="chart">Target chart</param>
    /// <param name="categoryAxisDataRange">Category axis data range</param>
    private static void SetCategoryData(Chart chart, string categoryAxisDataRange)
    {
        if (string.IsNullOrEmpty(categoryAxisDataRange) || chart.NSeries.Count == 0)
            return;

        chart.NSeries.CategoryData = categoryAxisDataRange;
    }

    /// <summary>
    ///     Adds data series to chart
    /// </summary>
    /// <param name="chart">Target chart</param>
    /// <param name="dataRange">Data range string (can be comma-separated for multiple series)</param>
    private static void AddDataSeries(Chart chart, string dataRange)
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
    ///     Adds a chart to the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing chartType, dataRange, optional categoryAxisDataRange, title, position</param>
    /// <returns>Success message with chart details</returns>
    private Task<string> AddChartAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var chartTypeStr = ArgumentHelper.GetString(arguments, "chartType");
            var dataRange = ArgumentHelper.GetString(arguments, "dataRange");
            var categoryAxisDataRange = ArgumentHelper.GetStringNullable(arguments, "categoryAxisDataRange");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var topRow = ArgumentHelper.GetIntNullable(arguments, "topRow");
            var leftColumn = ArgumentHelper.GetInt(arguments, "leftColumn", 0);
            var width = ArgumentHelper.GetInt(arguments, "width", 10);
            var height = ArgumentHelper.GetInt(arguments, "height", 15);

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var chartType = ParseChartType(chartTypeStr);

            int chartTopRow;
            if (topRow.HasValue)
            {
                chartTopRow = topRow.Value;
            }
            else
            {
                var dataRangeObj = ExcelHelper.CreateRange(worksheet.Cells, dataRange.Split(',')[0].Trim());
                chartTopRow = dataRangeObj.FirstRow + dataRangeObj.RowCount + 2;
            }

            var chartIndex =
                worksheet.Charts.Add(chartType, chartTopRow, leftColumn, chartTopRow + height, leftColumn + width);
            var chart = worksheet.Charts[chartIndex];

            AddDataSeries(chart, dataRange);
            SetCategoryData(chart, categoryAxisDataRange ?? "");

            if (!string.IsNullOrEmpty(title))
                chart.Title.Text = title;

            workbook.CalculateFormula();
            workbook.Save(outputPath);

            var result = $"Chart added with data range: {dataRange}";
            if (!string.IsNullOrEmpty(categoryAxisDataRange))
                result += $", X-axis: {categoryAxisDataRange}";
            result += $". Output: {outputPath}";
            return result;
        });
    }

    /// <summary>
    ///     Edits chart properties
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing chartIndex and various chart properties</param>
    /// <returns>Success message with applied changes</returns>
    private Task<string> EditChartAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var dataRange = ArgumentHelper.GetStringNullable(arguments, "dataRange");
            var categoryAxisDataRange = ArgumentHelper.GetStringNullable(arguments, "categoryAxisDataRange");
            var chartTypeStr = ArgumentHelper.GetStringNullable(arguments, "chartType");
            var showLegend = ArgumentHelper.GetBoolNullable(arguments, "showLegend");
            var legendPosition = ArgumentHelper.GetStringNullable(arguments, "legendPosition");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                throw new ArgumentException(
                    $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

            var chart = worksheet.Charts[chartIndex];
            var changes = new List<string>();

            if (!string.IsNullOrEmpty(title))
            {
                chart.Title.Text = title;
                changes.Add($"Title: {title}");
            }

            if (!string.IsNullOrEmpty(dataRange))
            {
                AddDataSeries(chart, dataRange);
                SetCategoryData(chart, categoryAxisDataRange ?? "");

                var rangeInfo = dataRange;
                if (!string.IsNullOrEmpty(categoryAxisDataRange))
                    rangeInfo += $", X-axis: {categoryAxisDataRange}";
                changes.Add($"Data range: {rangeInfo}");
            }

            if (!string.IsNullOrEmpty(chartTypeStr))
            {
                chart.Type = ParseChartType(chartTypeStr, chart.Type);
                changes.Add($"Chart type: {chartTypeStr}");
            }

            if (showLegend.HasValue)
            {
                chart.ShowLegend = showLegend.Value;
                changes.Add($"Legend: {(showLegend.Value ? "show" : "hide")}");
            }

            if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
            {
                chart.Legend.Position = ParseLegendPosition(legendPosition, chart.Legend.Position);
                changes.Add($"Legend position: {legendPosition}");
            }

            workbook.Save(outputPath);

            return changes.Count > 0
                ? $"Chart #{chartIndex} edited: {string.Join(", ", changes)}. Output: {outputPath}"
                : $"Chart #{chartIndex} no changes. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Deletes a chart from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing chartIndex</param>
    /// <returns>Success message with remaining chart count</returns>
    private Task<string> DeleteChartAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                throw new ArgumentException(
                    $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

            var chartName = worksheet.Charts[chartIndex].Name ?? $"Chart {chartIndex}";
            worksheet.Charts.RemoveAt(chartIndex);
            workbook.Save(outputPath);

            return
                $"Chart #{chartIndex} ({chartName}) deleted, {worksheet.Charts.Count} remaining. Output: {outputPath}";
        });
    }

    /// <summary>
    ///     Gets all charts from the worksheet
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <returns>JSON string with chart details including position and data sources</returns>
    private Task<string> GetChartsAsync(string path, int sheetIndex)
    {
        return Task.Run(() =>
        {
            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);
            var charts = worksheet.Charts;

            if (charts.Count == 0)
            {
                var emptyResult = new
                {
                    count = 0,
                    worksheetName = worksheet.Name,
                    items = Array.Empty<object>(),
                    message = "No charts found"
                };
                return JsonSerializer.Serialize(emptyResult, new JsonSerializerOptions { WriteIndented = true });
            }

            var chartList = new List<object>();
            for (var i = 0; i < charts.Count; i++)
            {
                var chart = charts[i];
                var seriesList = new List<object>();

                if (chart.NSeries is { Count: > 0 })
                    for (var j = 0; j < chart.NSeries.Count && j < 5; j++)
                    {
                        var series = chart.NSeries[j];
                        seriesList.Add(new
                        {
                            index = j,
                            name = series.Name ?? "(no name)",
                            valuesRange = series.Values ?? "",
                            categoryData = chart.NSeries.CategoryData
                        });
                    }

                chartList.Add(new
                {
                    index = i,
                    name = chart.Name ?? "(no name)",
                    type = chart.Type.ToString(),
                    location = new
                    {
                        upperLeftRow = chart.ChartObject.UpperLeftRow,
                        lowerRightRow = chart.ChartObject.LowerRightRow,
                        upperLeftColumn = chart.ChartObject.UpperLeftColumn,
                        lowerRightColumn = chart.ChartObject.LowerRightColumn
                    },
                    width = chart.ChartObject.Width,
                    height = chart.ChartObject.Height,
                    title = chart.Title?.Text,
                    legendEnabled = chart.Legend != null,
                    seriesCount = chart.NSeries?.Count ?? 0,
                    series = seriesList
                });
            }

            var result = new
            {
                count = charts.Count,
                worksheetName = worksheet.Name,
                items = chartList
            };

            return JsonSerializer.Serialize(result, new JsonSerializerOptions { WriteIndented = true });
        });
    }

    /// <summary>
    ///     Updates chart data range
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing chartIndex, dataRange, optional categoryAxisDataRange</param>
    /// <returns>Success message with new data range</returns>
    private Task<string> UpdateChartDataAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
            var dataRange = ArgumentHelper.GetString(arguments, "dataRange");
            var categoryAxisDataRange = ArgumentHelper.GetStringNullable(arguments, "categoryAxisDataRange");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                throw new ArgumentException(
                    $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

            var chart = worksheet.Charts[chartIndex];
            AddDataSeries(chart, dataRange);
            SetCategoryData(chart, categoryAxisDataRange ?? "");

            workbook.Save(outputPath);

            var result = $"Chart #{chartIndex} data updated to: {dataRange}";
            if (!string.IsNullOrEmpty(categoryAxisDataRange))
                result += $", X-axis: {categoryAxisDataRange}";
            result += $". Output: {outputPath}";
            return result;
        });
    }

    /// <summary>
    ///     Sets chart properties (title, legend)
    /// </summary>
    /// <param name="path">Excel file path</param>
    /// <param name="outputPath">Output file path</param>
    /// <param name="sheetIndex">Worksheet index (0-based)</param>
    /// <param name="arguments">JSON arguments containing chartIndex and various chart properties</param>
    /// <returns>Success message with applied changes</returns>
    private Task<string> SetChartPropertiesAsync(string path, string outputPath, int sheetIndex, JsonObject? arguments)
    {
        return Task.Run(() =>
        {
            var chartIndex = ArgumentHelper.GetInt(arguments, "chartIndex");
            var title = ArgumentHelper.GetStringNullable(arguments, "title");
            var removeTitle = ArgumentHelper.GetBool(arguments, "removeTitle", false);
            var legendVisible = ArgumentHelper.GetBoolNullable(arguments, "legendVisible");
            var legendPosition = ArgumentHelper.GetStringNullable(arguments, "legendPosition");

            using var workbook = new Workbook(path);
            var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

            if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
                throw new ArgumentException(
                    $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

            var chart = worksheet.Charts[chartIndex];
            var changes = new List<string>();

            if (removeTitle)
            {
                chart.Title.Text = "";
                changes.Add("Title removed");
            }
            else if (!string.IsNullOrEmpty(title))
            {
                chart.Title.Text = title;
                changes.Add($"Title: {title}");
            }

            if (legendVisible.HasValue)
            {
                chart.ShowLegend = legendVisible.Value;
                changes.Add($"Legend: {(legendVisible.Value ? "show" : "hide")}");
            }

            if (!string.IsNullOrEmpty(legendPosition) && chart.Legend != null)
            {
                chart.Legend.Position = ParseLegendPosition(legendPosition, chart.Legend.Position);
                changes.Add($"Legend position: {legendPosition}");
            }

            workbook.Save(outputPath);

            return changes.Count > 0
                ? $"Chart #{chartIndex} properties updated: {string.Join(", ", changes)}. Output: {outputPath}"
                : $"Chart #{chartIndex} no changes. Output: {outputPath}";
        });
    }
}