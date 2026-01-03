using System.ComponentModel;
using System.Text.Json;
using Aspose.Cells;
using Aspose.Cells.Charts;
using AsposeMcpServer.Core.Helpers;
using AsposeMcpServer.Core.Session;
using ModelContextProtocol.Server;

namespace AsposeMcpServer.Tools.Excel;

/// <summary>
///     Unified tool for managing Excel charts (add, edit, delete, get, update data, set properties)
/// </summary>
[McpServerToolType]
public class ExcelChartTool
{
    /// <summary>
    ///     Document session manager for in-memory editing support.
    /// </summary>
    private readonly DocumentSessionManager? _sessionManager;

    /// <summary>
    ///     Initializes a new instance of the <see cref="ExcelChartTool" /> class.
    /// </summary>
    /// <param name="sessionManager">Optional session manager for in-memory document editing.</param>
    public ExcelChartTool(DocumentSessionManager? sessionManager = null)
    {
        _sessionManager = sessionManager;
    }

    [McpServerTool(Name = "excel_chart")]
    [Description(@"Manage Excel charts. Supports 6 operations: add, edit, delete, get, update_data, set_properties.

Usage examples:
- Add chart: excel_chart(operation='add', path='book.xlsx', chartType='Column', dataRange='A1:B10', position='A12')
- Edit chart: excel_chart(operation='edit', path='book.xlsx', chartIndex=0, chartType='Line')
- Delete chart: excel_chart(operation='delete', path='book.xlsx', chartIndex=0)
- Get charts: excel_chart(operation='get', path='book.xlsx')
- Update data: excel_chart(operation='update_data', path='book.xlsx', chartIndex=0, dataRange='A1:C10')
- Set properties: excel_chart(operation='set_properties', path='book.xlsx', chartIndex=0, title='Chart Title')")]
    public string Execute(
        [Description("Operation: add, edit, delete, get, update_data, set_properties")]
        string operation,
        [Description("Excel file path (required if no sessionId)")]
        string? path = null,
        [Description("Session ID for in-memory editing")]
        string? sessionId = null,
        [Description("Output file path (file mode only)")]
        string? outputPath = null,
        [Description("Sheet index (0-based, default: 0)")]
        int sheetIndex = 0,
        [Description("Chart index (0-based, required for edit/delete/update_data/set_properties)")]
        int chartIndex = 0,
        [Description(
            "Chart type: Column, Bar, Line, Pie, Area, Scatter, Doughnut, Radar, Bubble, Cylinder, Cone, Pyramid")]
        string? chartType = null,
        [Description("Data range for chart values (e.g., 'B1:B10' or 'B1:C10' for multiple series)")]
        string? dataRange = null,
        [Description("Category axis (X-axis) data range (optional, e.g., 'A1:A10')")]
        string? categoryAxisDataRange = null,
        [Description("Chart title (optional)")]
        string? title = null,
        [Description("Top row index for chart position (0-based, optional, default: auto-detect)")]
        int? topRow = null,
        [Description("Left column index for chart position (0-based, default: 0)")]
        int leftColumn = 0,
        [Description("Chart width in columns (default: 10)")]
        int width = 10,
        [Description("Chart height in rows (default: 15)")]
        int height = 15,
        [Description("Show legend (optional, for edit/set_properties)")]
        bool? showLegend = null,
        [Description("Legend position: Bottom, Top, Left, Right (optional)")]
        string? legendPosition = null,
        [Description("Remove title (optional, for set_properties)")]
        bool removeTitle = false,
        [Description("Legend visibility (optional, for set_properties)")]
        bool? legendVisible = null)
    {
        using var ctx = DocumentContext<Workbook>.Create(_sessionManager, sessionId, path);

        return operation.ToLower() switch
        {
            "add" => AddChart(ctx, outputPath, sheetIndex, chartType, dataRange, categoryAxisDataRange, title, topRow,
                leftColumn, width, height),
            "edit" => EditChart(ctx, outputPath, sheetIndex, chartIndex, title, dataRange, categoryAxisDataRange,
                chartType, showLegend, legendPosition),
            "delete" => DeleteChart(ctx, outputPath, sheetIndex, chartIndex),
            "get" => GetCharts(ctx, sheetIndex),
            "update_data" => UpdateChartData(ctx, outputPath, sheetIndex, chartIndex, dataRange, categoryAxisDataRange),
            "set_properties" => SetChartProperties(ctx, outputPath, sheetIndex, chartIndex, title, removeTitle,
                legendVisible, legendPosition),
            _ => throw new ArgumentException($"Unknown operation: {operation}")
        };
    }

    /// <summary>
    ///     Parses chart type string to ChartType enum.
    /// </summary>
    /// <param name="chartTypeStr">The chart type string to parse.</param>
    /// <param name="defaultType">The default chart type if parsing fails.</param>
    /// <returns>The parsed ChartType enum value.</returns>
    private static ChartType ParseChartType(string? chartTypeStr, ChartType defaultType = ChartType.Column)
    {
        if (string.IsNullOrEmpty(chartTypeStr))
            return defaultType;

        return Enum.TryParse<ChartType>(chartTypeStr, true, out var result) ? result : defaultType;
    }

    /// <summary>
    ///     Parses legend position string to LegendPositionType enum.
    /// </summary>
    /// <param name="positionStr">The legend position string to parse.</param>
    /// <param name="defaultPosition">The default legend position if parsing fails.</param>
    /// <returns>The parsed LegendPositionType enum value.</returns>
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
    ///     Sets category data for chart series.
    /// </summary>
    /// <param name="chart">The chart to set category data for.</param>
    /// <param name="categoryAxisDataRange">The range for category axis data.</param>
    private static void SetCategoryData(Chart chart, string categoryAxisDataRange)
    {
        if (string.IsNullOrEmpty(categoryAxisDataRange) || chart.NSeries.Count == 0)
            return;

        chart.NSeries.CategoryData = categoryAxisDataRange;
    }

    /// <summary>
    ///     Adds data series to chart.
    /// </summary>
    /// <param name="chart">The chart to add data series to.</param>
    /// <param name="dataRange">The data range for the series.</param>
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
    ///     Adds a chart to the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="chartTypeStr">The chart type string.</param>
    /// <param name="dataRange">The data range for the chart.</param>
    /// <param name="categoryAxisDataRange">The category axis data range.</param>
    /// <param name="title">The chart title.</param>
    /// <param name="topRow">The top row position for the chart.</param>
    /// <param name="leftColumn">The left column position for the chart.</param>
    /// <param name="width">The chart width in columns.</param>
    /// <param name="height">The chart height in rows.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when dataRange is not provided.</exception>
    private static string AddChart(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        string? chartTypeStr, string? dataRange, string? categoryAxisDataRange, string? title,
        int? topRow, int leftColumn, int width, int height)
    {
        if (string.IsNullOrEmpty(dataRange))
            throw new ArgumentException("dataRange is required for add operation");

        var workbook = ctx.Document;
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

        ctx.Save(outputPath);

        var result = $"Chart added with data range: {dataRange}";
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
            result += $", X-axis: {categoryAxisDataRange}";
        result += $". {ctx.GetOutputMessage(outputPath)}";
        return result;
    }

    /// <summary>
    ///     Edits chart properties.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="chartIndex">The chart index.</param>
    /// <param name="title">The chart title.</param>
    /// <param name="dataRange">The data range for the chart.</param>
    /// <param name="categoryAxisDataRange">The category axis data range.</param>
    /// <param name="chartTypeStr">The chart type string.</param>
    /// <param name="showLegend">Whether to show the legend.</param>
    /// <param name="legendPosition">The legend position.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when chart index is out of range.</exception>
    private static string EditChart(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int chartIndex, string? title, string? dataRange, string? categoryAxisDataRange,
        string? chartTypeStr, bool? showLegend, string? legendPosition)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

        var chart = worksheet.Charts[chartIndex];
        List<string> changes = [];

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

        ctx.Save(outputPath);

        return changes.Count > 0
            ? $"Chart #{chartIndex} edited: {string.Join(", ", changes)}. {ctx.GetOutputMessage(outputPath)}"
            : $"Chart #{chartIndex} no changes. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Deletes a chart from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="chartIndex">The chart index to delete.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when chart index is out of range.</exception>
    private static string DeleteChart(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex, int chartIndex)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

        var chartName = worksheet.Charts[chartIndex].Name ?? $"Chart {chartIndex}";
        worksheet.Charts.RemoveAt(chartIndex);

        ctx.Save(outputPath);

        return
            $"Chart #{chartIndex} ({chartName}) deleted, {worksheet.Charts.Count} remaining. {ctx.GetOutputMessage(outputPath)}";
    }

    /// <summary>
    ///     Gets all charts from the worksheet.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <returns>A JSON string containing information about all charts.</returns>
    private static string GetCharts(DocumentContext<Workbook> ctx, int sheetIndex)
    {
        var workbook = ctx.Document;
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

        List<object> chartList = [];
        for (var i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            List<object> seriesList = [];

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
    }

    /// <summary>
    ///     Updates chart data range.
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="chartIndex">The chart index.</param>
    /// <param name="dataRange">The new data range for the chart.</param>
    /// <param name="categoryAxisDataRange">The category axis data range.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when dataRange is not provided or chart index is out of range.</exception>
    private static string UpdateChartData(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int chartIndex, string? dataRange, string? categoryAxisDataRange)
    {
        if (string.IsNullOrEmpty(dataRange))
            throw new ArgumentException("dataRange is required for update_data operation");

        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

        var chart = worksheet.Charts[chartIndex];
        AddDataSeries(chart, dataRange);
        SetCategoryData(chart, categoryAxisDataRange ?? "");

        ctx.Save(outputPath);

        var result = $"Chart #{chartIndex} data updated to: {dataRange}";
        if (!string.IsNullOrEmpty(categoryAxisDataRange))
            result += $", X-axis: {categoryAxisDataRange}";
        result += $". {ctx.GetOutputMessage(outputPath)}";
        return result;
    }

    /// <summary>
    ///     Sets chart properties (title, legend).
    /// </summary>
    /// <param name="ctx">The document context containing the workbook.</param>
    /// <param name="outputPath">The output file path.</param>
    /// <param name="sheetIndex">The worksheet index.</param>
    /// <param name="chartIndex">The chart index.</param>
    /// <param name="title">The chart title.</param>
    /// <param name="removeTitle">Whether to remove the chart title.</param>
    /// <param name="legendVisible">Whether the legend is visible.</param>
    /// <param name="legendPosition">The legend position.</param>
    /// <returns>A message indicating the result of the operation.</returns>
    /// <exception cref="ArgumentException">Thrown when chart index is out of range.</exception>
    private static string SetChartProperties(DocumentContext<Workbook> ctx, string? outputPath, int sheetIndex,
        int chartIndex, string? title, bool removeTitle, bool? legendVisible, string? legendPosition)
    {
        var workbook = ctx.Document;
        var worksheet = ExcelHelper.GetWorksheet(workbook, sheetIndex);

        if (chartIndex < 0 || chartIndex >= worksheet.Charts.Count)
            throw new ArgumentException(
                $"Chart index {chartIndex} is out of range (worksheet has {worksheet.Charts.Count} charts)");

        var chart = worksheet.Charts[chartIndex];
        List<string> changes = [];

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

        ctx.Save(outputPath);

        return changes.Count > 0
            ? $"Chart #{chartIndex} properties updated: {string.Join(", ", changes)}. {ctx.GetOutputMessage(outputPath)}"
            : $"Chart #{chartIndex} no changes. {ctx.GetOutputMessage(outputPath)}";
    }
}