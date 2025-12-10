using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelAddChartTool : IAsposeTool
{
    public string Description => "Add a chart to an Excel worksheet";

    public object InputSchema => new
    {
        type = "object",
        properties = new
        {
            path = new
            {
                type = "string",
                description = "Excel file path"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            chartType = new
            {
                type = "string",
                description = "Chart type (Column, Bar, Line, Pie, etc.)"
            },
            dataRange = new
            {
                type = "string",
                description = "Data range for chart (e.g., 'A1:B10')"
            },
            title = new
            {
                type = "string",
                description = "Chart title (optional)"
            }
        },
        required = new[] { "path", "chartType", "dataRange" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var chartTypeStr = arguments?["chartType"]?.GetValue<string>() ?? throw new ArgumentException("chartType is required");
        var dataRange = arguments?["dataRange"]?.GetValue<string>() ?? throw new ArgumentException("dataRange is required");
        var title = arguments?["title"]?.GetValue<string>();

        using var workbook = new Workbook(path);
        var worksheet = workbook.Worksheets[sheetIndex];

        var chartType = chartTypeStr.ToLower() switch
        {
            "column" => ChartType.Column,
            "bar" => ChartType.Bar,
            "line" => ChartType.Line,
            "pie" => ChartType.Pie,
            "area" => ChartType.Area,
            "scatter" => ChartType.Scatter,
            "doughnut" => ChartType.Doughnut,
            _ => ChartType.Column
        };

        int chartIndex = worksheet.Charts.Add(chartType, 5, 0, 20, 10);
        var chart = worksheet.Charts[chartIndex];
        chart.SetChartDataRange(dataRange, true);

        if (!string.IsNullOrEmpty(title))
        {
            chart.Title.Text = title;
        }

        workbook.Save(path);

        return await Task.FromResult($"Chart added to worksheet: {path}");
    }
}

