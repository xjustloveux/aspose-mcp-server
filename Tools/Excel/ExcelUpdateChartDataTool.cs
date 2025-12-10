using System.Text.Json.Nodes;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelUpdateChartDataTool : IAsposeTool
{
    public string Description => "Update chart data source range in an Excel worksheet";

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
            outputPath = new
            {
                type = "string",
                description = "Output file path (if not provided, overwrites input)"
            },
            sheetIndex = new
            {
                type = "number",
                description = "Sheet index (0-based, optional, default: 0)"
            },
            chartIndex = new
            {
                type = "number",
                description = "Chart index to update (0-based)"
            },
            dataRange = new
            {
                type = "string",
                description = "New data range for chart (e.g., 'A1:B10')"
            }
        },
        required = new[] { "path", "chartIndex", "dataRange" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var outputPath = arguments?["outputPath"]?.GetValue<string>() ?? path;
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required");
        var dataRange = arguments?["dataRange"]?.GetValue<string>() ?? throw new ArgumentException("dataRange is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var charts = worksheet.Charts;
        
        if (chartIndex < 0 || chartIndex >= charts.Count)
        {
            throw new ArgumentException($"圖表索引 {chartIndex} 超出範圍 (工作表共有 {charts.Count} 個圖表)");
        }

        var chart = charts[chartIndex];
        
        // Clear existing series
        chart.NSeries.Clear();
        
        // Parse data range - support multiple ranges separated by comma (e.g., "E2:E40,G2:G40")
        var ranges = dataRange.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        
        foreach (var range in ranges)
        {
            // Add series with the data range - Add returns the index
            int seriesIndex = chart.NSeries.Add(range, true);
            // Get the series object and set the values range explicitly
            var series = chart.NSeries[seriesIndex];
            series.Values = range;
        }
        
        // If no series were added, use SetChartDataRange as fallback
        if (chart.NSeries.Count == 0)
        {
            chart.SetChartDataRange(dataRange, true);
        }

        workbook.Save(outputPath);

        return await Task.FromResult($"成功更新圖表 #{chartIndex} 的數據源\n新數據範圍: {dataRange}\n輸出: {outputPath}");
    }
}

