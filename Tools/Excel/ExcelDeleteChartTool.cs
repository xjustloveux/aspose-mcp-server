using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteChartTool : IAsposeTool
{
    public string Description => "Delete a chart from an Excel worksheet";

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
            chartIndex = new
            {
                type = "number",
                description = "Chart index to delete (0-based)"
            }
        },
        required = new[] { "path", "chartIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var chartIndex = arguments?["chartIndex"]?.GetValue<int>() ?? throw new ArgumentException("chartIndex is required");

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
        var chartName = chart.Name ?? $"圖表 {chartIndex}";
        
        charts.RemoveAt(chartIndex);
        workbook.Save(path);
        
        var remainingCount = charts.Count;
        
        return await Task.FromResult($"成功刪除圖表 #{chartIndex} ({chartName})\n工作表剩餘圖表數: {remainingCount}\n輸出: {path}");
    }
}

