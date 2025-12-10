using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Charts;

namespace AsposeMcpServer.Tools;

public class ExcelGetChartsTool : IAsposeTool
{
    public string Description => "Get all charts information from an Excel worksheet";

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
            }
        },
        required = new[] { "path" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var charts = worksheet.Charts;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的圖表資訊 ===\n");
        result.AppendLine($"總圖表數: {charts.Count}\n");

        if (charts.Count == 0)
        {
            result.AppendLine("未找到圖表");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < charts.Count; i++)
        {
            var chart = charts[i];
            result.AppendLine($"【圖表 {i}】");
            result.AppendLine($"名稱: {chart.Name ?? "(無名稱)"}");
            result.AppendLine($"類型: {chart.Type}");
            result.AppendLine($"位置: 行 {chart.ChartObject.UpperLeftRow} - {chart.ChartObject.LowerRightRow}, 列 {chart.ChartObject.UpperLeftColumn} - {chart.ChartObject.LowerRightColumn}");
            result.AppendLine($"寬度: {chart.ChartObject.Width}");
            result.AppendLine($"高度: {chart.ChartObject.Height}");
            
            if (chart.NSeries != null && chart.NSeries.Count > 0)
            {
                result.AppendLine($"數據系列數: {chart.NSeries.Count}");
                for (int j = 0; j < chart.NSeries.Count && j < 5; j++) // Limit to first 5 series
                {
                    var series = chart.NSeries[j];
                    result.AppendLine($"  系列 {j}: {series.Name ?? "(無名稱)"}");
                }
            }
            
            if (chart.Title != null)
            {
                result.AppendLine($"標題: {chart.Title.Text}");
            }
            
            if (chart.Legend != null)
            {
                result.AppendLine($"圖例: 已啟用");
            }
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

