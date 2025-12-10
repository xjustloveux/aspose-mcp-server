using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteConditionalFormattingTool : IAsposeTool
{
    public string Description => "Delete conditional formatting from an Excel worksheet";

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
            conditionalFormattingIndex = new
            {
                type = "number",
                description = "Conditional formatting index to delete (0-based)"
            }
        },
        required = new[] { "path", "conditionalFormattingIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var conditionalFormattingIndex = arguments?["conditionalFormattingIndex"]?.GetValue<int>() ?? throw new ArgumentException("conditionalFormattingIndex is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var conditionalFormattings = worksheet.ConditionalFormattings;
        
        if (conditionalFormattingIndex < 0 || conditionalFormattingIndex >= conditionalFormattings.Count)
        {
            throw new ArgumentException($"條件格式索引 {conditionalFormattingIndex} 超出範圍 (工作表共有 {conditionalFormattings.Count} 個條件格式)");
        }

        conditionalFormattings.RemoveAt(conditionalFormattingIndex);
        workbook.Save(path);
        
        var remainingCount = conditionalFormattings.Count;
        
        return await Task.FromResult($"成功刪除條件格式 #{conditionalFormattingIndex}\n工作表剩餘條件格式數: {remainingCount}\n輸出: {path}");
    }
}

