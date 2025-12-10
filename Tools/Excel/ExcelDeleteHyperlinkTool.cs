using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteHyperlinkTool : IAsposeTool
{
    public string Description => "Delete a hyperlink from an Excel worksheet";

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
            hyperlinkIndex = new
            {
                type = "number",
                description = "Hyperlink index to delete (0-based)"
            }
        },
        required = new[] { "path", "hyperlinkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var hyperlinks = worksheet.Hyperlinks;
        
        if (hyperlinkIndex < 0 || hyperlinkIndex >= hyperlinks.Count)
        {
            throw new ArgumentException($"超連結索引 {hyperlinkIndex} 超出範圍 (工作表共有 {hyperlinks.Count} 個超連結)");
        }

        var hyperlink = hyperlinks[hyperlinkIndex];
        var address = hyperlink.Address ?? "";
        
        hyperlinks.RemoveAt(hyperlinkIndex);
        workbook.Save(path);
        
        var remainingCount = hyperlinks.Count;
        
        return await Task.FromResult($"成功刪除超連結 #{hyperlinkIndex}\n地址: {address}\n工作表剩餘超連結數: {remainingCount}\n輸出: {path}");
    }
}

