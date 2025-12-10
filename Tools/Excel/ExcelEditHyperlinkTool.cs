using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelEditHyperlinkTool : IAsposeTool
{
    public string Description => "Edit a hyperlink in an Excel worksheet";

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
                description = "Hyperlink index to edit (0-based)"
            },
            address = new
            {
                type = "string",
                description = "New hyperlink address (optional)"
            },
            textToDisplay = new
            {
                type = "string",
                description = "New display text (optional)"
            }
        },
        required = new[] { "path", "hyperlinkIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var hyperlinkIndex = arguments?["hyperlinkIndex"]?.GetValue<int>() ?? throw new ArgumentException("hyperlinkIndex is required");
        var address = arguments?["address"]?.GetValue<string>();
        var textToDisplay = arguments?["textToDisplay"]?.GetValue<string>();

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
        var oldAddress = hyperlink.Address ?? "";
        var oldText = hyperlink.TextToDisplay ?? "";
        
        if (!string.IsNullOrEmpty(address))
        {
            hyperlink.Address = address;
        }
        
        if (!string.IsNullOrEmpty(textToDisplay))
        {
            hyperlink.TextToDisplay = textToDisplay;
        }
        
        workbook.Save(path);
        
        var result = $"成功編輯超連結 #{hyperlinkIndex}\n";
        result += $"舊地址: {oldAddress}\n";
        result += $"新地址: {hyperlink.Address ?? oldAddress}\n";
        result += $"舊顯示文字: {oldText}\n";
        result += $"新顯示文字: {hyperlink.TextToDisplay ?? oldText}\n";
        result += $"輸出: {path}";
        
        return await Task.FromResult(result);
    }
}

