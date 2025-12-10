using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelGetHyperlinksTool : IAsposeTool
{
    public string Description => "Get all hyperlinks from an Excel worksheet";

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
        var hyperlinks = worksheet.Hyperlinks;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的超連結資訊 ===\n");
        result.AppendLine($"總超連結數: {hyperlinks.Count}\n");

        if (hyperlinks.Count == 0)
        {
            result.AppendLine("未找到超連結");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < hyperlinks.Count; i++)
        {
            var hyperlink = hyperlinks[i];
            result.AppendLine($"【超連結 {i}】");
            result.AppendLine($"地址: {hyperlink.Address ?? "(無)"}");
            result.AppendLine($"顯示文字: {hyperlink.TextToDisplay ?? "(無)"}");
            var area = hyperlink.Area;
            result.AppendLine($"位置: 行 {area.StartRow}-{area.EndRow}, 列 {area.StartColumn}-{area.EndColumn}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

