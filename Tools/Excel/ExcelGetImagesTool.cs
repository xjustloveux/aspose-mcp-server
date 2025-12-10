using System.Text.Json.Nodes;
using System.Text;
using Aspose.Cells;
using Aspose.Cells.Drawing;

namespace AsposeMcpServer.Tools;

public class ExcelGetImagesTool : IAsposeTool
{
    public string Description => "Get all images information from an Excel worksheet";

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
        var pictures = worksheet.Pictures;
        var result = new StringBuilder();

        result.AppendLine($"=== 工作表 '{worksheet.Name}' 的圖片資訊 ===\n");
        result.AppendLine($"總圖片數: {pictures.Count}\n");

        if (pictures.Count == 0)
        {
            result.AppendLine("未找到圖片");
            return await Task.FromResult(result.ToString());
        }

        for (int i = 0; i < pictures.Count; i++)
        {
            var picture = pictures[i];
            result.AppendLine($"【圖片 {i}】");
            result.AppendLine($"名稱: {picture.Name ?? "(無名稱)"}");
            result.AppendLine($"位置: 行 {picture.UpperLeftRow}-{picture.LowerRightRow}, 列 {picture.UpperLeftColumn}-{picture.LowerRightColumn}");
            result.AppendLine($"寬度: {picture.Width}");
            result.AppendLine($"高度: {picture.Height}");
            result.AppendLine($"原始寬度: {picture.OriginalWidth}");
            result.AppendLine($"原始高度: {picture.OriginalHeight}");
            result.AppendLine($"圖片類型: {picture.ImageType}");
            result.AppendLine();
        }

        return await Task.FromResult(result.ToString());
    }
}

