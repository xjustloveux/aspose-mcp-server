using System.Text.Json.Nodes;
using Aspose.Cells;

namespace AsposeMcpServer.Tools;

public class ExcelDeleteImageTool : IAsposeTool
{
    public string Description => "Delete an image from an Excel worksheet";

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
            imageIndex = new
            {
                type = "number",
                description = "Image index to delete (0-based)"
            }
        },
        required = new[] { "path", "imageIndex" }
    };

    public async Task<string> ExecuteAsync(JsonObject? arguments)
    {
        var path = arguments?["path"]?.GetValue<string>() ?? throw new ArgumentException("path is required");
        var sheetIndex = arguments?["sheetIndex"]?.GetValue<int>() ?? 0;
        var imageIndex = arguments?["imageIndex"]?.GetValue<int>() ?? throw new ArgumentException("imageIndex is required");

        using var workbook = new Workbook(path);
        
        if (sheetIndex < 0 || sheetIndex >= workbook.Worksheets.Count)
        {
            throw new ArgumentException($"工作表索引 {sheetIndex} 超出範圍 (共有 {workbook.Worksheets.Count} 個工作表)");
        }

        var worksheet = workbook.Worksheets[sheetIndex];
        var pictures = worksheet.Pictures;
        
        if (imageIndex < 0 || imageIndex >= pictures.Count)
        {
            throw new ArgumentException($"圖片索引 {imageIndex} 超出範圍 (工作表共有 {pictures.Count} 個圖片)");
        }

        var picture = pictures[imageIndex];
        var pictureName = picture.Name ?? $"圖片 {imageIndex}";
        
        pictures.RemoveAt(imageIndex);
        workbook.Save(path);
        
        var remainingCount = pictures.Count;
        
        return await Task.FromResult($"成功刪除圖片 #{imageIndex} ({pictureName})\n工作表剩餘圖片數: {remainingCount}\n輸出: {path}");
    }
}

